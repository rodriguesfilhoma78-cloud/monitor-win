"""
================================================================
 MONITOR WDO - server_wdo.py
 Sistema independente de monitoramento do Dolar Futuro (DOLFUT)
----------------------------------------------------------------
 Fork do Monitor WIN (server_win.py) - mesma arquitetura, ativo trocado.
 Pipeline : Profit Pro RTD -> Excel (VBA) -> dados_wdo.csv
            -> este servidor (FastAPI) -> WebSocket -> dashboard
 Extra    : macro (Brent BZ=F via Yahoo - trocado do S&P 500 do WIN em
            24/08/2026, correlaciona melhor com o dolar de pais exportador
            de commodity; Ouro GC=F e Juros EUA ^TNX via Yahoo, trocados em
            25/08/2026 no lugar do Dolar/DI redundantes - ver nota completa
            em MACRO_SYMBOLS abaixo) -> card MACRO do dashboard.
 Porta    : 8003 (roda em paralelo com o Monitor WIN na 8001 e com o
            sistema antigo "Mapa de Tendencia WDO" na 8000)
 Executar : python server_wdo.py
================================================================
Boas praticas aplicadas (heranca do server_win.py, que por sua vez
diferia do server_v2.py do sistema antigo "Mapa de Tendencia WDO"):
  1. Lifespan pattern (sem @app.on_event deprecated)
  2. Snapshot SQLite via run_in_executor (nao bloqueia o loop async)
  3. Niveis desacoplados em niveis.json + endpoints GET/POST /niveis
  4. Classes com responsabilidade unica (SRP - SOLID)
"""

import asyncio
import csv
import io
import json
import sqlite3
import statistics
import sys
import time
from collections import deque
from datetime import date, timedelta
from contextlib import asynccontextmanager
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Optional

import httpx
import uvicorn
from fastapi import FastAPI, WebSocket, WebSocketDisconnect
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, Response

# Ruido benigno do asyncio no Windows: o Proactor loga ConnectionResetError
# quando um WebSocket fecha abruptamente (aba fechada, refresh). Nao afeta
# o funcionamento, so polui o log - suprime so esse callback especifico.
if sys.platform == "win32":
    from asyncio.proactor_events import _ProactorBasePipeTransport
    _ProactorBasePipeTransport._call_connection_lost = lambda self, exc=None: None

import agente_wdo     # leitura de fluxo por IA (Google Gemini) - consumidor, nao fonte
import casado_wdo     # preco justo do WDO vs dolar a vista (o "casado")

# ----------------------------------------------------------------
# CONFIGURACAO
# ----------------------------------------------------------------
BASE_DIR      = Path(__file__).parent
DATA_DIR      = BASE_DIR
CSV_PATH      = DATA_DIR / "dados_wdo.csv"       # gerado pelo VBA (ExportarWDO)
BOOK_CSV      = DATA_DIR / "dados_book.csv"      # livro de ofertas (BOOK1)
TT_CSV        = DATA_DIR / "dados_tt.csv"        # fita / Times & Trades (T&T1)
VAP_CSV       = DATA_DIR / "dados_vap.csv"       # volume por preco (VAP1)
RANKING_CSV   = DATA_DIR / "ranking_acum.csv"    # acumulado p/ a aba AcumWDO (VBA le)
NIVEIS_PATH   = BASE_DIR / "niveis.json"         # niveis do dia (editavel)
DB_PATH       = BASE_DIR / "wdo_history.db"
DASHBOARD     = BASE_DIR / "dashboard_wdo.html"
POLL_INTERVAL = 1.0        # segundos entre leituras do CSV
LEITURA_AUTO_INTERVALO = 900   # segundos entre leituras de IA periodicas (15 min) -
                               # subiu de 3 p/ 15 min ao incluir macro no contexto
                               # (prompt maior, poupa cota do nivel gratuito)
JANELA_CORRETORAS_S = 30 * 60   # janela movel do ranking (quem agride AGORA,
                                # nao so o acumulado desde a abertura)
SERIE_CORRETORAS_S  = 60        # cadencia dos snapshots da serie temporal
                                # (corretoras_serie) - granularidade pro
                                # backtest de "quem virou de lado" depois do pregao
SNAPSHOT_EVERY = 2         # segundos entre snapshots no SQLite (resolucao
                           # do perfil de volume)
HOST, PORT    = "127.0.0.1", 8003

# Cobertura minima para um pregao virar base de pivots: o server tem que ter
# pego a abertura e o fechamento. Fora disso, H/L sao parciais e os pivots do
# dia seguinte saem errados.
PREGAO_ABERTURA_ATE = "10:00"    # primeiro tick tem que vir antes disso
PREGAO_FECHA_APOS   = "17:45"    # ultimo tick tem que vir depois disso

# --- Macro (Brent, Ouro, DXY, Juros EUA) --------------------------------
# Brent: Yahoo Finance. Em 24/08/2026 trocado o S&P 500 (driver de risco
# global, herdado do Monitor WIN) por BRENT (BZ=F): Brasil e exportador de
# petroleo (Petrobras), entao o termo de troca move o cambio - Brent sobe
# -> BRL tende a se fortalecer -> DOLFUT tende a CAIR (por isso a
# polaridade do selo do dashboard e' invertida pro Brent: alta = seta
# vermelha/contrario ao dolar, baixa = seta verde/favoravel). DXY continua
# como termometro de risco global.
#
# Em 25/08/2026, os dois cards redundantes/RTD foram trocados por
# commodities/juros globais que correlacionam melhor com o DOLFUT
# (decisao do usuario):
#   - "Dolar" (USD/BRL via Yahoo, proxy do proprio DOLFUT em tempo real
#     via RTD) -> OURO (GC=F). Ouro e cotado em USD e historicamente se
#     move de forma INVERSA a forca global do dolar (ouro sobe quando o
#     dolar enfraquece no mundo, e vice-versa). Em 17/09/2026 o usuario
#     INVERTEU o sinal desse card: agora ouro SOBE = seta verde/favoravel
#     ao DOLFUT, ouro CAI = seta vermelha/contraria (ver renderMacro() no
#     dashboard_wdo.html e _alinhamento_macro() no agente_wdo.py).
#   - "Juros DI" (DI futuro via RTD/Excel, juro domestico) -> JUROS EUA
#     (^TNX, treasury de 10 anos via Yahoo). Juro americano mais alto
#     atrai capital para os EUA e tende a fortalecer o dolar globalmente,
#     mas em 17/09/2026 o usuario INVERTEU o sinal desse card: agora juro
#     EUA SOBE = seta vermelha/contraria ao DOLFUT, cai = seta
#     verde/favoravel (ver renderMacro() e _alinhamento_macro()). Isso tambem
#     tira a dependencia do RTD/Excel (dados_macro_rtd.csv, exportado pelo
#     ExportarWDO.bas) para o card MACRO - o CSV continua sendo gerado
#     pela planilha mas nao e' mais lido por este servidor.
YAHOO_CHART   = ("https://query1.finance.yahoo.com/v8/finance/chart/{sym}"
                 "?range=1d&interval=15m")
MACRO_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
MACRO_POLL    = 30                            # segundos entre consultas
# DXY = dolar global (indice ICE contra cesta de moedas). Herdado do
# Monitor WIN, onde a correlacao (-0,36/ano) e com o IBOV; contra o
# proprio USD/BRL (DOLFUT aqui) a relacao e POSITIVA na maior parte do
# tempo (DXY sobe -> dolar tende a subir tambem, mesmo "fluxo p/
# emergentes"). O card MACRO nao foi recalibrado para essa polaridade -
# ver ressalva no README.
MACRO_SYMBOLS = {"brent": "BZ=F", "dxy": "DX-Y.NYB", "ouro": "GC=F",
                  "juros_us": "^TNX"}

# --- CASADO: preco justo do WDO vs dolar a vista (Caminho A, 27/08/2026) --
# dolar a vista (dolar comercial): AwesomeAPI, sem chave, cache de ~1 min.
# Fallback: Yahoo BRL=X (mesmo padrao dos simbolos macro). Ver casado_wdo.py.
SPOT_URL_AWESOME = "https://economia.awesomeapi.com.br/json/last/USD-BRL"
SPOT_SYM_YAHOO   = "BRL=X"
# dados_macro_rtd.csv (DI1*/DOLFUT via RTD/Excel): parou de alimentar o card
# MACRO em 25/08, mas VOLTOU a ser lido aqui (so o DI) para decompor o
# carrego bruto e estimar o cupom cambial implicito no card CASADO. O card
# MACRO continua 100% Yahoo. Se o arquivo estiver velho/ausente, o casado
# funciona sem o cupom implicito (campo fica nulo).
MACRO_RTD_CSV = DATA_DIR / "dados_macro_rtd.csv"
# Letras de mes dos codigos B3 (DI1F28 = Jan/2028).
_MES_COD_B3 = {"F": 1, "G": 2, "H": 3, "J": 4, "K": 5, "M": 6,
               "N": 7, "Q": 8, "U": 9, "V": 10, "X": 11, "Z": 12}


# ----------------------------------------------------------------
# MODELO DO TICK
# ----------------------------------------------------------------
@dataclass
class Tick:
    ultimo: Optional[float] = None
    abertura: Optional[float] = None
    maxima: Optional[float] = None
    minima: Optional[float] = None
    volume: Optional[float] = None
    agr_compra: Optional[float] = None
    agr_venda: Optional[float] = None
    vwap: Optional[float] = None
    fec_ant: Optional[float] = None     # fechamento oficial do pregao anterior (RTD "FEC")
    timestamp: str = ""

    @property
    def delta(self) -> Optional[float]:
        if self.agr_compra is None or self.agr_venda is None:
            return None
        return self.agr_compra - self.agr_venda


def _to_float(raw: str) -> Optional[float]:
    """Converte numeros no formato brasileiro ('175.110,00') ou US."""
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None
    # Formato BR: ponto = milhar, virgula = decimal
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


# ----------------------------------------------------------------
# LEITOR DO CSV (SRP: so le e converte)
# ----------------------------------------------------------------
class CsvReader:
    """Le a ultima linha valida de dados_win.csv exportado pelo VBA.

    Formato esperado (separador ';'):
    ultimo;abertura;maxima;minima;volume;agr_compra;agr_venda;vwap;timestamp
    """

    FIELDS = ["ultimo", "abertura", "maxima", "minima",
              "volume", "agr_compra", "agr_venda", "vwap", "fec_ant"]

    def __init__(self, path: Path):
        self.path = path
        self._last_mtime = 0.0

    def read_if_changed(self) -> Optional[Tick]:
        if not self.path.exists():
            return None
        mtime = self.path.stat().st_mtime
        if mtime == self._last_mtime:
            return None                      # nada novo
        self._last_mtime = mtime
        try:
            with open(self.path, encoding="utf-8-sig", errors="ignore") as f:
                rows = [r for r in csv.reader(f, delimiter=";") if r]
        except PermissionError:
            return None                      # Excel escrevendo no arquivo
        if not rows:
            return None
        row = rows[-1]
        # pula cabecalho se existir
        if row and not any(ch.isdigit() for ch in row[0]):
            if len(rows) < 2:
                return None
            row = rows[-2] if rows[-1] == row else rows[-1]
        values = {f: _to_float(row[i]) if i < len(row) else None
                  for i, f in enumerate(self.FIELDS)}
        if values["ultimo"] is None:
            return None
        return Tick(**values, timestamp=time.strftime("%H:%M:%S"))


# ----------------------------------------------------------------
# LEITOR DE FLUXO (livro + fita + VAP)
# ----------------------------------------------------------------
class FluxoReader:
    """Le dados_book.csv e dados_tt.csv e devolve metricas de fluxo.

    Do livro sai a pressao ESTATICA (quanto tem parado em cada lado); da fita
    sai a pressao DINAMICA (o que de fato foi executado e com que violencia).
    Sao coisas diferentes: livro cheio na compra com fita vendendo agressivo e
    justamente o cenario de absorcao.

    IMPORTANTE - a fita e AMOSTRA, nao contagem. O RTD publica so os ultimos
    21 negocios, e medido em 21/07/2026 esses 21 cobrem 0,56s de mediana
    (minimo 0,09s) contra um ciclo de export de 2s. Tentar somar negocio a
    negocio entre leituras subestima o fluxo em ~70% justamente nos momentos
    rapidos, que sao os que importam.

    A saida certa e a TAXA: com o span da janela, negocios/s e contratos/s sao
    estimativas nao-enviesadas e nao dependem de nenhuma leitura ser continua.
    Para contagem exata do dia use os acumulados do tick (agr_compra/agr_venda),
    que nao perdem nada; a fita serve para textura (tamanho de lote, ritmo).
    """

    JANELA_MIN_S = 0.1    # abaixo disso a taxa vira ruido amplificado
    AREA_VALOR = 0.70     # fracao do volume que define a area de valor
    ASSINATURA = 6        # negocios usados para casar a janela anterior

    def __init__(self, book_path: Path, tt_path: Path,
                 vap_path: Optional[Path] = None):
        self.book_path = book_path
        self.tt_path = tt_path
        self.vap_path = vap_path
        self._tt_anterior: list = []

    @staticmethod
    def _ler(path: Path) -> list:
        if not path.exists():
            return []
        try:
            with open(path, encoding="utf-8-sig", errors="ignore") as f:
                linhas = [r for r in csv.reader(f, delimiter=";") if r and any(r)]
        except PermissionError:
            return []                        # Excel escrevendo no arquivo
        return linhas[1:] if linhas else []   # descarta cabecalho

    def _book(self) -> Optional[dict]:
        """Melhor bid/ask, spread e profundidade acumulada de cada lado."""
        linhas = self._ler(self.book_path)
        compras, vendas = [], []
        for r in linhas:
            if len(r) < 8:
                continue
            qc, pc = _to_float(r[2]), _to_float(r[3])
            pv, qv = _to_float(r[4]), _to_float(r[5])
            if pc and qc:
                compras.append((pc, qc))
            if pv and qv:
                vendas.append((pv, qv))
        if not compras or not vendas:
            return None
        melhor_bid = max(p for p, _ in compras)
        melhor_ask = min(p for p, _ in vendas)
        qtd_bid = sum(q for p, q in compras if p == melhor_bid)
        qtd_ask = sum(q for p, q in vendas if p == melhor_ask)
        total = qtd_bid + qtd_ask
        return {
            "bid": melhor_bid, "ask": melhor_ask,
            "spread": melhor_ask - melhor_bid,
            "qtd_bid": qtd_bid, "qtd_ask": qtd_ask,
            "profundidade_bid": sum(q for _, q in compras),
            "profundidade_ask": sum(q for _, q in vendas),
            # >0 = mais oferta parada na compra; escala -1..+1
            "desequilibrio": round((qtd_bid - qtd_ask) / total, 3) if total else 0.0,
            "niveis": len(compras),
        }

    @staticmethod
    def _vol_abreviado(s: str) -> Optional[float]:
        """Volume do VAP vem abreviado: '620', '3,04k', '1,2M'.

        O RTD mistura numero puro (abaixo de mil) com sufixo k/M. Perde-se
        precisao no sufixo ('46,36k' = 46.360, resolucao de ~10 contratos),
        o que e irrelevante para a FORMA do perfil - que e o que se usa.
        """
        if s is None:
            return None
        t = str(s).strip()
        if not t or t == "-":
            return None
        mult = 1.0
        if t[-1] in "kK":
            mult, t = 1_000.0, t[:-1]
        elif t[-1] in "mM":
            mult, t = 1_000_000.0, t[:-1]
        t = t.replace(".", "").replace(",", ".") if mult == 1.0 else t.replace(",", ".")
        try:
            return float(t) * mult
        except ValueError:
            return None

    def _vap(self, preco_ref: Optional[float]) -> Optional[dict]:
        """Perfil de volume por preco: POC, area de valor e reparticao.

        POC = preco de maior volume negociado (onde o mercado aceitou negocio).
        Area de valor = faixa central que concentra 70% do volume, expandida a
        partir do POC pelo vizinho mais forte (definicao classica de Market
        Profile). Precos FORA da area de valor tendem a ser rejeitados.
        """
        if not self.vap_path:
            return None
        niveis = []
        for r in self._ler(self.vap_path):
            if len(r) < 2:
                continue
            p, v = _to_float(r[0]), self._vol_abreviado(r[1])
            if p and v:
                niveis.append((p, v))
        if not niveis:
            return None
        niveis.sort(key=lambda x: x[0])            # preco crescente
        total = sum(v for _, v in niveis)
        i_poc = max(range(len(niveis)), key=lambda i: niveis[i][1])
        poc = niveis[i_poc][0]

        # expande a partir do POC pelo lado de maior volume ate cobrir 70%
        lo = hi = i_poc
        acumulado = niveis[i_poc][1]
        alvo = total * self.AREA_VALOR
        while acumulado < alvo and (lo > 0 or hi < len(niveis) - 1):
            v_baixo = niveis[lo - 1][1] if lo > 0 else -1
            v_cima = niveis[hi + 1][1] if hi < len(niveis) - 1 else -1
            if v_cima == v_baixo:
                # Empate: expande os DOIS lados. Desempatar sempre para o mesmo
                # lado deformaria a area num perfil simetrico.
                hi += 1; lo -= 1
                acumulado += niveis[hi][1] + niveis[lo][1]
            elif v_cima > v_baixo:
                hi += 1
                acumulado += niveis[hi][1]
            else:
                lo -= 1
                acumulado += niveis[lo][1]
        out = {
            "poc": poc,
            "vah": niveis[hi][0],                  # topo da area de valor
            "val": niveis[lo][0],                  # base da area de valor
            "vap_total": total,
            "vap_niveis": len(niveis),
        }
        if preco_ref:
            acima = sum(v for p, v in niveis if p > preco_ref)
            out["vol_acima_pct"] = round(acima / total * 100, 1) if total else None
            out["dist_poc"] = preco_ref - poc
            out["dentro_area"] = out["val"] <= preco_ref <= out["vah"]
        return out

    @staticmethod
    def _hora_seg(h: str) -> Optional[float]:
        """'10:32:14.507' -> segundos do dia (a fita traz milissegundo)."""
        try:
            hh, mm, ss = h.strip().split(":")
            return int(hh) * 3600 + int(mm) * 60 + float(ss)
        except (ValueError, AttributeError):
            return None

    def _negocios_novos(self, atual: list) -> tuple:
        """Negocios que entraram na janela desde a leitura anterior.

        Casa o inicio da janela anterior (ASSINATURA negocios) dentro da
        atual: tudo antes do casamento e novo. So e confiavel porque a janela
        de 500 negocios cobre ~60s contra um ciclo de 2s; com a janela antiga
        de 21 (0,5s) isso perdia ~70% da fita — dai as TAXAS continuarem
        sendo a metrica de ritmo, e o dedupe servir ao ACUMULADO por corretora.
        """
        if not atual:
            return [], False
        if not self._tt_anterior:
            self._tt_anterior = atual
            return [], False                 # primeira leitura: so referencia
        marco = self._tt_anterior[:self.ASSINATURA]
        novos, perdida = atual, True
        if len(marco) >= self.ASSINATURA:
            for i in range(len(atual) - len(marco) + 1):
                if atual[i:i + len(marco)] == marco:
                    novos, perdida = atual[:i], False
                    break
        self._tt_anterior = atual
        return novos, perdida

    def ler(self, preco_ref: Optional[float] = None) -> Optional[dict]:
        book = self._book()
        vap = self._vap(preco_ref)

        janela_rows = [tuple(r[:5]) for r in self._ler(self.tt_path)
                       if len(r) >= 5]
        novos_rows, janela_perdida = self._negocios_novos(janela_rows)
        novos = []                           # negocios ineditos, ja parseados
        for r in novos_rows:
            q, p = _to_float(r[1]), _to_float(r[2])
            if q and p:
                novos.append({"hora": r[0], "qtd": q, "preco": p,
                              "comprador": r[3].strip(), "vendedor": r[4].strip()})

        qtds, horas, agr_compra, agr_venda = [], [], 0.0, 0.0
        for r in janela_rows:
            q, p = _to_float(r[1]), _to_float(r[2])
            t = self._hora_seg(r[0])
            if not q or not p:
                continue
            qtds.append(q)
            if t is not None:
                horas.append(t)
            # Agressor inferido pelo lado do livro em que o negocio saiu:
            # no ask = comprador atacou; no bid = vendedor atacou.
            if book:
                if p >= book["ask"]:
                    agr_compra += q
                elif p <= book["bid"]:
                    agr_venda += q
        executado = sum(qtds)
        # Span da amostra: base para converter a janela em taxa. Se todos os
        # negocios tem o mesmo carimbo (rajada), nao da para estimar ritmo.
        janela = (max(horas) - min(horas)) if len(horas) >= 2 else 0.0
        out = {
            "amostra_negocios": len(qtds),
            "amostra_contratos": executado,
            "janela_s": round(janela, 3),
            "negocios_s": round(len(qtds) / janela, 1) if janela > 0 else None,
            "contratos_s": round(executado / janela, 1) if janela > 0 else None,
            # Janela curta demais = taxa instavel (21 negocios em 0,02s viram
            # "875 neg/s", real para o instante mas nao sustentavel). Quem for
            # tirar media tem que ponderar por janela_s ou filtrar por isto.
            "taxa_confiavel": janela >= self.JANELA_MIN_S,
            # negocios ineditos desde a ultima leitura (para o acumulado por
            # corretora); consumidos pelo market_loop, nao vao no broadcast
            "novos": novos,
            "janela_perdida": janela_perdida,
            "lote_medio": round(executado / len(qtds), 1) if qtds else 0.0,
            "lote_max": max(qtds) if qtds else 0.0,
            "agr_compra_fita": agr_compra,
            "agr_venda_fita": agr_venda,
            # >0 = compra atacando; normalizado pelo executado na amostra
            "pressao_fita": round((agr_compra - agr_venda) / executado, 3)
                            if executado else 0.0,
            "timestamp": time.strftime("%H:%M:%S"),
        }
        if book:
            out.update(book)
        if vap:
            out.update(vap)
        return out if (book or qtds or vap) else None


# ----------------------------------------------------------------
# RANKING DE CORRETORAS (acumulado do dia, compilado da fita)
# ----------------------------------------------------------------
class RankingCorretoras:
    """Compila a fita em saldo por corretora ao longo do PREGAO inteiro.

    A janela do RTD (500 negocios ~ 60s) so mostra o instante — e a formula
    de ranking da planilha (DADOS!R21) usa essa janela como fallback. Aqui
    cada negocio inedito (dedupe do FluxoReader) soma no acumulado do dia:
    qtd e financeiro por corretora, nos dois lados. Reconstruido do SQLite
    no boot, entao restart do server nao zera o ranking.

    Mesma ressalva da nota na planilha: corretora agressora nao e posicao —
    XP vendendo 10k pode ser 200 clientes; o valor esta em ler dominancia e
    troca de mao (ex.: institucional entrando onde so havia varejo).

    Alem do acumulado do dia, mantem uma JANELA MOVEL (JANELA_CORRETORAS_S)
    via deque de eventos: o acumulado do dia inteiro reflete sobretudo a
    manha e fica lento pra mostrar quem esta agredindo AGORA; a janela
    responde isso, ao custo de ser mais ruidosa (um lote institucional
    grande pode virar o saldo da janela rapido). As duas se complementam.
    """

    def __init__(self):
        self.dia = date.today().isoformat()
        self._dados: dict = {}               # corretora -> metricas (dia)
        self._janela: dict = {}              # corretora -> metricas (janela movel)
        self._eventos: deque = deque()       # (ts, comprador, vendedor, qtd, fin)
        self.negocios_perdidos = 0           # ciclos com janela_perdida
        self._csv_bloqueado_desde: Optional[float] = None

    def _slot(self, dados: dict, nome: str) -> dict:
        return dados.setdefault(nome, {
            "qtd_compra": 0.0, "qtd_venda": 0.0,
            "fin_compra": 0.0, "fin_venda": 0.0, "negocios": 0})

    def processar(self, novos: list, janela_perdida: bool = False):
        hoje = date.today().isoformat()
        if hoje != self.dia:                 # virada de pregao zera o dia
            self.dia, self._dados = hoje, {}
            self._janela, self._eventos = {}, deque()
            self.negocios_perdidos = 0
        if janela_perdida:
            self.negocios_perdidos += 1
        agora = time.time()
        for n in novos:
            fin = n["qtd"] * n["preco"]
            c = self._slot(self._dados, n["comprador"])
            c["qtd_compra"] += n["qtd"]
            c["fin_compra"] += fin
            c["negocios"] += 1
            v = self._slot(self._dados, n["vendedor"])
            v["qtd_venda"] += n["qtd"]
            v["fin_venda"] += fin
            v["negocios"] += 1

            jc = self._slot(self._janela, n["comprador"])
            jc["qtd_compra"] += n["qtd"]
            jc["fin_compra"] += fin
            jc["negocios"] += 1
            jv = self._slot(self._janela, n["vendedor"])
            jv["qtd_venda"] += n["qtd"]
            jv["fin_venda"] += fin
            jv["negocios"] += 1
            self._eventos.append((agora, n["comprador"], n["vendedor"], n["qtd"], fin))
        self._podar_janela(agora)

    def _podar_janela(self, agora: float):
        """Remove da janela movel os eventos mais antigos que JANELA_CORRETORAS_S,
        descontando o que eles somaram (incremental - nunca reprocessa tudo)."""
        limite = agora - JANELA_CORRETORAS_S
        while self._eventos and self._eventos[0][0] < limite:
            _, comprador, vendedor, qtd, fin = self._eventos.popleft()
            jc = self._janela.get(comprador)
            if jc:
                jc["qtd_compra"] -= qtd
                jc["fin_compra"] -= fin
                jc["negocios"] -= 1
            jv = self._janela.get(vendedor)
            if jv:
                jv["qtd_venda"] -= qtd
                jv["fin_venda"] -= fin
                jv["negocios"] -= 1

    @staticmethod
    def _montar_ranking(dados: dict) -> list:
        total = sum(d["qtd_compra"] + d["qtd_venda"]
                    for d in dados.values()) or 1.0
        out = []
        for nome, d in dados.items():
            qtd_total = d["qtd_compra"] + d["qtd_venda"]
            if qtd_total <= 0:            # zerou pela poda da janela - nao exibe
                continue
            fin_total = d["fin_compra"] + d["fin_venda"]
            out.append({
                "corretora": nome,
                "saldo": round(d["qtd_compra"] - d["qtd_venda"], 0),
                "qtd_compra": d["qtd_compra"], "qtd_venda": d["qtd_venda"],
                "preco_medio": round(fin_total / qtd_total, 1) if qtd_total else None,
                "participacao_pct": round(qtd_total / total * 100, 1),
                "negocios": d["negocios"],
            })
        out.sort(key=lambda x: x["saldo"], reverse=True)
        return out

    def ranking(self) -> list:
        """Corretoras por saldo (compra - venda) DESDE A ABERTURA, maior comprador primeiro."""
        return self._montar_ranking(self._dados)

    def ranking_janela(self) -> list:
        """Corretoras por saldo dentro dos ultimos JANELA_CORRETORAS_S segundos -
        quem esta agredindo AGORA, nao desde a abertura."""
        self._podar_janela(time.time())
        return self._montar_ranking(self._janela)

    def snapshot_db(self) -> list:
        """Linhas para o upsert no SQLite (valores absolutos do dia)."""
        return [(self.dia, nome, d["qtd_compra"], d["qtd_venda"],
                 d["fin_compra"], d["fin_venda"], d["negocios"])
                for nome, d in self._dados.items()]

    def snapshot_serie(self, ts: str) -> list:
        """Linhas para o INSERT na serie temporal (corretoras_serie): um ponto
        no tempo do acumulado CUMULATIVO do dia por corretora. Guardar o
        cumulativo (nao a janela ja calculada) deixa reconstruir depois
        qualquer janela por diferenca entre dois pontos - nao fica presa ao
        tamanho de JANELA_CORRETORAS_S vigente no momento da gravacao."""
        return [(self.dia, ts, nome, d["qtd_compra"], d["qtd_venda"],
                 d["fin_compra"], d["fin_venda"], d["negocios"])
                for nome, d in self._dados.items()]

    def escrever_csv(self, path: Path):
        """Espelha o acumulado em CSV para o VBA colar na aba Acum.

        A formula do ranking na planilha (DADOS!R21) prefere a Acum quando ela
        tem dados - e assim o ranking VISIVEL na planilha passa a ser o do dia
        inteiro, nao o da janela de ~60s. Primeira linha carrega o dia: o VBA
        descarta arquivo de pregao anterior em vez de colar dado velho.
        Numeros inteiros e sem separador de milhar (Val() do VBA nao le
        virgula decimal BR nem sufixo).
        """
        linhas = [f"# dia={self.dia}",
                  "Corretora;QtdCompra;QtdVenda;FinCompra;FinVenda"]
        for nome, d in self._dados.items():
            linhas.append(f"{nome};{d['qtd_compra']:.0f};{d['qtd_venda']:.0f};"
                          f"{d['fin_compra']:.0f};{d['fin_venda']:.0f}")
        tmp = path.with_suffix(".tmp")
        tmp.write_text("\n".join(linhas), encoding="utf-8")
        self._replace_com_retry(tmp, path)

    def _replace_com_retry(self, tmp: Path, path: Path,
                            tentativas: int = 3, espera_s: float = 0.05):
        """Retry curto cobre a colisao rapida do AtualizarAcum (VBA). Se
        persistir, o mais provavel e o arquivo estar aberto manualmente em
        outro programa - retry nao ajuda nesse caso, so aguardar. Loga UMA
        vez por bloqueio (nao a cada ciclo de 1s)."""
        for tentativa in range(tentativas):
            try:
                tmp.replace(path)             # troca atomica (VBA nunca le pela metade)
                if self._csv_bloqueado_desde is not None:
                    dur = time.time() - self._csv_bloqueado_desde
                    print(f"[WDO] ranking_acum.csv liberado (ficou {dur:.0f}s bloqueado)")
                    self._csv_bloqueado_desde = None
                return
            except PermissionError:
                if tentativa < tentativas - 1:
                    time.sleep(espera_s * (2 ** tentativa))
        if self._csv_bloqueado_desde is None:
            self._csv_bloqueado_desde = time.time()
            print("[WDO] ranking_acum.csv bloqueado por outro processo - "
                  "confira se nao esta aberto no Excel/Bloco de notas")
        tmp.unlink(missing_ok=True)

    def carregar(self, linhas: list):
        """Restaura o acumulado do dia (boot apos restart intradiario)."""
        for dia, nome, qc, qv, fc, fv, n in linhas:
            if dia != self.dia:
                continue
            self._dados[nome] = {
                "qtd_compra": qc or 0.0, "qtd_venda": qv or 0.0,
                "fin_compra": fc or 0.0, "fin_venda": fv or 0.0,
                "negocios": int(n or 0)}


# ----------------------------------------------------------------
# MACRO: BRENT + OURO + DXY (Yahoo) e JUROS EUA (Yahoo)
# ----------------------------------------------------------------
class MacroFetcher:
    """Busca Brent, Ouro, DXY e Juros EUA (^TNX) no Yahoo Finance.

    Fonte unica e isolada aqui (mesmo desenho do BrentFetcher do PETR4):
    para trocar a fonte, basta reimplementar fetch_symbol().
    Em caso de erro mantem a ultima cotacao valida de cada simbolo.
    """

    def __init__(self):
        self.last: dict[str, dict] = {}      # chave -> ultima cotacao valida
        self.last_ok: float = 0.0

    async def fetch_symbol(self, client: httpx.AsyncClient,
                           sym: str) -> Optional[dict]:
        try:
            r = await client.get(YAHOO_CHART.format(sym=sym),
                                 headers=MACRO_HEADERS, timeout=10)
            r.raise_for_status()
            meta = r.json()["chart"]["result"][0]["meta"]
            preco = meta.get("regularMarketPrice")
            prev  = meta.get("chartPreviousClose") or meta.get("previousClose")
            if preco is None or not prev:
                return None
            return {
                "preco": round(float(preco), 4),
                "fech_ant": round(float(prev), 4),
                "var_pct": round((float(preco) / float(prev) - 1) * 100, 2),
                # diferenca em pontos-base (usado pelo juros_us, que e' uma
                # taxa - variacao percentual do valor da taxa nao faz
                # sentido pra ela; ignorado pelos demais simbolos)
                "var_bps": round((float(preco) - float(prev)) * 100, 1),
                "ts": time.strftime("%H:%M:%S"),
            }
        except Exception:
            return None                      # mantem a ultima cotacao valida

    async def fetch_all(self, client: httpx.AsyncClient) -> dict[str, dict]:
        for key, sym in MACRO_SYMBOLS.items():
            q = await self.fetch_symbol(client, sym)
            if q:
                self.last[key] = q
                self.last_ok = time.time()
        return self.last


# ----------------------------------------------------------------
# DOLAR A VISTA (spot) para o CASADO
# ----------------------------------------------------------------
class SpotFetcher:
    """Dolar comercial a vista: AwesomeAPI (sem chave) com fallback pro
    Yahoo BRL=X. Mesmo desenho do MacroFetcher: mantem a ultima cotacao
    valida e o instante dela (last_ok) - o card CASADO mostra a idade
    porque spot travado vira desvio falso (a vista correndo atras do
    futuro, nao fluxo).
    """

    def __init__(self):
        self.last: Optional[dict] = None
        self.last_ok: float = 0.0

    async def fetch(self, client: httpx.AsyncClient) -> Optional[dict]:
        q = await self._awesome(client) or await self._yahoo(client)
        if q:
            self.last = q
            self.last_ok = time.time()
        return self.last

    async def _awesome(self, client: httpx.AsyncClient) -> Optional[dict]:
        try:
            r = await client.get(SPOT_URL_AWESOME, headers=MACRO_HEADERS, timeout=10)
            r.raise_for_status()
            d = r.json()["USDBRL"]
            bid, ask = float(d["bid"]), float(d["ask"])
            mid = (bid + ask) / 2
            return {
                "rate": round(mid, 4),
                "bid": bid, "ask": ask,
                "var_pct": round(float(d.get("pctChange") or 0.0), 2),
                "fonte": "AwesomeAPI",
                "ts": d.get("create_date", "")[-8:],
            }
        except Exception:
            return None

    async def _yahoo(self, client: httpx.AsyncClient) -> Optional[dict]:
        try:
            r = await client.get(YAHOO_CHART.format(sym=SPOT_SYM_YAHOO),
                                 headers=MACRO_HEADERS, timeout=10)
            r.raise_for_status()
            meta = r.json()["chart"]["result"][0]["meta"]
            preco = meta.get("regularMarketPrice")
            prev = meta.get("chartPreviousClose") or meta.get("previousClose")
            if preco is None or not prev:
                return None
            return {
                "rate": round(float(preco), 4),
                "bid": None, "ask": None,
                "var_pct": round((float(preco) / float(prev) - 1) * 100, 2),
                "fonte": "Yahoo BRL=X",
                "ts": time.strftime("%H:%M:%S"),
            }
        except Exception:
            return None


class MacroRtdReader:
    """Le o DI curto de dados_macro_rtd.csv (DI1*/DOLFUT via RTD/Excel).

    Cache por mtime (mesmo padrao do CsvReader). So o DI interessa aqui -
    o card MACRO nao depende mais deste arquivo. Escolhe o contrato DI1 de
    menor vencimento futuro como proxy do juro domestico ate o vencimento
    do WDO (tenor curto, a curva quase nao inclina nesse trecho).
    """

    def __init__(self, path: Path):
        self.path = path
        self._mtime = 0.0
        self._di_anual: Optional[float] = None
        self._di_ticker: Optional[str] = None

    def _venc_di(self, tk: str) -> Optional[date]:
        try:
            mes = _MES_COD_B3.get(tk[3].upper())
            ano = 2000 + int(tk[4:6])
            return date(ano, mes, 1) if mes else None
        except (ValueError, IndexError):
            return None

    def di_anual(self) -> Optional[float]:
        """Taxa DI curta em fracao (0.1389). None se o arquivo sumiu."""
        if not self.path.exists():
            return None
        mtime = self.path.stat().st_mtime
        if mtime == self._mtime:
            return self._di_anual
        self._mtime = mtime
        try:
            with open(self.path, encoding="utf-8-sig", errors="ignore") as f:
                rows = list(csv.reader(f, delimiter=";"))
        except (PermissionError, OSError):
            return self._di_anual
        hoje = date.today()
        melhor: tuple = (None, None)     # (venc, taxa)
        for row in rows:
            if len(row) < 2 or not row[0].upper().startswith("DI1"):
                continue
            venc = self._venc_di(row[0].strip())
            taxa = _to_float(row[1])
            if venc is None or taxa is None or venc <= hoje:
                continue
            if melhor[0] is None or venc < melhor[0]:
                melhor = (venc, taxa)
                self._di_ticker = row[0].strip()
        self._di_anual = (melhor[1] / 100.0) if melhor[1] is not None else None
        return self._di_anual


# ----------------------------------------------------------------
# NIVEIS DO DIA (SRP: persistencia dos niveis)
# ----------------------------------------------------------------
class LevelStore:
    """Cache por mtime (mesmo padrao do CsvReader): so rele o arquivo
    quando ele mudou — o load() e chamado no loop quente a cada tick."""

    def __init__(self, path: Path):
        self.path = path
        self._cache: dict = {}
        self._mtime = 0.0

    def load(self) -> dict:
        if not self.path.exists():
            return {}
        mtime = self.path.stat().st_mtime
        if mtime != self._mtime:
            self._cache = json.loads(self.path.read_text(encoding="utf-8"))
            self._mtime = mtime
        return self._cache

    def save(self, data: dict) -> dict:
        self.path.write_text(
            json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8"
        )
        self._cache = data
        self._mtime = self.path.stat().st_mtime
        return data


# ----------------------------------------------------------------
# NIVEIS AUTOMATICOS (pivot points classicos + Central Pivot Range)
# ----------------------------------------------------------------
TICK_WDO = 0.5   # tick do DOLFUT (WIN usa 5 pontos; WDO usa 0,5 ponto)


def _tick_wdo(v: float) -> float:
    """Arredonda para o multiplo de 0,5 mais proximo (tick do WDO)."""
    return round(round(v / TICK_WDO) * TICK_WDO, 1)


def calcular_niveis_do_dia(base: dict, contrato: str) -> dict:
    """Pivot points classicos a partir do OHLC do pregao anterior.

    P  = (H + L + C) / 3          (pivot central)
    R1 = 2P - L    S1 = 2P - H
    R2 = P + (H-L) S2 = P - (H-L)
    R3 = H + 2(P-L) S3 = L - 2(H-P)

    Zona decisiva = Central Pivot Range (CPR):
      BC = (H + L) / 2   TC = 2P - BC   (zona entre BC e TC)
    """
    h, l, c = base["maxima"], base["minima"], base["fechamento"]
    aviso = " (cobertura parcial)" if base.get("cobertura_parcial") else ""
    p  = (h + l + c) / 3
    bc = (h + l) / 2
    tc = 2 * p - bc
    r1, s1 = 2 * p - l,       2 * p - h
    r2, s2 = p + (h - l),     p - (h - l)
    r3, s3 = h + 2 * (p - l), l - 2 * (h - p)
    return {
        "data_pregao": date.today().isoformat(),
        "base_dia": base["dia"],
        "contrato": contrato,
        "fonte": f"Pivots automáticos (OHLC {base['dia']}){aviso}",
        "pivot": _tick_wdo(p),
        "resistencias": [_tick_wdo(r1), _tick_wdo(r2), _tick_wdo(r3)],
        "suportes":     [_tick_wdo(s1), _tick_wdo(s2), _tick_wdo(s3)],
        "alvos_compra": [_tick_wdo(r1), _tick_wdo(r2), _tick_wdo(r3)],
        "alvos_venda":  [_tick_wdo(s1), _tick_wdo(s2), _tick_wdo(s3)],
        "zona_decisiva": {"min": _tick_wdo(min(bc, tc)), "max": _tick_wdo(max(bc, tc))},
        "ladder": {"min": _tick_wdo(s3), "max": _tick_wdo(r3)},
        "maxima_ant": h,
        "minima_ant": l,
        "fechamento_ant": c,
    }


def dia_coberto(primeiro_ts: Optional[str], ultimo_ts: Optional[str]) -> bool:
    """O server esteve no ar da abertura ao fechamento deste pregao?

    So um dia coberto de ponta a ponta tem H/L confiaveis para virar base
    dos pivots do dia seguinte.
    """
    if not primeiro_ts or not ultimo_ts:
        return False
    return (primeiro_ts[:5] <= PREGAO_ABERTURA_ATE
            and ultimo_ts[:5] >= PREGAO_FECHA_APOS)


_vwap_avisado = False


def vwap_plausivel(t: Tick) -> Optional[float]:
    """Devolve o VWAP so se ele for um PRECO; senao None (nao grava lixo).

    Blindagem contra o campo RTD errado chegar na coluna do VWAP: em 09/07/2026
    o VBA passou a ler a agressao (~950.000) na vaga do VWAP e isso foi parar
    no daily_ohlc sem ninguem perceber. VWAP tem que cair dentro do range do
    dia; damos 1% de folga para arredondamento.
    """
    global _vwap_avisado
    if t.vwap is None:
        return None
    if t.maxima is None or t.minima is None:
        return t.vwap
    if t.minima * 0.99 <= t.vwap <= t.maxima * 1.01:
        _vwap_avisado = False
        return t.vwap
    if not _vwap_avisado:
        print(f"[WDO] ALERTA: vwap={t.vwap:,.2f} fora do range do dia "
              f"({t.minima:,.2f}-{t.maxima:,.2f}) - descartado. Conferir a "
              f"coluna do campo RTD 67 no ExportarWDO.bas")
        _vwap_avisado = True
    return None


# ----------------------------------------------------------------
# HISTORICO SQLITE (snapshot nao-bloqueante)
# ----------------------------------------------------------------
class SnapshotDB:
    CAMPOS_FLUXO = ("bid", "ask", "spread", "qtd_bid", "qtd_ask",
                    "desequilibrio", "profundidade_bid", "profundidade_ask",
                    "amostra_negocios", "amostra_contratos", "janela_s",
                    "negocios_s", "contratos_s", "lote_medio", "lote_max",
                    "agr_compra_fita", "agr_venda_fita", "pressao_fita",
                    "poc", "vah", "val", "vap_total", "vol_acima_pct",
                    "dist_poc")

    def __init__(self, path: Path):
        self.path = path
        self._init()

    def _init(self):
        with sqlite3.connect(self.path) as con:
            con.execute("""
                CREATE TABLE IF NOT EXISTS snapshots (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ts TEXT, ultimo REAL, abertura REAL, maxima REAL,
                    minima REAL, volume REAL, agr_compra REAL,
                    agr_venda REAL, vwap REAL
                )""")
            con.execute("""
                CREATE TABLE IF NOT EXISTS daily_ohlc (
                    dia TEXT PRIMARY KEY,
                    abertura REAL, maxima REAL, minima REAL,
                    fechamento REAL, vwap REAL
                )""")
            con.execute("""
                CREATE TABLE IF NOT EXISTS eventos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    dia TEXT, ts TEXT, evento TEXT, direcao TEXT,
                    nivel REAL, delta_ema REAL, msg TEXT
                )""")
            con.execute("""
                CREATE TABLE IF NOT EXISTS niveis_hist (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    dia TEXT, ts TEXT, origem TEXT, dados TEXT
                )""")
            # Registro manual do trader (painel "REGISTRO DO TRADER"): cada
            # entrada guarda o contexto de mercado do instante -> dataset de
            # estudo do ativo (cruzar com confluencia/backtest).
            con.execute("""
                CREATE TABLE IF NOT EXISTS operacoes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ts TEXT, dia TEXT, hora TEXT,
                    lado TEXT, preco REAL, motivo TEXT, nota TEXT,
                    abertura REAL, maxima REAL, minima REAL,
                    volume REAL, delta REAL, vwap REAL
                )""")
            con.execute("""
                CREATE TABLE IF NOT EXISTS fluxo (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    dia TEXT, ts TEXT,
                    bid REAL, ask REAL, spread REAL,
                    qtd_bid REAL, qtd_ask REAL, desequilibrio REAL,
                    profundidade_bid REAL, profundidade_ask REAL,
                    amostra_negocios INTEGER, amostra_contratos REAL,
                    janela_s REAL, negocios_s REAL, contratos_s REAL,
                    lote_medio REAL, lote_max REAL,
                    agr_compra_fita REAL, agr_venda_fita REAL,
                    pressao_fita REAL
                )""")
            con.execute("""
                CREATE TABLE IF NOT EXISTS corretoras (
                    dia TEXT, corretora TEXT,
                    qtd_compra REAL, qtd_venda REAL,
                    fin_compra REAL, fin_venda REAL,
                    negocios INTEGER,
                    PRIMARY KEY (dia, corretora)
                )""")
            # Serie temporal do ranking (um snapshot do CUMULATIVO do dia a
            # cada SERIE_CORRETORAS_S) - a tabela 'corretoras' acima so guarda
            # o ultimo estado (sobrescrito); esta guarda o historico intradia,
            # pra depois do pregao reconstruir por diferenca entre dois pontos
            # qual corretora virou de lado e quando (ver ranking_janela do
            # RankingCorretoras, que faz isso em tempo real com janela fixa).
            con.execute("""
                CREATE TABLE IF NOT EXISTS corretoras_serie (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    dia TEXT, ts TEXT, corretora TEXT,
                    qtd_compra REAL, qtd_venda REAL,
                    fin_compra REAL, fin_venda REAL,
                    negocios INTEGER
                )""")
            con.execute("""
                CREATE TABLE IF NOT EXISTS macro_snapshots (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    dia TEXT, ts TEXT,
                    brent REAL, brent_var REAL,
                    dxy REAL, dxy_var REAL,
                    ouro REAL, ouro_var REAL,
                    juros_us REAL, juros_us_var_bps REAL
                )""")
            # Leituras de IA (Gemini): antes so iam pro WS e se perdiam - sem
            # historico nao da pra medir taxa de acerto (ver
            # backtest_confluencia.py). gatilho_tipo: confluencia/periodica/manual.
            # vies e do modelo; vies_mercado/forca_mercado sao o consenso
            # CALCULADO (agente_wdo.vies_consolidado), pra poder medir os dois
            # separado.
            con.execute("""
                CREATE TABLE IF NOT EXISTS leituras (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    dia TEXT, ts TEXT,
                    gatilho_tipo TEXT, gatilho_nivel REAL, gatilho_direcao TEXT,
                    vies TEXT, vies_mercado TEXT, forca_mercado INTEGER,
                    resumo TEXT, evidencias TEXT, alertas TEXT, ressalvas TEXT,
                    preco REAL, modelo TEXT, cache INTEGER
                )""")
            # migracao: tabela criada com S&P 500 (herdado do WIN) -> Brent
            # (24/08/2026, ver nota do MACRO_SYMBOLS no topo do arquivo)
            for old, new in (("sp500", "brent"), ("sp500_var", "brent_var")):
                try:
                    con.execute(f"ALTER TABLE macro_snapshots "
                                f"RENAME COLUMN {old} TO {new}")
                except sqlite3.OperationalError:
                    pass                     # coluna ja renomeada/inexistente
            # DXY adicionado depois da criacao original da tabela
            for col in ("dxy", "dxy_var"):
                try:
                    con.execute(f"ALTER TABLE macro_snapshots ADD COLUMN {col} REAL")
                except sqlite3.OperationalError:
                    pass                     # coluna ja existe
            # CASADO (27/08/2026): preco justo do WDO vs dolar a vista.
            # spot = taxa a vista (5.16), wdo_pts = ultimo do futuro, du =
            # dias uteis ate o vencimento da frente. diferencial/carrego/
            # desvio em PONTOS. A calibracao do carrego (casado_calib_sync)
            # regride diferencial~du sobre a coluna `diferencial` + `du`.
            for col in ("spot", "spot_var", "wdo_pts", "diferencial", "du",
                        "carrego_justo", "desvio", "cupom_impl"):
                try:
                    con.execute(f"ALTER TABLE macro_snapshots ADD COLUMN {col} REAL")
                except sqlite3.OperationalError:
                    pass                     # coluna ja existe
            # migracao: Dolar (USD/BRL Yahoo/DOLFUT RTD, redundante com o
            # preco principal) -> Ouro; DI (RTD) -> Juros EUA (25/08/2026,
            # ver nota do MACRO_SYMBOLS no topo do arquivo)
            for old, new in (("dolar", "ouro"), ("dolar_var", "ouro_var"),
                              ("di", "juros_us"),
                              ("di_var_bps", "juros_us_var_bps")):
                try:
                    con.execute(f"ALTER TABLE macro_snapshots "
                                f"RENAME COLUMN {old} TO {new}")
                except sqlite3.OperationalError:
                    pass                     # coluna ja renomeada/inexistente
            # snapshots antigos nao tinham a coluna de data
            try:
                con.execute("ALTER TABLE snapshots ADD COLUMN dia TEXT")
            except sqlite3.OperationalError:
                pass                     # coluna ja existe
            # Indices por dia: snapshots (88k+ linhas) e fluxo (238k+) cresciam
            # sem indice e todo endpoint que filtra WHERE dia=? (historico,
            # ranking, backtest) fazia table scan crescente. PRECISA vir depois
            # do ALTER TABLE snapshots ADD COLUMN dia acima - banco novo (sem
            # migracao previa) nao tem a coluna ainda nesse ponto e o CREATE
            # INDEX falhava com "no such column: dia" (bug herdado do
            # server_win.py, so nao aparecia la porque o win_history.db em
            # producao ja tinha a coluna de uma migracao anterior).
            con.execute("CREATE INDEX IF NOT EXISTS idx_snapshots_dia_ts "
                        "ON snapshots(dia, ts)")
            con.execute("CREATE INDEX IF NOT EXISTS idx_fluxo_dia_ts "
                        "ON fluxo(dia, ts)")
            con.execute("CREATE INDEX IF NOT EXISTS idx_eventos_dia "
                        "ON eventos(dia)")
            con.execute("CREATE INDEX IF NOT EXISTS idx_macro_dia "
                        "ON macro_snapshots(dia)")
            con.execute("CREATE INDEX IF NOT EXISTS idx_leituras_dia "
                        "ON leituras(dia)")
            con.execute("CREATE INDEX IF NOT EXISTS idx_corretoras_serie_dia_corretora "
                        "ON corretoras_serie(dia, corretora, ts)")
            # campos do VAP adicionados depois da criacao original da tabela fluxo
            for col in ("poc", "vah", "val", "vap_total", "vol_acima_pct", "dist_poc"):
                try:
                    con.execute(f"ALTER TABLE fluxo ADD COLUMN {col} REAL")
                except sqlite3.OperationalError:
                    pass                     # coluna ja existe
            # campos de amostra/janela da fita adicionados depois da criacao
            # original: tabelas antigas (com negocios/contratos/fita_estourou)
            # nao os tinham -> INSERT falhava e derrubava o market_loop (22/07).
            for col, tipo in (("amostra_negocios", "INTEGER"),
                              ("amostra_contratos", "REAL"), ("janela_s", "REAL"),
                              ("negocios_s", "REAL"), ("contratos_s", "REAL")):
                try:
                    con.execute(f"ALTER TABLE fluxo ADD COLUMN {col} {tipo}")
                except sqlite3.OperationalError:
                    pass                     # coluna ja existe
            # cobertura do pregao (para preferir pivots de dias completos)
            for col, tipo in (("primeiro_ts", "TEXT"), ("ultimo_ts", "TEXT"),
                              ("valido", "INTEGER")):
                try:
                    con.execute(f"ALTER TABLE daily_ohlc ADD COLUMN {col} {tipo}")
                except sqlite3.OperationalError:
                    pass                 # coluna ja existe
            self._recalcular_cobertura(con)

    @staticmethod
    def _recalcular_cobertura(con):
        """Refaz primeiro_ts/ultimo_ts/valido de todos os dias pelos snapshots.

        Roda a cada boot: e a fonte da verdade sobre cobertura, entao um dia
        gravado por uma versao antiga do server (ou um server que caiu antes
        do fechamento) e reclassificado sozinho, sem manutencao manual.
        """
        for (dia,) in con.execute("SELECT dia FROM daily_ohlc").fetchall():
            ini, fim = con.execute(
                "SELECT MIN(ts), MAX(ts) FROM snapshots WHERE dia=?",
                (dia,)).fetchone()
            con.execute(
                "UPDATE daily_ohlc SET primeiro_ts=?, ultimo_ts=?, valido=? "
                "WHERE dia=?",
                (ini, fim, 1 if dia_coberto(ini, fim) else 0, dia))

    def save_sync(self, t: Tick):
        """Chamado via run_in_executor - roda em thread separada."""
        vwap = vwap_plausivel(t)
        with sqlite3.connect(self.path) as con:
            con.execute(
                "INSERT INTO snapshots (dia,ts,ultimo,abertura,maxima,minima,"
                "volume,agr_compra,agr_venda,vwap) VALUES (?,?,?,?,?,?,?,?,?,?)",
                (date.today().isoformat(), t.timestamp, t.ultimo, t.abertura,
                 t.maxima, t.minima, t.volume, t.agr_compra, t.agr_venda, vwap),
            )
            # Consolida o OHLC do dia (ultimo tick do pregao = fechamento) e
            # marca a janela de cobertura (primeiro/ultimo ts vistos hoje) -
            # ponta a ponta pode servir de base para os pivots (ver dia_valido).
            hoje = date.today().isoformat()
            con.execute(
                "INSERT INTO daily_ohlc (dia,abertura,maxima,minima,fechamento,"
                "vwap,primeiro_ts,ultimo_ts,valido) "
                "VALUES (?,?,?,?,?,?,?,?,0) "
                "ON CONFLICT(dia) DO UPDATE SET "
                "abertura=excluded.abertura, maxima=excluded.maxima, "
                "minima=excluded.minima, fechamento=excluded.fechamento, "
                "vwap=COALESCE(excluded.vwap, daily_ohlc.vwap), "
                "ultimo_ts=excluded.ultimo_ts",
                (hoje, t.abertura, t.maxima, t.minima, t.ultimo, vwap,
                 t.timestamp, t.timestamp),
            )
            row = con.execute(
                "SELECT primeiro_ts, ultimo_ts FROM daily_ohlc WHERE dia=?",
                (hoje,)).fetchone()
            con.execute("UPDATE daily_ohlc SET valido=? WHERE dia=?",
                        (1 if dia_coberto(row[0], row[1]) else 0, hoje))

    def salvar_fluxo_sync(self, f: dict):
        campos = self.CAMPOS_FLUXO
        with sqlite3.connect(self.path) as con:
            con.execute(
                "INSERT INTO fluxo (dia,ts," + ",".join(campos) + ") VALUES ("
                + ",".join("?" * (len(campos) + 2)) + ")",
                (date.today().isoformat(), f.get("timestamp"),
                 *(f.get(c) for c in campos)))

    def salvar_corretoras_sync(self, linhas: list):
        """Chamado via run_in_executor - roda em thread separada."""
        with sqlite3.connect(self.path) as con:
            con.executemany(
                "INSERT INTO corretoras (dia,corretora,qtd_compra,qtd_venda,"
                "fin_compra,fin_venda,negocios) VALUES (?,?,?,?,?,?,?) "
                "ON CONFLICT(dia,corretora) DO UPDATE SET "
                "qtd_compra=excluded.qtd_compra, qtd_venda=excluded.qtd_venda, "
                "fin_compra=excluded.fin_compra, fin_venda=excluded.fin_venda, "
                "negocios=excluded.negocios", linhas)

    def carregar_corretoras_sync(self, dia: str) -> list:
        with sqlite3.connect(self.path) as con:
            return con.execute(
                "SELECT dia,corretora,qtd_compra,qtd_venda,fin_compra,"
                "fin_venda,negocios FROM corretoras WHERE dia=?",
                (dia,)).fetchall()

    def salvar_corretoras_serie_sync(self, linhas: list):
        """Chamado via run_in_executor - roda em thread separada. Cada
        chamada INSERE um ponto novo (nao e upsert - e serie temporal)."""
        with sqlite3.connect(self.path) as con:
            con.executemany(
                "INSERT INTO corretoras_serie (dia,ts,corretora,qtd_compra,"
                "qtd_venda,fin_compra,fin_venda,negocios) "
                "VALUES (?,?,?,?,?,?,?,?)", linhas)

    def carregar_corretoras_serie_sync(self, dia: str,
                                        corretora: Optional[str] = None) -> list:
        """Serie temporal do dia, ordenada por ts - base pro backtest de
        'quem virou de lado'. Filtra por corretora quando informado."""
        with sqlite3.connect(self.path) as con:
            con.row_factory = sqlite3.Row
            if corretora:
                rows = con.execute(
                    "SELECT * FROM corretoras_serie WHERE dia=? AND corretora=? "
                    "ORDER BY ts", (dia, corretora)).fetchall()
            else:
                rows = con.execute(
                    "SELECT * FROM corretoras_serie WHERE dia=? ORDER BY ts",
                    (dia,)).fetchall()
            return [dict(r) for r in rows]

    def ohlc_anterior(self, hoje: str) -> Optional[dict]:
        """OHLC do ultimo pregao ANTES de 'hoje' (pula fim de semana).

        Regra: pega o pregao mais RECENTE que seja coerente (C dentro de
        [L,H]) e cuja ABERTURA foi capturada (primeiro_ts <= abertura). O
        fechamento usado e o ULTIMO tick salvo daquele dia. Assim um dia que
        o server nao cobriu ate o fim (valido=0) NAO e descartado: pivots
        recentes valem mais que um dia velho perfeito. Quando a cobertura
        nao foi ate o fechamento oficial, marca 'cobertura_parcial' e o
        dashboard avisa. So exige a abertura porque sem ela o H/L fica
        inutilizavel (ex.: dia que so rodou de tarde).
        """
        coerente = ("SELECT * FROM daily_ohlc WHERE dia < ? "
                    "AND maxima IS NOT NULL AND minima IS NOT NULL "
                    "AND fechamento IS NOT NULL "
                    "AND fechamento BETWEEN minima AND maxima ")
        with sqlite3.connect(self.path) as con:
            con.row_factory = sqlite3.Row
            row = con.execute(
                coerente + "AND primeiro_ts IS NOT NULL AND primeiro_ts <= ? "
                "ORDER BY dia DESC LIMIT 1", (hoje, PREGAO_ABERTURA_ATE)
            ).fetchone()
            if row is None:
                # Nenhum dia com a abertura capturada: cai no mais recente
                # apenas coerente (ultimo recurso) em vez de ficar sem niveis.
                row = con.execute(
                    coerente + "ORDER BY dia DESC LIMIT 1", (hoje,)).fetchone()
        if row is None:
            return None
        return dict(row, cobertura_parcial=(row["valido"] != 1))

    def corrigir_fechamento(self, dia: str, fechamento: float) -> bool:
        """Substitui o fechamento gravado pelo oficial (RTD FEC).

        Se o FEC oficial cair fora do range H/L gravado, o problema nao e o
        fechamento: e o H/L, que ficou parcial porque o server nao cobriu o
        pregao inteiro. Nesse caso marca o dia como invalido em vez de gravar
        um OHLC impossivel.
        """
        with sqlite3.connect(self.path) as con:
            row = con.execute(
                "SELECT maxima, minima FROM daily_ohlc WHERE dia=?",
                (dia,)).fetchone()
            if row and row[0] is not None and row[1] is not None:
                if not (row[1] <= fechamento <= row[0]):
                    con.execute(
                        "UPDATE daily_ohlc SET valido=0 WHERE dia=?", (dia,))
                    print(f"[WDO] {dia}: FEC oficial {fechamento:,.2f} fora do "
                          f"range gravado ({row[1]:,.2f}-{row[0]:,.2f}) - "
                          f"cobertura parcial, dia marcado como invalido")
                    return False
            con.execute("UPDATE daily_ohlc SET fechamento=? WHERE dia=?",
                        (fechamento, dia))
        return True

    def ajustar_ohlc_parcial(self, dia: str, maxima: float,
                             minima: float, fechamento: float) -> None:
        """Grava o FEC oficial num pregao PARCIAL e estende H/L p/ inclui-lo.

        So faz sentido para dias sem cobertura ate o fim (valido=0): o
        fechamento oficial e um preco que negociou, entao a maxima real foi
        >= FEC e a minima real <= FEC. Estender o range e correto (nao
        fabrica dado) e evita descartar o FEC quando o tick capturado antes
        de o server sair ficou aquem do fechamento real. Mantem valido=0 -
        o dia continua marcado como cobertura parcial.
        """
        with sqlite3.connect(self.path) as con:
            con.execute(
                "UPDATE daily_ohlc SET maxima=?, minima=?, fechamento=? "
                "WHERE dia=?", (maxima, minima, fechamento, dia))

    # ---- REGISTRO DO TRADER -------------------------------------------
    def salvar_operacao_sync(self, d: dict) -> int:
        with sqlite3.connect(self.path) as con:
            cur = con.execute(
                "INSERT INTO operacoes (ts,dia,hora,lado,preco,motivo,nota,"
                "abertura,maxima,minima,volume,delta,vwap) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (time.strftime("%Y-%m-%dT%H:%M:%S"),
                 d.get("data_pregao") or date.today().isoformat(),
                 d.get("hora"), str(d.get("lado", "")).lower(),
                 d.get("preco"), d.get("motivo"), d.get("nota"),
                 d.get("abertura"), d.get("maxima"), d.get("minima"),
                 d.get("volume"), d.get("delta"), d.get("vwap")))
        return cur.lastrowid

    def listar_operacoes_sync(self, dia: Optional[str] = None) -> list:
        with sqlite3.connect(self.path) as con:
            con.row_factory = sqlite3.Row
            if dia:
                rows = con.execute(
                    "SELECT * FROM operacoes WHERE dia=? ORDER BY id DESC",
                    (dia,)).fetchall()
            else:
                rows = con.execute(
                    "SELECT * FROM operacoes ORDER BY id DESC").fetchall()
        return [dict(r) for r in rows]

    def remover_operacao_sync(self, op_id: int) -> bool:
        with sqlite3.connect(self.path) as con:
            cur = con.execute("DELETE FROM operacoes WHERE id=?", (op_id,))
        return cur.rowcount > 0

    def save_evento_sync(self, ev: dict):
        """Persiste sinal de confluencia/divergencia para analise futura."""
        with sqlite3.connect(self.path) as con:
            con.execute(
                "INSERT INTO eventos (dia,ts,evento,direcao,nivel,delta_ema,msg) "
                "VALUES (?,?,?,?,?,?,?)",
                (date.today().isoformat(), time.strftime("%H:%M:%S"),
                 ev.get("evento"), ev.get("direcao"), ev.get("nivel"),
                 ev.get("delta_ema"), ev.get("msg")))

    def salvar_leitura_sync(self, leitura: dict, gatilho: Optional[dict] = None,
                            preco: Optional[float] = None):
        """Persiste a leitura de IA (vies do modelo + vies_mercado calculado)
        pra dar dado ao backtest_confluencia.py - ver nota da tabela em _init."""
        gatilho = gatilho or {}
        vm = leitura.get("vies_mercado") or {}
        with sqlite3.connect(self.path) as con:
            con.execute(
                "INSERT INTO leituras (dia,ts,gatilho_tipo,gatilho_nivel,"
                "gatilho_direcao,vies,vies_mercado,forca_mercado,resumo,"
                "evidencias,alertas,ressalvas,preco,modelo,cache) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (date.today().isoformat(), time.strftime("%H:%M:%S"),
                 gatilho.get("evento") or gatilho.get("tipo") or "manual",
                 gatilho.get("nivel"), gatilho.get("direcao"),
                 leitura.get("vies"), vm.get("vies"), vm.get("forca"),
                 leitura.get("resumo"),
                 json.dumps(leitura.get("evidencias") or [], ensure_ascii=False),
                 json.dumps(leitura.get("alertas") or [], ensure_ascii=False),
                 json.dumps(leitura.get("ressalvas") or [], ensure_ascii=False),
                 preco, leitura.get("modelo"), int(bool(leitura.get("cache")))))

    def historico_similar_sync(self, vies_mercado: str, forca_min: int = 2,
                               horizonte_min: int = 15, limite: int = 60
                               ) -> Optional[dict]:
        """Taxa de acerto REAL de leituras passadas com o MESMO vies_mercado
        calculado (mesma direcao, mesmo nivel de consenso >= forca_min) -
        espelha a logica de backtest_confluencia.backtest_leituras, mas so
        para o vies atual e sob demanda (le `leituras`+`snapshots`, nao
        escreve nada, nao muda o que ja roda). Da pra IA um numero de
        calibracao de confianca em vez de tratar toda leitura como sem
        passado. None se nao houver caso comparavel ainda (tabela nova ou
        poucos dias de coleta)."""
        if vies_mercado not in ("compra", "venda"):
            return None
        direcao_alta = (vies_mercado == "compra")
        with sqlite3.connect(self.path) as con:
            linhas = con.execute(
                "SELECT dia, ts, preco FROM leituras "
                "WHERE vies_mercado=? AND forca_mercado>=? "
                "ORDER BY dia DESC, ts DESC LIMIT ?",
                (vies_mercado, forca_min, limite)).fetchall()
            acertos = erros = neutros = sem_dado = 0
            for dia, ts, preco in linhas:
                baseline = preco
                if baseline is None:
                    row = con.execute(
                        "SELECT ultimo FROM snapshots WHERE dia=? AND ts>=? "
                        "AND ultimo IS NOT NULL ORDER BY ts LIMIT 1",
                        (dia, ts)).fetchone()
                    baseline = row[0] if row else None
                h, m, s = (int(x) for x in ts.split(":"))
                total = h * 3600 + m * 60 + s + horizonte_min * 60
                h2, resto = divmod(total % 86400, 3600)
                m2, s2 = divmod(resto, 60)
                ts_fut = f"{h2:02d}:{m2:02d}:{s2:02d}"
                row = con.execute(
                    "SELECT ultimo FROM snapshots WHERE dia=? AND ts>=? "
                    "AND ultimo IS NOT NULL ORDER BY ts LIMIT 1",
                    (dia, ts_fut)).fetchone()
                futuro = row[0] if row else None
                if baseline is None or futuro is None:
                    sem_dado += 1
                    continue
                delta = futuro - baseline
                if delta == 0:
                    neutros += 1
                elif (delta > 0) == direcao_alta:
                    acertos += 1
                else:
                    erros += 1
        base = acertos + erros
        if base == 0:
            return None
        return {
            "vies": vies_mercado, "forca_min": forca_min,
            "horizonte_min": horizonte_min,
            "n": base, "acerto_pct": round(100 * acertos / base, 1),
            "amostra_total": len(linhas),
        }

    def save_macro_sync(self, brent: Optional[dict], dxy: Optional[dict],
                        ouro: Optional[dict], juros_us: Optional[dict],
                        casado: Optional[dict] = None):
        """Historico macro para analise de correlacao com o WDO (+ casado)."""
        g = lambda d, k: d.get(k) if d else None
        c = casado or {}
        with sqlite3.connect(self.path) as con:
            con.execute(
                "INSERT INTO macro_snapshots (dia,ts,brent,brent_var,"
                "dxy,dxy_var,ouro,ouro_var,juros_us,juros_us_var_bps,"
                "spot,spot_var,wdo_pts,diferencial,du,carrego_justo,desvio,"
                "cupom_impl) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (date.today().isoformat(), time.strftime("%H:%M:%S"),
                 g(brent, "preco"), g(brent, "var_pct"),
                 g(dxy, "preco"), g(dxy, "var_pct"),
                 g(ouro, "preco"), g(ouro, "var_pct"),
                 g(juros_us, "preco"), g(juros_us, "var_bps"),
                 c.get("spot_rate"), c.get("spot_var"), c.get("wdo_pts"),
                 c.get("diferencial"), c.get("du"), c.get("carrego_justo"),
                 c.get("desvio"), c.get("cupom_impl_pct")))

    def casado_calib_sync(self, dias: int = 40) -> Optional[dict]:
        """Carrego 'justo' ajustado do historico: OLS diferencial~du com 1
        ponto por pregao (o primeiro diferencial do dia, que carrega menos
        ruido de fluxo). None ate MIN_DIAS_CALIB pregoes com a coluna nova.
        """
        corte = (date.today() - timedelta(days=dias)).isoformat()
        with sqlite3.connect(self.path) as con:
            rows = con.execute(
                "SELECT dia, du, diferencial FROM macro_snapshots "
                "WHERE diferencial IS NOT NULL AND du IS NOT NULL "
                "AND dia >= ? ORDER BY dia, id", (corte,)).fetchall()
        por_dia: dict = {}
        for dia, du, dif in rows:
            if dia not in por_dia:            # primeiro (abertura) de cada dia
                por_dia[dia] = (du, dif)
        return casado_wdo.regredir_carrego(list(por_dia.values()))

    def casado_desvio_hist_sync(self, hora_hhmm: str,
                                janela_min: int = 45,
                                dias: int = 8) -> list:
        """Desvios recentes na MESMA faixa de horario (+/- janela_min),
        pros ultimos `dias` pregoes - base do z-score do card CASADO."""
        from datetime import datetime as _dt
        try:
            t = _dt.strptime(hora_hhmm[:5], "%H:%M")
        except ValueError:
            return []
        lo = (t - timedelta(minutes=janela_min)).strftime("%H:%M")
        hi = (t + timedelta(minutes=janela_min)).strftime("%H:%M")
        corte = (date.today() - timedelta(days=dias)).isoformat()
        with sqlite3.connect(self.path) as con:
            rows = con.execute(
                "SELECT desvio FROM macro_snapshots "
                "WHERE desvio IS NOT NULL AND dia >= ? "
                "AND substr(ts,1,5) BETWEEN ? AND ? "
                "ORDER BY id DESC LIMIT 500", (corte, lo, hi)).fetchall()
        return [r[0] for r in rows]

    def log_niveis_sync(self, dados: dict, origem: str):
        """Registra cada versao dos niveis do dia (auto, refinado ou manual)."""
        with sqlite3.connect(self.path) as con:
            con.execute(
                "INSERT INTO niveis_hist (dia,ts,origem,dados) VALUES (?,?,?,?)",
                (date.today().isoformat(), time.strftime("%H:%M:%S"),
                 origem, json.dumps(dados, ensure_ascii=False)))

    def perfil_sync(self, dia: str, bucket: float = 1.0) -> dict:
        """Perfil de volume do pregao: distribui o volume negociado entre
        snapshots consecutivos no bucket de preco em que ocorreu.

        Aproximacao: todo o volume do intervalo e atribuido ao preco do
        snapshot que o fechou — quanto menor SNAPSHOT_EVERY, mais fiel.

        Retorna POC (preco de maior acumulo), value area (~70% do volume,
        expandida a partir do POC) e o histograma completo por bucket.

        bucket default = 1,0 ponto (herdado 50 do WIN, cujo range diario em
        pontos e ~50-100x maior que o do WDO - com tick de 0,5, 1 ponto ja
        da granularidade fina o bastante sem virar ruido).
        """
        with sqlite3.connect(self.path) as con:
            rows = con.execute(
                "SELECT ultimo, volume FROM snapshots WHERE dia=? "
                "AND ultimo IS NOT NULL AND volume IS NOT NULL "
                "ORDER BY id", (dia,)).fetchall()
            cobertura = con.execute(
                "SELECT MIN(ts), MAX(ts) FROM snapshots WHERE dia=?",
                (dia,)).fetchone()
        hist: dict[float, float] = {}
        prev_vol = None
        for ultimo, volume in rows:
            if prev_vol is not None:
                dv = volume - prev_vol
                if dv > 0:                    # ignora reset/duplicata
                    b = round(round(ultimo / bucket) * bucket, 2)
                    hist[b] = hist.get(b, 0.0) + dv
            prev_vol = volume
        base = {"dia": dia, "bucket": bucket, "amostras": len(rows),
                "cobertura": {"ini": cobertura[0], "fim": cobertura[1]}}
        if not hist:
            return {**base, "total": 0, "poc": None,
                    "va_min": None, "va_max": None, "perfil": []}
        total = sum(hist.values())
        poc = max(hist, key=lambda k: hist[k])
        # Value area: expande a partir do POC pelo vizinho de maior volume
        ordered = sorted(hist)
        i = j = ordered.index(poc)
        acc = hist[poc]
        while acc < 0.70 * total and (i > 0 or j < len(ordered) - 1):
            below = hist[ordered[i - 1]] if i > 0 else -1.0
            above = hist[ordered[j + 1]] if j < len(ordered) - 1 else -1.0
            if above >= below:
                j += 1; acc += hist[ordered[j]]
            else:
                i -= 1; acc += hist[ordered[i]]
        perfil = [{"preco": p, "vol": hist[p],
                   "pct": round(hist[p] / total * 100, 2)} for p in ordered]
        return {**base, "total": total, "poc": poc,
                "va_min": ordered[i], "va_max": ordered[j], "perfil": perfil}

    def plano_ativacao_sync(self, dia: str, zmin: Optional[float],
                            zmax: Optional[float]) -> dict:
        """Hora REAL em que o gatilho do plano foi ativado no pregao.

        Compra ativa quando a maxima do dia alcanca a ZD-topo; venda
        quando a minima alcanca a ZD-base. Como maxima/minima sao extremos
        cumulativos (monotonicos), o primeiro snapshot em que ja haviam
        cruzado marca o horario do rompimento (resolucao ~SNAPSHOT_EVERY).

        Calculado dos snapshots -> sobrevive a F5 e a reinicio do server no
        mesmo dia, e reflete o horario de mercado (nao o de carga da pagina).
        Recalcula sozinho se a zona decisiva mudar (niveis refinados).
        """
        out = {"compra": None, "venda": None}
        with sqlite3.connect(self.path) as con:
            if zmax is not None:
                r = con.execute(
                    "SELECT MIN(ts) FROM snapshots WHERE dia=? AND maxima>=?",
                    (dia, zmax)).fetchone()
                out["compra"] = r[0] if r and r[0] else None
            if zmin is not None:
                r = con.execute(
                    "SELECT MIN(ts) FROM snapshots WHERE dia=? AND minima<=?",
                    (dia, zmin)).fetchone()
                out["venda"] = r[0] if r and r[0] else None
        return out

    def historico_sync(self, dia: str) -> dict:
        """Pacote completo de um pregao: OHLC, niveis, eventos e snapshots."""
        with sqlite3.connect(self.path) as con:
            con.row_factory = sqlite3.Row
            ohlc = con.execute(
                "SELECT * FROM daily_ohlc WHERE dia=?", (dia,)).fetchone()
            nivs = [dict(r) for r in con.execute(
                "SELECT * FROM niveis_hist WHERE dia=? ORDER BY id", (dia,))]
            evs = [dict(r) for r in con.execute(
                "SELECT * FROM eventos WHERE dia=? ORDER BY id", (dia,))]
            snaps = [dict(r) for r in con.execute(
                "SELECT * FROM snapshots WHERE dia=? ORDER BY id", (dia,))]
        for n in nivs:
            n["dados"] = json.loads(n["dados"])
        return {"dia": dia, "ohlc": dict(ohlc) if ohlc else None,
                "niveis": nivs, "eventos": evs, "snapshots": snaps}


# ----------------------------------------------------------------
# GERENCIADOR DE CONEXOES WEBSOCKET
# ----------------------------------------------------------------
class ConnectionManager:
    def __init__(self):
        self.clients: list[WebSocket] = []

    async def connect(self, ws: WebSocket):
        await ws.accept()
        self.clients.append(ws)

    def disconnect(self, ws: WebSocket):
        if ws in self.clients:
            self.clients.remove(ws)

    async def broadcast(self, payload: dict):
        dead = []
        for ws in self.clients:
            try:
                await ws.send_json(payload)
            except Exception:
                dead.append(ws)
        for ws in dead:
            self.disconnect(ws)


# ----------------------------------------------------------------
# MOTOR DE CONFLUENCIA
# ----------------------------------------------------------------
class ConfluenceEngine:
    """Combina o MAPA (niveis estaticos do infografico) com o FLUXO
    (agressao em tempo real).

    Sinal de CONFLUENCIA = duas fontes INDEPENDENTES concordando:
      1. Estrutura : preco cruzou um nivel monitorado do niveis.json
                     (resistencias + suportes + alvos + zona decisiva,
                     deduplicados por preco)
      2. Fluxo     : EMA do fluxo INCREMENTAL de agressao confirma a
                     direcao E esta acelerando

    FLUXO INCREMENTAL: agr_compra/agr_venda do RTD sao ACUMULADOS do dia,
    entao delta = compra - venda tambem e acumulado. O que confirma um
    rompimento e o fluxo NOVO (delta_t - delta_{t-1}), nao o saldo do dia
    — um dia vendedor (delta acumulado negativo) pode ter rompimento de
    alta legitimo se a agressao compradora dominar AGORA.

    THRESHOLD ADAPTATIVO: em vez de numero magico fixo, o minimo para
    confirmar e K_STD x desvio-padrao movel do proprio fluxo — o sinal
    se recalibra sozinho ao regime de volume do dia/contrato (herdado do
    Monitor WIN, onde o delta opera em escala ~1000x maior que a do WDO;
    o desenho adaptativo existe justamente para nao precisar de um
    numero fixo por contrato).

    Persistencia: o cruzamento so vira sinal apos PERSIST_TICKS ticks
    consecutivos alem do nivel (mesma filosofia do SignalGuard do WDO).
    """

    EMA_ALPHA     = 0.15      # suavizacao do fluxo incremental
    PERSIST_TICKS = 3         # ticks consecutivos alem do nivel
    COOLDOWN_S    = 120       # segundos entre sinais no mesmo nivel
    K_STD         = 1.5       # threshold = K_STD x std movel do fluxo
    MIN_SAMPLES   = 30        # amostras de fluxo antes de confirmar (~1 min)
    FLOW_WINDOW   = 300       # janela movel do std (~10 min de ticks)

    def __init__(self, level_store: "LevelStore"):
        self.levels = level_store
        self.flow_ema: Optional[float] = None
        self.flow_ema_prev: Optional[float] = None
        self.prev_delta: Optional[float] = None   # delta ACUMULADO anterior
        self.prev_price: Optional[float] = None
        self._flows: deque[float] = deque(maxlen=self.FLOW_WINDOW)
        self._persist: dict[str, int] = {}     # chave nivel -> ticks contados
        self._last_fire: dict[str, float] = {}

    def reset_flow(self):
        """Virada de pregao: o acumulado do RTD zera, descartar estado."""
        self.flow_ema = self.flow_ema_prev = self.prev_delta = None
        self._flows.clear()

    def _update_delta(self, tick: Tick):
        d = tick.delta                       # acumulado do dia
        if d is None:
            return
        if self.prev_delta is None:
            self.prev_delta = d              # primeiro tick: sem incremento
            return
        flow = d - self.prev_delta           # fluxo novo deste tick
        self.prev_delta = d
        self._flows.append(flow)
        self.flow_ema_prev = self.flow_ema
        if self.flow_ema is None:
            self.flow_ema = flow
        else:
            self.flow_ema = self.EMA_ALPHA * flow + (1 - self.EMA_ALPHA) * self.flow_ema

    def _threshold(self) -> Optional[float]:
        """K_STD x desvio-padrao movel do fluxo; None enquanto aquece."""
        if len(self._flows) < self.MIN_SAMPLES:
            return None
        std = statistics.pstdev(self._flows)
        return self.K_STD * std if std > 0 else None

    def _flow_confirms(self, direction: str) -> bool:
        """Fluxo confirma se a EMA aponta na direcao, acima do threshold
        adaptativo, E esta acelerando."""
        thr = self._threshold()
        if thr is None or self.flow_ema is None:
            return False
        accelerating = (
            self.flow_ema_prev is not None
            and abs(self.flow_ema) > abs(self.flow_ema_prev)
        )
        if direction == "up":
            return self.flow_ema > thr and accelerating
        return self.flow_ema < -thr and accelerating

    def check(self, tick: Tick, fluxo: Optional[dict] = None) -> list[dict]:
        """Retorna lista de eventos de confluencia detectados neste tick.

        `fluxo` (vah/val/dentro_area) fica ~1 ciclo (POLL_INTERVAL) atrasado
        em relacao ao tick neste ponto do market_loop, pois e lido depois do
        bloco de confluencia. Sem problema pra VAH/VAL (nao muda tick a
        tick) - so documentar.
        """
        events: list[dict] = []
        self._update_delta(tick)
        price, prev = tick.ultimo, self.prev_price
        self.prev_price = price
        if price is None or prev is None:
            return events

        # Lista unificada de niveis monitorados: R/S, alvos e zona decisiva
        # sao semanticamente a mesma coisa (precos de decisao). Sets fazem
        # o dedupe — no modo automatico alvos == R/S e nao duplicam sinal.
        cfg  = self.levels.load()
        zona = cfg.get("zona_decisiva") or {}
        ups   = set(cfg.get("resistencias", [])) | set(cfg.get("alvos_compra", []))
        downs = set(cfg.get("suportes", []))     | set(cfg.get("alvos_venda", []))
        if zona.get("max"): ups.add(zona["max"])
        if zona.get("min"): downs.add(zona["min"])

        now = time.time()
        for lvl in sorted(ups):
            events += self._track(f"up:{lvl}", price > lvl, "up", lvl, now, fluxo)
        for lvl in sorted(downs, reverse=True):
            events += self._track(f"dn:{lvl}", price < lvl, "down", lvl, now, fluxo)
        return events

    def _track(self, key: str, beyond: bool, direction: str,
               lvl: float, now: float, fluxo: Optional[dict] = None) -> list[dict]:
        if not beyond:
            self._persist[key] = 0
            return []
        self._persist[key] = self._persist.get(key, 0) + 1
        if self._persist[key] != self.PERSIST_TICKS:
            return []                                  # ainda sem persistencia
        if now - self._last_fire.get(key, 0) < self.COOLDOWN_S:
            return []                                  # em cooldown
        contexto_va = self._contexto_value_area(fluxo)
        if not self._flow_confirms(direction):
            # rompeu MAS o fluxo nao confirma -> alerta de divergencia
            self._last_fire[key] = now
            return [{
                "evento": "divergencia", "direcao": direction, "nivel": lvl,
                "delta_ema": round(self.flow_ema or 0, 1),
                "contexto_value_area": contexto_va,
                "msg": f"Rompeu {lvl:.0f} SEM confirmacao de fluxo "
                       f"({contexto_va or 'sem VAH/VAL no momento'})"
            }]
        self._last_fire[key] = now
        return [{
            "evento": "confluencia", "direcao": direction, "nivel": lvl,
            "delta_ema": round(self.flow_ema, 1),
            "contexto_value_area": contexto_va,
            "msg": f"CONFLUENCIA: {'rompimento' if direction=='up' else 'perda'} "
                   f"de {lvl:.0f} confirmado pelo fluxo (EMA {self.flow_ema:+.0f}, "
                   f"{contexto_va or 'sem VAH/VAL'})"
        }]

    @staticmethod
    def _contexto_value_area(fluxo: Optional[dict]) -> Optional[str]:
        """Iniciativa/Responsiva (Dalton, cap. 25). Usa vah/val so como
        NIVEIS DE PRECO (a forma do perfil) - nunca vap_total, que e
        conhecidamente inconsistente no feed do RTD (mesma ressalva do
        agente_wdo.py e do docstring de FluxoReader._vap)."""
        if not fluxo or fluxo.get("vah") is None:
            return None
        return ("iniciativa (fora da area de valor)" if not fluxo.get("dentro_area")
                else "responsiva (ainda dentro da area de valor)")


# ----------------------------------------------------------------
# APP + LIFESPAN
# ----------------------------------------------------------------
reader   = CsvReader(CSV_PATH)
levels   = LevelStore(NIVEIS_PATH)
db       = SnapshotDB(DB_PATH)
manager  = ConnectionManager()
engine   = ConfluenceEngine(levels)
macro    = MacroFetcher()
spot     = SpotFetcher()
macro_rtd = MacroRtdReader(MACRO_RTD_CSV)
fluxo_reader = FluxoReader(BOOK_CSV, TT_CSV, VAP_CSV)
ranking = RankingCorretoras()
ranking.carregar(db.carregar_corretoras_sync(ranking.dia))   # sobrevive a restart
last_tick: Optional[Tick] = None
last_macro: Optional[dict] = None
last_casado: Optional[dict] = None
last_plano_ativacao: Optional[dict] = None
last_fluxo: Optional[dict] = None
_leitura_auto_em_andamento = False   # evita empilhar chamadas ao Gemini
# CASADO: calibracao do carrego (recalculada 1x/dia) + diferencial de
# abertura do dia (fallback enquanto nao ha calibracao suficiente).
_casado_calib: Optional[dict] = None
_casado_calib_dia: Optional[str] = None
_casado_ref: dict = {"dia": None, "dif_abertura": None}


def atualizar_niveis_automaticos(force: bool = False) -> Optional[dict]:
    """Recalcula os niveis se o niveis.json ainda for de outro pregao.

    Nao sobrescreve niveis definidos manualmente HOJE (data_pregao == hoje),
    a menos que force=True. Sem OHLC anterior no historico, mantem o manual.
    """
    cfg = levels.load()
    hoje = date.today().isoformat()
    base = db.ohlc_anterior(hoje)
    if base is None:
        return None
    if not force and cfg.get("data_pregao") == hoje:
        # Ja ha niveis de hoje. So refaz se forem AUTOMATICOS e a base de
        # OHLC tiver mudado - ex.: monitor reaberto depois de o pregao
        # anterior ter fechado, agora com o ultimo tick salvo daquele dia.
        # Niveis definidos MANUALMENTE hoje sao preservados.
        fonte_auto = str(cfg.get("fonte", "")).startswith("Pivots autom")
        if not (fonte_auto and cfg.get("base_dia") != base["dia"]):
            return None
    novo = calcular_niveis_do_dia(base, cfg.get("contrato", "WDO"))
    levels.save(novo)
    db.log_niveis_sync(novo, "auto")
    return novo


def refinar_niveis_com_fec(fec: float) -> Optional[dict]:
    """Ao abrir no dia seguinte, o Excel/RTD ja traz o FEC (fechamento
    oficial do pregao anterior). Usa esse valor para acertar a base dos
    pivots e reemitir regua, distancia aos niveis e plano do dia:

    - Dia anterior COMPLETO (cobertura ate 17:45): so troca o C se divergir
      >= 1 tick. Se o FEC cair fora do H/L gravado, o problema e o H/L ->
      marca invalido e refaz do ultimo pregao confiavel.
    - Dia anterior PARCIAL (NAO foi ate 17:45): o H/L pode estar truncado
      no ponto em que o server saiu. Adota o FEC como fechamento e ESTENDE
      o range para inclui-lo, para os niveis refletirem o fechamento oficial
      em vez do ultimo tick capturado antes do encerramento.

    So age sobre niveis automaticos; niveis manuais de hoje sao preservados.
    """
    cfg = levels.load()
    hoje = date.today().isoformat()
    if cfg.get("data_pregao") != hoje:
        return None
    if not str(cfg.get("fonte", "")).startswith("Pivots autom"):
        return None                          # niveis manuais: nao mexe
    base = db.ohlc_anterior(hoje)
    if base is None or abs(base["fechamento"] - fec) < TICK_WDO:
        return None                          # ja esta correto (< 1 tick)

    if base.get("cobertura_parcial"):
        # Pregao que nao fechou 17:45: confia no FEC e estende H/L se preciso.
        novo_h = max(base["maxima"], fec)
        novo_l = min(base["minima"], fec)
        db.ajustar_ohlc_parcial(base["dia"], novo_h, novo_l, fec)
        base = dict(base, maxima=novo_h, minima=novo_l, fechamento=fec)
    elif not db.corrigir_fechamento(base["dia"], fec):
        # Dia dito completo mas FEC fora do H/L: H/L nao era confiavel.
        # Aquele dia acabou de ser marcado invalido; refaz do proximo.
        return atualizar_niveis_automaticos(force=True)
    else:
        base = dict(base, fechamento=fec)

    novo = calcular_niveis_do_dia(base, cfg.get("contrato", "WDO"))
    novo["fonte"] += " · C=FEC oficial"
    levels.save(novo)
    db.log_niveis_sync(novo, "auto_fec")
    return novo


async def market_loop():
    """Loop principal: le CSV -> confluencia -> broadcast -> snapshot."""
    global last_tick, last_plano_ativacao, last_fluxo
    last_snapshot = 0.0
    last_ranking_bcast = 0.0
    last_serie_snapshot = 0.0
    last_leitura_auto = 0.0
    loop = asyncio.get_running_loop()
    dia_atual = date.today()
    fec_conferido = False
    novo = atualizar_niveis_automaticos()
    if novo:
        print(f"[WDO] Niveis do dia recalculados: {novo['fonte']}")
    print(f"[WDO] Loop iniciado. Observando {CSV_PATH.name} a cada {POLL_INTERVAL}s")
    while True:
        try:
            # Virada de dia com o server no ar: recalcula e avisa os dashboards
            if date.today() != dia_atual:
                dia_atual = date.today()
                fec_conferido = False
                engine.reset_flow()      # acumulados do RTD zeram no novo pregao
                novo = atualizar_niveis_automaticos()
                if novo:
                    print(f"[WDO] Niveis do dia recalculados: {novo['fonte']}")
                    await manager.broadcast({"evento": "niveis_atualizados"})
            tick = reader.read_if_changed()
            if tick:
                # Primeiro tick com FEC do dia: confere o C usado nos pivots
                if not fec_conferido and tick.fec_ant:
                    fec_conferido = True
                    novo = refinar_niveis_com_fec(tick.fec_ant)
                    if novo:
                        print(f"[WDO] Pivots refinados com FEC oficial "
                              f"({tick.fec_ant:.0f}): {novo['fonte']}")
                        await manager.broadcast({"evento": "niveis_atualizados"})
                last_tick = tick
                await manager.broadcast(asdict(tick))
                # Motor de confluencia: mapa (niveis) x fluxo (delta)
                for ev in engine.check(tick, fluxo=last_fluxo):
                    print(f"[WDO] {ev['msg']}")
                    await manager.broadcast(ev)
                    # persiste o sinal para analise historica
                    await loop.run_in_executor(None, db.save_evento_sync, ev)
                    if ev["evento"] == "confluencia":
                        # dispara a leitura de IA sozinha (background - nao
                        # espera a resposta da API para seguir lendo o proximo tick)
                        asyncio.create_task(gerar_leitura_automatica(ev))
                now = time.time()
                if now - last_snapshot >= SNAPSHOT_EVERY:
                    last_snapshot = now
                    # NAO bloqueia o event loop (corrige debito do server_v2)
                    await loop.run_in_executor(None, db.save_sync, tick)
                    # Hora real de ativacao do gatilho do plano (dos snapshots).
                    zona = (levels.load().get("zona_decisiva") or {})
                    ativ = await loop.run_in_executor(
                        None, db.plano_ativacao_sync, date.today().isoformat(),
                        zona.get("min"), zona.get("max"))
                    if ativ != last_plano_ativacao:
                        last_plano_ativacao = ativ
                        await manager.broadcast({"evento": "plano_ativacao", **ativ})
                # Leitura de IA periodica (alem do gatilho de confluencia): mantem
                # o painel atualizado em pregao parado, sem confluencia. A cada
                # ~15 min - contexto inclui macro (prompt maior).
                if now - last_leitura_auto >= LEITURA_AUTO_INTERVALO:
                    last_leitura_auto = now
                    asyncio.create_task(gerar_leitura_automatica(
                        {"tipo": "periodica", "msg": "leitura periodica"}))
            # Fluxo (livro + fita): lido todo ciclo porque a janela da fita e
            # curta (21 negocios) - pular ciclo significa perder negocio.
            fl = fluxo_reader.ler(last_tick.ultimo if last_tick else None)
            if fl:
                # compila o acumulado por corretora ANTES do broadcast (o campo
                # 'novos' e consumo interno, nao vai para os dashboards)
                novos = fl.pop("novos", [])
                perdida = fl.pop("janela_perdida", False)
                ranking.processar(novos, perdida)
                if novos:
                    await loop.run_in_executor(
                        None, db.salvar_corretoras_sync, ranking.snapshot_db())
                    await loop.run_in_executor(
                        None, ranking.escrever_csv, RANKING_CSV)
                last_fluxo = fl
                await manager.broadcast({"evento": "fluxo", **fl})
                if fl.get("amostra_negocios") or fl.get("bid"):
                    await loop.run_in_executor(None, db.salvar_fluxo_sync, fl)
                # ranking por WS a cada ~15s (payload maior, nao precisa de 2s)
                now_rk = time.time()
                if novos and now_rk - last_ranking_bcast >= 15:
                    last_ranking_bcast = now_rk
                    await manager.broadcast(
                        {"evento": "ranking", "dia": ranking.dia,
                         "corretoras": ranking.ranking()[:12],
                         "corretoras_janela": ranking.ranking_janela()[:12],
                         "janela_min": JANELA_CORRETORAS_S // 60,
                         "ciclos_perdidos": ranking.negocios_perdidos})
                # serie temporal do ranking (backtest de "quem virou de lado"
                # depois do pregao) - cadencia mais espacada, nao precisa
                # acompanhar o broadcast
                if novos and now_rk - last_serie_snapshot >= SERIE_CORRETORAS_S:
                    last_serie_snapshot = now_rk
                    await loop.run_in_executor(
                        None, db.salvar_corretoras_serie_sync,
                        ranking.snapshot_serie(last_tick.timestamp if last_tick
                                               else time.strftime("%H:%M:%S")))
        except Exception as e:
            # Blindagem: um erro num passo secundario (persistencia, leitura de
            # fluxo, etc.) NUNCA pode derrubar o loop e congelar o dashboard.
            # Loga e segue para o proximo ciclo. (Bug real 22/07: schema drift
            # da tabela fluxo matava o loop silenciosamente.)
            import traceback
            print(f"[WDO] ERRO no ciclo do market_loop (seguindo): "
                  f"{type(e).__name__}: {e}")
            traceback.print_exc()
        await asyncio.sleep(POLL_INTERVAL)


def _calcular_casado_sync() -> Optional[dict]:
    """Monta o pacote do casado a partir do ultimo tick do WDO + spot +
    DI + calibracao. Roda no executor (le CSV do DI, faz o z-score)."""
    global _casado_calib, _casado_calib_dia, _casado_ref
    if last_tick is None or last_tick.ultimo is None or spot.last is None:
        return None
    hoje = date.today().isoformat()

    # calibracao do carrego: 1x/dia
    if _casado_calib_dia != hoje:
        _casado_calib_dia = hoje
        try:
            _casado_calib = db.casado_calib_sync()
        except Exception as e:
            print(f"[WDO] casado_calib falhou (seguindo sem): {e}")
            _casado_calib = None

    venc = casado_wdo.vencimento_frente()
    du = casado_wdo.dias_uteis_ate(venc)
    wdo_pts = float(last_tick.ultimo)
    spot_rate = spot.last["rate"]
    diferencial = wdo_pts - spot_rate * casado_wdo.PONTOS_POR_DOLAR

    # diferencial de abertura (fallback de carrego; so captura cedo)
    if _casado_ref.get("dia") != hoje:
        _casado_ref = {"dia": hoje, "dif_abertura": None}
    if _casado_ref["dif_abertura"] is None and time.strftime("%H:%M") <= "10:00":
        _casado_ref["dif_abertura"] = round(diferencial, 1)

    hist = db.casado_desvio_hist_sync(time.strftime("%H:%M"))
    pkg = casado_wdo.calcular(
        wdo_pts=wdo_pts, spot_rate=spot_rate,
        dif_abertura=_casado_ref["dif_abertura"],
        calib=_casado_calib, di_anual=macro_rtd.di_anual(),
        du=du, venc=venc, desvio_hist=hist)
    if pkg:
        pkg["spot_var"] = spot.last.get("var_pct")
        pkg["spot_fonte"] = spot.last.get("fonte")
        pkg["idade_s"] = (round(time.time() - spot.last_ok, 1)
                          if spot.last_ok else None)
        pkg["calibrado"] = _casado_calib is not None
    return pkg


async def macro_loop():
    """Loop paralelo: Brent + DXY + Ouro + Juros EUA (Yahoo) + CASADO
    (dolar a vista via AwesomeAPI vs ultimo do WDO).

    Transmite o pacote consolidado via WS a cada MACRO_POLL segundos e
    persiste no SQLite para estudo de correlacao com o WDO.
    """
    global last_macro, last_casado
    loop = asyncio.get_running_loop()
    print(f"[WDO] Macro loop iniciado ({', '.join(MACRO_SYMBOLS.values())}"
          f" + casado) a cada {MACRO_POLL}s")
    async with httpx.AsyncClient() as client:
        while True:
            quotes = await macro.fetch_all(client)
            await spot.fetch(client)
            casado_pkg = await loop.run_in_executor(None, _calcular_casado_sync)
            if casado_pkg:
                last_casado = casado_pkg
            if quotes or last_casado:
                last_macro = {
                    "evento": "macro",
                    "brent": quotes.get("brent"),
                    "dxy": quotes.get("dxy"),
                    "ouro": quotes.get("ouro"),
                    "juros_us": quotes.get("juros_us"),
                    "casado": last_casado,
                    "ts": time.strftime("%H:%M:%S"),
                }
                await manager.broadcast(last_macro)
                await loop.run_in_executor(
                    None, db.save_macro_sync, quotes.get("brent"),
                    quotes.get("dxy"), quotes.get("ouro"),
                    quotes.get("juros_us"), last_casado)
            await asyncio.sleep(MACRO_POLL)


@asynccontextmanager
async def lifespan(app: FastAPI):
    t1 = asyncio.create_task(market_loop())
    t2 = asyncio.create_task(macro_loop())
    yield
    t1.cancel()
    t2.cancel()


app = FastAPI(title="Monitor WDO", lifespan=lifespan)
app.add_middleware(
    CORSMiddleware, allow_origins=["*"],
    allow_methods=["*"], allow_headers=["*"],
)


# ----------------------------------------------------------------
# ROTAS
# ----------------------------------------------------------------
@app.get("/")
async def dashboard():
    # no-store: o HTML muda com frequencia (layout) e os dados sao ao vivo -
    # todo refresh deve puxar a versao nova, sem depender de Ctrl+F5.
    return FileResponse(DASHBOARD, headers={"Cache-Control": "no-store, max-age=0"})


@app.get("/manifest.json")
async def manifest():
    return FileResponse(BASE_DIR / "manifest.json", media_type="application/manifest+json")


@app.get("/sw.js")
async def service_worker():
    return FileResponse(BASE_DIR / "sw.js", media_type="application/javascript")


@app.get("/icon.svg")
async def icon():
    return FileResponse(BASE_DIR / "icon.svg", media_type="image/svg+xml")


@app.get("/niveis")
async def get_niveis():
    return JSONResponse(levels.load())


@app.post("/niveis")
async def set_niveis(data: dict):
    """Atualiza os niveis do dia sem reiniciar o servidor.
    Ex.: curl -X POST http://127.0.0.1:8003/niveis -H "Content-Type: application/json" -d @niveis.json
    Depois de salvar, avisa os dashboards conectados para recarregar.
    """
    saved = levels.save(data)
    db.log_niveis_sync(saved, "manual")
    await manager.broadcast({"evento": "niveis_atualizados"})
    return saved


@app.post("/niveis/recalcular")
async def recalcular_niveis():
    """Forca o recalculo dos niveis pelo OHLC do pregao anterior.
    Util se voce editou o niveis.json manualmente e quer voltar aos pivots.
    """
    novo = atualizar_niveis_automaticos(force=True)
    if novo is None:
        return JSONResponse(
            {"erro": "sem OHLC de pregao anterior no historico"}, status_code=404)
    await manager.broadcast({"evento": "niveis_atualizados"})
    return novo


@app.get("/historico/{dia}")
async def get_historico(dia: str):
    """Pacote completo de um pregao (dia = AAAA-MM-DD):
    OHLC consolidado, versoes dos niveis, sinais de confluencia/divergencia
    e snapshots de 30s. Base para analise/backtest.
    """
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(None, db.historico_sync, dia)


@app.get("/perfil/{dia}")
async def get_perfil(dia: str, bucket: float = 1.0):
    """Perfil de volume do pregao (dia = AAAA-MM-DD, ?bucket=1.0):
    POC, value area (~70%) e histograma por faixa de preco.
    Base da 'regua por acumulo' — regioes onde o mercado realmente negociou.
    """
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(None, lambda: db.perfil_sync(dia, bucket))


@app.get("/ultimo")
async def get_ultimo():
    return asdict(last_tick) if last_tick else {"status": "aguardando dados"}


@app.get("/fluxo")
async def get_fluxo():
    """Livro (pressao parada) + fita (pressao executada) do ultimo ciclo."""
    return last_fluxo or {"status": "aguardando dados"}


@app.get("/ranking")
async def get_ranking():
    """Acumulado do dia por corretora, compilado da fita (T&T1), mais a
    janela movel (ultimos JANELA_CORRETORAS_S segundos).

    saldo > 0 = corretora comprou mais do que vendeu no agregado do pregao.
    'corretoras' e desde a abertura (tendencia do dia); 'corretoras_janela'
    e so a janela recente (quem esta agredindo AGORA - mais ruidoso).
    ciclos_perdidos > 0 indica trechos de fita que passaram sem serem vistos
    (janela do RTD transbordou entre leituras) - o acumulado e um piso.
    """
    return {"dia": ranking.dia, "corretoras": ranking.ranking(),
            "corretoras_janela": ranking.ranking_janela(),
            "janela_min": JANELA_CORRETORAS_S // 60,
            "ciclos_perdidos": ranking.negocios_perdidos}


async def _contexto_leitura_atual() -> dict:
    """Compila o contexto atual para o agente - usado tanto pelo endpoint
    /leitura (clique manual) quanto pelo gatilho automatico de confluencia."""
    loop = asyncio.get_running_loop()
    dia = date.today().isoformat()
    hist = await loop.run_in_executor(None, db.historico_sync, dia)
    ctx = agente_wdo.montar_contexto(
        last_tick, last_fluxo,
        {"corretoras": ranking.ranking(),
         "corretoras_janela": ranking.ranking_janela(),
         "ciclos_perdidos": ranking.negocios_perdidos},
        levels.load(), hist.get("eventos") or [], hist.get("ohlc"),
        macro=last_macro, casado=last_casado)
    vm = ctx.get("vies_mercado")
    if vm and vm.get("vies") in ("compra", "venda") and (vm.get("forca") or 0) >= 2:
        try:
            historico = await loop.run_in_executor(
                None, db.historico_similar_sync, vm["vies"])
            if historico:
                ctx["historico_similar"] = historico
        except Exception as e:
            print(f"[WDO] Falha ao calcular historico_similar (seguindo): {e}")
    return ctx


@app.get("/leitura")
async def get_leitura(forcar: bool = False):
    """Leitura de fluxo por IA (Google Gemini, nivel gratuito): vies,
    evidencias (numeros do proprio contexto), alertas e ressalvas - nunca
    recomendacao de entrada/saida. Consome os mesmos dados dos outros
    endpoints (fluxo, ranking, niveis, eventos do dia); nao e fonte nova.

    Sob demanda (botao do dashboard), com cache de ~45s para nao gerar
    chamada nova a cada clique repetido. ?forcar=true ignora o cache.
    Tambem disparada sozinha em cada confluencia (ver gerar_leitura_automatica).
    """
    if not agente_wdo.disponivel():
        return JSONResponse(
            {"erro": "GEMINI_API_KEY nao configurada no ambiente do servidor - "
                     "crie uma chave em aistudio.google.com e defina a variavel"},
            status_code=503)
    contexto = await _contexto_leitura_atual()
    loop = asyncio.get_running_loop()
    try:
        leitura = await loop.run_in_executor(
            None, agente_wdo.gerar_leitura, contexto, forcar)
    except RuntimeError as e:
        return JSONResponse({"erro": str(e)}, status_code=502)
    leitura["vies_mercado"] = contexto.get("vies_mercado")
    try:
        await loop.run_in_executor(
            None, db.salvar_leitura_sync, leitura, {"tipo": "manual"},
            last_tick.ultimo if last_tick else None)
    except Exception as e:
        print(f"[WDO] Falha ao persistir leitura manual em `leituras` (seguindo): {e}")
    return leitura


async def gerar_leitura_automatica(gatilho: dict):
    """Dispara a leitura de IA sozinha quando o motor de confluencia
    detecta CONVERGENCIA entre estrutura (nivel) e fluxo (delta_ema) -
    o proprio significado de 'confluencia' no ConfluenceEngine. Nao roda
    em divergencia (rompimento SEM confirmacao de fluxo), so em confluencia
    confirmada.

    Roda em background (nao bloqueia o market_loop - a chamada de rede pode
    levar alguns segundos). Pula silenciosamente se ja houver uma leitura em
    andamento ou a chave do Gemini nao estiver configurada, para nao
    empilhar chamadas nem estourar a cota do nivel gratuito.
    """
    global _leitura_auto_em_andamento
    if _leitura_auto_em_andamento or not agente_wdo.disponivel():
        return
    _leitura_auto_em_andamento = True
    try:
        contexto = await _contexto_leitura_atual()
        loop = asyncio.get_running_loop()
        leitura = await loop.run_in_executor(
            None, agente_wdo.gerar_leitura, contexto, False)   # respeita o cache de 45s
        leitura["gatilho"] = gatilho
        leitura["vies_mercado"] = contexto.get("vies_mercado")
        await manager.broadcast({"evento": "leitura_auto", **leitura})
        try:
            await loop.run_in_executor(
                None, db.salvar_leitura_sync, leitura, gatilho,
                last_tick.ultimo if last_tick else None)
        except Exception as e:
            # Persistencia e secundaria - uma falha aqui nao pode derrubar o
            # gatilho automatico nem virar excecao perdida no asyncio.create_task.
            print(f"[WDO] Falha ao persistir leitura em `leituras` (seguindo): {e}")
        print(f"[WDO] Leitura automatica gerada (gatilho: {gatilho.get('msg')})")
    except RuntimeError as e:
        print(f"[WDO] Leitura automatica falhou: {e}")
    finally:
        _leitura_auto_em_andamento = False


@app.get("/plano_ativacao")
async def get_plano_ativacao():
    """Hora real de ativacao dos gatilhos de compra/venda do plano hoje."""
    return last_plano_ativacao or {"compra": None, "venda": None}


@app.get("/macro")
async def get_macro():
    """Ultimo pacote macro (Brent, DXY, Ouro, Juros EUA, casado) + idade do dado."""
    if last_macro is None:
        return {"status": "aguardando dados"}
    return {**last_macro,
            "idade_s": round(time.time() - macro.last_ok, 1)
            if macro.last_ok else None}


@app.get("/casado")
async def get_casado():
    """Preco justo do WDO por arbitragem vs o dolar a vista (o 'casado').

    diferencial = WDO - pronto (pontos); preco_justo = pronto + carrego;
    desvio = WDO - preco_justo. rotulo = esticado / neutro / atrasado pelo
    z-score do desvio. fonte_carrego diz se o carrego veio do historico
    calibrado ou so da abertura do dia. idade_s alto = spot travado.
    """
    if last_casado is None:
        return {"status": "aguardando dados (WDO + dolar a vista)"}
    return last_casado


@app.post("/operacao")
async def post_operacao(data: dict):
    """Registra uma operação manual do trader (painel REGISTRO DO TRADER)."""
    loop = asyncio.get_running_loop()
    op_id = await loop.run_in_executor(None, lambda: db.salvar_operacao_sync(data))
    return {"ok": True, "id": op_id}


@app.get("/operacoes")
async def get_operacoes(dia: Optional[str] = None):
    """Lista operações registradas (mais recentes primeiro; ?dia=AAAA-MM-DD)."""
    loop = asyncio.get_running_loop()
    ops = await loop.run_in_executor(None, lambda: db.listar_operacoes_sync(dia))
    return {"operacoes": ops}


@app.delete("/operacao/{op_id}")
async def delete_operacao(op_id: int):
    loop = asyncio.get_running_loop()
    n = await loop.run_in_executor(None, lambda: db.remover_operacao_sync(op_id))
    return {"ok": n}


@app.get("/operacoes.csv")
async def get_operacoes_csv():
    """Exporta o histórico de operações em CSV (para estudo/backtest)."""
    loop = asyncio.get_running_loop()
    ops = await loop.run_in_executor(None, lambda: db.listar_operacoes_sync())
    cols = ("id", "ts", "dia", "hora", "lado", "preco", "motivo", "nota",
            "abertura", "maxima", "minima", "volume", "delta", "vwap")
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=cols, extrasaction="ignore")
    w.writeheader()
    for r in ops:
        w.writerow(r)
    body = buf.getvalue().encode("utf-8-sig")
    return Response(content=body, media_type="text/csv; charset=utf-8",
                    headers={"Content-Disposition":
                             'attachment; filename="operacoes_wdo.csv"'})


@app.websocket("/ws")
async def ws_endpoint(ws: WebSocket):
    await manager.connect(ws)
    # envia estado atual imediatamente ao conectar
    if last_tick:
        await ws.send_json(asdict(last_tick))
    if last_macro:
        await ws.send_json(last_macro)
    if last_plano_ativacao:
        await ws.send_json({"evento": "plano_ativacao", **last_plano_ativacao})
    try:
        while True:
            await ws.receive_text()   # mantem a conexao viva
    except WebSocketDisconnect:
        manager.disconnect(ws)


if __name__ == "__main__":
    print("=" * 60)
    print(" MONITOR WDO - http://127.0.0.1:8003")
    print(" Dashboard:   http://127.0.0.1:8003/")
    print(" Niveis:      http://127.0.0.1:8003/niveis")
    print(" Macro:       http://127.0.0.1:8003/macro")
    print(" Casado:      http://127.0.0.1:8003/casado")
    print("=" * 60)
    uvicorn.run(app, host=HOST, port=PORT, log_level="warning")
