"""
================================================================
 MONITOR WIN - server_win.py
 Sistema independente de monitoramento do Mini Indice (WINFUT)
----------------------------------------------------------------
 Pipeline : Profit Pro RTD -> Excel (VBA) -> dados_win.csv
            -> este servidor (FastAPI) -> WebSocket -> dashboard
 Extra    : macro (S&P 500 ES=F via Yahoo; DI futuro + DOLFUT via
            RTD/Excel em dados_macro_rtd.csv, dolar Yahoo como fallback)
            -> card MACRO do dashboard
 Porta    : 8001 (roda em paralelo com o server do WDO na 8000)
 Executar : python server_win.py
================================================================
Boas praticas aplicadas (diferencas vs. server_v2.py do WDO):
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
import time
from collections import deque
from datetime import date
from contextlib import asynccontextmanager
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Optional

import httpx
import uvicorn
from fastapi import FastAPI, WebSocket, WebSocketDisconnect
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, Response

import agente_win     # leitura de fluxo por IA (Google Gemini) - consumidor, nao fonte

# ----------------------------------------------------------------
# CONFIGURACAO
# ----------------------------------------------------------------
BASE_DIR      = Path(__file__).parent
# 2026-07-22: ponte encerrada. O VBA (Modulo1 + ModuloBlueChips) foi repontuado
# de Apps\monitor_win para Day trade\monitor_win, entao os CSVs de entrada voltam
# a ser lidos daqui (BASE_DIR), projeto todo numa pasta so.
DATA_DIR      = BASE_DIR
CSV_PATH      = DATA_DIR / "dados_win.csv"       # gerado pelo VBA (ExportarWIN)
BLUE_CHIPS_CSV = DATA_DIR / "dados_blue_chips.csv"  # gerado pelo VBA (ExportarBlueChips)
BOOK_CSV      = DATA_DIR / "dados_book.csv"      # livro de ofertas (BOOK0)
TT_CSV        = DATA_DIR / "dados_tt.csv"        # fita / Times & Trades (T&T0)
VAP_CSV       = DATA_DIR / "dados_vap.csv"       # volume por preco (VAP0)
RANKING_CSV   = DATA_DIR / "ranking_acum.csv"    # acumulado p/ a aba Acum (VBA le)
NIVEIS_PATH   = BASE_DIR / "niveis.json"         # niveis do dia (editavel)
DB_PATH       = BASE_DIR / "win_history.db"
DASHBOARD     = BASE_DIR / "dashboard_win.html"
POLL_INTERVAL = 1.0        # segundos entre leituras do CSV
SNAPSHOT_EVERY = 2         # segundos entre snapshots no SQLite (resolucao
                           # do perfil de volume; ~12k linhas/dia no WIN)
HOST, PORT    = "127.0.0.1", 8001

# Cobertura minima para um pregao virar base de pivots: o server tem que ter
# pego a abertura e o fechamento. Fora disso, H/L sao parciais e os pivots do
# dia seguinte saem errados.
PREGAO_ABERTURA_ATE = "09:15"    # primeiro tick tem que vir antes disso
PREGAO_FECHA_APOS   = "17:45"    # ultimo tick tem que vir depois disso

# Peso aproximado de cada blue chip no IBOV (atualizar periodicamente -
# nao ha fonte RTD para isso, e um dado de composicao do indice).
PESO_IBOV = {
    "VALE3": 11.2, "PETR4": 7.8, "ITUB4": 6.5, "BBDC4": 3.1, "BBAS3": 2.4,
}

# --- Macro (S&P 500, Dolar, DI) ---------------------------------------
# S&P 500 e Dolar: Yahoo Finance. O S&P e o driver global de risco que
# mais move o IBOV (correlacao positiva); usamos o E-mini futuro ES=F
# porque negocia quase 24h — o indice a vista (^GSPC) fica parado antes
# da abertura de NY. DI futuro: nao existe no Yahoo; vem do Profit via
# RTD/Excel em dados_macro_rtd.csv (ver ExportarWIN.bas).
YAHOO_CHART   = ("https://query1.finance.yahoo.com/v8/finance/chart/{sym}"
                 "?range=1d&interval=15m")
MACRO_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
MACRO_POLL    = 30                            # segundos entre consultas
# DXY = dolar global (indice ICE contra cesta de moedas). Correlacao
# diaria com o IBOV (-0,36 em 1 ano) mais forte que a do USD/BRL e quase
# independente dele (r~0,18) — carrega o "fluxo p/ emergentes" que o
# cambio BRL sozinho nao mostra. DXY sobe -> IBOV tende a cair.
MACRO_SYMBOLS = {"sp500": "ES=F", "dxy": "DX-Y.NYB", "dolar": "USDBRL=X"}
MACRO_RTD_CSV = DATA_DIR / "dados_macro_rtd.csv"  # VBA (ExportarMacroRTD) - ver PONTE acima
RTD_MAX_AGE   = 180                           # s sem update = desatualizado


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
    """

    def __init__(self):
        self.dia = date.today().isoformat()
        self._dados: dict = {}               # corretora -> metricas
        self.negocios_perdidos = 0           # ciclos com janela_perdida

    def _slot(self, nome: str) -> dict:
        return self._dados.setdefault(nome, {
            "qtd_compra": 0.0, "qtd_venda": 0.0,
            "fin_compra": 0.0, "fin_venda": 0.0, "negocios": 0})

    def processar(self, novos: list, janela_perdida: bool = False):
        hoje = date.today().isoformat()
        if hoje != self.dia:                 # virada de pregao zera o dia
            self.dia, self._dados = hoje, {}
            self.negocios_perdidos = 0
        if janela_perdida:
            self.negocios_perdidos += 1
        for n in novos:
            fin = n["qtd"] * n["preco"]
            c = self._slot(n["comprador"])
            c["qtd_compra"] += n["qtd"]
            c["fin_compra"] += fin
            c["negocios"] += 1
            v = self._slot(n["vendedor"])
            v["qtd_venda"] += n["qtd"]
            v["fin_venda"] += fin
            v["negocios"] += 1

    def ranking(self) -> list:
        """Corretoras por saldo (compra - venda), maior comprador primeiro."""
        total = sum(d["qtd_compra"] + d["qtd_venda"]
                    for d in self._dados.values()) or 1.0
        out = []
        for nome, d in self._dados.items():
            qtd_total = d["qtd_compra"] + d["qtd_venda"]
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

    def snapshot_db(self) -> list:
        """Linhas para o upsert no SQLite (valores absolutos do dia)."""
        return [(self.dia, nome, d["qtd_compra"], d["qtd_venda"],
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
        tmp.replace(path)                    # troca atomica (VBA nunca le pela metade)

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
# LEITOR DO CSV DE BLUE CHIPS (fluxo das acoes que compoem o IBOV)
# ----------------------------------------------------------------
class BlueChipsReader:
    """Le dados_blue_chips.csv (gerado por ExportarBlueChips.bas) e
    classifica o fluxo de cada ativo pela dominancia de agressao.

    Formato esperado (separador ';', 1 linha por ativo):
    ticker;ultimo;abertura;maxima;minima;fec_ant;agr_compra;agr_venda;vwap;volume;timestamp
    """

    FIELDS = ["ticker", "ultimo", "abertura", "maxima", "minima", "fec_ant",
              "agr_compra", "agr_venda", "vwap", "volume", "timestamp"]
    DOMINANCIA_MIN = 5.0   # % de dominancia abaixo disso = "neutro" (ruido)

    def __init__(self, path: Path, alt_paths: tuple = ()):
        # Le do arquivo MAIS RECENTE entre os candidatos. Blindagem 22/07: o
        # macro ModuloBlueChips pode estar gravando na pasta antiga (Apps) ou na
        # nova (Day trade) dependendo de qual loop OnTime esta vivo; seguir o
        # mais fresco faz o painel funcionar nos dois casos e se auto-corrige
        # sozinho quando o Excel reiniciar com o caminho certo compilado.
        self.paths = [p for p in (path, *alt_paths) if p]
        self._last_mtime = 0.0

    def _fresh_path(self) -> Optional[Path]:
        cand = [(p.stat().st_mtime, p) for p in self.paths if p.exists()]
        return max(cand, key=lambda t: t[0])[1] if cand else None

    def read_if_changed(self) -> Optional[list[dict]]:
        path = self._fresh_path()
        if path is None:
            return None
        mtime = path.stat().st_mtime
        if mtime == self._last_mtime:
            return None
        self._last_mtime = mtime
        try:
            with open(path, encoding="utf-8-sig", errors="ignore") as f:
                rows = [r for r in csv.reader(f, delimiter=";") if r]
        except PermissionError:
            return None                      # Excel escrevendo no arquivo
        out = []
        for row in rows[1:]:                 # pula cabecalho
            if len(row) < len(self.FIELDS) or not row[0]:
                continue
            d = dict(zip(self.FIELDS, row))
            ultimo = _to_float(d["ultimo"])
            fec = _to_float(d["fec_ant"])
            ac = _to_float(d["agr_compra"]) or 0.0
            av = _to_float(d["agr_venda"]) or 0.0
            delta = ac - av
            dominancia = abs(delta) / (ac + av) * 100 if (ac + av) else 0.0
            fluxo = "neutro"
            if dominancia > self.DOMINANCIA_MIN:
                fluxo = "compra" if delta > 0 else "venda"
            var_pct = ((ultimo - fec) / fec * 100) if (ultimo and fec) else None
            out.append({
                "ticker": d["ticker"], "ultimo": ultimo, "var_pct": var_pct,
                "peso_ibov": PESO_IBOV.get(d["ticker"]), "fluxo": fluxo,
                "dominancia": round(dominancia, 1),
            })
        return out or None


# ----------------------------------------------------------------
# MACRO: S&P 500 + DOLAR (Yahoo) e DI FUTURO (RTD via CSV)
# ----------------------------------------------------------------
class MacroFetcher:
    """Busca S&P 500 e Dolar no Yahoo Finance.

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


class RtdMacroReader:
    """Le dados_macro_rtd.csv exportado pelo VBA (DI futuros + DOLFUT).

    Formato esperado (separador ';', 1 linha por ativo):
    ticker;ultimo;fec_ant;volume;timestamp

    DI: o VBA exporta TODOS os DI1* da planilha; aqui vence o de maior
    volume (contrato mais liquido = referencia do juro futuro; a escolha
    acompanha a rolagem sem mexer em codigo). Taxa em % a.a.;
    variacao em bps = (ultimo - fec_ant) x 100.

    DOLFUT: cotado em pontos B3 (R$ por US$1000) -> /1000 = R$/US$.
    Tempo real do pregao, preferido ao spot do Yahoo (fallback).
    """

    def __init__(self, path: Path):
        self.path = path

    def read(self) -> dict:
        if not self.path.exists():
            return {}
        try:
            idade = time.time() - self.path.stat().st_mtime
            with open(self.path, encoding="utf-8-sig", errors="ignore") as f:
                rows = [r for r in csv.reader(f, delimiter=";") if r]
        except PermissionError:
            return {}                        # Excel escrevendo no arquivo
        stale = idade > RTD_MAX_AGE
        out: dict = {}
        best_di = None
        for row in rows[1:]:                 # pula cabecalho
            if len(row) < 4:
                continue
            tk = row[0].strip()
            ultimo = _to_float(row[1])
            fec    = _to_float(row[2])
            vol    = _to_float(row[3]) or 0.0
            ts     = row[4] if len(row) > 4 else ""
            if ultimo is None:
                continue
            if tk.startswith("DI1"):
                if best_di is None or vol > best_di["_vol"]:
                    best_di = {
                        "ticker": tk, "taxa": ultimo, "fec_ant": fec,
                        "var_bps": round((ultimo - fec) * 100, 1) if fec else None,
                        "ts": ts, "desatualizado": stale, "_vol": vol,
                    }
            elif tk == "DOLFUT" and fec:
                out["dolar"] = {
                    "preco": round(ultimo / 1000, 4),
                    "fech_ant": round(fec / 1000, 4),
                    "var_pct": round((ultimo / fec - 1) * 100, 2),
                    "ts": ts, "fonte": "DOLFUT · Profit RTD",
                    "desatualizado": stale,
                }
        if best_di:
            best_di.pop("_vol")
            out["di"] = best_di
        return out


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
def _tick5(v: float) -> int:
    """Arredonda para o multiplo de 5 mais proximo (tick do WIN)."""
    return int(round(v / 5.0) * 5)


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
        "contrato": contrato,
        "fonte": f"Pivots automáticos (OHLC {base['dia']}){aviso}",
        "pivot": _tick5(p),
        "resistencias": [_tick5(r1), _tick5(r2), _tick5(r3)],
        "suportes":     [_tick5(s1), _tick5(s2), _tick5(s3)],
        "alvos_compra": [_tick5(r1), _tick5(r2), _tick5(r3)],
        "alvos_venda":  [_tick5(s1), _tick5(s2), _tick5(s3)],
        "zona_decisiva": {"min": _tick5(min(bc, tc)), "max": _tick5(max(bc, tc))},
        "ladder": {"min": _tick5(s3), "max": _tick5(r3)},
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
        print(f"[WIN] ALERTA: vwap={t.vwap:,.0f} fora do range do dia "
              f"({t.minima:,.0f}-{t.maxima:,.0f}) - descartado. Conferir a "
              f"coluna do campo RTD 67 no ExportarWIN.bas")
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
            con.execute("""
                CREATE TABLE IF NOT EXISTS macro_snapshots (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    dia TEXT, ts TEXT,
                    sp500 REAL, sp500_var REAL,
                    dxy REAL, dxy_var REAL,
                    dolar REAL, dolar_var REAL,
                    di REAL, di_var_bps REAL
                )""")
            # migracao: tabelas criadas com Brent (17/07 cedo) -> S&P 500
            for old, new in (("brent", "sp500"), ("brent_var", "sp500_var")):
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
            # snapshots antigos nao tinham a coluna de data
            try:
                con.execute("ALTER TABLE snapshots ADD COLUMN dia TEXT")
            except sqlite3.OperationalError:
                pass                     # coluna ja existe
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

    def ohlc_anterior(self, hoje: str) -> Optional[dict]:
        """OHLC do ultimo pregao ANTES de 'hoje' (pula fim de semana).

        Prefere pregoes cobertos de ponta a ponta (valido=1), porque H/L
        truncados geram pivots errados. Se nao houver nenhum, cai no dia mais
        recente que ao menos seja coerente (C dentro de [L,H]) e marca
        'cobertura_parcial' - o dashboard avisa em vez de ficar sem niveis.
        """
        base = ("SELECT * FROM daily_ohlc WHERE dia < ? "
                "AND maxima IS NOT NULL AND minima IS NOT NULL "
                "AND fechamento IS NOT NULL "
                "AND fechamento BETWEEN minima AND maxima ")
        with sqlite3.connect(self.path) as con:
            con.row_factory = sqlite3.Row
            row = con.execute(
                base + "AND valido = 1 ORDER BY dia DESC LIMIT 1", (hoje,)
            ).fetchone()
            if row:
                return dict(row, cobertura_parcial=False)
            row = con.execute(
                base + "ORDER BY dia DESC LIMIT 1", (hoje,)).fetchone()
        return dict(row, cobertura_parcial=True) if row else None

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
                    print(f"[WIN] {dia}: FEC oficial {fechamento:,.0f} fora do "
                          f"range gravado ({row[1]:,.0f}-{row[0]:,.0f}) - "
                          f"cobertura parcial, dia marcado como invalido")
                    return False
            con.execute("UPDATE daily_ohlc SET fechamento=? WHERE dia=?",
                        (fechamento, dia))
        return True

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

    def save_macro_sync(self, sp500: Optional[dict], dxy: Optional[dict],
                        dolar: Optional[dict], di: Optional[dict]):
        """Historico macro para analise de correlacao com o WIN."""
        g = lambda d, k: d.get(k) if d else None
        with sqlite3.connect(self.path) as con:
            con.execute(
                "INSERT INTO macro_snapshots (dia,ts,sp500,sp500_var,"
                "dxy,dxy_var,dolar,dolar_var,di,di_var_bps) "
                "VALUES (?,?,?,?,?,?,?,?,?,?)",
                (date.today().isoformat(), time.strftime("%H:%M:%S"),
                 g(sp500, "preco"), g(sp500, "var_pct"),
                 g(dxy, "preco"), g(dxy, "var_pct"),
                 g(dolar, "preco"), g(dolar, "var_pct"),
                 g(di, "taxa"), g(di, "var_bps")))

    def log_niveis_sync(self, dados: dict, origem: str):
        """Registra cada versao dos niveis do dia (auto, refinado ou manual)."""
        with sqlite3.connect(self.path) as con:
            con.execute(
                "INSERT INTO niveis_hist (dia,ts,origem,dados) VALUES (?,?,?,?)",
                (date.today().isoformat(), time.strftime("%H:%M:%S"),
                 origem, json.dumps(dados, ensure_ascii=False)))

    def perfil_sync(self, dia: str, bucket: int = 50) -> dict:
        """Perfil de volume do pregao: distribui o volume negociado entre
        snapshots consecutivos no bucket de preco em que ocorreu.

        Aproximacao: todo o volume do intervalo e atribuido ao preco do
        snapshot que o fechou — quanto menor SNAPSHOT_EVERY, mais fiel.

        Retorna POC (preco de maior acumulo), value area (~70% do volume,
        expandida a partir do POC) e o histograma completo por bucket.
        """
        with sqlite3.connect(self.path) as con:
            rows = con.execute(
                "SELECT ultimo, volume FROM snapshots WHERE dia=? "
                "AND ultimo IS NOT NULL AND volume IS NOT NULL "
                "ORDER BY id", (dia,)).fetchall()
            cobertura = con.execute(
                "SELECT MIN(ts), MAX(ts) FROM snapshots WHERE dia=?",
                (dia,)).fetchone()
        hist: dict[int, float] = {}
        prev_vol = None
        for ultimo, volume in rows:
            if prev_vol is not None:
                dv = volume - prev_vol
                if dv > 0:                    # ignora reset/duplicata
                    b = int(round(ultimo / bucket) * bucket)
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
    se recalibra ao regime de volume do dia (o delta do WIN opera em
    escala ~1000x maior que o do WDO).

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

    def check(self, tick: Tick) -> list[dict]:
        """Retorna lista de eventos de confluencia detectados neste tick."""
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
            events += self._track(f"up:{lvl}", price > lvl, "up", lvl, now)
        for lvl in sorted(downs, reverse=True):
            events += self._track(f"dn:{lvl}", price < lvl, "down", lvl, now)
        return events

    def _track(self, key: str, beyond: bool, direction: str,
               lvl: float, now: float) -> list[dict]:
        if not beyond:
            self._persist[key] = 0
            return []
        self._persist[key] = self._persist.get(key, 0) + 1
        if self._persist[key] != self.PERSIST_TICKS:
            return []                                  # ainda sem persistencia
        if now - self._last_fire.get(key, 0) < self.COOLDOWN_S:
            return []                                  # em cooldown
        if not self._flow_confirms(direction):
            # rompeu MAS o fluxo nao confirma -> alerta de divergencia
            self._last_fire[key] = now
            return [{
                "evento": "divergencia", "direcao": direction, "nivel": lvl,
                "delta_ema": round(self.flow_ema or 0, 1),
                "msg": f"Rompeu {lvl:.0f} SEM confirmacao de fluxo (possivel violino)"
            }]
        self._last_fire[key] = now
        return [{
            "evento": "confluencia", "direcao": direction, "nivel": lvl,
            "delta_ema": round(self.flow_ema, 1),
            "msg": f"CONFLUENCIA: {'rompimento' if direction=='up' else 'perda'} "
                   f"de {lvl:.0f} confirmado pelo fluxo (EMA {self.flow_ema:+.0f})"
        }]


# ----------------------------------------------------------------
# APP + LIFESPAN
# ----------------------------------------------------------------
reader   = CsvReader(CSV_PATH)
blue_chips_reader = BlueChipsReader(
    BLUE_CHIPS_CSV,
    alt_paths=(Path(r"C:\Users\rodri\OneDrive\Apps\monitor_win\dados_blue_chips.csv"),))
levels   = LevelStore(NIVEIS_PATH)
db       = SnapshotDB(DB_PATH)
manager  = ConnectionManager()
engine   = ConfluenceEngine(levels)
macro    = MacroFetcher()
rtd_macro = RtdMacroReader(MACRO_RTD_CSV)
fluxo_reader = FluxoReader(BOOK_CSV, TT_CSV, VAP_CSV)
ranking = RankingCorretoras()
ranking.carregar(db.carregar_corretoras_sync(ranking.dia))   # sobrevive a restart
last_tick: Optional[Tick] = None
last_blue_chips: Optional[dict] = None
last_macro: Optional[dict] = None
last_plano_ativacao: Optional[dict] = None
last_fluxo: Optional[dict] = None
_leitura_auto_em_andamento = False   # evita empilhar chamadas ao Gemini


def atualizar_niveis_automaticos(force: bool = False) -> Optional[dict]:
    """Recalcula os niveis se o niveis.json ainda for de outro pregao.

    Nao sobrescreve niveis definidos manualmente HOJE (data_pregao == hoje),
    a menos que force=True. Sem OHLC anterior no historico, mantem o manual.
    """
    cfg = levels.load()
    hoje = date.today().isoformat()
    if not force and cfg.get("data_pregao") == hoje:
        return None
    base = db.ohlc_anterior(hoje)
    if base is None:
        return None
    novo = calcular_niveis_do_dia(base, cfg.get("contrato", "WIN"))
    levels.save(novo)
    db.log_niveis_sync(novo, "auto")
    return novo


def refinar_niveis_com_fec(fec: float) -> Optional[dict]:
    """Troca o fechamento gravado pelo FEC oficial do RTD e recalcula.

    So age se os niveis de hoje forem automaticos e o C usado divergir
    do oficial em 1 tick ou mais (protege contra server desligado antes
    do fechamento no dia anterior).
    """
    cfg = levels.load()
    hoje = date.today().isoformat()
    if cfg.get("data_pregao") != hoje:
        return None
    if not str(cfg.get("fonte", "")).startswith("Pivots autom"):
        return None                          # niveis manuais: nao mexe
    base = db.ohlc_anterior(hoje)
    if base is None or abs(base["fechamento"] - fec) < 5:
        return None                          # ja esta correto (< 1 tick)
    if not db.corrigir_fechamento(base["dia"], fec):
        # Cobertura parcial: aquele dia acabou de ser marcado invalido.
        # Refaz os niveis a partir do ultimo pregao realmente confiavel.
        return atualizar_niveis_automaticos(force=True)
    base = dict(base, fechamento=fec)
    novo = calcular_niveis_do_dia(base, cfg.get("contrato", "WIN"))
    novo["fonte"] += " · C=FEC oficial"
    levels.save(novo)
    db.log_niveis_sync(novo, "auto_fec")
    return novo


async def market_loop():
    """Loop principal: le CSV -> confluencia -> broadcast -> snapshot."""
    global last_tick, last_blue_chips, last_plano_ativacao, last_fluxo
    last_snapshot = 0.0
    last_ranking_bcast = 0.0
    loop = asyncio.get_running_loop()
    dia_atual = date.today()
    fec_conferido = False
    novo = atualizar_niveis_automaticos()
    if novo:
        print(f"[WIN] Niveis do dia recalculados: {novo['fonte']}")
    print(f"[WIN] Loop iniciado. Observando {CSV_PATH.name} a cada {POLL_INTERVAL}s")
    while True:
        try:
            # Virada de dia com o server no ar: recalcula e avisa os dashboards
            if date.today() != dia_atual:
                dia_atual = date.today()
                fec_conferido = False
                engine.reset_flow()      # acumulados do RTD zeram no novo pregao
                novo = atualizar_niveis_automaticos()
                if novo:
                    print(f"[WIN] Niveis do dia recalculados: {novo['fonte']}")
                    await manager.broadcast({"evento": "niveis_atualizados"})
            tick = reader.read_if_changed()
            if tick:
                # Primeiro tick com FEC do dia: confere o C usado nos pivots
                if not fec_conferido and tick.fec_ant:
                    fec_conferido = True
                    novo = refinar_niveis_com_fec(tick.fec_ant)
                    if novo:
                        print(f"[WIN] Pivots refinados com FEC oficial "
                              f"({tick.fec_ant:.0f}): {novo['fonte']}")
                        await manager.broadcast({"evento": "niveis_atualizados"})
                last_tick = tick
                await manager.broadcast(asdict(tick))
                # Motor de confluencia: mapa (niveis) x fluxo (delta)
                for ev in engine.check(tick):
                    print(f"[WIN] {ev['msg']}")
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
                         "ciclos_perdidos": ranking.negocios_perdidos})
            bc = blue_chips_reader.read_if_changed()
            if bc:
                positivas = sum(1 for a in bc if a["fluxo"] == "compra")
                negativas = sum(1 for a in bc if a["fluxo"] == "venda")
                vies = ("compra" if positivas > negativas else
                        "venda" if negativas > positivas else "neutro")
                payload = {"evento": "blue_chips", "ativos": bc, "vies": vies,
                           "positivas": positivas, "negativas": negativas}
                last_blue_chips = payload
                await manager.broadcast(payload)
        except Exception as e:
            # Blindagem: um erro num passo secundario (persistencia, leitura de
            # fluxo, etc.) NUNCA pode derrubar o loop e congelar o dashboard.
            # Loga e segue para o proximo ciclo. (Bug real 22/07: schema drift
            # da tabela fluxo matava o loop silenciosamente.)
            import traceback
            print(f"[WIN] ERRO no ciclo do market_loop (seguindo): "
                  f"{type(e).__name__}: {e}")
            traceback.print_exc()
        await asyncio.sleep(POLL_INTERVAL)


async def macro_loop():
    """Loop paralelo: S&P 500 + Dolar no Yahoo, DI no CSV do RTD.

    Transmite o pacote consolidado via WS a cada MACRO_POLL segundos e
    persiste no SQLite para estudo de correlacao com o WIN.
    """
    global last_macro
    loop = asyncio.get_running_loop()
    print(f"[WIN] Macro loop iniciado ({', '.join(MACRO_SYMBOLS.values())} "
          f"+ DI/DOLFUT via {MACRO_RTD_CSV.name}) a cada {MACRO_POLL}s")
    async with httpx.AsyncClient() as client:
        while True:
            quotes = await macro.fetch_all(client)
            rtd = rtd_macro.read()
            di = rtd.get("di")
            # Dolar: DOLFUT em tempo real tem prioridade; Yahoo (spot,
            # delay) e o fallback quando o RTD falta ou esta parado.
            dolar_rtd = rtd.get("dolar")
            dolar = (dolar_rtd if dolar_rtd and not dolar_rtd["desatualizado"]
                     else quotes.get("dolar"))
            if quotes or di or dolar:
                last_macro = {
                    "evento": "macro",
                    "sp500": quotes.get("sp500"),
                    "dxy": quotes.get("dxy"),
                    "dolar": dolar,
                    "di": di,
                    "ts": time.strftime("%H:%M:%S"),
                }
                await manager.broadcast(last_macro)
                await loop.run_in_executor(
                    None, db.save_macro_sync,
                    quotes.get("sp500"), quotes.get("dxy"), dolar, di)
            await asyncio.sleep(MACRO_POLL)


@asynccontextmanager
async def lifespan(app: FastAPI):
    t1 = asyncio.create_task(market_loop())
    t2 = asyncio.create_task(macro_loop())
    yield
    t1.cancel()
    t2.cancel()


app = FastAPI(title="Monitor WIN", lifespan=lifespan)
app.add_middleware(
    CORSMiddleware, allow_origins=["*"],
    allow_methods=["*"], allow_headers=["*"],
)


# ----------------------------------------------------------------
# ROTAS
# ----------------------------------------------------------------
@app.get("/")
async def dashboard():
    return FileResponse(DASHBOARD)


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
    Ex.: curl -X POST http://127.0.0.1:8001/niveis -H "Content-Type: application/json" -d @niveis.json
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
async def get_perfil(dia: str, bucket: int = 50):
    """Perfil de volume do pregao (dia = AAAA-MM-DD, ?bucket=50):
    POC, value area (~70%) e histograma por faixa de preco.
    Base da 'regua por acumulo' — regioes onde o mercado realmente negociou.
    """
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(None, lambda: db.perfil_sync(dia, bucket))


@app.get("/ultimo")
async def get_ultimo():
    return asdict(last_tick) if last_tick else {"status": "aguardando dados"}


@app.get("/blue_chips")
async def get_blue_chips():
    return last_blue_chips or {"status": "aguardando dados"}


@app.get("/fluxo")
async def get_fluxo():
    """Livro (pressao parada) + fita (pressao executada) do ultimo ciclo."""
    return last_fluxo or {"status": "aguardando dados"}


@app.get("/ranking")
async def get_ranking():
    """Acumulado do dia por corretora, compilado da fita (T&T0).

    saldo > 0 = corretora comprou mais do que vendeu no agregado do pregao.
    ciclos_perdidos > 0 indica trechos de fita que passaram sem serem vistos
    (janela do RTD transbordou entre leituras) - o acumulado e um piso.
    """
    return {"dia": ranking.dia, "corretoras": ranking.ranking(),
            "ciclos_perdidos": ranking.negocios_perdidos}


async def _contexto_leitura_atual() -> dict:
    """Compila o contexto atual para o agente - usado tanto pelo endpoint
    /leitura (clique manual) quanto pelo gatilho automatico de confluencia."""
    loop = asyncio.get_running_loop()
    dia = date.today().isoformat()
    hist = await loop.run_in_executor(None, db.historico_sync, dia)
    return agente_win.montar_contexto(
        last_tick, last_fluxo,
        {"corretoras": ranking.ranking(),
         "ciclos_perdidos": ranking.negocios_perdidos},
        levels.load(), hist.get("eventos") or [], hist.get("ohlc"))


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
    if not agente_win.disponivel():
        return JSONResponse(
            {"erro": "GEMINI_API_KEY nao configurada no ambiente do servidor - "
                     "crie uma chave em aistudio.google.com e defina a variavel"},
            status_code=503)
    contexto = await _contexto_leitura_atual()
    loop = asyncio.get_running_loop()
    try:
        leitura = await loop.run_in_executor(
            None, agente_win.gerar_leitura, contexto, forcar)
    except RuntimeError as e:
        return JSONResponse({"erro": str(e)}, status_code=502)
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
    if _leitura_auto_em_andamento or not agente_win.disponivel():
        return
    _leitura_auto_em_andamento = True
    try:
        contexto = await _contexto_leitura_atual()
        loop = asyncio.get_running_loop()
        leitura = await loop.run_in_executor(
            None, agente_win.gerar_leitura, contexto, False)   # respeita o cache de 45s
        leitura["gatilho"] = gatilho
        await manager.broadcast({"evento": "leitura_auto", **leitura})
        print(f"[WIN] Leitura automatica gerada (gatilho: {gatilho.get('msg')})")
    except RuntimeError as e:
        print(f"[WIN] Leitura automatica falhou: {e}")
    finally:
        _leitura_auto_em_andamento = False


@app.get("/plano_ativacao")
async def get_plano_ativacao():
    """Hora real de ativacao dos gatilhos de compra/venda do plano hoje."""
    return last_plano_ativacao or {"compra": None, "venda": None}


@app.get("/macro")
async def get_macro():
    """Ultimo pacote macro (S&P 500, Dolar, DI) + idade do dado em segundos."""
    if last_macro is None:
        return {"status": "aguardando dados"}
    return {**last_macro,
            "idade_s": round(time.time() - macro.last_ok, 1)
            if macro.last_ok else None}


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
                             'attachment; filename="operacoes_win.csv"'})


@app.websocket("/ws")
async def ws_endpoint(ws: WebSocket):
    await manager.connect(ws)
    # envia estado atual imediatamente ao conectar
    if last_tick:
        await ws.send_json(asdict(last_tick))
    if last_blue_chips:
        await ws.send_json(last_blue_chips)
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
    print(" MONITOR WIN - http://127.0.0.1:8001")
    print(" Dashboard:   http://127.0.0.1:8001/")
    print(" Niveis:      http://127.0.0.1:8001/niveis")
    print(" Macro:       http://127.0.0.1:8001/macro")
    print("=" * 60)
    uvicorn.run(app, host=HOST, port=PORT, log_level="warning")
