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
from fastapi.responses import FileResponse, JSONResponse

# ----------------------------------------------------------------
# CONFIGURACAO
# ----------------------------------------------------------------
BASE_DIR      = Path(__file__).parent
CSV_PATH      = BASE_DIR / "dados_win.csv"       # gerado pelo VBA (ExportarWIN)
BLUE_CHIPS_CSV = BASE_DIR / "dados_blue_chips.csv"  # gerado pelo VBA (ExportarBlueChips)
NIVEIS_PATH   = BASE_DIR / "niveis.json"         # niveis do dia (editavel)
DB_PATH       = BASE_DIR / "win_history.db"
DASHBOARD     = BASE_DIR / "dashboard_win.html"
POLL_INTERVAL = 1.0        # segundos entre leituras do CSV
SNAPSHOT_EVERY = 2         # segundos entre snapshots no SQLite (resolucao
                           # do perfil de volume; ~12k linhas/dia no WIN)
HOST, PORT    = "127.0.0.1", 8001

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
MACRO_RTD_CSV = BASE_DIR / "dados_macro_rtd.csv"  # VBA (ExportarMacroRTD)
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

    def __init__(self, path: Path):
        self.path = path
        self._last_mtime = 0.0

    def read_if_changed(self) -> Optional[list[dict]]:
        if not self.path.exists():
            return None
        mtime = self.path.stat().st_mtime
        if mtime == self._last_mtime:
            return None
        self._last_mtime = mtime
        try:
            with open(self.path, encoding="utf-8-sig", errors="ignore") as f:
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
    p  = (h + l + c) / 3
    bc = (h + l) / 2
    tc = 2 * p - bc
    r1, s1 = 2 * p - l,       2 * p - h
    r2, s2 = p + (h - l),     p - (h - l)
    r3, s3 = h + 2 * (p - l), l - 2 * (h - p)
    return {
        "data_pregao": date.today().isoformat(),
        "contrato": contrato,
        "fonte": f"Pivots automáticos (OHLC {base['dia']})",
        "pivot": _tick5(p),
        "resistencias": [_tick5(r1), _tick5(r2), _tick5(r3)],
        "suportes":     [_tick5(s1), _tick5(s2), _tick5(s3)],
        "alvos_compra": [_tick5(r1), _tick5(r2), _tick5(r3)],
        "alvos_venda":  [_tick5(s1), _tick5(s2), _tick5(s3)],
        "zona_decisiva": {"min": _tick5(min(bc, tc)), "max": _tick5(max(bc, tc))},
        "ladder": {"min": _tick5(s3), "max": _tick5(r3)},
    }


# ----------------------------------------------------------------
# HISTORICO SQLITE (snapshot nao-bloqueante)
# ----------------------------------------------------------------
class SnapshotDB:
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

    def save_sync(self, t: Tick):
        """Chamado via run_in_executor - roda em thread separada."""
        with sqlite3.connect(self.path) as con:
            con.execute(
                "INSERT INTO snapshots (dia,ts,ultimo,abertura,maxima,minima,"
                "volume,agr_compra,agr_venda,vwap) VALUES (?,?,?,?,?,?,?,?,?,?)",
                (date.today().isoformat(), t.timestamp, t.ultimo, t.abertura,
                 t.maxima, t.minima, t.volume, t.agr_compra, t.agr_venda, t.vwap),
            )
            # Consolida o OHLC do dia (ultimo tick do pregao = fechamento).
            # E a base para o calculo automatico dos niveis do dia seguinte.
            con.execute(
                "INSERT INTO daily_ohlc (dia,abertura,maxima,minima,fechamento,vwap) "
                "VALUES (?,?,?,?,?,?) "
                "ON CONFLICT(dia) DO UPDATE SET "
                "abertura=excluded.abertura, maxima=excluded.maxima, "
                "minima=excluded.minima, fechamento=excluded.fechamento, "
                "vwap=excluded.vwap",
                (date.today().isoformat(), t.abertura, t.maxima, t.minima,
                 t.ultimo, t.vwap),
            )

    def ohlc_anterior(self, hoje: str) -> Optional[dict]:
        """OHLC do ultimo pregao ANTES de 'hoje' (pula fim de semana)."""
        with sqlite3.connect(self.path) as con:
            con.row_factory = sqlite3.Row
            row = con.execute(
                "SELECT * FROM daily_ohlc WHERE dia < ? "
                "AND maxima IS NOT NULL AND minima IS NOT NULL "
                "AND fechamento IS NOT NULL "
                "ORDER BY dia DESC LIMIT 1", (hoje,)
            ).fetchone()
        return dict(row) if row else None

    def corrigir_fechamento(self, dia: str, fechamento: float):
        """Substitui o fechamento gravado pelo oficial (RTD FEC)."""
        with sqlite3.connect(self.path) as con:
            con.execute("UPDATE daily_ohlc SET fechamento=? WHERE dia=?",
                        (fechamento, dia))

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
blue_chips_reader = BlueChipsReader(BLUE_CHIPS_CSV)
levels   = LevelStore(NIVEIS_PATH)
db       = SnapshotDB(DB_PATH)
manager  = ConnectionManager()
engine   = ConfluenceEngine(levels)
macro    = MacroFetcher()
rtd_macro = RtdMacroReader(MACRO_RTD_CSV)
last_tick: Optional[Tick] = None
last_blue_chips: Optional[dict] = None
last_macro: Optional[dict] = None
last_plano_ativacao: Optional[dict] = None


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
    db.corrigir_fechamento(base["dia"], fec)
    base = dict(base, fechamento=fec)
    novo = calcular_niveis_do_dia(base, cfg.get("contrato", "WIN"))
    novo["fonte"] += " · C=FEC oficial"
    levels.save(novo)
    db.log_niveis_sync(novo, "auto_fec")
    return novo


async def market_loop():
    """Loop principal: le CSV -> confluencia -> broadcast -> snapshot."""
    global last_tick, last_blue_chips, last_plano_ativacao
    last_snapshot = 0.0
    loop = asyncio.get_running_loop()
    dia_atual = date.today()
    fec_conferido = False
    novo = atualizar_niveis_automaticos()
    if novo:
        print(f"[WIN] Niveis do dia recalculados: {novo['fonte']}")
    print(f"[WIN] Loop iniciado. Observando {CSV_PATH.name} a cada {POLL_INTERVAL}s")
    while True:
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
