# -*- coding: utf-8 -*-
"""
Sync incremental do win_history.db (SQLite local) -> Azure SQL Database.

- Tabelas append-only (fluxo, snapshots, macro_snapshots, eventos, niveis_hist,
  operacoes): empurra so as linhas novas (id > watermark salvo em sync_state.json),
  via MERGE ... WHEN NOT MATCHED THEN INSERT (equivalente ao ON CONFLICT DO NOTHING).
- Tabelas mutaveis pequenas (corretoras, daily_ohlc): reenviadas por inteiro com
  MERGE ... WHEN MATCHED UPDATE / WHEN NOT MATCHED INSERT (upsert).

Conexao: variavel de ambiente AZURE_SQL_URL, no formato de connection string ODBC:
  Driver={ODBC Driver 18 for SQL Server};Server=tcp:<servidor>.database.windows.net,1433;
  Database=<banco>;Uid=<usuario>;Pwd=<senha>;Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;
  (Azure Portal -> seu banco -> Connection strings -> aba ODBC)
Alternativa: arquivo config.local na mesma pasta com a connection string numa unica linha.

Requer:  py -m pip install pyodbc
         + ODBC Driver 18 for SQL Server instalado no Windows.

Rodar:  py sync_azure_sql.py
"""
import os
import sys
import json
import sqlite3
from pathlib import Path

try:
    import pyodbc
except ImportError:
    sys.exit("pyodbc nao instalado. Rode:  py -m pip install pyodbc")

AQUI = Path(__file__).resolve().parent
DB_PATH = Path(os.environ.get("WIN_DB_PATH", AQUI.parent / "win_history.db"))
STATE_PATH = AQUI / "sync_state.json"

# colunas nvarchar(max): sem isso o fast_executemany dimensiona o buffer
# pela 1a linha do lote e trunca silenciosamente as linhas maiores.
LONGTEXT_COLS = {"msg", "dados", "motivo", "nota"}


def setinputsizes_for(cursor, cols):
    if any(c in LONGTEXT_COLS for c in cols):
        cursor.setinputsizes([(pyodbc.SQL_WVARCHAR, 0, 0) if c in LONGTEXT_COLS else None
                               for c in cols])
CONFIG_LOCAL = AQUI / "config.local"
BATCH = 1000

# tabelas append-only keyed por id -> (colunas)
APPEND_ONLY = {
    "fluxo": ["id", "dia", "ts", "bid", "ask", "spread", "qtd_bid", "qtd_ask",
              "desequilibrio", "profundidade_bid", "profundidade_ask",
              "lote_medio", "lote_max", "agr_compra_fita",
              "agr_venda_fita", "pressao_fita", "poc", "vah",
              "val", "vap_total", "vol_acima_pct", "dist_poc", "amostra_negocios",
              "amostra_contratos", "janela_s", "negocios_s", "contratos_s"],
    "snapshots": ["id", "ts", "ultimo", "abertura", "maxima", "minima", "volume",
                  "agr_compra", "agr_venda", "vwap", "dia"],
    "macro_snapshots": ["id", "dia", "ts", "sp500", "sp500_var", "dolar",
                        "dolar_var", "di", "di_var_bps", "dxy", "dxy_var"],
    "eventos": ["id", "dia", "ts", "evento", "direcao", "nivel", "delta_ema", "msg"],
    "niveis_hist": ["id", "dia", "ts", "origem", "dados"],
    "operacoes": ["id", "ts", "dia", "hora", "lado", "preco", "motivo", "nota",
                  "abertura", "maxima", "minima", "volume", "delta", "vwap"],
}

# tabelas mutaveis keyed por PK composta/simples -> (colunas, chave)
UPSERT = {
    "corretoras": (["dia", "corretora", "qtd_compra", "qtd_venda", "fin_compra",
                    "fin_venda", "negocios"], ["dia", "corretora"]),
    "daily_ohlc": (["dia", "abertura", "maxima", "minima", "fechamento", "vwap",
                    "primeiro_ts", "ultimo_ts", "valido"], ["dia"]),
}


def get_conn_str():
    url = os.environ.get("AZURE_SQL_URL", "").strip()
    if not url and CONFIG_LOCAL.exists():
        url = CONFIG_LOCAL.read_text(encoding="utf-8").strip()
    if not url:
        sys.exit("AZURE_SQL_URL nao definida (env var) nem config.local encontrado.")
    return url


def load_state():
    if STATE_PATH.exists():
        return json.loads(STATE_PATH.read_text(encoding="utf-8"))
    return {}


def save_state(state):
    STATE_PATH.write_text(json.dumps(state, indent=2), encoding="utf-8")


def make_temp(pconn, tabela, cols):
    tmp = f"#{tabela}_tmp"
    collist = ", ".join(cols)
    with pconn.cursor() as pg:
        pg.execute(f"SELECT {collist} INTO {tmp} FROM dbo.{tabela} WHERE 1=0")
    return tmp


def sync_append_only(sconn, pconn, state):
    lite = sconn.cursor()
    for tabela, cols in APPEND_ONLY.items():
        wm = int(state.get(tabela, -1))
        collist = ", ".join(cols)
        lite.execute(f"select {collist} from {tabela} where id > ? order by id", (wm,))
        total, novo_wm = 0, wm
        tmp = make_temp(pconn, tabela, cols)
        placeholders = ", ".join(["?"] * len(cols))
        with pconn.cursor() as pg:
            pg.fast_executemany = True
            insert_tmp = f"insert into {tmp} ({collist}) values ({placeholders})"
            setinputsizes_for(pg, cols)
            while True:
                rows = lite.fetchmany(BATCH)
                if not rows:
                    break
                pg.executemany(insert_tmp, rows)
                total += len(rows)
                novo_wm = max(novo_wm, max(r[0] for r in rows))
            oncols = ", ".join(f"src.{c}" for c in cols)
            pg.execute(
                f"merge dbo.{tabela} as tgt "
                f"using {tmp} as src on tgt.id = src.id "
                f"when not matched then insert ({collist}) values ({oncols});"
            )
        pconn.commit()
        state[tabela] = novo_wm
        save_state(state)
        print(f"  {tabela:16s} +{total:>7d} linhas (id ate {novo_wm})")


def sync_upsert(sconn, pconn, _state):
    lite = sconn.cursor()
    for tabela, (cols, key) in UPSERT.items():
        collist = ", ".join(cols)
        lite.execute(f"select {collist} from {tabela}")
        rows = lite.fetchall()
        tmp = make_temp(pconn, tabela, cols)
        placeholders = ", ".join(["?"] * len(cols))
        setcols = [c for c in cols if c not in key]
        oncond = " and ".join(f"tgt.{k} = src.{k}" for k in key)
        oncols = ", ".join(f"src.{c}" for c in cols)
        with pconn.cursor() as pg:
            pg.fast_executemany = True
            setinputsizes_for(pg, cols)
            pg.executemany(f"insert into {tmp} ({collist}) values ({placeholders})", rows)
            setclause = ", ".join(f"tgt.{c}=src.{c}" for c in setcols)
            when_matched = f"when matched then update set {setclause} " if setcols else ""
            pg.execute(
                f"merge dbo.{tabela} as tgt "
                f"using {tmp} as src on {oncond} "
                f"{when_matched}"
                f"when not matched then insert ({collist}) values ({oncols});"
            )
        pconn.commit()
        print(f"  {tabela:16s}  {len(rows):>7d} linhas (upsert)")


def main():
    if not DB_PATH.exists():
        sys.exit(f"SQLite nao encontrado: {DB_PATH}")
    conn_str = get_conn_str()
    print(f"SQLite : {DB_PATH}")
    print(f"Azure SQL: {conn_str.split('Server=')[-1].split(';')[0]}")  # nao imprime a senha
    state = load_state()
    sconn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    try:
        with pyodbc.connect(conn_str, timeout=30) as pconn:
            print("append-only:")
            sync_append_only(sconn, pconn, state)
            print("upsert:")
            sync_upsert(sconn, pconn, state)
    finally:
        sconn.close()
    save_state(state)
    print("OK - sync concluido.")


if __name__ == "__main__":
    main()
