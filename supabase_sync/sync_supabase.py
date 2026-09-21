# -*- coding: utf-8 -*-
"""
Sync incremental do wdo_history.db (SQLite local) -> Supabase (Postgres).

Fork de sync_supabase.py do Monitor WIN: MESMO projeto/credencial Supabase,
mas grava em tabelas com prefixo wdo_ (wdo_fluxo, wdo_snapshots, etc.) -
tabelas PROPRIAS, criadas via migration em 24/08/2026, pra nao colidir/
misturar linha com o WIN (que usa as tabelas sem prefixo). Watermark
(sync_state.json) tambem e' proprio desta pasta - os ids do wdo_history.db
comecam do zero, independentes dos ids do WIN.

- Tabelas append-only (fluxo, snapshots, macro_snapshots, eventos, niveis_hist,
  operacoes): empurra so as linhas novas (id > watermark salvo em sync_state.json)
  com INSERT ... ON CONFLICT (id) DO NOTHING.
- Tabelas mutaveis pequenas (corretoras, daily_ohlc): reenviadas por inteiro com
  ON CONFLICT DO UPDATE (upsert).

Conexao: variavel de ambiente SUPABASE_DB_URL (a MESMA do Monitor WIN -
mesmo projeto Supabase, tabelas diferentes)
  ex: postgresql://postgres.<ref>:<senha>@aws-0-<regiao>.pooler.supabase.com:6543/postgres
  (Supabase Dashboard -> Project Settings -> Database -> Connection string -> URI)
Alternativa: arquivo config.local na mesma pasta com a URL numa unica linha.

Rodar:  py sync_supabase.py
"""
import os
import sys
import json
import sqlite3
from pathlib import Path

try:
    import psycopg
except ImportError:
    sys.exit("psycopg nao instalado. Rode:  py -m pip install \"psycopg[binary]\"")

AQUI = Path(__file__).resolve().parent
DB_PATH = Path(os.environ.get("WDO_DB_PATH", AQUI.parent / "wdo_history.db"))
STATE_PATH = AQUI / "sync_state.json"
CONFIG_LOCAL = AQUI / "config.local"
BATCH = 1000

# tabelas append-only keyed por id -> (colunas locais (sqlite), tabela remota)
APPEND_ONLY = {
    "fluxo": (["id", "dia", "ts", "bid", "ask", "spread", "qtd_bid", "qtd_ask",
              "desequilibrio", "profundidade_bid", "profundidade_ask",
              "lote_medio", "lote_max", "agr_compra_fita",
              "agr_venda_fita", "pressao_fita", "poc", "vah",
              "val", "vap_total", "vol_acima_pct", "dist_poc", "amostra_negocios",
              "amostra_contratos", "janela_s", "negocios_s", "contratos_s"], "wdo_fluxo"),
    "snapshots": (["id", "ts", "ultimo", "abertura", "maxima", "minima", "volume",
                  "agr_compra", "agr_venda", "vwap", "dia"], "wdo_snapshots"),
    # casado (27/08/2026): as 8 colunas do card CASADO ja existem no
    # wdo_macro_snapshots remoto (ALTER TABLE ... ADD COLUMN <col> double
    # precision aplicado em 27/08/2026) e vao no INSERT abaixo.
    "macro_snapshots": (["id", "dia", "ts", "brent", "brent_var", "ouro",
                        "ouro_var", "juros_us", "juros_us_var_bps", "dxy", "dxy_var",
                        "spot", "spot_var", "wdo_pts", "diferencial", "du",
                        "carrego_justo", "desvio", "cupom_impl"], "wdo_macro_snapshots"),
    "eventos": (["id", "dia", "ts", "evento", "direcao", "nivel", "delta_ema", "msg"], "wdo_eventos"),
    "niveis_hist": (["id", "dia", "ts", "origem", "dados"], "wdo_niveis_hist"),
    "operacoes": (["id", "ts", "dia", "hora", "lado", "preco", "motivo", "nota",
                  "abertura", "maxima", "minima", "volume", "delta", "vwap"], "wdo_operacoes"),
}

# tabelas mutaveis keyed por PK composta/simples -> (colunas, chave, tabela remota)
UPSERT = {
    "corretoras": (["dia", "corretora", "qtd_compra", "qtd_venda", "fin_compra",
                    "fin_venda", "negocios"], ["dia", "corretora"], "wdo_corretoras"),
    "daily_ohlc": (["dia", "abertura", "maxima", "minima", "fechamento", "vwap",
                    "primeiro_ts", "ultimo_ts", "valido"], ["dia"], "wdo_daily_ohlc"),
}


def get_db_url():
    url = os.environ.get("SUPABASE_DB_URL", "").strip()
    if not url and CONFIG_LOCAL.exists():
        url = CONFIG_LOCAL.read_text(encoding="utf-8").strip()
    if not url:
        sys.exit("SUPABASE_DB_URL nao definida (env var) nem config.local encontrado.")
    return url


def load_state():
    if STATE_PATH.exists():
        return json.loads(STATE_PATH.read_text(encoding="utf-8"))
    return {}


def save_state(state):
    STATE_PATH.write_text(json.dumps(state, indent=2), encoding="utf-8")


def sync_append_only(sconn, pconn, state):
    lite = sconn.cursor()
    for tabela_local, (cols, tabela_remota) in APPEND_ONLY.items():
        wm = int(state.get(tabela_local, -1))
        collist = ", ".join(cols)
        lite.execute(f"select {collist} from {tabela_local} where id > ? order by id", (wm,))
        placeholders = ", ".join(["%s"] * len(cols))
        insert = (f"insert into public.{tabela_remota} ({collist}) values ({placeholders}) "
                  f"on conflict (id) do nothing")
        total, novo_wm = 0, wm
        with pconn.cursor() as pg:
            while True:
                rows = lite.fetchmany(BATCH)
                if not rows:
                    break
                pg.executemany(insert, rows)
                total += len(rows)
                novo_wm = max(novo_wm, max(r[0] for r in rows))
        pconn.commit()
        state[tabela_local] = novo_wm
        print(f"  {tabela_remota:20s} +{total:>7d} linhas (id ate {novo_wm})")


def sync_upsert(sconn, pconn, _state):
    lite = sconn.cursor()
    for tabela_local, (cols, key, tabela_remota) in UPSERT.items():
        collist = ", ".join(cols)
        lite.execute(f"select {collist} from {tabela_local}")
        rows = lite.fetchall()
        placeholders = ", ".join(["%s"] * len(cols))
        setcols = [c for c in cols if c not in key]
        setclause = ", ".join(f"{c}=excluded.{c}" for c in setcols)
        conflict = ("do nothing" if not setcols
                    else f"do update set {setclause}")
        insert = (f"insert into public.{tabela_remota} ({collist}) values ({placeholders}) "
                  f"on conflict ({', '.join(key)}) {conflict}")
        with pconn.cursor() as pg:
            pg.executemany(insert, rows)
        pconn.commit()
        print(f"  {tabela_remota:20s}  {len(rows):>7d} linhas (upsert)")


def main():
    if not DB_PATH.exists():
        sys.exit(f"SQLite nao encontrado: {DB_PATH}")
    url = get_db_url()
    print(f"SQLite : {DB_PATH}")
    print(f"Supabase: {url.split('@')[-1]}")   # nao imprime a senha
    state = load_state()
    sconn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    try:
        with psycopg.connect(url, connect_timeout=30) as pconn:
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
