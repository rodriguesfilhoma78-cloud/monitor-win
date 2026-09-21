"""
================================================================
 MONITOR WDO - backtest_corretoras.py
 Backtest retroativo de "quem virou de lado" durante o pregao
----------------------------------------------------------------
 Le a serie temporal (corretoras_serie, gravada a cada SERIE_CORRETORAS_S -
 ver server_wdo.py) e reconstroi, para cada corretora, o saldo de uma
 JANELA MOVEL (default 30min, mesmo tamanho do ranking_janela ao vivo) em
 cada instante do dia - por DIFERENCA entre dois snapshots do acumulado
 cumulativo, nao reprocessando a fita.

 Aponta VIRADAS DE LADO: quando o saldo da janela cruza de compra pra venda
 (ou o contrario) com magnitude relevante dos dois lados - ex.: corretora
 comprando forte de manha e vendendo forte de tarde. Compara tambem o saldo
 do DIA INTEIRO com o saldo da ULTIMA janela, pra achar corretoras cujo
 comportamento recente diverge da tendencia do dia (o mesmo par top5/
 top5_janela que a IA recebe ao vivo, aqui olhado retroativamente e pra
 toda corretora, nao so o top5).

 IMPORTANTE: "virada" aqui e leitura de FLUXO AGRESSOR, nao de posicao -
 mesma ressalva do ranking ao vivo (RankingCorretoras): uma corretora e
 muitos clientes, o sinal util e dominancia/troca de mao, nao certeza de
 quem esta posicionado. saldo_dia usa o ULTIMO snapshot gravado no dia,
 entao pode ficar levemente defasado do fechamento exato (ate
 SERIE_CORRETORAS_S de atraso).

 Roda standalone, sem dependencias externas:
    python backtest_corretoras.py [--dia AAAA-MM-DD] [--janela-min 30]
                                   [--min-lote 50]
================================================================
"""
from __future__ import annotations

import argparse
import json
import sqlite3
from bisect import bisect_right
from pathlib import Path
from typing import Optional

DB_PATH = Path(__file__).parent / "wdo_history.db"
JANELA_MIN_DEFAULT = 30      # mesmo tamanho do JANELA_CORRETORAS_S ao vivo
MIN_LOTE_DEFAULT = 50        # saldo minimo (contratos) dos dois lados pra uma
                             # virada contar - filtra ruido de lotes pequenos


def _ts_para_s(ts: str) -> int:
    h, m, s = (int(x) for x in ts.split(":"))
    return h * 3600 + m * 60 + s


def _s_para_ts(segundos: int) -> str:
    h, resto = divmod(segundos % 86400, 3600)
    m, s = divmod(resto, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def _dia_mais_recente(con: sqlite3.Connection) -> Optional[str]:
    row = con.execute("SELECT MAX(dia) FROM corretoras_serie").fetchone()
    return row[0] if row and row[0] else None


def _carregar_serie(con: sqlite3.Connection, dia: str) -> dict:
    """corretora -> lista de (ts_s, qtd_compra, qtd_venda) ordenada por ts."""
    linhas = con.execute(
        "SELECT ts, corretora, qtd_compra, qtd_venda FROM corretoras_serie "
        "WHERE dia=? ORDER BY corretora, ts", (dia,)).fetchall()
    por_corretora: dict = {}
    for ts, corretora, qc, qv in linhas:
        por_corretora.setdefault(corretora, []).append((_ts_para_s(ts), qc, qv))
    return por_corretora


def _saldo_janela(pontos: list, i: int, janela_s: int) -> Optional[float]:
    """Saldo da janela terminando no snapshot i, por diferenca com o
    snapshot mais proximo de (ts_i - janela_s). None se a janela ainda nao
    'enche' (comeco do pregao, sem snapshot antigo suficiente)."""
    ts_i, qc_i, qv_i = pontos[i]
    limite = ts_i - janela_s
    if pontos[0][0] > limite:
        return None
    ts_list = [p[0] for p in pontos]
    j = max(bisect_right(ts_list, limite) - 1, 0)
    _, qc_ref, qv_ref = pontos[j]
    return (qc_i - qc_ref) - (qv_i - qv_ref)


def analisar_corretora(nome: str, pontos: list, janela_s: int,
                        min_lote: float) -> dict:
    """Serie de saldo_janela ao longo do dia + viradas de lado detectadas."""
    serie = [(ts, _saldo_janela(pontos, i, janela_s))
             for i, (ts, _, _) in enumerate(pontos)]
    serie_valida = [(ts, s) for ts, s in serie if s is not None]

    # Compara cada ponto RELEVANTE (|saldo| >= min_lote) contra o ultimo
    # ponto relevante visto, nao so o vizinho imediato - senao uma sequencia
    # que passa por perto de zero (900 -> 0 -> -1400) escondia a virada, ja
    # que nenhum PAR ADJACENTE isolado teria os dois lados acima do minimo.
    viradas = []
    ultimo: Optional[tuple] = None      # (ts, saldo) do ultimo ponto relevante
    for ts, s in serie_valida:
        if abs(s) < min_lote:
            continue
        if ultimo is not None and (ultimo[1] > 0) != (s > 0):
            viradas.append({
                "de_ts": _s_para_ts(ultimo[0]), "para_ts": _s_para_ts(ts),
                "saldo_antes": round(ultimo[1], 0), "saldo_depois": round(s, 0),
            })
        ultimo = (ts, s)

    _, qc_ult, qv_ult = pontos[-1]
    saldo_dia = qc_ult - qv_ult
    saldo_janela_final = serie_valida[-1][1] if serie_valida else None
    divergiu = (saldo_janela_final is not None and saldo_dia != 0
                and (saldo_janela_final > 0) != (saldo_dia > 0)
                and abs(saldo_janela_final) >= min_lote)
    return {
        "corretora": nome, "saldo_dia": round(saldo_dia, 0),
        "saldo_janela_final": (round(saldo_janela_final, 0)
                                if saldo_janela_final is not None else None),
        "divergiu_do_dia": divergiu,
        "n_viradas": len(viradas), "viradas": viradas,
    }


def main():
    parser = argparse.ArgumentParser(
        description="Backtest de virada de lado por corretora (corretoras_serie)")
    parser.add_argument("--dia", help="AAAA-MM-DD (default: dia mais recente na serie)")
    parser.add_argument("--janela-min", type=int, default=JANELA_MIN_DEFAULT)
    parser.add_argument("--min-lote", type=float, default=MIN_LOTE_DEFAULT)
    args = parser.parse_args()

    if not DB_PATH.exists():
        print(f"Banco nao encontrado: {DB_PATH}")
        return
    con = sqlite3.connect(DB_PATH)
    try:
        tabelas = {r[0] for r in con.execute(
            "SELECT name FROM sqlite_master WHERE type='table'")}
        if "corretoras_serie" not in tabelas:
            print("Tabela corretoras_serie nao existe ainda - rode o server "
                  "(server_wdo.py) durante um pregao com a serie temporal "
                  "ativa antes de rodar este backtest.")
            return

        dia = args.dia or _dia_mais_recente(con)
        if not dia:
            print("Nenhum dado em corretoras_serie ainda - sem pregao gravado.")
            return

        janela_s = args.janela_min * 60
        serie_por_corretora = _carregar_serie(con, dia)
        if not serie_por_corretora:
            print(f"Sem dados de corretoras_serie para o dia {dia}.")
            return

        resultados = [analisar_corretora(nome, pontos, janela_s, args.min_lote)
                      for nome, pontos in serie_por_corretora.items()]
        resultados.sort(key=lambda r: abs(r["saldo_dia"]), reverse=True)

        divergentes = [r for r in resultados if r["divergiu_do_dia"]]
        com_viradas = [r for r in resultados if r["n_viradas"] > 0]

        print("=" * 72)
        print(f" BACKTEST DE VIRADA DE LADO POR CORRETORA - dia {dia}")
        print(f" janela = {args.janela_min}min | lote minimo p/ contar virada = {args.min_lote}")
        print(" 'virada' = saldo da janela cruzou de compra pra venda (ou o")
        print(" contrario) durante o pregao. Corretora agressora NAO e")
        print(" posicao - o sinal util e dominancia/troca de mao.")
        print("=" * 72)
        print(f"\n{len(resultados)} corretoras no dia | "
              f"{len(divergentes)} divergindo do saldo do dia | "
              f"{len(com_viradas)} com pelo menos 1 virada")

        print(f"\n{'corretora':<24} | {'saldo dia':>10} | {'saldo janela':>13} | "
              f"{'diverge?':>8} | {'viradas':>7}")
        for r in resultados:
            sj = r["saldo_janela_final"]
            sj_txt = f"{sj:.0f}" if sj is not None else "-"
            print(f"{r['corretora']:<24} | {r['saldo_dia']:>10.0f} | {sj_txt:>13} | "
                  f"{'SIM' if r['divergiu_do_dia'] else 'nao':>8} | {r['n_viradas']:>7}")

        if com_viradas:
            print("\n### Detalhe das viradas")
            for r in com_viradas:
                print(f"\n  {r['corretora']} ({r['n_viradas']} virada(s)):")
                for v in r["viradas"]:
                    print(f"    {v['de_ts']} (saldo {v['saldo_antes']:.0f}) -> "
                          f"{v['para_ts']} (saldo {v['saldo_depois']:.0f})")

        out_path = Path(__file__).parent / "backtest_corretoras_resultado.json"
        out_path.write_text(
            json.dumps({"dia": dia, "janela_min": args.janela_min,
                        "min_lote": args.min_lote, "resultados": resultados},
                       ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"\nResumo salvo em {out_path.name}")
    finally:
        con.close()


if __name__ == "__main__":
    main()
