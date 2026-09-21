# -*- coding: utf-8 -*-
"""
================================================================
 MONITOR WIN - supabase_sync/verificar_pregao.py
 Verificacao pontual do sync Supabase apos um pregao real
----------------------------------------------------------------
 So leitura - nao mexe em win_history.db nem em sync_state.json.

 Criado 2026-08-22, depois do backup no Azure SQL ter sido
 desativado (tarefa `MonitorWinAzureSql` desabilitada, erro 40613
 persistente) e o Supabase ter virado o UNICO backup em nuvem ativo
 do Monitor WIN. A checagem de 22/08 foi feita num sabado sem dado
 novo - este script confere de verdade num dia de pregao.

 Compara duas coisas:
 1) Watermark salvo em sync_state.json vs MAX(id) real de cada
    tabela no SQLite local - gap 0 em todas = sync 100% em dia.
 2) Linhas NOVAS em sync.log desde o baseline (arquivo
    _baseline_check_<data>.json, gravado ANTES do pregao que esta
    sendo checado) - conta quantos "OK - sync concluido" e quantos
    "Traceback" apareceram desde entao. Sem baseline, cai no
    fallback de examinar so o ultimo bloco do log.

 Rodar manual:  py verificar_pregao.py [baseline.json] > relatorio.txt
 (a Tarefa Agendada unica criada em 22/08 ja roda isso automatico
 as 18:45 de 24/08/2026, redirecionando pro .txt.)
================================================================
"""
from __future__ import annotations

import json
import sqlite3
import sys
import time
from pathlib import Path

AQUI = Path(__file__).resolve().parent
DB_PATH = AQUI.parent / "win_history.db"
STATE_PATH = AQUI / "sync_state.json"
LOG_PATH = AQUI / "sync.log"


def checar_watermarks() -> tuple[list[str], bool]:
    linhas = ["## 1) Watermark (sync_state.json) vs MAX(id) real no SQLite\n"]
    if not STATE_PATH.exists():
        linhas.append(f"ERRO: {STATE_PATH.name} nao encontrado.\n")
        return linhas, False
    if not DB_PATH.exists():
        linhas.append(f"ERRO: {DB_PATH.name} nao encontrado.\n")
        return linhas, False

    state = json.loads(STATE_PATH.read_text(encoding="utf-8"))
    con = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    tudo_em_dia = True
    linhas.append(f"{'tabela':16} {'watermark':>12} {'max id local':>14} {'pendente':>10}\n")
    for tabela, watermark in state.items():
        try:
            row = con.execute(f"SELECT MAX(id) FROM {tabela}").fetchone()
        except sqlite3.OperationalError as e:
            linhas.append(f"  {tabela}: ERRO ao consultar ({e})\n")
            tudo_em_dia = False
            continue
        maxid = row[0] or 0
        pendente = maxid - watermark
        if pendente != 0:
            tudo_em_dia = False
        linhas.append(f"{tabela:16} {watermark:>12} {maxid:>14} {pendente:>10}\n")
    con.close()
    linhas.append(
        "\nOK: todas as tabelas com pendente=0 (sync em dia).\n"
        if tudo_em_dia else
        "\nATENCAO: pelo menos uma tabela com pendente != 0 - sync NAO esta em dia.\n"
    )
    return linhas, tudo_em_dia


def checar_log(baseline_path: Path | None) -> tuple[list[str], bool]:
    linhas = ["\n## 2) sync.log desde o baseline\n"]
    if not LOG_PATH.exists():
        linhas.append(f"ERRO: {LOG_PATH.name} nao encontrado.\n")
        return linhas, False

    todas = LOG_PATH.read_text(encoding="utf-8", errors="replace").splitlines()
    offset = 0
    if baseline_path and baseline_path.exists():
        baseline = json.loads(baseline_path.read_text(encoding="utf-8"))
        offset = baseline.get("log_lines_baseline", 0)
        linhas.append(
            f"Baseline: {baseline_path.name} (gravado em "
            f"{baseline.get('criado_em', '?')}, {offset} linhas)\n")
    else:
        linhas.append(
            "Sem baseline - examinando so o ultimo bloco do log (fallback).\n")
        # fallback: acha o inicio do ultimo bloco ("SQLite :" mais recente)
        for i in range(len(todas) - 1, -1, -1):
            if todas[i].startswith("SQLite :"):
                offset = i
                break

    novas = todas[offset:]
    n_ok = sum(1 for l in novas if l.strip() == "OK - sync concluido.")
    n_erro = sum(1 for l in novas if l.strip() == "Traceback (most recent call last):")
    linhas.append(f"Linhas novas desde o baseline: {len(novas)}\n")
    linhas.append(f"  'OK - sync concluido' novos: {n_ok}\n")
    linhas.append(f"  'Traceback' novos:           {n_erro}\n")

    log_ok = n_ok > 0 and n_erro == 0
    if log_ok:
        linhas.append("OK: pelo menos 1 sync novo com sucesso, nenhuma falha nova.\n")
    elif n_erro > 0:
        linhas.append("ATENCAO: apareceu Traceback novo no log - ver detalhe abaixo.\n")
        linhas.append("\n--- ultimas 40 linhas novas do log ---\n")
        linhas.extend(l + "\n" for l in novas[-40:])
    else:
        linhas.append(
            "ATENCAO: nenhum 'OK - sync concluido' novo desde o baseline - "
            "a tarefa agendada pode nao ter rodado.\n")
    return linhas, log_ok


def main():
    baseline_arg = Path(sys.argv[1]) if len(sys.argv) > 1 else None
    if baseline_arg is None:
        candidatos = sorted(AQUI.glob("_baseline_check_*.json"))
        baseline_arg = candidatos[-1] if candidatos else None

    print("=" * 64)
    print(" VERIFICACAO DO SYNC SUPABASE - Monitor WIN")
    print(f" Gerado em: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 64)

    linhas_wm, wm_ok = checar_watermarks()
    linhas_log, log_ok = checar_log(baseline_arg)

    print("".join(linhas_wm))
    print("".join(linhas_log))

    print("\n" + "=" * 64)
    if wm_ok and log_ok:
        print(" VEREDITO: PASS - Supabase sincronizou o pregao sem lacunas.")
    else:
        print(" VEREDITO: FAIL - ver detalhes acima (watermark e/ou log com problema).")
    print("=" * 64)


if __name__ == "__main__":
    main()
