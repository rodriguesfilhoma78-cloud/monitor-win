"""
================================================================
 MONITOR WDO - backtest_confluencia.py
 Backtest retroativo de assertividade (sem dependencias externas)
----------------------------------------------------------------
 Responde: quando o motor local (ConfluenceEngine, eventos de
 confluencia/divergencia) ou a leitura de IA (tabela `leituras`, nova -
 ver server_wdo.py) apontam uma direcao, o preco realmente continuou
 naquela direcao nos minutos seguintes?

 Cruza `eventos`/`leituras` (o "palpite" e o instante) com `snapshots`
 (preco tick a tick) pelo (dia, ts) - os dois ja existem no banco, isso e
 so leitura, nao muda nada ao vivo. Roda standalone:

    python backtest_confluencia.py

 IMPORTANTE sobre o que este numero significa (e o que nao significa):
 - "acerto" aqui e so "o preco se moveu na direcao apontada" - NAO e
   lucro/prejuizo de uma operacao real (sem custo, sem stop, sem gestao).
   E o baseline mais honesto que da pra tirar do que ja esta gravado.
 - A tabela `leituras` foi criada nesta sessao (11/08/2026) - o bloco de
   leituras so vai ter volume depois de alguns dias de coleta. O bloco de
   `eventos` (confluencia/divergencia) ja tem historico desde 03/07/2026 e
   e a baseline que da pra olhar hoje.
================================================================
"""
from __future__ import annotations

import json
import sqlite3
from pathlib import Path
from typing import Optional

DB_PATH = Path(__file__).parent / "wdo_history.db"
HORIZONTES_MIN = (5, 15, 30)   # janelas de continuacao checadas apos o sinal


def _ts_mais_minutos(ts: str, minutos: int) -> str:
    """'HH:MM:SS' + minutos -> 'HH:MM:SS' (sem cruzar meia-noite de verdade,
    so precisa ser comparavel dentro do mesmo pregao)."""
    h, m, s = (int(x) for x in ts.split(":"))
    total = h * 3600 + m * 60 + s + minutos * 60
    h2, resto = divmod(total % 86400, 3600)
    m2, s2 = divmod(resto, 60)
    return f"{h2:02d}:{m2:02d}:{s2:02d}"


def _preco_em_ou_apos(con: sqlite3.Connection, dia: str, ts: str) -> Optional[float]:
    """Primeiro preco de `snapshots` no dia com ts >= o pedido - None se o
    pregao acabou antes (sem dado, nao conta como erro)."""
    row = con.execute(
        "SELECT ultimo FROM snapshots WHERE dia=? AND ts>=? AND ultimo IS NOT NULL "
        "ORDER BY ts LIMIT 1", (dia, ts)).fetchone()
    return row[0] if row else None


def _novo_placar() -> dict:
    return {h: {"acertos": 0, "erros": 0, "neutros": 0, "sem_dado": 0}
            for h in HORIZONTES_MIN}


def _bate_direcao(delta: float, direcao_alta: bool) -> Optional[bool]:
    """None = empate exato (preco nao se moveu - nem acerto nem erro)."""
    if delta == 0:
        return None
    return (delta > 0) if direcao_alta else (delta < 0)


def backtest_eventos(con: sqlite3.Connection) -> tuple[dict, list[dict]]:
    """Confluencia/divergencia do ConfluenceEngine (`eventos`, 231 linhas
    desde 03/07/2026): a `direcao` do evento e o lado do rompimento no
    instante `ts` - o backtest checa se o preco continuou nesse lado."""
    linhas = con.execute(
        "SELECT dia, ts, evento, direcao, nivel FROM eventos "
        "WHERE evento IN ('confluencia','divergencia') ORDER BY dia, ts"
    ).fetchall()
    placar = {"confluencia": _novo_placar(), "divergencia": _novo_placar()}
    detalhes = []
    for dia, ts, evento, direcao, nivel in linhas:
        baseline = _preco_em_ou_apos(con, dia, ts)
        linha = {"dia": dia, "ts": ts, "evento": evento, "direcao": direcao,
                 "nivel": nivel, "baseline": baseline}
        for h in HORIZONTES_MIN:
            futuro = _preco_em_ou_apos(con, dia, _ts_mais_minutos(ts, h))
            chave = f"delta_{h}min"
            if baseline is None or futuro is None:
                placar[evento][h]["sem_dado"] += 1
                linha[chave] = None
                continue
            delta = futuro - baseline
            linha[chave] = round(delta, 1)
            ok = _bate_direcao(delta, direcao_alta=(direcao == "up"))
            if ok is None:
                placar[evento][h]["neutros"] += 1
            elif ok:
                placar[evento][h]["acertos"] += 1
            else:
                placar[evento][h]["erros"] += 1
        detalhes.append(linha)
    return placar, detalhes


def backtest_leituras(con: sqlite3.Connection) -> tuple[dict, dict, int]:
    """Leituras de IA (tabela `leituras`, criada 11/08/2026): compara o
    `vies` do MODELO e, separado, o `vies_mercado` CALCULADO (macro +
    var% do WDO, so quando forca>=2 = consenso real) contra o preco
    depois da leitura. Ignora vies 'misto'/'indefinido' (sem direcao pra
    testar)."""
    try:
        linhas = con.execute(
            "SELECT dia, ts, vies, vies_mercado, forca_mercado, preco FROM leituras "
            "ORDER BY dia, ts").fetchall()
    except sqlite3.OperationalError:
        # tabela `leituras` e nova (11/08/2026) - so existe depois do server
        # rodar _init() uma vez com o server_wdo.py atualizado.
        return _novo_placar(), _novo_placar(), 0
    placar_ia = _novo_placar()
    placar_mercado = _novo_placar()
    total = len(linhas)
    for dia, ts, vies, vies_mercado, forca, preco in linhas:
        baseline = preco if preco is not None else _preco_em_ou_apos(con, dia, ts)
        futuros = {h: _preco_em_ou_apos(con, dia, _ts_mais_minutos(ts, h))
                   for h in HORIZONTES_MIN}

        def marcar(placar, vies_alvo):
            if vies_alvo not in ("compra", "venda"):
                return
            for h in HORIZONTES_MIN:
                futuro = futuros[h]
                if baseline is None or futuro is None:
                    placar[h]["sem_dado"] += 1
                    continue
                ok = _bate_direcao(futuro - baseline, direcao_alta=(vies_alvo == "compra"))
                if ok is None:
                    placar[h]["neutros"] += 1
                elif ok:
                    placar[h]["acertos"] += 1
                else:
                    placar[h]["erros"] += 1

        marcar(placar_ia, vies)
        marcar(placar_mercado, vies_mercado if (forca or 0) >= 2 else None)
    return placar_ia, placar_mercado, total


def _taxa(c: dict) -> Optional[float]:
    base = c["acertos"] + c["erros"]
    return round(100 * c["acertos"] / base, 1) if base else None


def _imprimir_placar(titulo: str, placar_por_horizonte: dict):
    print(f"\n{titulo}")
    print(f"  {'horizonte':>10} | {'acertos':>7} | {'erros':>6} | "
          f"{'neutros':>7} | {'sem dado':>8} | {'taxa de acerto':>14}")
    for h in HORIZONTES_MIN:
        c = placar_por_horizonte[h]
        taxa = _taxa(c)
        taxa_txt = f"{taxa}%" if taxa is not None else "—"
        print(f"  {str(h)+'min':>10} | {c['acertos']:>7} | {c['erros']:>6} | "
              f"{c['neutros']:>7} | {c['sem_dado']:>8} | {taxa_txt:>14}")


def main():
    if not DB_PATH.exists():
        print(f"Banco nao encontrado: {DB_PATH}")
        return
    con = sqlite3.connect(DB_PATH)
    try:
        print("=" * 64)
        print(" BACKTEST DE ASSERTIVIDADE - Monitor WDO")
        print(" 'acerto' = preco se moveu na direcao apontada N minutos")
        print(" depois. NAO e resultado de operacao real (sem custo/stop).")
        print("=" * 64)

        placar_eventos, detalhes = backtest_eventos(con)
        n_conf = sum(sum(c.values()) for c in placar_eventos["confluencia"].values()) // len(HORIZONTES_MIN)
        n_div = sum(sum(c.values()) for c in placar_eventos["divergencia"].values()) // len(HORIZONTES_MIN)
        print(f"\n### Motor local (ConfluenceEngine) - {n_conf} confluencia, "
              f"{n_div} divergencia, 03/07 a hoje")
        _imprimir_placar("CONFLUENCIA (fluxo confirmou o rompimento)",
                          placar_eventos["confluencia"])
        _imprimir_placar("DIVERGENCIA (rompimento SEM confirmacao de fluxo)",
                          placar_eventos["divergencia"])

        placar_ia, placar_mercado, total_leituras = backtest_leituras(con)
        print(f"\n### Leitura de IA (tabela `leituras`, nova - {total_leituras} "
              f"linhas registradas ate agora)")
        if total_leituras == 0:
            print("  Sem dado ainda - a tabela `leituras` comecou a ser gravada "
                  "nesta sessao (11/08/2026). Rode este script de novo depois "
                  "de alguns dias de pregao pra ter uma taxa de acerto real.")
        else:
            _imprimir_placar("VIES DO MODELO (Gemini)", placar_ia)
            _imprimir_placar("VIES DE MERCADO CALCULADO (forca >= 2: macro + "
                              "WDO concordando)", placar_mercado)

        resumo = {
            "eventos": {tipo: {f"{h}min": placar_eventos[tipo][h] for h in HORIZONTES_MIN}
                        for tipo in ("confluencia", "divergencia")},
            "leituras_ia": {f"{h}min": placar_ia[h] for h in HORIZONTES_MIN},
            "leituras_vies_mercado": {f"{h}min": placar_mercado[h] for h in HORIZONTES_MIN},
            "total_leituras": total_leituras,
        }
        out_path = Path(__file__).parent / "backtest_resultado.json"
        out_path.write_text(json.dumps(resumo, ensure_ascii=False, indent=2),
                            encoding="utf-8")
        print(f"\nResumo salvo em {out_path.name}")
    finally:
        con.close()


if __name__ == "__main__":
    main()
