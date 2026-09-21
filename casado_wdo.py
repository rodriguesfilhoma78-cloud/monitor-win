"""
================================================================
 MONITOR WDO - casado_wdo.py
 Preco justo do WDO por arbitragem contra o dolar a vista (o "casado")
----------------------------------------------------------------
 Caminho A (decisao do usuario em 27/08/2026): sem mexer na planilha.

   preco_justo_WDO = pronto_pts + carrego
   diferencial     = WDO_real - pronto_pts        (o que o trader olha)
   desvio_de_fluxo = WDO_real - preco_justo_WDO

 - pronto (dolar a vista): AwesomeAPI (economia.awesomeapi.com.br),
   dolar comercial, sem chave, cache de ~1 min. Fallback: Yahoo BRL=X.
   Buscado pelo SpotFetcher no server_wdo.py.
 - DI (juro domestico): ja vem no dados_macro_rtd.csv (DI1F*/DI1N*),
   lido pelo MacroRtdReader no server_wdo.py. Usado so pra decompor o
   carrego bruto e estimar o cupom implicito - NAO entra no preco justo.
 - cupom cambial: nao existe fonte gratuita em tempo real (B3 so vende o
   dado; brapi.dev e' 15 min atrasado e so acoes). Por isso o carrego
   "justo" e' AJUSTADO DO HISTORICO: regressao diaria de
   `diferencial ~ du` sobre os ultimos pregoes (server_wdo: casado_calib).
   Enquanto nao ha historico suficiente, cai pro carrego observado na
   ABERTURA do proprio dia (desvio vira "movimento do diferencial desde a
   abertura", que ja e' um proxy de fluxo util).

 Ressalvas (ver conversa de 27/08 e README):
 - AwesomeAPI atrasa/trava fora do horario de Londres/NY: parte do desvio
   pode ser so o pronto correndo atras do futuro. `idade_s` no payload.
 - Carrego fitado so vale depois de MIN_DIAS_CALIB pregoes com a coluna
   nova. Antes disso o rotulo diz "carrego: abertura de hoje".
 - Virada de vencimento (1o dia util do mes): du despenca, diferencial
   pula. O server trata a data de rolagem, mas no proprio dia da virada o
   sinal fica ruim.
 - Intervencao do BC, PTAX do ultimo dia util e fim de mes distorcem o
   diferencial - mesmo "modelo quebrado" do card MACRO.
================================================================
"""
from __future__ import annotations

from datetime import date, timedelta
from typing import Optional

# WDO/DOL sao cotados em reais por US$ 1.000 -> pronto_pts = taxa * 1000.
PONTOS_POR_DOLAR = 1000.0

# Historico minimo (pregoes) pra confiar na regressao diferencial~du.
MIN_DIAS_CALIB = 8

# z-score do desvio a partir do qual o rotulo deixa de ser "neutro".
Z_ESTICADO = 1.5

# --- Feriados B3 (bolsa fechada) --------------------------------------
# B3 fica em Sao Paulo, entao inclui os feriados estaduais/municipais que
# a bolsa observa (25/01 aniversario de SP, 09/07 Revolucao de 32) e o
# ultimo dia util do ano (31/12, bolsa nao negocia).
# MANUTENCAO: atualizar todo fim de ano. Se faltar um feriado aqui, o `du`
# sai 1 dia maior e o carrego bruto/cupom implicito ficam levemente off -
# nao quebra nada, so envelhece a precisao.
_FERIADOS_B3 = {
    # 2026
    "2026-01-01", "2026-01-25", "2026-02-16", "2026-02-17",
    "2026-04-03", "2026-04-21", "2026-05-01", "2026-06-04",
    "2026-07-09", "2026-09-07", "2026-10-12", "2026-11-02",
    "2026-11-15", "2026-11-20", "2026-12-25", "2026-12-31",
    # 2027
    "2027-01-01", "2027-01-25", "2027-02-08", "2027-02-09",
    "2027-03-26", "2027-04-21", "2027-05-01", "2027-05-27",
    "2027-07-09", "2027-09-07", "2027-10-12", "2027-11-02",
    "2027-11-15", "2027-11-20", "2027-12-25", "2027-12-31",
}


def _e_dia_util(d: date) -> bool:
    return d.weekday() < 5 and d.isoformat() not in _FERIADOS_B3


def _primeiro_dia_util(ano: int, mes: int) -> date:
    d = date(ano, mes, 1)
    while not _e_dia_util(d):
        d += timedelta(days=1)
    return d


def vencimento_frente(hoje: Optional[date] = None) -> date:
    """Vencimento do contrato de dolar da frente (DOLFUT_F_0 do Profit).

    O futuro de dolar vence no 1o dia util do mes do contrato, e o
    contrato negociado durante um mes qualquer e' o do mes SEGUINTE
    (em agosto negocia-se DOLU26, que vence 1o dia util de setembro).
    No proprio dia da rolagem, ja rola pro mes seguinte.
    """
    hoje = hoje or date.today()
    ano = hoje.year + (1 if hoje.month == 12 else 0)
    mes = 1 if hoje.month == 12 else hoje.month + 1
    v = _primeiro_dia_util(ano, mes)
    if hoje >= v:                      # dia da rolagem (ou depois): pula um mes
        ano = ano + (1 if mes == 12 else 0)
        mes = 1 if mes == 12 else mes + 1
        v = _primeiro_dia_util(ano, mes)
    return v


def dias_uteis_ate(venc: date, hoje: Optional[date] = None) -> int:
    """Dias uteis (B3) de hoje ate o vencimento, exclusivo no vencimento.

    Ex.: se venc e' amanha e amanha e' dia util -> 1.
    """
    hoje = hoje or date.today()
    du, d = 0, hoje
    while d < venc:
        if _e_dia_util(d):
            du += 1
        d += timedelta(days=1)
    return du


def regredir_carrego(pontos: list[tuple[float, float]]) -> Optional[dict]:
    """OLS simples `diferencial = a + b*du` sobre (du, diferencial) de
    pregoes passados (1 ponto por dia - o diferencial da abertura, que
    carrega menos ruido de fluxo intradiario).

    Devolve {slope, intercept, n, r2} ou None se amostra insuficiente ou
    sem variacao de du (ex.: todos os dias no mesmo ponto do mes).
    """
    pts = [(x, y) for x, y in pontos if x is not None and y is not None]
    n = len(pts)
    if n < MIN_DIAS_CALIB:
        return None
    sx = sum(x for x, _ in pts)
    sy = sum(y for _, y in pts)
    mx, my = sx / n, sy / n
    sxx = sum((x - mx) ** 2 for x, _ in pts)
    sxy = sum((x - mx) * (y - my) for x, y in pts)
    if sxx <= 1e-9:
        return None
    b = sxy / sxx
    a = my - b * mx
    syy = sum((y - my) ** 2 for _, y in pts)
    r2 = (sxy ** 2 / (sxx * syy)) if syy > 1e-9 else 0.0
    return {"slope": round(b, 4), "intercept": round(a, 3),
            "n": n, "r2": round(r2, 3)}


def calcular(wdo_pts: Optional[float], spot_rate: Optional[float],
             dif_abertura: Optional[float] = None,
             calib: Optional[dict] = None,
             di_anual: Optional[float] = None,
             du: Optional[int] = None,
             venc: Optional[date] = None,
             desvio_hist: Optional[list[float]] = None) -> Optional[dict]:
    """Monta o pacote do casado. Retorna None se faltar WDO ou pronto.

    - wdo_pts   : ultimo do WDO (pontos, ja na escala do futuro)
    - spot_rate : dolar a vista (ex.: 5.1606) - vira spot_pts = *1000
    - dif_abertura : diferencial capturado na abertura de hoje (fallback
                     de carrego enquanto nao ha calibracao)
    - calib     : saida de regredir_carrego() (carrego fitado do historico)
    - di_anual  : juro DI curto (ex.: 0.1389) - so pra cupom implicito
    - du, venc  : dias uteis / data de vencimento (o server calcula e passa)
    - desvio_hist : desvios recentes na MESMA faixa de horario, pro z-score
    """
    if wdo_pts is None or spot_rate is None:
        return None
    if venc is None:
        venc = vencimento_frente()
    if du is None:
        du = dias_uteis_ate(venc)

    spot_pts = spot_rate * PONTOS_POR_DOLAR
    diferencial = wdo_pts - spot_pts

    if calib and du > 0:
        carrego_justo = calib["slope"] * du + calib["intercept"]
        fonte_carrego = f"histórico ({calib['n']}d, R²={calib['r2']:.2f})"
    elif dif_abertura is not None:
        carrego_justo = dif_abertura
        fonte_carrego = "abertura de hoje (sem histórico)"
    else:
        carrego_justo = diferencial            # dia 1, 1o tick: desvio ~ 0
        fonte_carrego = "1º tick (sem referência)"

    preco_justo = spot_pts + carrego_justo
    desvio = wdo_pts - preco_justo

    # Cupom cambial implicito no diferencial atual (so exibicao):
    #   diferencial ~= spot_pts * (DI - cupom) * du/252
    #   => cupom ~= DI - diferencial/spot_pts * 252/du
    # du muito curto (virada de vencimento) amplifica ruido no cupom -
    # nao vale a pena exibir um numero que pula 5 p.p. por 1 ponto de fita.
    cupom_impl = None
    if di_anual is not None and du >= 3 and spot_pts > 0:
        cupom_impl = di_anual - (diferencial / spot_pts) * (252.0 / du)

    z = None
    rotulo = "neutro"
    if desvio_hist:
        amostra = [x for x in desvio_hist if x is not None]
        if len(amostra) >= 10:
            m = sum(amostra) / len(amostra)
            var = sum((x - m) ** 2 for x in amostra) / len(amostra)
            sd = var ** 0.5
            if sd > 1e-6:
                z = (desvio - m) / sd
                if z >= Z_ESTICADO:
                    rotulo = "esticado"          # futuro rico vs drivers
                elif z <= -Z_ESTICADO:
                    rotulo = "atrasado"          # futuro barato vs drivers

    return {
        "spot_rate": round(spot_rate, 4),
        "spot_pts": round(spot_pts, 1),
        "wdo_pts": round(wdo_pts, 1),
        "diferencial": round(diferencial, 1),
        "carrego_justo": round(carrego_justo, 1),
        "preco_justo": round(preco_justo, 1),
        "desvio": round(desvio, 1),
        "desvio_z": round(z, 2) if z is not None else None,
        "rotulo": rotulo,
        "cupom_impl_pct": round(cupom_impl * 100, 2) if cupom_impl is not None else None,
        "di_curto_pct": round(di_anual * 100, 2) if di_anual is not None else None,
        "du": du,
        "vencimento": venc.isoformat(),
        "fonte_carrego": fonte_carrego,
    }


def voto_vies(casado: Optional[dict]) -> Optional[int]:
    """Voto pro vies_consolidado (agente_wdo): +1 fluxo comprador
    estrutural no dolar, -1 vendedor. So vota quando o desvio esta
    claramente esticado/atrasado (|z| >= Z_ESTICADO); fora disso nao
    opina (o diferencial no range normal nao e' sinal de direcao).
    """
    if not casado or casado.get("desvio_z") is None:
        return None
    z = casado["desvio_z"]
    if z >= Z_ESTICADO:
        return 1
    if z <= -Z_ESTICADO:
        return -1
    return None
