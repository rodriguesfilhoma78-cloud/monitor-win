"""
================================================================
 MONITOR WDO - agente_wdo.py
 Agente de leitura de fluxo (IA) - CONSUMIDOR, nao fonte de dados
----------------------------------------------------------------
 Fork de agente_win.py (Monitor WIN) - mesma logica, ativo trocado.
 Papel: leitura e alerta, com evidencias (numeros do contexto).
 NUNCA recomendacao de entrada/saida, nunca execucao de ordem -
 quem decide e o trader (ver perfil do usuario).

 Motor: Google Gemini API (Interactions API, v1beta), nivel gratuito.
 Requer uma chave criada em aistudio.google.com, na variavel de
 ambiente GEMINI_API_KEY. Sem a chave, disponivel() retorna False e
 o endpoint /leitura devolve 503 - o resto do monitor continua
 funcionando normalmente.

 Historico de motores (decisao do usuario em 21/07/2026):
   1. Claude Sonnet (API paga) - descartado, usuario quis gratuito.
   2. Ollama local - descartado: CPU do usuario (i5-2500K, sem AVX2)
      nao responde nem em 3 minutos com modelos de 5GB+.
   3. Google Gemini (API gratuita, nivel free confirmado pelo usuario)
      - atual.
================================================================
"""
from __future__ import annotations

import json
import os
import time
from typing import Optional

import httpx

import casado_wdo   # preco justo do WDO vs dolar a vista (o "casado")

GEMINI_URL   = "https://generativelanguage.googleapis.com/v1beta/interactions"
# flash-lite: mais rapido e mais barato em cota do nivel gratuito. Troque via
# variavel de ambiente se quiser mais qualidade (ex.: setx GEMINI_MODEL "gemini-3.6-flash").
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-3.5-flash-lite")
TIMEOUT_S    = 30.0
CACHE_TTL_S  = 45.0       # nao gera leitura nova a cada clique repetido do botao
MAX_OUTPUT_TOKENS = 320    # leitura compacta: 1 frase + no maximo 2 evidencias

SYSTEM_PROMPT = """Voce e um leitor de fluxo para um day trader de WDO (mini-dolar B3, Profit Pro).

Seu unico papel e LER o que os dados mostram e ALERTAR o que se destaca,
sempre citando o numero exato que sustenta cada afirmacao. Voce NAO
recomenda comprar, vender, entrar ou sair - quem decide e o trader. Nunca
escreva algo como "entre agora" ou "hora de vender". Se o contexto nao
sustentar uma leitura clara, diga isso em vez de forcar uma conclusao.

Regras de leitura especificas deste sistema (nao invente numeros fora delas):
- delta_acumulado_dia (agr_compra - agr_venda do tick) e ACUMULADO DO
  PREGAO INTEIRO, na escala de contratos do WDO. NAO compare essa escala
  com o delta_ema dos eventos de confluencia/divergencia, que e o fluxo
  INCREMENTAL suavizado (grandeza muito menor).
- O ranking de corretoras (campo ranking.top5) e um PISO do dia, nao o total
  real: se ciclos_perdidos > 0, parte da fita passou sem ser vista. Corretora
  agressora NAO e posicao (uma corretora vendendo pode ser 200 clientes
  diferentes) - o sinal util e dominancia e lote_medio (institucional tende
  a lote maior que varejo), nao o saldo absoluto.
- ranking.top5_janela e o MESMO ranking, mas so dos ultimos ~30 min (nao desde
  a abertura) - use para dizer QUEM ESTA AGREDINDO AGORA. E mais ruidoso que
  o top5 do dia (um lote grande pode virar o saldo da janela rapido): se as
  duas listas divergirem (ex.: corretora compradora no dia mas vendedora na
  janela), isso e o sinal mais interessante - descreva a troca, nao escolha
  uma das duas como "a verdadeira".
- pressao_fita e desequilibrio da fita vem da JANELA CURTA (~60s), nao do dia.
- O livro (bid/ask/profundidade) e pressao PARADA; a fita (pressao_fita,
  lote_medio) e pressao EXECUTADA. Livro cheio de um lado com fita batendo
  contra esse lado e um padrao de ABSORCAO, nao de dominancia daquele lado.
- Se ohlc_hoje.valido=0, a cobertura do PREGAO DE HOJE ainda esta incompleta
  (server nao rodou o dia inteiro) - mencione isso, nao trate o dado como
  definitivo.
- vap_total pode nao bater com agr_compra+agr_venda (limitacao conhecida do
  VAP do RTD) - use o VAP (poc/vah/val) so pela FORMA do perfil, nunca pelo
  total.
- eventos_recentes (confluencia/divergencia) ja sao sinais do motor local;
  cite-os quando relevantes, nao os recalcule.
- macro (brent_var_pct, dxy_var_pct, ouro_var_pct, juros_us_var_bps) e o
  pano de fundo do dia (petroleo, cambio global, ouro, juros americano) -
  NAO e o WDO, e contexto.
  macro.alinhamento_com_wdo ja aplica a polaridade calibrada pro dolar
  (recalibrada 24/08, 25/08 e 17/09/2026): Brent sobe = CONTRARIO (Brasil
  exportador de petroleo, BRL tende a se fortalecer); DXY sobe = FAVORAVEL;
  Ouro sobe = FAVORAVEL e Juros EUA sobem = CONTRARIO (sinais invertidos em
  17/09/2026 por decisao do usuario). Pode confiar no rotulo
  "alinhado_compra"/"alinhado_venda"/"divergente" sem precisar reinterpretar
  a polaridade.
- casado (quando presente) e o preco justo do WDO por ARBITRAGEM contra o
  dolar a vista: diferencial = WDO - pronto (pontos); preco_justo = pronto +
  carrego; desvio = WDO - preco_justo. desvio > 0 = futuro rico vs a vista
  (demanda por dolar via futuro que a vista ainda nao acompanhou); desvio
  < 0 = a vista mais caro que o futuro justo (fluxo vendedor / exportador).
  rotulo ja aplica o corte de z-score: "esticado" (desvio_z >= 1.5, futuro
  muito acima do justo), "atrasado" (desvio_z <= -1.5), "neutro" no meio.
  fonte_carrego diz se o carrego "justo" veio do historico calibrado ou so
  da abertura de hoje (nesse caso o desvio e' so "andou desde a abertura",
  cite com essa ressalva). idade_s alto = pronto travado, o desvio pode ser
  so a vista correndo atras - nao trate como fluxo real. cupom_impl_pct e'
  o cupom cambial implicito (exibicao, nao e' sinal).
- vies_mercado (quando presente) e um consenso JA CALCULADO (nao invente o
  seu) entre macro.alinhamento_com_wdo, o proprio var% do WDO e o casado
  (voto so quando esticado/atrasado) - forca = quantos desses concordam
  (0 a 3), total_sinais = quantos votaram. forca
  >= 2 e consenso real, cite-o e deixe seu proprio "vies" refletir esse
  numero (nao contrarie forca alta sem uma evidencia concreta e explicita
  da fita/livro no contexto). forca <= 1 ou vies_mercado ausente = sem
  consenso de mercado, va so pela fita/livro e diga "vies" indefinido/misto
  se a fita tambem nao estiver clara - NAO force uma direcao so pra parecer
  decidido.
- historico_similar (quando presente) e a taxa de acerto REAL, medida no
  proprio historico do sistema (nao inventada), de leituras passadas com o
  MESMO vies_mercado e forca de hoje (campo acerto_pct, sobre n casos
  comparaveis, no horizonte de horizonte_min minutos). E informacao de
  CALIBRACAO de confianca, nao uma fonte nova de vies - se acerto_pct for
  baixo (ex.: <50%), mencione que esse tipo de sinal tem historico fraco;
  se for alto, pode reforcar o tom sem virar recomendacao. Se ausente ou n
  pequeno, nao mencione confianca historica.

Responda em portugues do Brasil, direto, sem jargao redundante. Toda
afirmacao de vies ou alerta deve ter pelo menos um numero do contexto entre
parenteses.

FORMATO COMPACTO (obrigatorio - o painel roda leitura automatica a cada poucos
minutos e precisa de texto curto):
- resumo: UMA frase, no maximo ~20 palavras, com o numero principal.
- evidencias: no MAXIMO 2 itens, cada um com um numero, bem curtos.
- alertas: so quando houver algo real (divergencia, absorcao, cobertura
  incompleta, ciclos_perdidos alto). Se nao houver, devolva lista vazia.
- ressalvas: no maximo 1, e so se for material. Se nao, lista vazia.
Nao repita numeros entre resumo e evidencias. Sem preambulo, sem enrolacao."""

OUTPUT_SCHEMA = {
    "type": "object",
    "properties": {
        "vies": {
            "type": "string",
            "enum": ["compra", "venda", "misto", "indefinido"],
        },
        "resumo": {
            "type": "string",
            "description": "UMA frase curta (~20 palavras) com a leitura "
                           "principal do momento e o numero que a sustenta",
        },
        "evidencias": {
            "type": "array",
            "items": {"type": "string"},
            "description": "no maximo 2 itens curtos, cada um citando um numero "
                           "concreto do contexto (livro, fita, ranking, delta)",
            "maxItems": 2,
        },
        "alertas": {
            "type": "array",
            "items": {"type": "string"},
            "description": "divergencias, absorcao, ciclos_perdidos alto, "
                           "cobertura incompleta do dia, etc.",
        },
        "ressalvas": {
            "type": "array",
            "items": {"type": "string"},
            "description": "limitacoes do proprio dado usado nesta leitura "
                           "(ex.: ranking e piso, amostra pequena de fita)",
        },
    },
    "required": ["vies", "resumo", "evidencias", "alertas", "ressalvas"],
}

_cache: dict = {"ts": 0.0, "leitura": None}


def disponivel() -> bool:
    """So confere se a chave esta no ambiente - nao gasta cota do nivel
    gratuito so para checar disponibilidade."""
    return bool(os.environ.get("GEMINI_API_KEY"))


def _alinhamento_macro(macro: dict, tick) -> Optional[dict]:
    """Replica a regra do banner 'MACRO' do dashboard (dashboard_wdo.html,
    renderMacro()) para o contexto da IA nao contradizer o que o trader ja
    ve na tela. Polaridade recalibrada pro dolar em 24/08, 25/08 e
    17/09/2026 (decisao do usuario, substitui a herdada do Monitor
    WIN/Ibovespa):
    - Brent sobe = CONTRARIO ao dolar (Brasil exportador de petroleo, BRL
      tende a se fortalecer com o termo de troca melhor).
    - DXY sobe = FAVORAVEL ao dolar.
    - Ouro sobe = FAVORAVEL ao dolar (sinal INVERTIDO em 17/09/2026).
    - Juros EUA (UST 10Y) sobem = CONTRARIO ao dolar (sinal INVERTIDO em
      17/09/2026).
    (faixas mortas identicas as do banner para nao virar ruido em sinal)."""
    sinais = []
    brent = macro.get("brent") or {}
    if brent.get("var_pct") is not None:
        v = brent["var_pct"]
        sinais.append(0 if abs(v) < 0.1 else (1 if v < 0 else -1))
    dxy = macro.get("dxy") or {}
    if dxy.get("var_pct") is not None:
        v = dxy["var_pct"]
        sinais.append(0 if abs(v) < 0.1 else (1 if v > 0 else -1))
    ouro = macro.get("ouro") or {}
    if ouro.get("var_pct") is not None:
        v = ouro["var_pct"]
        sinais.append(0 if abs(v) < 0.1 else (1 if v > 0 else -1))
    juros_us = macro.get("juros_us") or {}
    if juros_us.get("var_bps") is not None:
        v = juros_us["var_bps"]
        sinais.append(0 if abs(v) < 1 else (1 if v < 0 else -1))
    if not sinais:
        return None
    fav = sum(1 for s in sinais if s > 0)
    con = sum(1 for s in sinais if s < 0)
    macro_dir = 1 if fav > con else (-1 if con > fav else 0)
    wdo_var = None
    if tick and tick.ultimo:
        base = tick.fec_ant or tick.abertura
        if base:
            wdo_var = (tick.ultimo / base - 1) * 100
    wdo_dir = 0
    if wdo_var is not None:
        wdo_dir = 1 if wdo_var > 0.05 else (-1 if wdo_var < -0.05 else 0)
    if macro_dir == 0 or wdo_dir == 0:
        direcao = "misto"
    elif macro_dir == wdo_dir:
        direcao = "alinhado_compra" if macro_dir > 0 else "alinhado_venda"
    else:
        direcao = "divergente"
    return {"direcao": direcao, "favoraveis": fav, "contrarios": con,
            "wdo_var_pct": round(wdo_var, 2) if wdo_var is not None else None}


def vies_consolidado(alinhamento: Optional[dict],
                     casado: Optional[dict] = None) -> Optional[dict]:
    """Viés de mercado CALCULADO (nao pelo modelo): concordancia entre o
    pano de fundo macro (`alinhamento_com_wdo`, ja calculado por
    `_alinhamento_macro`) e o proprio movimento do WDO no dia (embutido no
    `alinhamento`, campo wdo_var_pct).

    Isso da pra IA um numero pronto em vez de pedir pra ela "sentir" o
    consenso a cada chamada (mais assertivo/consistente entre leituras) e
    cria um sinal DETERMINISTICO, separado do texto livre do modelo, que da
    pra logar em `leituras` e medir taxa de acerto sozinho
    (`backtest_confluencia.py`).

    votos: +1 macro alinhado com compra / WDO subindo / casado esticado,
           -1 o espelho. Sinais "misto"/"neutro"/ausentes nao votam.
           (O Monitor WIN somava um 3o voto de blue chips - aqui a 3a vaga
           e' do casado, que so vota quando o desvio esta claramente
           esticado/atrasado; ver casado_wdo.voto_vies.)
    """
    votos = []
    if alinhamento and alinhamento["direcao"] in ("alinhado_compra", "alinhado_venda"):
        votos.append(1 if alinhamento["direcao"] == "alinhado_compra" else -1)
    wdo_var = alinhamento.get("wdo_var_pct") if alinhamento else None
    if wdo_var is not None:
        if wdo_var > 0.05:
            votos.append(1)
        elif wdo_var < -0.05:
            votos.append(-1)
    voto_casado = casado_wdo.voto_vies(casado)
    if voto_casado is not None:
        votos.append(voto_casado)
    if not votos:
        return None
    compra, venda = votos.count(1), votos.count(-1)
    if compra > venda:
        vies, forca = "compra", compra
    elif venda > compra:
        vies, forca = "venda", venda
    else:
        vies, forca = "misto", max(compra, venda)
    return {
        "vies": vies, "forca": forca, "total_sinais": len(votos),
        "macro": alinhamento["direcao"] if alinhamento else None,
        "wdo_var_pct": wdo_var,
        "casado": (casado.get("rotulo") if voto_casado is not None else None),
    }


def montar_contexto(tick, fluxo: Optional[dict], ranking_dados: dict,
                     niveis: dict, eventos_hoje: list,
                     ohlc_hoje: Optional[dict],
                     macro: Optional[dict] = None,
                     casado: Optional[dict] = None) -> dict:
    """Resume o estado atual do monitor num JSON compacto - so o que o
    modelo precisa, para nao gastar tokens/cota a toa."""
    ctx: dict = {"hora": time.strftime("%H:%M:%S")}
    if tick:
        ctx["tick"] = {
            "ultimo": tick.ultimo, "abertura": tick.abertura,
            "maxima": tick.maxima, "minima": tick.minima,
            "agr_compra": tick.agr_compra, "agr_venda": tick.agr_venda,
            "delta_acumulado_dia": tick.delta, "vwap": tick.vwap,
        }
    if fluxo:
        ctx["fluxo"] = {k: v for k, v in fluxo.items()
                        if k not in ("novos", "janela_perdida")}
    if ranking_dados and ranking_dados.get("corretoras"):
        ctx["ranking"] = {
            "top5": ranking_dados["corretoras"][:5],
            "top5_janela": ranking_dados.get("corretoras_janela", [])[:5],
            "ciclos_perdidos": ranking_dados.get("ciclos_perdidos", 0),
        }
    if niveis:
        ctx["niveis"] = {
            "zona_decisiva": niveis.get("zona_decisiva"),
            "resistencias": niveis.get("resistencias"),
            "suportes": niveis.get("suportes"),
            "fonte_pivots": niveis.get("fonte"),   # ja carrega "(cobertura parcial)" se for o caso
        }
    if eventos_hoje:
        ctx["eventos_recentes"] = eventos_hoje[-8:]
    if ohlc_hoje:
        ctx["ohlc_hoje"] = {k: ohlc_hoje.get(k) for k in
                            ("abertura", "maxima", "minima", "valido")
                            if k in ohlc_hoje}
    alinhamento = None
    if macro:
        ctx["macro"] = {
            "brent_var_pct": (macro.get("brent") or {}).get("var_pct"),
            "dxy_var_pct": (macro.get("dxy") or {}).get("var_pct"),
            "ouro_var_pct": (macro.get("ouro") or {}).get("var_pct"),
            "juros_us_var_bps": (macro.get("juros_us") or {}).get("var_bps"),
        }
        alinhamento = _alinhamento_macro(macro, tick)
        if alinhamento:
            ctx["macro"]["alinhamento_com_wdo"] = alinhamento
    if casado:
        ctx["casado"] = {k: casado.get(k) for k in (
            "diferencial", "preco_justo", "desvio", "desvio_z", "rotulo",
            "du", "cupom_impl_pct", "idade_s", "fonte_carrego") if k in casado}
    vm = vies_consolidado(alinhamento, casado)
    if vm:
        ctx["vies_mercado"] = vm
    return ctx


def _normalizar(bruto: dict) -> dict:
    """O schema e reforcado pelo response_format, mas nao custa nada
    blindar contra chave ausente/fora do enum em vez de quebrar o
    dashboard."""
    def _lista(v) -> list:
        if isinstance(v, list):
            return [str(x) for x in v if x is not None]
        return [str(v)] if v else []

    vies = bruto.get("vies")
    if vies not in ("compra", "venda", "misto", "indefinido"):
        vies = "indefinido"
    return {
        "vies": vies,
        "resumo": str(bruto.get("resumo") or "O modelo nao devolveu um resumo valido."),
        "evidencias": _lista(bruto.get("evidencias")),
        "alertas": _lista(bruto.get("alertas")),
        "ressalvas": _lista(bruto.get("ressalvas")),
    }


def _extrair_texto(body: dict) -> str:
    """A Interactions API devolve o texto dentro de steps[].content[] -
    'output_text' e uma conveniencia adicionada pelos SDKs oficiais, nao
    existe na resposta REST crua."""
    texto = ""
    for step in body.get("steps") or []:
        if step.get("type") != "model_output":
            continue
        for bloco in step.get("content") or []:
            if bloco.get("type") == "text":
                texto += bloco.get("text", "")
    return texto


def gerar_leitura(contexto: dict, forcar: bool = False) -> dict:
    """Chama o Gemini com o contexto e devolve a leitura estruturada.

    Cacheia por CACHE_TTL_S para o botao do dashboard nao disparar uma
    chamada nova a cada clique repetido (poupa cota do nivel gratuito).
    `forcar=True` ignora o cache. Levanta RuntimeError com mensagem
    apresentavel ao usuario em caso de chave ausente, limite de cota
    (429) ou resposta invalida - o endpoint converte isso em HTTP 503/502.
    """
    now = time.time()
    if not forcar and _cache["leitura"] and now - _cache["ts"] < CACHE_TTL_S:
        out = dict(_cache["leitura"])
        out["cache"] = True
        return out
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError(
            "GEMINI_API_KEY nao configurada no ambiente do servidor - "
            "crie uma chave em aistudio.google.com e defina a variavel.")

    payload = {
        "model": GEMINI_MODEL,
        "system_instruction": SYSTEM_PROMPT,
        "input": ("Contexto atual do WDO (JSON):\n"
                 + json.dumps(contexto, ensure_ascii=False)),
        "response_format": {
            "type": "text",
            "mime_type": "application/json",
            "schema": OUTPUT_SCHEMA,
        },
        "generation_config": {"max_output_tokens": MAX_OUTPUT_TOKENS},
    }
    headers = {"x-goog-api-key": api_key, "Content-Type": "application/json"}
    try:
        r = httpx.post(GEMINI_URL, json=payload, headers=headers, timeout=TIMEOUT_S)
        r.raise_for_status()
    except httpx.ConnectError:
        raise RuntimeError("Nao foi possivel conectar a API do Gemini (sem internet?).")
    except httpx.TimeoutException:
        raise RuntimeError(f"Gemini demorou mais de {TIMEOUT_S:.0f}s para responder.")
    except httpx.HTTPStatusError as e:
        detalhe = str(e)
        try:
            detalhe = e.response.json().get("error", {}).get("message", detalhe)
        except Exception:
            pass
        if e.response.status_code == 429:
            raise RuntimeError(
                "Limite de cota do nivel gratuito do Gemini atingido - "
                "aguarde um pouco e tente de novo.")
        if e.response.status_code in (401, 403):
            raise RuntimeError(f"Chave do Gemini invalida ou sem permissao: {detalhe}")
        raise RuntimeError(f"Erro do Gemini ({e.response.status_code}): {detalhe}")

    body = r.json()
    if body.get("status") == "failed":
        raise RuntimeError(f"Gemini falhou ao gerar a leitura: {body.get('status')}")

    texto = _extrair_texto(body)
    if not texto:
        raise RuntimeError(f"Resposta vazia do Gemini (status={body.get('status')}).")
    try:
        bruto = json.loads(texto)
    except json.JSONDecodeError:
        bruto = {"resumo": texto.strip()[:600], "ressalvas": [
            "O modelo nao devolveu o formato JSON esperado - leitura pode estar incompleta."]}

    leitura = _normalizar(bruto)
    leitura["gerado_em"] = time.strftime("%H:%M:%S")
    leitura["cache"] = False
    leitura["modelo"] = GEMINI_MODEL
    _cache["ts"], _cache["leitura"] = now, leitura
    return leitura
