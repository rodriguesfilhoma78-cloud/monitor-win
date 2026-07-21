"""
================================================================
 MONITOR WIN - agente_win.py
 Agente de leitura de fluxo (IA) - CONSUMIDOR, nao fonte de dados
----------------------------------------------------------------
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

GEMINI_URL   = "https://generativelanguage.googleapis.com/v1beta/interactions"
# flash-lite: mais rapido e mais barato em cota do nivel gratuito. Troque via
# variavel de ambiente se quiser mais qualidade (ex.: setx GEMINI_MODEL "gemini-3.6-flash").
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-3.5-flash-lite")
TIMEOUT_S    = 30.0
CACHE_TTL_S  = 45.0       # nao gera leitura nova a cada clique repetido do botao
MAX_OUTPUT_TOKENS = 900

SYSTEM_PROMPT = """Voce e um leitor de fluxo para um day trader de WIN (mini-indice B3, Profit Pro).

Seu unico papel e LER o que os dados mostram e ALERTAR o que se destaca,
sempre citando o numero exato que sustenta cada afirmacao. Voce NAO
recomenda comprar, vender, entrar ou sair - quem decide e o trader. Nunca
escreva algo como "entre agora" ou "hora de vender". Se o contexto nao
sustentar uma leitura clara, diga isso em vez de forcar uma conclusao.

Regras de leitura especificas deste sistema (nao invente numeros fora delas):
- delta_acumulado_dia (agr_compra - agr_venda do tick) e ACUMULADO DO
  PREGAO INTEIRO, na escala de milhares/milhoes de contratos do WIN. NAO
  compare essa escala com o delta_ema dos eventos de confluencia/divergencia,
  que e o fluxo INCREMENTAL suavizado (grandeza muito menor).
- O ranking de corretoras (campo ranking.top5) e um PISO do dia, nao o total
  real: se ciclos_perdidos > 0, parte da fita passou sem ser vista. Corretora
  agressora NAO e posicao (uma corretora vendendo pode ser 200 clientes
  diferentes) - o sinal util e dominancia e lote_medio (institucional tende
  a lote maior que varejo), nao o saldo absoluto.
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

Responda em portugues do Brasil, direto, sem jargao redundante. Toda
afirmacao de vies ou alerta deve ter pelo menos um numero do contexto entre
parenteses."""

OUTPUT_SCHEMA = {
    "type": "object",
    "properties": {
        "vies": {
            "type": "string",
            "enum": ["compra", "venda", "misto", "indefinido"],
        },
        "resumo": {
            "type": "string",
            "description": "1-2 frases com a leitura principal do momento",
        },
        "evidencias": {
            "type": "array",
            "items": {"type": "string"},
            "description": "cada item cita um numero concreto do contexto "
                           "(livro, fita, ranking, delta) que sustenta a leitura",
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


def montar_contexto(tick, fluxo: Optional[dict], ranking_dados: dict,
                     niveis: dict, eventos_hoje: list,
                     ohlc_hoje: Optional[dict]) -> dict:
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
        "input": ("Contexto atual do WIN (JSON):\n"
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
