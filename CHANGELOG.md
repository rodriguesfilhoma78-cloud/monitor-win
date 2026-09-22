# Changelog — Monitor WIN

> [!important] Esta pasta é o **Monitor WDO**, forkada do Monitor WIN em
> 2026-08-24 (adaptação de scripts para o DOLFUT: ativo, tick de 0,5 ponto,
> porta 8003, blue chips removido — ver `README.md`). Este changelog é o
> histórico ORIGINAL do Monitor WIN, herdado como referência de como cada
> peça do código evoluiu; novas entradas específicas do WDO vão no topo,
> abaixo desta nota.

Histórico do que foi feito, mais recente primeiro.
Notas de sessão detalhadas: `Sessão 2026-08-24 — Criação do Monitor WDO.md`
na pasta `Day trade` (vault Obsidian).

## 2026-09-22 — Mini Indice (WIN) passa a usar preco real via RTD/Excel

A pedido do usuario, trocado o proxy `^BVSP` (Ibovespa a vista, Yahoo) pelo
preco REAL do futuro WIN (WINFUTV), lido via RTD do Profit Pro na mesma
planilha `WDO_Master_RTD.xlsm`. A linha WINFUTV ja existia na aba DADOS com
RTD completo (mesmo layout de DOLFUT); so faltava exportar.

- `ExportarWDO.bas` (Modulo3 ao vivo): nova `ExportarMacroRTDWDO()`, chamada
  de dentro de `ExportarWDO()`, escreve `dados_macro_rtd.csv` com DI1*/DOLFUT/
  WINFUTV (colunas fixas D=ultimo, H=fec_ant, J=volume). Corrige de quebra um
  bug preexistente: esse CSV vinha sendo escrito pelo Modulo1 (WIN) na pasta
  `monitor_win\`, NAO na pasta `Dolar monitor\` que o `server_wdo.py` le -
  o arquivo certo ficou parado desde 24/08/2026 e o cupom implicito do card
  CASADO rodava com DI desatualizado havia quase um mes sem erro visivel.
- `server_wdo.py`: `MacroRtdReader` ganhou `mini_indice()` (preco/fech_ant/
  var_pct/ts a partir da linha WINFUTV), no mesmo formato do
  `MacroFetcher.fetch_symbol()`. `MACRO_SYMBOLS` perdeu a entrada
  `"mini_indice": "^BVSP"` (nao busca mais no Yahoo); `macro_loop()` usa
  `macro_rtd.mini_indice()` no lugar de `quotes.get("mini_indice")`.
- `dashboard_wdo.html` / `agente_wdo.py`: comentarios e a legenda do card
  (`WINFUT · RTD` no lugar de `^BVSP · Yahoo`) atualizados; polaridade
  (indice sobe = contrario ao dolar) nao mudou, so a fonte do preco.
- **Pendente, NAO aplicado ao vivo**: o `IniciarServidorWDO()` do mesmo
  modulo aponta pra porta 8002 (devia ser 8003) e usa `pythonw.exe` cru (via
  PATH, que e' o stub da Microsoft Store e falha em silencio via automacao
  COM) em vez do caminho real do interpretador. Uma tentativa de corrigir os
  dois no mesmo modulo, na mesma sessao, quebrou a compilacao do VBA ao vivo
  por um motivo nao diagnosticado (o bloco sozinho compila limpo isolado,
  mas nao dentro deste modulo) e derrubou o Excel (crash + autorecover,
  ~15-30min de RTD parado nos dois monitores ate reabrir). Revertido de
  proposito - ver comentario detalhado no `ExportarWDO.bas`.

## 2026-09-21 — Mini Indice (WIN) adicionado ao card MACRO

A pedido do usuario, quinta perna do card MACRO: Mini Indice, usando `^BVSP`
(Ibovespa a vista via Yahoo) como proxy - nao ha simbolo Yahoo pro futuro WIN
em si. Polaridade CONTRARIA ao DOLFUT, mesmo sentido do Brent: indice sobe =
apetite a risco atrai fluxo pro Brasil, BRL se fortalece = seta
vermelha/contraria; indice cai = seta verde/favoravel.

- `MACRO_SYMBOLS` (`server_wdo.py`) ganhou `"mini_indice": "^BVSP"`; migracao
  de schema adiciona `mini_indice`/`mini_indice_var` em `macro_snapshots`;
  `save_macro_sync()` e `macro_loop()` persistem e transmitem o novo campo.
- `dashboard_wdo.html` (`renderMacro()`): nova linha do card com a mesma
  logica de seta invertida do Brent (`s = v>0 ? -1 : 1`).
- `agente_wdo.py`: `_alinhamento_macro()` inclui o novo sinal na votacao de
  alinhamento macro x WDO; `montar_contexto()` expoe `mini_indice_var_pct`;
  prompt da IA documenta a polaridade pra nao contradizer o card.
- Auditoria (mesma sessao, a pedido do usuario): confirmado que a IA usa
  fluxo completo (livro+fita+VAP+ranking), macro (futuro) + casado (a vista)
  e o `vies_mercado` calculado, e que `evidencias`/`alertas`/`ressalvas` estao
  sendo persistidos de verdade em `leituras` (257 linhas reais checadas no
  banco). `backtest_confluencia.py` rodado de novo: ainda 0 eventos de
  `confluencia` (o fix de `ACCEL_TOLERANCE` desta mesma sessao ainda nao teve
  pregao pra gerar dado novo) e acuracia direcional seguindo perto de
  cara-ou-coroa (43-51% conforme o sinal e horizonte) - ver nota de sessao no
  vault Obsidian pros numeros completos.

## 2026-09-21 — Instrumentacao: detecta T&T0 (Fita) pausado no Profit Pro

Investigado por que `corretoras`/`corretoras_serie` ficaram sem gravar nenhuma
linha de 16 a 20/09: com o pregao aberto, comparando `fluxo` campo a campo,
BOOK0 (bid/ask) e VAP0 (POC/VAH/VAL) estavam normais o dia inteiro, mas
`amostra_negocios` (T&T0) ficou zerado das 09:09 as 14:07 - o `dados_tt.csv`
mostrava `RTD Pausado - Selecione uma das abas linkadas` no lugar de negocios.
Causa: o Profit Pro so atualiza a Fita/T&T0 enquanto aquela aba especifica
esta ativa/selecionada - Livro e VAP nao tem essa restricao, por isso so o
ranking de corretoras (que depende 100% do T&T0) apaga, sem nenhum erro no
resto do sistema.

- `FluxoReader.ler()` (`server_wdo.py`) ganhou deteccao: `tt_pausado` = true
  quando a janela do T&T0 tem linhas mas nenhuma virou negocio valido (o
  sinal do estado "RTD Pausado", diferente de arquivo vazio/pregao fechado).
  Novo campo `tt_pausado_desde` (hora do 1o ciclo pausado, reseta quando volta).
- Log (`[WDO] T&T0 (Fita) RTD pausado ha Xs - selecione a aba...`) so depois
  de `TT_PAUSADO_AVISO_S=180s` pausado, pra nao logar blip de 1-2 ciclos; loga
  a recuperacao tambem, com a duracao total do episodio.
- `tt_pausado`/`tt_pausado_desde` vao no broadcast `evento: fluxo` (todo ciclo,
  ~2s) - NAO no `evento: ranking`, que so dispara com `novos != []` e ficaria
  mudo bem na hora que mais importa avisar.
- `dashboard_wdo.html` ainda nao tem nenhum painel de corretoras/fluxo bruto
  (so `/ranking` e o WS existem, sem UI) - o aviso fica disponivel pra quem
  consumir o WS/`/ranking`, mas nao aparece visualmente no dashboard ainda.
- Validado isolado (sem restart do servidor ao vivo): leitura real do CSV
  atual confirma `tt_pausado=false`; simulacao com CSV sintetico de linhas
  "RTD Pausado" confirma `tt_pausado=true` com `tt_pausado_desde` estavel
  entre leituras consecutivas.

## 2026-09-21 — ConfluenceEngine: relaxa condicao de "acelerando"

`backtest_confluencia.py` contra o `wdo_history.db` real (03/07 a 21/09) mostrou
0 eventos de `confluencia` contra 238 de `divergencia` - o motor nunca
confirmava um rompimento. Causa: `_flow_confirms` exigia `abs(flow_ema) >
abs(flow_ema_prev)` estrito, ou seja a EMA do fluxo tinha que estar CRESCENDO
no tick exato em que a persistencia (`PERSIST_TICKS=3`) fechava - coincidencia
rara demais pra acontecer na pratica.

- Nova constante `ACCEL_TOLERANCE = 0.8`: em vez de exigir crescimento estrito,
  so exige que a EMA esteja em pelo menos 80% do pico anterior (nao esta
  "claramente perdendo forca"). Permite confirmar confluencia logo apos o pico
  do impulso, nao so enquanto ainda esta subindo.
- Validado isolado (sem depender do pregao ao vivo): pico que desacelera pra
  85% do valor anterior agora confirma (antes era rejeitado mesmo bem acima do
  threshold); desaceleracao forte (75%) continua rejeitada; caso "ainda
  acelerando" continua confirmando como antes.
- Pendente: validar com dado real do proximo pregao (rodar
  `backtest_confluencia.py` de novo depois de alguns dias e conferir se
  `confluencia` passa a aparecer e se a taxa de acerto e melhor que os ~40-47%
  medidos so com `divergencia`).

## 2026-09-17 — Card MACRO: sinais de Ouro e Juros EUA invertidos

Decisão do usuário: inverter a polaridade das duas pernas no contexto macro
do dólar. Brent e DXY ficaram como estavam.

- **Ouro (GC=F)**: sobe = **FAVORÁVEL** ao DOLFUT (era contrário), cai =
  contrário.
- **Juros EUA (UST 10Y, ^TNX)**: sobem = **CONTRÁRIO** ao DOLFUT (era
  favorável), caem = favorável.
- Mexido em `dashboard_wdo.html` (`renderMacro()` — setas e selo de
  alinhamento) e `agente_wdo.py` (`_alinhamento_macro()` + texto do prompt
  que descreve a polaridade pra IA). Como o selo de alinhamento e o
  `vies_mercado` contam essas setas, o voto macro muda junto.
- Faixas mortas inalteradas (ouro 0,1%; juros 1 bp). Comentários de
  referência em `server_wdo.py` e `README.md` atualizados.

## 2026-09-13 — Janela móvel, série temporal e backtest de virada de lado por corretora

Nota de sessão: `Sessão 2026-09-13 — Janela móvel, série temporal e backtest de virada de lado por corretora.md`. Mesma implementação do Monitor WIN, espelhada aqui (`server_wdo.py`/`agente_wdo.py`).

- `RankingCorretoras` ganhou janela móvel de 30min (`JANELA_CORRETORAS_S`,
  poda incremental via deque) ao lado do acumulado do dia - `ranking_janela()`
  novo. WS/`GET /ranking` trazem `corretoras_janela`; contexto da IA
  (`agente_wdo.py`) ganhou `ranking.top5_janela`.
- Nova tabela `corretoras_serie` (SQLite): snapshot do cumulativo por
  corretora a cada 60s (`SERIE_CORRETORAS_S`) - permite reconstruir qualquer
  janela depois do pregão por diferença entre dois pontos.
- `backtest_corretoras.py` (novo, standalone): lê `corretoras_serie` e aponta
  viradas de lado (cruzamento de sinal do saldo da janela) e divergência
  entre saldo do dia e saldo da última janela.
- Testado só com dados sintéticos (fim de semana, sem fita real) - pendente
  validar no pregão de 2026-09-14.

## 2026-08-27 — Card CASADO: preço justo do WDO vs dólar à vista

Caminho A (sem mexer na planilha), decisão do usuário. Objetivo: dar uma
leitura de **de que lado está o fluxo** comparando o futuro com o dólar à
vista — não é gatilho de entrada.

- **`casado_wdo.py`** (novo): fórmula do carrego (`preço_justo = pronto +
  carrego`, `diferencial = WDO − pronto`, `desvio = WDO − preço_justo`),
  calendário B3 (feriados 2026–2027 hardcoded — manutenção anual),
  `vencimento_frente()` (1º dia útil do mês seguinte, trata a rolagem),
  `dias_uteis_ate()`, OLS `diferencial ~ du` (`regredir_carrego`),
  `calcular()` e `voto_vies()` (3º voto do `vies_mercado`, só quando
  `|z| ≥ 1,5`).
- **Dólar à vista**: `SpotFetcher` no `server_wdo.py` — AwesomeAPI
  (`economia.awesomeapi.com.br/json/last/USD-BRL`, dólar comercial, sem
  chave, cache ~1 min), fallback Yahoo `BRL=X`. Buscado no `macro_loop` a
  cada 30 s junto com o macro.
- **DI de volta**: `MacroRtdReader` relê `dados_macro_rtd.csv` (só o DI1
  de menor vencimento) — usado **apenas** pra decompor o carrego bruto e
  o cupom cambial implícito (exibição). Cupom cambial não tem fonte
  gratuita em tempo real (B3 só vende; brapi.dev é 15 min atrasado e só
  ações), por isso o carrego "justo" é **ajustado do histórico**:
  regressão diária sobre ≥ `MIN_DIAS_CALIB` (8) pregões da coluna nova.
  Antes disso, cai pro diferencial da abertura do dia (`fonte_carrego` no
  payload diz qual está valendo).
- **`macro_snapshots`** (SQLite): +8 colunas (`spot`, `spot_var`,
  `wdo_pts`, `diferencial`, `du`, `carrego_justo`, `desvio`,
  `cupom_impl`), migração automática no mesmo padrão das anteriores.
  `save_macro_sync` estendido; novos `casado_calib_sync` (OLS 1 ponto/dia)
  e `casado_desvio_hist_sync` (z-score por faixa de horário).
- **`dashboard_wdo.html`**: card **CASADO — FUTURO vs À VISTA** na coluna
  direita, abaixo do MACRO. Diferencial, preço justo, desvio (colorido),
  à vista + idade do dado, cupom implícito, e selo
  ESTICADO/NEUTRO/ATRASADO pelo z-score. Evento WS `macro` agora carrega
  `casado`.
- **`agente_wdo.py`**: bloco `casado` no contexto da IA + regra no
  `SYSTEM_PROMPT`; `vies_consolidado` ganhou o voto do casado (a vaga do
  voto de blue chips removido do fork).
- **Novo endpoint** `GET /casado`.
- **Ressalvas** (ver conversa 27/08): AwesomeAPI trava fora do horário de
  Londres/NY → parte do desvio pode ser só a vista correndo atrás
  (`idade_s` no payload); carrego fitado só vale depois de ~2 semanas;
  virada de vencimento, intervenção do BC e PTAX de fim de mês distorcem
  o diferencial.
- **Pendente**: `supabase_sync/sync_supabase.py` **não** foi alterado —
  as 8 colunas novas do `macro_snapshots` só replicam pro Supabase depois
  de `ALTER TABLE wdo_macro_snapshots ADD COLUMN ...` no remoto E incluir
  os nomes na lista `APPEND_ONLY["macro_snapshots"]` (TODO comentado no
  arquivo). Incluir sem a migration remota trava o sync inteiro.

## 2026-08-25 — Card MACRO: Dólar→Ouro e Juros DI→Juros EUA

- Trocados os dois cards redundantes/via RTD do painel MACRO por
  commodities/juros globais que correlacionam melhor com o DOLFUT
  (decisão do usuário):
  - **Dólar (USD/BRL via Yahoo, proxy do próprio DOLFUT via RTD)** →
    **Ouro (GC=F, Yahoo)**, com polaridade **invertida** igual ao Brent
    (ouro cai = favorável/seta verde, ouro sobe = contrário/seta
    vermelha) — ouro é cotado em USD e se move de forma inversa à força
    global do dólar.
  - **Juros DI (DI futuro via RTD/Excel)** → **Juros EUA (UST 10Y, ^TNX
    via Yahoo)**, mantendo a polaridade direta que o DI já tinha (sobe =
    favorável/seta verde) — juro americano mais alto atrai capital pros
    EUA e tende a fortalecer o dólar globalmente.
- `server_wdo.py`: removida a classe `RtdMacroReader` e a leitura de
  `dados_macro_rtd.csv` no `macro_loop` — os quatro símbolos do card
  MACRO (Brent, DXY, Ouro, Juros EUA) agora vêm todos do Yahoo Finance.
  A planilha pode continuar exportando o CSV sem problema, só não há
  mais consumidor Python dele.
- Migração automática das colunas do `macro_snapshots` (SQLite):
  `dolar`/`dolar_var`→`ouro`/`ouro_var`, `di`/`di_var_bps`→
  `juros_us`/`juros_us_var_bps` (mesmo padrão usado pra `sp500`→`brent`
  em 24/08). `supabase_sync/sync_supabase.py` atualizado pra sincronizar
  as colunas novas — a tabela remota `wdo_macro_snapshots` no Supabase
  ainda precisa da migração equivalente antes do próximo sync rodar sem
  erro.
- `dashboard_wdo.html` e `agente_wdo.py` (contexto da IA e
  `_alinhamento_macro`) atualizados pra usar as chaves/rótulos novos.

## 2026-08-24 — Monitor WDO criado (fork do WIN)

- Cópia integral de `monitor_win` → `Dolar monitor`, adaptada pro DOLFUT:
  ativo (`WINFUTV`→`DOLFUT`), tick (5→0,5 ponto), porta (8001→**8003**,
  8002 já era da Agenda Econômica), painel Blue Chips removido, histórico
  (`wdo_history.db`/`niveis.json`) zerado de propósito.
- Fluxo (livro/fita/VAP): não foram criadas janelas novas no Profit — o
  usuário reapontou as janelas existentes (`BOOK0`/`T&T0`/`VAP0`) do
  WINFUTV pro DOLPRO. **Efeito colateral: o Monitor WIN ficou sem fluxo**
  até alguém abrir janelas novas e separadas pra ele.
- Card MACRO recalibrado pro dólar: Brent (BZ=F) no lugar do S&P 500,
  polaridade invertida (alta=contrário/vermelho); DXY/Dólar/DI com
  polaridade direta (positivo=favorável/verde) — oposto da lógica
  herdada do WIN/Ibovespa.
- Bug corrigido: `/niveis` vazio (`{}`, dia 1 sem OHLC anterior) travava
  a régua/últimos ticks/plano do dia em silêncio — `loadNiveis()` agora
  trata como "sem dado real" e usa o fallback.
- App PWA "Monitor WDO — Tempo Real" instalado no Edge (app-id
  `lnnhenchgamoiedplfmhhpgkoejjbdab`), auto-start no login, adicionado ao
  `INICIAR_TRADE.bat`.
- Sync pro Supabase: tabelas próprias `wdo_*` (mesmo projeto do WIN, sem
  misturar linha), tarefa agendada `MonitorWdoSupabase` (13:00 e 18:30).
- Gotcha de infra: `pythonw`/`python` no PATH resolvem pro alias da
  Microsoft Store (`WindowsApps\`), que falha em silêncio quando chamado
  via `WScript.Shell.Run` (automação COM do VBA) — `IniciarServidorWDO`
  agora usa o caminho real
  (`AppData\Local\Python\pythoncore-3.14-64\pythonw.exe`).

## 2026-08-22

### Leitura de IA calibrada por histórico real (`historico_similar`)
- Origem: análise do framework TradingAgents (Tauric Research) por pedido do
  usuário - a maior parte não se aplica (debate multiagente Bull/Bear/Risk
  pra decidir Comprar/Vender contradiz a decisão já tomada de o agente nunca
  recomendar entrada/saída; analistas de fundamentos/notícias/social não têm
  equivalente no WIN). A única ideia aproveitada: o padrão de puxar
  situações passadas parecidas antes de decidir, adaptado sem ChromaDB nem
  embeddings - só uma query no que já existe.
- `SnapshotDB.historico_similar_sync()` (novo, `server_win.py`): dado o
  `vies_mercado` calculado agora (compra/venda, força ≥2), busca as até 60
  leituras passadas mais recentes com o MESMO `vies_mercado`/força na tabela
  `leituras`, e mede se o preço bateu a direção 15min depois via `snapshots`
  - mesma lógica de `backtest_confluencia.backtest_leituras()`, mas sob
  demanda e só para o caso atual, não recalcula nada que já roda. Puramente
  leitura (nenhum INSERT/UPDATE), protegido por try/except em
  `_contexto_leitura_atual()` para nunca derrubar `/leitura` se a query
  falhar (ex.: lock do SQLite durante escrita concorrente do poller).
- `agente_win.py`: novo campo `historico_similar` no contexto (quando
  presente) e nova regra no `SYSTEM_PROMPT` explicando que é calibração de
  confiança medida (`acerto_pct` sobre `n` casos), não uma fonte nova de
  viés - não muda `OUTPUT_SCHEMA` nem a lógica de `montar_contexto()`.
- Testado contra o `win_history.db` real (só leitura, contagem de
  `leituras`/`snapshots` conferida idêntica antes/depois): 140 leituras
  acumuladas desde 11/08, `historico_similar_sync("compra")` → 58 casos
  comparáveis / 46,6% de acerto em 15min; `"venda"` → 48 casos / 45,8%.
  Wiring ponta a ponta testado a frio (sem tick/macro ao vivo - fim de
  semana, servidor parado) e com `vies_mercado` forçado para simular o
  caminho quente; falha simulada de banco (`database is locked`) confirmada
  como não-fatal. Não testado contra a API real do Gemini (sem
  `GEMINI_API_KEY` neste ambiente) - só valida de fato no próximo pregão.

## 2026-08-13

### Cutoff de abertura do pivô automático relaxado para 10h
- `PREGAO_ABERTURA_ATE` (`server_win.py`) subiu de `"09:30"` para `"10:00"`:
  o pivô automático travava em "cobertura parcial" nos dias em que o
  Excel/Profit só começava a mandar tick um pouco depois das 09:30. Esse
  limite também é reusado em `dia_valido()` e na busca do último dia válido
  como base de pivô, então o efeito é consistente em todo o cálculo.
- Muda só na próxima vez que o Monitor WIN for iniciado (server ficou rodando
  com o valor antigo em memória até então, por decisão do usuário).

## 2026-08-11

### Tabela `leituras` + backtest retroativo de assertividade
- Diagnóstico: a leitura de IA (Gemini) nunca era persistida no banco - só
  ia pro WS/dashboard e se perdia. Sem histórico não dava pra medir taxa de
  acerto nenhuma, muito menos evoluir/comparar. Nova tabela `leituras`
  (`server_win.py`, `SnapshotDB._init`) grava toda leitura (manual,
  periódica ou por confluência): `vies` do modelo, `vies_mercado` calculado,
  resumo/evidências/alertas/ressalvas, gatilho, preço no momento.
- `backtest_confluencia.py` (novo, standalone, sem dependências externas):
  cruza `eventos`/`leituras` com `snapshots` (preço 5/15/30min depois) pra
  medir se a direção apontada realmente se confirmou. Rodado contra os 231
  eventos já existentes (03/07 a hoje): confluência 44-89% de acerto
  conforme o horizonte (amostra pequena, 9 eventos), divergência 45-57%
  (222 eventos - perto de aleatório, esperado já que "divergência" é
  justamente ausência de confirmação de fluxo). Baseline honesta - "acerto"
  aqui é só o preço ter se movido na direção certa, não é resultado de
  operação real.
- Índices novos em `snapshots`, `fluxo`, `eventos`, `macro_snapshots` (por
  `dia`) - essas tabelas (88k/238k linhas) não tinham índice e cresciam sem
  ele; toda query `WHERE dia=?` (histórico, backtest) fazia table scan.

### Viés de mercado calculado (macro + blue chips + preço do WIN)
- `agente_win.vies_consolidado()`: combina `alinhamento_com_win` (macro),
  `blue_chips.vies_agregado` e o var% do próprio WIN num consenso
  determinístico (`vies` + `forca` de 0 a 3 = quantos concordam). Isso dá
  pra IA um número pronto em vez de "sentir" o consenso a cada chamada
  (mais consistente entre leituras) e cria um sinal separado do texto livre
  do modelo, medível sozinho no backtest. Prompt do Gemini atualizado pra
  citar `vies_mercado` e não contrariar força >=2 sem evidência concreta da
  fita/livro.
- Dashboard (`dashboard_win.html`, `renderConclusao()`): quando a leitura
  chega (automática por confluência ou manual), mostra um banner de
  "CONCLUSÃO" cruzando o viés do modelo com o `vies_mercado` calculado
  (concordam = confluente, divergem = alerta, força <2 = sem consenso),
  reaproveitando o estilo do banner macro (`.align-banner`) já existente.

## 2026-08-03

### Blue chips: Fluxo trocado de agressão (agr_compra/venda) para Var% de preço
- Usuário reportou PETR4 e BBAS3 com Var% negativo mas Fluxo "compra" verde.
  Diferente das recorrências de 30/31-07 (colunas erradas via RTD), dessa vez
  os números de agressão batiam certo em escala (ex.: PETR4 agr_compra
  4.199.100 vs agr_venda 3.399.300) — era uma divergência real entre duas
  métricas diferentes: fluxo de ordens do dia (agressão) vs. variação de
  preço vs. fechamento anterior.
- Decisão do usuário: Fluxo deve sempre bater com o sinal do Var% (preço
  caiu = venda, subiu = compra). `BlueChipsReader` (`server_win.py`) agora
  classifica `fluxo` por `var_pct` em vez de `agr_compra - agr_venda`
  (`VAR_MIN = 0.1%` de zona neutra). `dominancia` (só usada pra largura da
  barrinha visual) virou `abs(var_pct) * 6`. Leitura de agr_compra/venda no
  CSV continua existindo, só parou de alimentar essa classificação.
- Servidor reiniciado e conferido ao vivo no dashboard: as 5 blue chips
  batendo Fluxo com Var% (ex.: VALE3 -2,29% → venda, ITUB4 +0,67% → compra).

## 2026-07-31

### Blue chips: ColPorCampo portado do Módulo1 (WIN) — blindagem definitiva
- Depois da recorrência do bug abaixo (mesmo problema, 2 dias seguidos, com
  causas ligeiramente diferentes), portado o padrão `ColPorCampo` do
  `Módulo1`/WIN pro `ModuloBlueChips`: cada coluna (ULT/ABE/MAX/MIN/FEC/VOL/
  98=compra/99=venda/67=vwap) agora é localizada pela PRÓPRIA fórmula RTD da
  linha do ativo (`"TICKER_B_0","CODIGO"`) em vez de índice fixo, com a
  coluna atual como fallback só se a fórmula não for encontrada. Sufixo RTD
  das ações é `_B_0` (o do WIN é `_F_0` — futuro vs. ação). Testado nos 5
  ativos: valores continuam batendo (compra/venda na mesma escala, VWAP como
  preço real) e o painel no dashboard confere com a direção do preço.
  `ExportarBlueChips.bas` em disco e o módulo ao vivo no Excel sincronizados.

### Blue chips com viés de venda falso-positivo — RECORRÊNCIA da correção de 30/07
- O mesmo bug do dia anterior voltou a aparecer no painel (fluxo "venda"
  mesmo com preço subindo). Causa: a correção de 30/07 só tinha sido salva
  no arquivo `ExportarBlueChips.bas` em disco, nunca reimportada para o
  módulo VBA (`ModuloBlueChips`) que estava de fato rodando dentro do
  `WDO_Master_RTD.xlsm` já aberto — editar o `.bas` sozinho não é
  suficiente, o módulo ao vivo continua com o código antigo até ser
  reimportado/recolado no VBE.
- Corrigido de novo (mesmas colunas: 27=compra, 29=venda, 30=vwap,
  10=volume), desta vez aplicando a mutação diretamente no `CodeModule` do
  workbook aberto via COM (`DeleteLines` + `AddFromString`). Religados os 4
  loops do workbook (`IniciarExportWIN`, `IniciarExportPETR4`,
  `IniciarExportPetroVale`, `IniciarExportBlueChips`) logo em seguida, por
  causa do reset de variáveis de módulo que essa mutação causa (ver nota de
  30/07 abaixo). Conferido no dashboard: fluxo agora bate com a direção do
  preço (ex.: VALE3 caindo → venda, PETR4 subindo → compra).

## 2026-07-30

### Pivots presos no OHLC de 23/07
- `PREGAO_ABERTURA_ATE` relaxado de `"09:15"` para `"09:30"` — dias em que o
  Excel/Profit abriam um pouco depois desse corte (24, 27, 29/07) ficavam
  sempre descartados como base de pivots, caindo num pregão mais antigo que
  o necessário.

### Blue chips com viés de venda falso-positivo (5/5 sempre)
- `ModuloBlueChips` (VBA) lia `agr_compra`/`agr_venda` de colunas fixas
  (X24/Z26) que deslocaram quando a planilha ganhou colunas de MACD — a
  coluna 24 virou "Prior Cote" (preço) e a 26 virou "Saldo Acumulado de
  Agressão" (saldo, não volume). Corrigido para as colunas reais
  (27=compra, 29=venda, 30=vwap, 10=volume). `ExportarBlueChips.bas` e
  `ExportarWIN.bas` em disco resincronizados com o código real do Excel
  (estavam desatualizados e apontavam pra um caminho OneDrive inexistente).

### Leitura de IA (Gemini) ganha contexto macro + blue chips
- `agente_win.montar_contexto()` agora recebe `macro` (S&P500/DXY/dólar/DI)
  e `blue_chips`; calcula `alinhamento_com_win` com a MESMA regra do banner
  "MACRO — IMPACTO NO IBOV" do dashboard, pra IA não contradizer o que já
  aparece na tela. `LEITURA_AUTO_INTERVALO` subiu de 180s (3min) pra 900s
  (15min) — prompt maior, poupa cota do nível gratuito. Confluência continua
  disparando leitura na hora, sem esperar esse intervalo.

> Descoberta no meio do trabalho: a sessão vinha operando por engano numa
> cópia órfã duplicada do projeto (`Day trade\Day trade\monitor_win`, CSV
> congelado). Corrigido — servidor certo confirmado no ar. Nota detalhada:
> `Sessão 2026-07-30 — Pivots, blue chips e IA com macro`.

## 2026-07-24

### Leitura de IA automática (periódica) + texto compacto
- **Decisão do usuário**: o agente de IA deve rodar sozinho, com textos curtos
  e objetivos. A automação já era parcial (disparava em cada `confluencia`);
  faltava cobrir pregão parado e enxugar o texto.
- **Gatilho periódico** no `market_loop`: nova constante
  `LEITURA_AUTO_INTERVALO = 180` (~3 min, topo do `server_win.py`) +
  `last_leitura_auto`. Dispara `gerar_leitura_automatica({"tipo":"periodica"})`
  em background — respeita o lock `_leitura_auto_em_andamento` e o cache de 45s,
  então não empilha chamadas nem estoura cota. ~180 leituras/dia num pregão
  9h–18h: folga confortável no free tier do Gemini. Confluência segue disparando.
- **Modo compacto** (`agente_win.py`): `SYSTEM_PROMPT` com bloco "FORMATO
  COMPACTO" (resumo em 1 frase ~20 palavras, máx. 2 evidências, alertas/ressalvas
  só quando materiais); `maxItems:2` no schema; `MAX_OUTPUT_TOKENS` 900→320.
- **Dashboard**: `renderLeitura` distingue origem "confluência" vs "periódica";
  alerta no feed só na confluência (a periódica atualiza em silêncio, sem virar
  spam a cada 3 min).
- Filosofia mantida: leitura+alerta com evidências, nunca recomendação de
  entrada; nenhuma fonte nova. `py_compile` OK. **Vale no próximo restart do
  server.** Nota detalhada: `Sessão 2026-07-24 — Leitura de IA automática e compacta`.

## 2026-07-23

### Régua / Plano do dia / Distância aos níveis defasados — base dos pivôs 2 dias atrasada
- **Sintoma**: dashboard com `Pivots automáticos (OHLC 2026-07-21)` no dia 23/07;
  preço (178.350) acima de TODOS os níveis (R3 em 177.370 a −980 pts).
- **Causa raiz**: `ohlc_anterior` preferia o último dia `valido=1`. O pregão de
  22/07 ficou `valido=0` porque o server parou de capturar às 16:40 (antes de
  `PREGAO_FECHA_APOS=17:45`), então foi pulado e usou 21/07 — cujo range
  (173.5k–175.3k) não tinha relação com o mercado atual (22/07 fechou 178.780).
- **Correção imediata (ao vivo, sem reiniciar)**: 22/07 marcado `valido=1` no
  `win_history.db` (dados coerentes) + `POST /niveis/recalcular`. Níveis
  passaram para OHLC 22/07 (Pivot 177.350 · Zona 176.630–178.065 · R1 180.450
  … S3 170.900), coerentes com o preço.

### `ohlc_anterior` passa a usar o pregão mais recente (último tick salvo)
- Antes preferia o último dia coberto ponta a ponta (`valido=1`), o que deixava
  a base velha quando um dia recente não fechava 17:45. Agora usa o pregão
  **mais recente** cuja abertura foi capturada (`primeiro_ts <= 09:15`) e
  coerente (C dentro de [L,H]), com o **último tick salvo** como fechamento.
  Cobertura parcial (`valido=0`) não descarta mais o dia — só marca
  `cobertura_parcial` (dashboard avisa). Fallback: nenhum dia com abertura
  capturada → cai no mais recente coerente.

### Níveis auto recalculam ao reabrir o monitor se a base mudou
- `calcular_niveis_do_dia` grava `base_dia` no `niveis.json`.
  `atualizar_niveis_automaticos` (chamado no start do `market_loop`) refaz os
  níveis automáticos quando `base_dia` mudou, mesmo já sendo do dia. Níveis
  definidos MANUALMENTE hoje continuam preservados.

### FEC oficial acerta e estende dia parcial
- Ao abrir no dia seguinte, o Excel/RTD traz o FEC (fechamento oficial do
  pregão anterior, coluna 8 do `dados_win.csv`). `refinar_niveis_com_fec` (1º
  tick do dia) agora ramifica: dia **completo** → só troca o C (fora do H/L →
  marca inválido e refaz do próximo); dia **parcial** → adota o FEC como
  fechamento e **estende H/L** para incluí-lo (`ajustar_ohlc_parcial`) — máxima
  real ≥ FEC e mínima ≤ FEC, não fabrica dado. Régua, distância e plano saem do
  mesmo pacote de pivôs → atualizam juntos e vão pro dashboard via WS.
- **Validação**: FEC de 22/07 = 178.775 vs último tick salvo 178.780 (1 tick) —
  confirmou a abordagem "último tick salvo".

> Mudanças de código valem no próximo restart do server (o processo no ar é o
> código antigo, já com os níveis certos via recálculo ao vivo). Recomendação:
> deixar o server no ar até depois das 17:45 para o dia entrar como cobertura
> completa e dispensar a correção via FEC.

## 2026-07-22

### Correção crítica: dashboard "conectado" mas sem atualizar ("Aguardando ticks")
- **Schema drift da tabela `fluxo` matava o `market_loop`**. Todo tick, o
  `INSERT` em `salvar_fluxo_sync` referenciava colunas (`amostra_negocios`,
  `amostra_contratos`, `janela_s`, `negocios_s`, `contratos_s`) que a tabela
  existente não tinha (`CREATE TABLE IF NOT EXISTS` não altera tabela pronta; a
  migração dessas colunas ficou faltando, provável perda na reconstrução por
  bytecode de 21/07). A exceção derrubava a task do loop silenciosamente — só o
  tick inicial da conexão chegava, `macro` seguia por rodar em loop à parte.
  Fix: migração `ALTER TABLE fluxo ADD COLUMN` das colunas faltantes.
- **Blindagem**: corpo do `while` do `market_loop` agora em `try/except` — um
  erro de persistência/leitura secundária loga e segue, nunca mais congela o
  stream de ticks.

### Descasamento de pasta VBA↔server (pós-mudança p/ Day trade)
- O macro do Excel gravava os CSVs em `OneDrive\Apps\monitor_win` (caminho
  antigo) enquanto o server lia de `Day trade\monitor_win`. Caminhos do
  `ExportarWIN.bas`/`Modulo1` e `ExportarBlueChips.bas`/`ModuloBlueChips`
  repontuados p/ `Day trade`. (O WIN foi corrigido na macro viva; blue chips
  ver abaixo.)
- **Blue chips Var% zerada**: server lia arquivo velho (Day trade, de ontem)
  porque o `ModuloBlueChips` vivo seguia gravando em `Apps` (salvar o `Const`
  não recompila macro em execução via `OnTime`). `BlueChipsReader` agora lê do
  arquivo de **mtime mais recente** entre Day trade e Apps — funciona já e se
  auto-corrige no próximo restart limpo do Excel.

### Dashboard
- `LEITURA DO FLUXO (IA)` e `REGISTRO DO TRADER` movidos da coluna direita para
  o espaço vazio da coluna central (lado a lado). Direita fica com
  MACRO / DISTÂNCIA / PLANO.

## 2026-07-21

### Fluxo: livro de ofertas (BOOK0) e fita (T&T0) entram na esteira
- O inventário da planilha (`inventario_planilha.md`) achou três tópicos RTD
  já assinados e nunca exportados: `BOOK0` (livro, A22:H39, 18 níveis),
  `T&T0` (fita, J22:N42, 21 negócios com **milissegundo**) e `VAP0` (volume
  at price, W22:X70 — ainda não ligado; atenção: nesse os nomes dos campos
  são invertidos, `VOL` devolve preço e `PRC` devolve volume).
- `ExportarWIN.bas`: `ExportarFluxo` grava `dados_book.csv` e `dados_tt.csv`
  a cada ciclo. As âncoras dos blocos são localizadas pela fórmula do índice
  0 (`AlvoRTD`) e ficam em cache — varrer a planilha a cada 2s seria caro, e
  se o bloco mudar de lugar a leitura seguinte relocaliza sozinha.
- `server_win.py`: `FluxoReader` + tabela `fluxo` + `GET /fluxo` + evento WS
  `fluxo`. Do livro sai a pressão **parada** (melhor bid/ask, spread,
  profundidade, desequilíbrio); da fita sai a pressão **executada** (lote
  médio/máximo, ritmo, agressor inferido pelo lado do livro em que o negócio
  saiu).
- **A fita é AMOSTRA, não contagem.** Medido hoje: os 21 negócios da janela
  do RTD cobrem 0,56s de mediana (mínimo 0,09s) contra ciclo de export de 2s.
  Somar negócio a negócio entre leituras subestimaria o fluxo em ~70%
  justamente nos momentos rápidos. A saída correta é a **taxa**: com o span
  da janela (`janela_s`), `negocios_s` e `contratos_s` são estimativas
  não-enviesadas que não dependem de continuidade entre leituras. Contagem
  exata do dia continua vindo dos acumulados (`agr_compra`/`agr_venda`).
- `taxa_confiavel` marca janelas < 0,1s, em que a divisão amplifica ruído
  (21 negócios em 0,02s viram "875 neg/s" — real para o instante, mas não
  sustentável). Quem tirar média tem que ponderar por `janela_s`.
- 25 testes do `FluxoReader` (livro, taxa, rajada de mesmo carimbo, CSV vazio).

### Ponte Acum: o ranking DA PLANILHA passa a ser o do dia inteiro
- `RankingCorretoras.escrever_csv` espelha o acumulado em `ranking_acum.csv`
  (escrita atômica via .tmp + replace; números inteiros — `Val()` do VBA não
  lê vírgula BR). Primeira linha `# dia=...`: o VBA limpa a Acum em vez de
  colar dado de pregão anterior.
- VBA `AtualizarAcum` (a cada 15 ciclos ≈ 30s) cola o CSV na aba `Acum` —
  a fórmula DADOS!R21 prefere a Acum e troca sozinha da janela de ~60s para
  o acumulado do pregão. Silencioso em qualquer falha; roda manual por
  Alt+F8. Testado ponta a ponta ao vivo (fórmula virou para o acumulado).
- Até o próximo boot do server o CSV não é realimentado (o processo no ar é
  o código antigo) — a ponte fica plenamente ativa amanhã na largada.

### Ranking de corretoras compilado do dia inteiro
- Usuário criou na planilha: T&T0 ampliado para **500 negócios** (linhas
  20-521, RTD entrega até o índice 499 ≈ 60s de fita), bloco "Ranking de
  Corretoras" em `DADOS!R19:U40` (fórmula LET/FILTER/SORT) e aba **`Acum`**
  (vazia, só cabeçalhos). A fórmula prefere a `Acum`; sem ela, cai no cálculo
  sobre a janela viva — ou seja, o ranking da planilha mostra só ~60s.
- **Fix no VBA**: a altura dos blocos só era detectada ao relocalizar a
  âncora — estender o bloco com a âncora intacta deixava o export travado na
  altura velha (100 linhas de um bloco de 500). `AlturaOk` agora confere a
  borda do bloco a cada ciclo (2 células) e redetecta quando muda.
- `FluxoReader`: dedupe de negócios entre leituras voltou (`_negocios_novos`,
  assinatura de 6) — agora confiável porque a janela de 500 cobre ~60s contra
  ciclo de 2s. Campos internos `novos`/`janela_perdida` consumidos pelo loop.
- **`RankingCorretoras`**: acumula qtd/financeiro por corretora nos dois
  lados, o pregão inteiro. Tabela `corretoras` (upsert por dia+corretora),
  reload no boot (restart não zera), reset na virada de pregão.
  `GET /ranking` + evento WS `ranking` (top 12, a cada 15s).
- Validação ao vivo (45s): 439 negócios compilados, 0 janelas perdidas.
  XP +1000 contratos líquidos, BTG −596, Itaú −463 em só 10 negócios (lote
  médio 46 — padrão institucional).
- `ciclos_perdidos` expõe quando a fita transbordou sem ser vista: o
  acumulado é um **piso**, não o total exato da B3.
- 19 testes novos (dedupe, acumulador, upsert/reload, perda).

### Blocos ampliados na planilha + perfil de volume (VAP0) ligado
- Blocos estendidos na aba DADOS: **T&T0 de 21 → 100 negócios** (linhas
  22-121) e **VAP0 de 49 → 299 níveis** (linhas 22-320). BOOK0 segue com 18.
- Efeito medido na fita: a janela subiu de **0,56s para ~3,4s** de mediana —
  agora cobre com folga o ciclo de export de 2s, e as taxas ficaram estáveis
  (`taxa_confiavel=True` em todas as leituras da verificação).
- `ExportarWIN.bas`: as alturas dos blocos passaram a ser **detectadas**
  (`AlturaBloco`) em vez de constantes. Acrescentar linhas de RTD na planilha
  passa a funcionar sozinho, sem mexer no código.
- `ExportarVAP` grava `dados_vap.csv` (preço;volume).
- `FluxoReader._vap`: **POC** (preço de maior volume), **área de valor** (70%
  do volume, expandida a partir do POC pelo vizinho mais forte; em empate
  expande os dois lados para não deformar perfil simétrico), `vol_acima_pct`,
  `dist_poc` e `dentro_area`. Colunas novas na tabela `fluxo`.
- `_vol_abreviado`: o volume do VAP vem **abreviado** (`620`, `3,04k`,
  `1,2M`), misturando número puro com sufixo. O parser antigo devolveria
  vazio em tudo acima de mil. Perde-se resolução (~10 contratos no sufixo k),
  irrelevante para a forma do perfil.
- 25 testes novos (parser abreviado, POC, área de valor, casos degenerados).

### VWAP, volume e agressão furados desde 09/07 — módulo VBA desatualizado
- **Causa raiz**: o `ExportarWIN.bas` do disco ganhou a blindagem
  `ColPorCampo` em 18/07, mas o módulo **nunca foi importado de volta** para
  o `WDO_Master_RTD.xlsm`. A planilha seguiu rodando a versão com colunas
  fixas (`24/26/27/28`), e as colunas do RTD tinham mudado em ~09/07. O CSV
  saía com `volume`=campo 100 (delta), `agr_compra`=campo 86 (Prior Cote),
  `agr_venda`=campo 103 (delta) e `vwap`=campo 98 (agressão ≈ 950.000).
- **Correção**: Módulo1 substituído ao vivo via COM (`PararExportWIN` →
  troca → `IniciarExportWIN`), backup em `Modulo1.backup_20260721_095650.bas`.
  `ColPorCampo` confirmado resolvendo certo: ULT→4, VOL→10, 98→27, 99→29,
  67→30. VWAP voltou a ser preço (174.781 dentro do range do dia).
- **Blindagem** (`vwap_plausivel`): o VWAP só é gravado se cair dentro do
  range do dia (folga de 1%); fora disso vira `NULL` e loga alerta. Um campo
  RTD trocado não contamina mais o histórico em silêncio.

### Pivots calculados sobre pregão de cobertura parcial
- **Causa raiz**: `daily_ohlc` gravava H/L só da janela em que o server
  esteve no ar. Com `corrigir_fechamento` escrevendo o FEC oficial (do dia
  inteiro) num H/L parcial, o C caía fora do range — 07/07, 09/07 e 10/07
  ficaram com OHLC impossível e serviram de base para os pivots seguintes.
- `daily_ohlc` ganhou `primeiro_ts`, `ultimo_ts` e `valido`; `dia_coberto`
  exige primeiro tick até 09:15 e último após 17:45. `ohlc_anterior` prefere
  dias válidos e, se não houver nenhum, cai no mais recente coerente
  marcando `cobertura parcial` na `fonte` (o dashboard avisa em vez de ficar
  sem níveis). `corrigir_fechamento` recusa gravar C fora de [L,H] e marca o
  dia inválido; `refinar_niveis_com_fec` refaz os níveis a partir do último
  pregão confiável.
- `_recalcular_cobertura` roda a cada boot e reclassifica todos os dias pelos
  snapshots — dias gravados por versões antigas se corrigem sozinhos.
- Limpeza do `win_history.db` (backup `win_history.antes_correcao_*.db`):
  removidos 07/07, 09/07 e 10/07 (C fora do range) e 17/07 (OHLC duplicado
  de 18/07, sem snapshots); 2 `vwap` contaminados anulados; 146 snapshots
  órfãos sem data apagados.
- **Atenção**: nenhum dos 6 pregões restantes tem cobertura completa — o
  server só grava enquanto está no ar. Até fechar um pregão inteiro
  (09:00→18:00), os pivots seguem saindo com aviso de cobertura parcial.

## 2026-07-17

### Card MACRO — Brent, Dólar e Juros DI (impacto no IBOV)
- `server_win.py`: novo `macro_loop` paralelo — Brent (`BZ=F`) via Yahoo
  a cada 30s (mesmo desenho do monitor PETR4); DI futuro e DOLFUT via
  `dados_macro_rtd.csv` do Profit RTD (`RtdMacroReader`). O DI escolhido
  é o de maior volume (contrato mais líquido — acompanha a rolagem);
  Dólar prefere DOLFUT em tempo real (pontos/1000 = R$) com Yahoo
  `USDBRL=X` como fallback se o RTD parar (>180s). Evento WS `macro`,
  rota `GET /macro`, tabela `macro_snapshots` no SQLite (correlação
  macro × WIN).
- `ExportarWIN.bas`: novo `ExportarMacroRTD` — varre a coluna A da DADOS
  e exporta todas as linhas `DI1*` + `DOLFUT` (último, FEC, volume) a
  cada ciclo de 2s. Módulo1 do `WDO_Master_RTD.xlsm` atualizado via COM
  na própria sessão (PararExportWIN → substituição → IniciarExportWIN).
- `dashboard_win.html`: card "MACRO — IMPACTO NO IBOV" com Brent em
  destaque, linhas Brent/Dólar/DI com seta de impacto no índice
  (Brent↑, Dólar↓ e DI↓ favorecem) e banner de alinhamento
  favorável/contrário/misto com a variação do WIN (vs. FEC).

### Plano: hora REAL de ativação do gatilho (servidor)
- O `✅ ativado hh:mm` mostrava a hora em que a página abria e zerava no
  F5 (estado só no navegador). Agora o servidor calcula a hora do
  rompimento a partir dos snapshots do pregão (`plano_ativacao_sync`:
  1º snapshot em que a máxima/mínima cruzou a ZD↑/ZD↓), transmite via WS
  (`evento: plano_ativacao`), expõe em `GET /plano_ativacao` e envia no
  connect. Sobrevive a F5 e a reinício do server no mesmo dia; reflete a
  hora de mercado; recalcula sozinho se a zona decisiva mudar.
- Dashboard usa `S.planoAtiv` do servidor (removida a lógica local
  `trigC/trigV`); o simulador carimba localmente (sem servidor por trás).

### Macro: + DXY (dólar global) — estudo de correlação
- Estudo de 1 ano de retornos diários vs IBOV (`corr_macro.py`): DXY
  r=−0,36 (mais forte que USD/BRL −0,24) e quase independente dele
  (r≈0,18) → adiciona sinal, não repete. Brent −0,23 instável
  (validou a remoção); minério ~0 no diário (fica no Petro&Vale);
  VIX −0,39 e EEM +0,46 (semi-circular) como candidatos futuros.
- Card agora: S&P 500 · **DXY** · Dólar · DI. DXY via `DX-Y.NYB` (ICE),
  DXY↑ → contra o índice. Nova coluna `dxy/dxy_var` em `macro_snapshots`
  (migração `ALTER … ADD COLUMN`).

### Macro: Brent → S&P 500 (driver global do IBOV)
- Card MACRO troca o Brent pelo **S&P 500** (E-mini futuro `ES=F` via
  Yahoo — negocia quase 24h, o índice à vista `^GSPC` fica parado antes
  da abertura de NY). Mesma lógica: S&P↑ (risk-on) favorece o índice.
- `server_win.py`: `MACRO_SYMBOLS` sp500=ES=F; payload/WS/`GET /macro`
  usam a chave `sp500`; tabela `macro_snapshots` renomeia colunas
  `brent/brent_var` → `sp500/sp500_var` (migração `ALTER … RENAME`).
- `dashboard_win.html`: linha "🇺🇸 S&P 500 · ES=F · Yahoo".

### Fix: Blue Chips não atualizava (auto-start faltando)
- Causa raiz: o `Workbook_Open` (em EstaPastaDeTrabalho do
  `WDO_Master_RTD.xlsm`) iniciava WIN, PETR4 e PetroVale, mas **nunca
  chamava `IniciarExportBlueChips`** — o `dados_blue_chips.csv` ficava
  congelado no último pregão em que o export foi ligado à mão. O próprio
  cabeçalho do `ModuloBlueChips` já pedia essa linha; ela nunca foi
  adicionada. Corrigido via COM: `OnTime +16s "IniciarExportBlueChips"`.
- `ExportarBlueChips.bas`: módulo exportado para o disco (faltava no
  repo — só existia dentro da planilha). **Salvar o .xlsm** para o
  Workbook_Open persistir.
- `dashboard_win.html`: tabela Blue Chips ganhou espaçamento entre
  colunas (`.bctbl`, padding 8px + `white-space:nowrap` nos números) —
  na coluna central mais estreita os valores colavam (−2,35%11,2%).

### Selo de alinhamento no topo do card MACRO
- O preço grande do Brent saiu do cabeçalho do card (segue na linha
  dele); no lugar entrou o selo de viés estilo monitor PETR4:
  ✔ Alinhados COMPRA (verde) / ✔ Alinhados VENDA (vermelho) /
  ⚠ Divergentes / ◆ Vento macro misto — cruzando a maioria das setas
  macro com a direção do WIN (var vs FEC, faixa morta ±0,05%).

### Layout: Alertas e Blue Chips na coluna central
- Os cards Alertas e Blue Chips saíram da coluna direita para o espaço
  vazio abaixo dos stats na coluna central, lado a lado (`.row2`, grid
  2×1 que empilha abaixo de 900px). Coluna direita ficou: MACRO →
  Distância → Plano do dia.

### Início automático via .bat (redundância ao auto-start do VBA)
- `iniciar_win.bat`: mesmo padrão do `iniciar_petr4.bat` — testa a porta
  8001 e só sobe `python server_win.py` (console minimizado) se não houver
  server no ar. Idempotente: pode rodar quantas vezes quiser.
- `Monitor WIN — servidor.vbs` na pasta Inicializar do Windows
  (`shell:startup`): roda o .bat oculto no logon. Os atalhos que já
  existiam lá só abrem os dashboards (apps do Edge) — nenhum subia o
  servidor; cobria só o caminho VBA, que já falhou silencioso (09/07).

### Risco + flag nos níveis atingidos (padrão do monitor PETR4)
- "Distância aos níveis": nível tocado pela máxima/mínima do dia fica
  riscado com ✅ (R/S e bordas da zona decisiva).
- "Plano do dia": gatilho Compra/Venda ganha `✅ ativado hh:mm` quando a
  máxima/mínima cruza a zona decisiva (alvos já riscavam desde antes).

## 2026-07-03

### Início automático via Excel
- `ExportarWIN.bas`: novo `IniciarServidorWIN` — testa a porta 8001 (500ms);
  se não responde, sobe `pythonw server_win.py` sem janela de console.
  Nunca reinicia um server já no ar (preserva histórico intradiário).
- `IniciarExportWIN` (chamado pelo `Workbook_Open` do `WDO_Master_RTD.xlsm`)
  agora sobe servidor + exportação. Sequência do pregão: Profit → Excel
  com macros → ~10s → http://127.0.0.1:8001
- Cópia antiga do módulo no Desktop (`dashboard-wdo\monitor_win`) sincronizada.

### ConfluenceEngine reformulado (server_win.py)
- **Fluxo incremental**: EMA sobre `delta_t − delta_{t−1}` em vez do delta
  acumulado do dia (os agregados do RTD são cumulativos — o delta operava
  em dezenas de milhares contra um `MIN_DELTA = 50` fixo, que não filtrava nada).
- **Threshold adaptativo**: `K_STD (1.5) × desvio-padrão móvel` do fluxo
  (janela 300 amostras ≈ 10 min, aquecimento 30 ≈ 1 min).
- **Níveis unificados**: resistências + suportes + alvos + zona decisiva
  em conjunto único com dedupe por preço.
- `reset_flow()` na virada de pregão; `LevelStore` com cache por mtime.
- Teste: dia vendedor (acumulado −11k) com rajada compradora no R1 →
  confluência confirmada (fluxo +3.611 > threshold 2.794); engine antigo
  descartaria o sinal.

### Perfil de volume (régua por acúmulo)
- `SNAPSHOT_EVERY`: 30s → **2s**.
- Novo endpoint `GET /perfil/{dia}?bucket=50`: POC, value area (~70%),
  histograma por faixa de preço, `amostras` e `cobertura` (perfil de dia
  com coleta parcial é enviesado — não confiar sem checar a cobertura).

### Dashboard (dashboard_win.html)
- Régua: range efetivo que só **expande** quando o preço sai dele
  (marcador não gruda mais na borda em dia de tendência).
- Faixas de calor de volume na régua + linha POC (refresh 2 min).
- Distância aos níveis: inclui ZD↑/ZD↓ (zona decisiva) e destaca o
  nível mais próximo.
- Plano do dia dinâmico: alvo tocado pela máxima/mínima marcado com ✅,
  linhas de invalidação derivadas da zona decisiva.

> Mudanças de backend entram em vigor no próximo start do servidor
> (abertura de 04/07 via início automático). Dashboard vale com F5.

## 2026-07-02

- Sistema consertado e evoluído (níveis automáticos por pivots/CPR com
  refino pelo FEC oficial, histórico SQLite, eventos de confluência).
  Ver nota "Sessão 2026-07-02 — Monitor WIN consertado e evoluído".
