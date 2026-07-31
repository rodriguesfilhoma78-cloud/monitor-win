# Changelog — Monitor WIN

Histórico do que foi feito, mais recente primeiro.
Notas de sessão detalhadas: vault Obsidian `cerebelo\Day trade`.

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
