# Changelog — Monitor WIN

Histórico do que foi feito, mais recente primeiro.
Notas de sessão detalhadas: vault Obsidian `cerebelo\Day trade`.

## 2026-07-18

### Registro de operações do trader (fase 1 do aprendizado por imitação)
- Objetivo de longo prazo: dataset de "situação de mercado + decisão do
  trader + resultado" para descobrir os padrões por trás das decisões
  (primeiro estatística, depois talvez ML). Fase 1 = coleta.
- `server_win.py`: tabela `operacoes` no SQLite (entrada, saída,
  resultado em pts, `contexto` JSON com a FOTO do mercado no instante da
  entrada: tick + níveis + macro + blue chips + fluxo EMA + plano).
  Rotas: `POST /operacoes` (abre; preço/contexto capturados no servidor,
  uma operação por vez), `POST /operacoes/fechar` (fecha e calcula
  resultado — perdas gravadas igual: sem elas o dataset vira viés de
  sobrevivência), `GET /operacoes` (dia + aberta), `GET /operacoes/stats`
  (taxa de acerto, média e total por tipo+motivo). Evento WS
  `operacoes_atualizadas` sincroniza os dashboards.
- `frontend/`: card "REGISTRO DO TRADER" — select de motivo (rompimento,
  pullback, fluxo, reversão, vwap, outro) + botões ▲ COMPREI / ▼ VENDI;
  com operação aberta mostra painel com P&L ao vivo e ✖ SAÍ; lista das
  operações do dia com resultado. Registro bloqueado no modo simulador
  (dados sintéticos poluiriam o dataset).
- Nota livre opcional (`nota_entrada`/`nota_saida`, até 500 chars): o
  trader escreve o raciocínio da entrada e da saída com as próprias
  palavras; o campo troca de contexto conforme o estado (entrando/
  saindo) e as notas aparecem como 🗒 com tooltip na lista do dia.
  Insumo para análise por IA dos padrões de raciocínio × resultado.

### Reorganização de pastas + dashboard separado em HTML/CSS/JS
- Estrutura nova: `frontend/` (dashboard_win.html + style.css + app.js),
  `excel/` (módulos .bas), `docs/` (este changelog). `server_win.py` e os
  arquivos de runtime (CSVs, niveis.json, .db) seguem na RAIZ — é onde o
  VBA grava (caminho fixo nos .bas) e onde o server lê; não mover.
- `server_win.py`: `FRONTEND_DIR`, rotas `/style.css` e `/app.js`;
  HTML referencia por caminho relativo (modo `file://` preservado).
- `requirements.txt`: `uvicorn` → `uvicorn[standard]` — o uvicorn puro
  não tem lib de WebSocket; em máquina nova o `/ws` falhava silencioso
  ("Unsupported upgrade request") e o dashboard ficava em "reconectando".

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
