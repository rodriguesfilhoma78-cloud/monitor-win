# Changelog — Monitor WIN

Histórico do que foi feito, mais recente primeiro.
Notas de sessão detalhadas: vault Obsidian `cerebelo\Day trade`.

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
