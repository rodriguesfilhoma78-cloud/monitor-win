# Monitor WDO — Sistema independente de monitoramento do Dólar Futuro

Fork do **Monitor WIN** (pasta `monitor_win`), adaptado para o **DOLFUT**
(mini-dólar). Mesma arquitetura, mesmo motor de confluência, mesmos
pivots automáticos — só o ativo, a porta e o tick mudaram. Roda em
paralelo com o Monitor WIN (porta **8001**) e com o sistema antigo
"Mapa de Tendência WDO" em `profit_wdo_map` (porta **8000**), na porta
**8003**.

## O que mudou em relação ao Monitor WIN

- **Ativo**: `WINFUTV` → `DOLFUT` (linha da planilha `DADOS` localizada
  automaticamente pela coluna A, como no WIN).
- **Tick**: 5 pontos (WIN) → **0,5 ponto** (WDO). Preços agora têm casa
  decimal em todo o pipeline (CSV, banco, dashboard).
- **Porta**: 8001 → **8003**.
- **Painel Blue Chips**: removido (era específico do Ibovespa/WIN).
- **Fluxo (livro/fita/perfil)**: lê os tópicos RTD **BOOK0/T&T0/VAP0** —
  os MESMOS que o WIN usava. Em 24/08/2026 o usuário reapontou essas
  janelas de Livro de Ofertas/Fita/VAP no Profit do WINFUTV para o
  DOLPRO (dólar): o tópico RTD não muda quando o ativo exibido na
  janela muda, então as fórmulas já existentes na planilha passaram a
  refletir o DOLFUT sozinhas, sem precisar colar nada de novo no Excel.
  **Consequência: o Monitor WIN fica sem fluxo (livro/fita/VAP) até
  alguém abrir janelas novas e separadas pra ele** — decisão consciente
  do usuário nesta sessão, não bug. Se precisar do fluxo do WIN de
  volta, abra uma janela nova de Livro/Fita/VAP (que vira BOOK1/T&T1/
  VAP1) e linke pro WIN.
- **Ranking de corretoras (aba da planilha)**: cola em **`AcumWDO`** em
  vez de `Acum` (que já é do WIN) — crie essa aba se quiser usar esse
  recurso; sem ela, a colagem falha em silêncio (mesmo comportamento do
  WIN quando a aba não existe).
- **Card MACRO**: recalibrado pro dólar em 24/08/2026 e novamente em
  25/08/2026. **Brent (BZ=F, Yahoo)** no lugar do S&P 500 — Brasil é
  exportador de petróleo, então o termo de troca move o câmbio: Brent
  sobe → BRL tende a se fortalecer → **contrário** ao dólar (seta
  vermelha); Brent cai → **favorável** (seta verde). **DXY** segue a
  lógica direta: positivo = favorável (seta verde). Em 25/08/2026 os dois
  cards redundantes/via RTD foram trocados: **Dólar (USD/BRL via
  Yahoo/DOLFUT RTD)** → **Ouro (GC=F, Yahoo)**; **Juros DI (RTD)** →
  **Juros EUA (UST 10Y, ^TNX via Yahoo)**. **Em 17/09/2026 os sinais
  desses dois cards foram INVERTIDOS a pedido do usuário**: agora **ouro
  sobe → favorável/seta verde**, ouro cai → contrário/seta vermelha; e
  **Juros EUA sobem → contrário/seta vermelha**, caem → favorável/seta
  verde (antes era o oposto: ouro seguia a polaridade invertida do Brent
  por ser cotado em USD, e o juro americano seguia a lógica direta do DI).
  Polaridade oposta à herdada do WIN/Ibovespa, já corrigida em
  `dashboard_wdo.html` (`renderMacro()`) e `agente_wdo.py`
  (`_alinhamento_macro()`).
  Isso também tirou a dependência do RTD/Excel (`dados_macro_rtd.csv`)
  para o card MACRO — a planilha pode continuar gerando o arquivo, só não
  há mais consumidor Python dele.
- **Card CASADO** (27/08/2026, `casado_wdo.py`): preço justo do WDO por
  **arbitragem contra o dólar à vista**. `diferencial = WDO − pronto` (em
  pontos), `preço justo = pronto + carrego`, `desvio = WDO − preço justo`.
  Dólar à vista via **AwesomeAPI** (`economia.awesomeapi.com.br`, dólar
  comercial, sem chave, cache ~1 min; fallback Yahoo `BRL=X`). O `du` até
  o vencimento da frente (1º dia útil do mês seguinte) e o cupom cambial
  implícito usam o DI curto lido de volta do `dados_macro_rtd.csv` — o
  card CASADO é o **único** consumidor Python desse CSV hoje; se o arquivo
  sumir, o casado roda sem o cupom. **Cupom cambial não tem fonte
  gratuita em tempo real**, então o carrego "justo" é **ajustado do
  histórico**: regressão diária `diferencial ~ du` sobre os últimos
  pregões (≥ 8 dias; antes disso cai pro diferencial da abertura do dia).
  O selo ESTICADO/ATRASADO vem do z-score do desvio vs. a mesma faixa de
  horário dos últimos pregões, e alimenta um 3º voto no `vies_mercado`
  (só quando `|z| ≥ 1,5`). Ressalvas: AwesomeAPI trava fora do horário de
  Londres/NY (campo `idade_s` no payload), o carrego fitado só vale depois
  de ~2 semanas de coluna nova, e virada de vencimento / intervenção do BC
  / PTAX de fim de mês distorcem o diferencial. Endpoint: `/casado`.
- **Dados históricos zerados**: `wdo_history.db`, `niveis.json` e os CSVs
  de dados foram limpos na criação desta cópia (continham dados do WIN,
  em outra escala de preço — deixá-los teria produzido pivots e níveis
  completamente errados no primeiro boot).

## Estrutura da pasta

```
Dolar monitor\
├── server_wdo.py        # Servidor FastAPI + WebSocket + API de níveis (porta 8003)
├── agente_wdo.py         # Leitura de fluxo por IA (Gemini)
├── dashboard_wdo.html    # Dashboard (servido pelo próprio server em /)
├── niveis.json           # Níveis do dia (editável — sem reiniciar o server)
├── ExportarWDO.bas       # Módulo VBA para o Excel (exporta dados_wdo.csv)
├── requirements.txt
├── dados_wdo.csv         # (gerado pelo VBA em tempo de execução)
└── wdo_history.db        # (gerado automaticamente — snapshots SQLite)
```

## Instalação (uma vez)

```bash
cd "C:\Users\rodri\Desktop\Day trade\Dolar monitor"
pip install -r requirements.txt
```

No Excel (`WDO_Master_RTD.xlsm`):
1. A linha do **DOLFUT** na aba DADOS já existe (localizada automaticamente
   pela coluna A, função `LinhaDoAtivoWDO` no `ModuloWDO`), com todas as
   colunas RTD (D=último, E=abertura, F=máxima, G=mínima, H=fec. anterior,
   J=volume, AA=agr.compra, AC=agr.venda, AD=VWAP) já formuladas.
2. Importe o módulo `ExportarWDO.bas` na planilha: `Alt+F11` → botão
   direito no projeto → Importar arquivo → `ExportarWDO.bas`.
3. Adicione a chamada `IniciarExportWDO` ao `Workbook_Open` (o mesmo
   evento que já chama `IniciarExportWIN`), ou rode manualmente via
   `Alt+F8` → `IniciarExportWDO`.
4. **Fluxo**: já está pronto — as janelas de Livro/Fita/VAP no Profit
   foram reapontadas do WINFUTV para o DOLPRO (dólar), e como o tópico
   RTD (`BOOK0`/`T&T0`/`VAP0`) não muda quando o ativo exibido muda, as
   fórmulas que já existiam na planilha (herdadas do WIN) já refletem o
   DOLFUT. Nada a fazer no Excel para isso funcionar. Se quiser o fluxo
   do WIN de volta, abra uma janela NOVA de Livro/Fita/VAP no Profit
   (vira `BOOK1`/`T&T1`/`VAP1`) e linke-a pro WIN.
5. **Ranking de corretoras (opcional)**: crie uma aba `AcumWDO` na
   planilha se quiser o ranking acumulado do dia colado automaticamente.
6. Confira o `CAMINHO_CSV` no topo do módulo se mudar a pasta.

## Sequência de inicialização (dia de pregão)

1. Abrir Profit Pro
2. Abrir Excel (`WDO_Master_RTD.xlsm`) e **habilitar macros** se solicitado —
   ~10s depois a exportação do WDO inicia sozinha **e o `server_wdo.py`
   sobe automaticamente em segundo plano** (sem janela), caso a porta
   8003 ainda não esteja respondendo. Um server já no ar nunca é
   reiniciado (preserva o histórico intradiário).
   (fallback manual: Alt+F8 → `IniciarExportWDO`)
3. Navegador: **http://127.0.0.1:8003**

> Assim como no WIN: **não reinicie o server após as 9h** se quiser
> preservar o histórico intradiário de ticks no dashboard aberto.

## Níveis do dia

**Automático (padrão):** na primeira inicialização do dia (ou na virada de
data com o server no ar), os níveis são recalculados por **pivot points
clássicos** sobre o OHLC do pregão anterior (gravado em `daily_ohlc` no
`wdo_history.db`):

```
P  = (H+L+C)/3
R1 = 2P-L      S1 = 2P-H
R2 = P+(H-L)   S2 = P-(H-L)
R3 = H+2(P-L)  S3 = L-2(H-P)
Zona decisiva = Central Pivot Range: entre BC=(H+L)/2 e TC=2P-BC
```

Tudo arredondado ao tick de **0,5 ponto** do WDO. A `fonte` no cabeçalho
do dashboard indica quando os níveis são automáticos.

O **C (fechamento)** usa o valor gravado no fim do pregão anterior; no
primeiro tick do dia o servidor confere contra o **FEC oficial** do RTD
(coluna H da planilha, exportada como `fec_ant` no CSV) e recalcula se
divergir — protege contra server desligado antes das 18h.

**Manual (override):** edite `niveis.json` e envie — enquanto o
`data_pregao` for o de hoje, o automático não sobrescreve:

```bash
curl -X POST http://127.0.0.1:8003/niveis -H "Content-Type: application/json" -d @niveis.json
```

**Voltar ao automático:** `curl -X POST http://127.0.0.1:8003/niveis/recalcular`

## Endpoints

| Rota      | Método | Descrição                                  |
|-----------|--------|--------------------------------------------|
| `/`       | GET    | Dashboard                                  |
| `/ws`     | WS     | Stream de ticks em tempo real              |
| `/niveis` | GET    | Níveis atuais (JSON)                       |
| `/niveis` | POST   | Atualiza níveis + notifica dashboards      |
| `/niveis/recalcular` | POST | Força recálculo pelos pivots do pregão anterior |
| `/historico/{dia}` | GET | Pacote do pregão (OHLC, níveis, sinais, snapshots) p/ análise |
| `/perfil/{dia}` | GET | Perfil de volume: POC, value area e histograma por preço (`?bucket=1.0`) |
| `/ultimo` | GET    | Último tick recebido (debug rápido)        |
| `/macro`  | GET    | Último pacote macro (Brent, DXY, Ouro, Juros EUA, casado) + idade |
| `/casado` | GET    | Preço justo do WDO vs dólar à vista: diferencial, carrego, desvio, z-score |

## Perfil de volume (régua por acúmulo)

Os snapshots (a cada **2s**) alimentam o perfil de volume do pregão:
o volume negociado entre snapshots consecutivos é atribuído à faixa de
preço (bucket de **1 ponto**, ajustado do 50 do WIN para a escala menor
do WDO) em que ocorreu. O dashboard pinta as faixas de maior acúmulo
como bandas azuis na régua + linha **POC** (preço com mais negócios),
atualizadas a cada 2 min. Pivots = projeção do dia anterior; acúmulo =
memória real de onde o mercado negociou — os dois convivem.

## Testar sem o pregão

Abra o dashboard e clique em **▶ Simulador** — gera ticks sintéticos
para validar alertas, régua e viés.

## O que NÃO foi tocado (compartilhado com o Monitor WIN)

- `dados_macro_rtd.csv` continua sendo mantido pelo `ExportarWIN.bas` do
  Monitor WIN (a lista DI1*+DOLFUT não mudou). O card MACRO não lê mais
  esse arquivo desde 25/08/2026, mas o **card CASADO voltou a lê-lo**
  desde 27/08/2026 — só o DI curto, só pra decompor o carrego e estimar o
  cupom implícito. Se o CSV envelhecer ou sumir, o casado roda sem o
  cupom (campo fica nulo); o resto do dashboard não é afetado.
- `azure_sync/` e `supabase_sync/` são cópias das pastas do Monitor WIN,
  não adaptadas — rodá-las aqui escreveria no(s) mesmo(s) destino(s)/
  tabela(s) do backup do WIN. Não mexer sem antes apontar para um destino
  próprio do WDO.
- `inventario_planilha.md` e `mapa_campos_rtd.md` documentam a descoberta
  original feita para o WIN — ainda válidos como referência de como os
  campos RTD funcionam, mas os exemplos citam o WINFUTV.
