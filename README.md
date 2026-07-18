# Monitor WIN — Sistema independente de monitoramento do Mini Índice

Sistema completo e separado do Mapa de Tendência WDO. Roda em paralelo
(porta **8001**, enquanto o WDO usa a 8000).

## Estrutura da pasta

```
monitor_win\
├── server_win.py            # Servidor FastAPI + WebSocket + API de níveis (back-end)
├── iniciar_win.bat          # Sobe o server se a porta 8001 estiver livre
├── requirements.txt
├── frontend\                # Dashboard (servido pelo próprio server em /)
│   ├── dashboard_win.html   #   estrutura da página
│   ├── style.css            #   estilos
│   └── app.js               #   lógica (WebSocket, régua, alertas, macro…)
├── excel\                   # Módulos VBA (importar no Excel; não rodam daqui)
│   ├── ExportarWIN.bas      #   exporta dados_win.csv + dados_macro_rtd.csv
│   └── ExportarBlueChips.bas#   exporta dados_blue_chips.csv
├── docs\
│   └── CHANGELOG.md
├── niveis.json              # Níveis do dia (editável — sem reiniciar o server)
├── dados_win.csv            # (gerado pelo VBA em tempo de execução — na RAIZ)
└── win_history.db           # (gerado automaticamente — snapshots SQLite)
```

> **Importante:** os CSVs de runtime, o `niveis.json` e o `.db` continuam na
> **raiz** — é onde o VBA grava (caminho fixo nos `.bas`) e onde o server lê.
> Não mova esses arquivos para dentro das pastas.

## Instalação (uma vez)

```bash
cd C:\Users\rodri\OneDrive\Apps\monitor_win
pip install -r requirements.txt
```

No Excel (`WDO_Master_RTD.xlsm`):
1. A linha do **WINFUTV** na aba DADOS é localizada automaticamente pela
   coluna A (imune a inserção/remoção de linhas). Colunas usadas:
   D=último, E=abertura, F=máxima, G=mínima, X=agr.compra, Z=agr.venda,
   AA=VWAP, AB=volume.
2. O módulo `ExportarWIN.bas` já está importado (Módulo2) e a exportação
   inicia sozinha ao abrir a planilha (evento `Workbook_Open`).
3. Confira o `CAMINHO_CSV` no topo do módulo se mudar a pasta.

## Sequência de inicialização (dia de pregão)

1. Abrir Profit Pro
2. Abrir Excel (`WDO_Master_RTD.xlsm`) e **habilitar macros** se solicitado —
   ~10s depois a exportação do WIN inicia sozinha **e o `server_win.py`
   sobe automaticamente em segundo plano** (sem janela), caso a porta
   8001 ainda não esteja respondendo. Um server já no ar nunca é
   reiniciado (preserva o histórico intradiário).
   (fallback manual: Alt+F8 → `IniciarExportWIN`)
3. Navegador: **http://127.0.0.1:8001**

> Assim como no WDO: **não reinicie o server após as 9h** se quiser
> preservar o histórico intradiário de ticks no dashboard aberto.

## Níveis do dia

**Automático (padrão):** na primeira inicialização do dia (ou na virada de
data com o server no ar), os níveis são recalculados por **pivot points
clássicos** sobre o OHLC do pregão anterior (gravado em `daily_ohlc` no
`win_history.db`):

```
P  = (H+L+C)/3
R1 = 2P-L      S1 = 2P-H
R2 = P+(H-L)   S2 = P-(H-L)
R3 = H+2(P-L)  S3 = L-2(H-P)
Zona decisiva = Central Pivot Range: entre BC=(H+L)/2 e TC=2P-BC
```

Tudo arredondado ao tick de 5 pontos. A `fonte` no cabeçalho do dashboard
indica quando os níveis são automáticos.

O **C (fechamento)** usa o valor gravado no fim do pregão anterior; no
primeiro tick do dia o servidor confere contra o **FEC oficial** do RTD
(coluna H da planilha, exportada como `fec_ant` no CSV) e recalcula se
divergir — protege contra server desligado antes das 18h.

**Manual (override):** edite `niveis.json` e envie — enquanto o
`data_pregao` for o de hoje, o automático não sobrescreve:

```bash
curl -X POST http://127.0.0.1:8001/niveis -H "Content-Type: application/json" -d @niveis.json
```

**Voltar ao automático:** `curl -X POST http://127.0.0.1:8001/niveis/recalcular`

## Endpoints

| Rota      | Método | Descrição                                  |
|-----------|--------|--------------------------------------------|
| `/`       | GET    | Dashboard                                  |
| `/ws`     | WS     | Stream de ticks em tempo real              |
| `/niveis` | GET    | Níveis atuais (JSON)                       |
| `/niveis` | POST   | Atualiza níveis + notifica dashboards      |
| `/niveis/recalcular` | POST | Força recálculo pelos pivots do pregão anterior |
| `/historico/{dia}` | GET | Pacote do pregão (OHLC, níveis, sinais, snapshots) p/ análise |
| `/perfil/{dia}` | GET | Perfil de volume: POC, value area e histograma por preço (`?bucket=50`) |
| `/operacoes` | GET | Operações do trader no dia + a aberta (se houver) |
| `/operacoes` | POST | Registra entrada (`{"tipo":"compra"\|"venda","motivo":"..."}`) |
| `/operacoes/fechar` | POST | Fecha a operação aberta e grava o resultado |
| `/operacoes/stats` | GET | Taxa de acerto / média / total por tipo+motivo |
| `/ultimo` | GET    | Último tick recebido (debug rápido)        |

## Perfil de volume (régua por acúmulo)

Os snapshots (a cada **2s**) alimentam o perfil de volume do pregão:
o volume negociado entre snapshots consecutivos é atribuído à faixa de
preço (bucket de 50 pts) em que ocorreu. O dashboard pinta as faixas de
maior acúmulo como bandas azuis na régua + linha **POC** (preço com mais
negócios), atualizadas a cada 2 min. Pivots = projeção do dia anterior;
acúmulo = memória real de onde o mercado negociou — os dois convivem.

## Registro de operações (fase 1 do aprendizado por imitação)

Card **REGISTRO DO TRADER** no dashboard. A cada entrada real, o trader
escolhe o motivo (rompimento, pullback, fluxo…) e clica **▲ COMPREI** ou
**▼ VENDI**; ao encerrar, **✖ SAÍ**. Regras de ouro:

- **Registrar TODAS as operações, ganhe ou perca.** Só as vencedoras =
  viés de sobrevivência; as perdas ensinam tanto quanto os ganhos.
- O preço e o **contexto completo** (tick, níveis, macro, blue chips,
  fluxo EMA, plano) são capturados **pelo servidor** no instante do
  clique e gravados na tabela `operacoes` do `win_history.db`.
- **Nota livre (opcional, vale ouro):** campo de texto para o raciocínio
  da entrada ("fluxo virou na VWAP…") e da saída ("saí no alvo / no
  medo / no stop"). Não é obrigatório — a tag é o mínimo; a nota é o
  bônus que a análise por IA lê depois para agrupar padrões de
  raciocínio e cruzar com o resultado.
- Uma operação aberta por vez; registro **bloqueado no simulador**.

Com o dataset acumulado, `GET /operacoes/stats` responde com números o
que cada padrão do trader rende (taxa de acerto, média em pts). É a base
das fases seguintes: análise de padrões → copiloto que aponta contextos
similares → (só com edge provado) discussão de automação.

## Testar sem o pregão

Abra o dashboard e clique em **▶ Simulador** — gera ticks sintéticos
para validar alertas, régua e viés.
