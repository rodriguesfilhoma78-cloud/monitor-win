-- =====================================================================
-- Monitor WIN -> Azure SQL Database  |  schema espelhando win_history.db (SQLite)
-- Rodar 1x no Azure: Portal -> seu banco -> Query editor (ou SSMS) -> colar -> Run
-- Idempotente: pode rodar de novo sem quebrar (IF OBJECT_ID ... IS NULL).
-- =====================================================================

-- fluxo: fita/fluxo do book (a tabela grande, ~2s de cadencia)
IF OBJECT_ID('dbo.fluxo', 'U') IS NULL
CREATE TABLE dbo.fluxo (
    id                bigint primary key,
    dia               nvarchar(20),
    ts                nvarchar(30),
    bid               float,
    ask               float,
    spread            float,
    qtd_bid           float,
    qtd_ask           float,
    desequilibrio     float,
    profundidade_bid  float,
    profundidade_ask  float,
    negocios          bigint,
    contratos         float,
    lote_medio        float,
    lote_max          float,
    agr_compra_fita   float,
    agr_venda_fita    float,
    pressao_fita      float,
    fita_estourou     bigint,
    poc               float,
    vah               float,
    val               float,
    vap_total         float,
    vol_acima_pct     float,
    dist_poc          float,
    amostra_negocios  bigint,
    amostra_contratos float,
    janela_s          float,
    negocios_s        float,
    contratos_s       float
);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_fluxo_dia' AND object_id = OBJECT_ID('dbo.fluxo'))
CREATE INDEX idx_fluxo_dia ON dbo.fluxo (dia);

-- snapshots: preco tick a tick
IF OBJECT_ID('dbo.snapshots', 'U') IS NULL
CREATE TABLE dbo.snapshots (
    id          bigint primary key,
    ts          nvarchar(30),
    ultimo      float,
    abertura    float,
    maxima      float,
    minima      float,
    volume      float,
    agr_compra  float,
    agr_venda   float,
    vwap        float,
    dia         nvarchar(20)
);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_snapshots_dia' AND object_id = OBJECT_ID('dbo.snapshots'))
CREATE INDEX idx_snapshots_dia ON dbo.snapshots (dia);

-- macro_snapshots: S&P, dolar, DI, DXY
IF OBJECT_ID('dbo.macro_snapshots', 'U') IS NULL
CREATE TABLE dbo.macro_snapshots (
    id         bigint primary key,
    dia        nvarchar(20),
    ts         nvarchar(30),
    sp500      float,
    sp500_var  float,
    dolar      float,
    dolar_var  float,
    di         float,
    di_var_bps float,
    dxy        float,
    dxy_var    float
);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_macro_dia' AND object_id = OBJECT_ID('dbo.macro_snapshots'))
CREATE INDEX idx_macro_dia ON dbo.macro_snapshots (dia);

-- eventos: sinais/alertas
IF OBJECT_ID('dbo.eventos', 'U') IS NULL
CREATE TABLE dbo.eventos (
    id        bigint primary key,
    dia       nvarchar(20),
    ts        nvarchar(30),
    evento    nvarchar(100),
    direcao   nvarchar(20),
    nivel     float,
    delta_ema float,
    msg       nvarchar(max)
);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_eventos_dia' AND object_id = OBJECT_ID('dbo.eventos'))
CREATE INDEX idx_eventos_dia ON dbo.eventos (dia);

-- niveis_hist: niveis (dados = JSON em texto)
IF OBJECT_ID('dbo.niveis_hist', 'U') IS NULL
CREATE TABLE dbo.niveis_hist (
    id     bigint primary key,
    dia    nvarchar(20),
    ts     nvarchar(30),
    origem nvarchar(50),
    dados  nvarchar(max)
);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_niveis_dia' AND object_id = OBJECT_ID('dbo.niveis_hist'))
CREATE INDEX idx_niveis_dia ON dbo.niveis_hist (dia);

-- operacoes: SUAS operacoes (hoje vazia; pronta pra usar depois)
IF OBJECT_ID('dbo.operacoes', 'U') IS NULL
CREATE TABLE dbo.operacoes (
    id       bigint primary key,
    ts       nvarchar(30),
    dia      nvarchar(20),
    hora     nvarchar(20),
    lado     nvarchar(10),
    preco    float,
    motivo   nvarchar(max),
    nota     nvarchar(max),
    abertura float,
    maxima   float,
    minima   float,
    volume   float,
    delta    float,
    vwap     float
);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'idx_operacoes_dia' AND object_id = OBJECT_ID('dbo.operacoes'))
CREATE INDEX idx_operacoes_dia ON dbo.operacoes (dia);

-- corretoras: agressao por corretora/dia (PK composta: dia+corretora)
IF OBJECT_ID('dbo.corretoras', 'U') IS NULL
CREATE TABLE dbo.corretoras (
    dia        nvarchar(20),
    corretora  nvarchar(100),
    qtd_compra float,
    qtd_venda  float,
    fin_compra float,
    fin_venda  float,
    negocios   bigint,
    primary key (dia, corretora)
);

-- daily_ohlc: OHLC diario (PK: dia)
IF OBJECT_ID('dbo.daily_ohlc', 'U') IS NULL
CREATE TABLE dbo.daily_ohlc (
    dia         nvarchar(20) primary key,
    abertura    float,
    maxima      float,
    minima      float,
    fechamento  float,
    vwap        float,
    primeiro_ts nvarchar(30),
    ultimo_ts   nvarchar(30),
    valido      bigint
);
