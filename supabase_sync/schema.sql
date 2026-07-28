-- =====================================================================
-- Monitor WIN -> Supabase  |  schema espelhando win_history.db (SQLite)
-- Rodar 1x no Supabase: Dashboard -> SQL Editor -> colar -> Run
-- Idempotente: pode rodar de novo sem quebrar (IF NOT EXISTS).
-- =====================================================================

-- fluxo: fita/fluxo do book (a tabela grande, ~2s de cadencia)
create table if not exists public.fluxo (
    id                bigint primary key,
    dia               text,
    ts                text,
    bid               double precision,
    ask               double precision,
    spread            double precision,
    qtd_bid           double precision,
    qtd_ask           double precision,
    desequilibrio     double precision,
    profundidade_bid  double precision,
    profundidade_ask  double precision,
    negocios          bigint,
    contratos         double precision,
    lote_medio        double precision,
    lote_max          double precision,
    agr_compra_fita   double precision,
    agr_venda_fita    double precision,
    pressao_fita      double precision,
    fita_estourou     bigint,
    poc               double precision,
    vah               double precision,
    val               double precision,
    vap_total         double precision,
    vol_acima_pct     double precision,
    dist_poc          double precision,
    amostra_negocios  bigint,
    amostra_contratos double precision,
    janela_s          double precision,
    negocios_s        double precision,
    contratos_s       double precision
);
create index if not exists idx_fluxo_dia on public.fluxo (dia);

-- snapshots: preco tick a tick
create table if not exists public.snapshots (
    id          bigint primary key,
    ts          text,
    ultimo      double precision,
    abertura    double precision,
    maxima      double precision,
    minima      double precision,
    volume      double precision,
    agr_compra  double precision,
    agr_venda   double precision,
    vwap        double precision,
    dia         text
);
create index if not exists idx_snapshots_dia on public.snapshots (dia);

-- macro_snapshots: S&P, dolar, DI, DXY
create table if not exists public.macro_snapshots (
    id         bigint primary key,
    dia        text,
    ts         text,
    sp500      double precision,
    sp500_var  double precision,
    dolar      double precision,
    dolar_var  double precision,
    di         double precision,
    di_var_bps double precision,
    dxy        double precision,
    dxy_var    double precision
);
create index if not exists idx_macro_dia on public.macro_snapshots (dia);

-- eventos: sinais/alertas
create table if not exists public.eventos (
    id        bigint primary key,
    dia       text,
    ts        text,
    evento    text,
    direcao   text,
    nivel     double precision,
    delta_ema double precision,
    msg       text
);
create index if not exists idx_eventos_dia on public.eventos (dia);

-- niveis_hist: niveis (dados = JSON em texto)
create table if not exists public.niveis_hist (
    id     bigint primary key,
    dia    text,
    ts     text,
    origem text,
    dados  text
);
create index if not exists idx_niveis_dia on public.niveis_hist (dia);

-- operacoes: SUAS operacoes (hoje vazia; pronta pra usar depois)
create table if not exists public.operacoes (
    id       bigint primary key,
    ts       text,
    dia      text,
    hora     text,
    lado     text,
    preco    double precision,
    motivo   text,
    nota     text,
    abertura double precision,
    maxima   double precision,
    minima   double precision,
    volume   double precision,
    delta    double precision,
    vwap     double precision
);
create index if not exists idx_operacoes_dia on public.operacoes (dia);

-- corretoras: agressao por corretora/dia (PK composta: dia+corretora)
create table if not exists public.corretoras (
    dia        text,
    corretora  text,
    qtd_compra double precision,
    qtd_venda  double precision,
    fin_compra double precision,
    fin_venda  double precision,
    negocios   bigint,
    primary key (dia, corretora)
);

-- daily_ohlc: OHLC diario (PK: dia)
create table if not exists public.daily_ohlc (
    dia        text primary key,
    abertura   double precision,
    maxima     double precision,
    minima     double precision,
    fechamento double precision,
    vwap       double precision,
    primeiro_ts text,
    ultimo_ts   text,
    valido      bigint
);
