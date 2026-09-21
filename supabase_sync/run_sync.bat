@echo off
REM Sync fim-de-dia do Monitor WDO -> Supabase (tabelas wdo_*)
REM Sem Tarefa Agendada propria ainda - rodar a mao ou criar uma nova
REM (ex.: "MonitorWdoSupabase") separada da "MonitorWinSupabase" do WIN.
cd /d "%~dp0"
py sync_supabase.py >> sync.log 2>&1
