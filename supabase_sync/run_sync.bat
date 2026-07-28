@echo off
REM Sync fim-de-dia do Monitor WIN -> Supabase
REM Usado pela Tarefa Agendada "MonitorWinSupabase" e pode ser rodado a mao.
cd /d "%~dp0"
py sync_supabase.py >> sync.log 2>&1
