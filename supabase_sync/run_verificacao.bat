@echo off
REM Verificacao unica pos-pregao do sync Supabase - Tarefa Agendada
REM "MonitorWinSupabaseCheck24Ago" (uma unica vez, 24/08/2026 18:45).
cd /d "%~dp0"
py verificar_pregao.py "_baseline_check_20260824.json" > "verificacao_2026-08-24.txt" 2>&1
