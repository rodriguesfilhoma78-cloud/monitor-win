@echo off
REM Sync fim-de-dia do Monitor WIN -> Azure SQL Database
REM Pode ser usado por uma Tarefa Agendada (ex: "MonitorWinAzureSql") ou rodado a mao.
cd /d "%~dp0"
py sync_azure_sql.py >> sync.log 2>&1
