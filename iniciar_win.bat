@echo off
rem Sobe o Monitor WIN em segundo plano e abre o dashboard no navegador.
cd /d "%~dp0"

start "Monitor WIN" /min python server_win.py > server_win.log 2> server_win.err.log

echo Iniciando Monitor WIN em http://127.0.0.1:8001 ...
timeout /t 3 /nobreak >nul
start "" "http://127.0.0.1:8001/"
exit /b 0
