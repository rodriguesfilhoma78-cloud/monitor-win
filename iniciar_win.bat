@echo off
rem Sobe o Monitor WIN em segundo plano e abre o dashboard no navegador.
cd /d "%~dp0"

curl -s -o nul -m 1 http://127.0.0.1:8001/ultimo
if %ERRORLEVEL% EQU 0 (
    echo Monitor WIN ja esta no ar - nao inicio outro processo.
    goto abrir_dashboard
)
start "Monitor WIN" /min python server_win.py > server_win.log 2> server_win.err.log
echo Iniciando Monitor WIN em http://127.0.0.1:8001 ...
timeout /t 3 /nobreak >nul

:abrir_dashboard
rem Abre o APP instalado (janela PWA), nao a aba do navegador.
start "" "C:\Program Files (x86)\Microsoft\Edge\Application\msedge_proxy.exe" --profile-directory=Default --app-id=lpjoagiglbnkkhicpnomcbehdihdoikl --app-url=http://127.0.0.1:8001/
exit /b 0
