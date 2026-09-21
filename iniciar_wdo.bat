@echo off
rem Sobe o Monitor WDO em segundo plano e abre o dashboard no navegador.
cd /d "%~dp0"

curl -s -o nul -m 1 http://127.0.0.1:8003/ultimo
if %ERRORLEVEL% EQU 0 (
    echo Monitor WDO ja esta no ar - nao inicio outro processo.
    goto abrir_dashboard
)
start "Monitor WDO" /min python server_wdo.py > server_wdo.log 2> server_wdo.err.log
echo Iniciando Monitor WDO em http://127.0.0.1:8003 ...
timeout /t 3 /nobreak >nul

:abrir_dashboard
rem Abre o APP instalado (janela PWA), nao a aba do navegador.
start "" "C:\Program Files (x86)\Microsoft\Edge\Application\msedge_proxy.exe" --profile-directory=Default --app-id=lnnhenchgamoiedplfmhhpgkoejjbdab --app-url=http://127.0.0.1:8003/
exit /b 0
