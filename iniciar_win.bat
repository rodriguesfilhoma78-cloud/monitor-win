@echo off
title Monitor WIN (porta 8001)
cd /d "%~dp0"

rem === nao reinicia se ja estiver no ar (mesma filosofia do README) ===
netstat -ano | findstr ":8001" | findstr "LISTENING" >nul 2>&1
if %errorlevel%==0 (
    echo Monitor WIN ja esta rodando na porta 8001. Nada a fazer.
    "%SystemRoot%\System32\timeout.exe" /t 3 /nobreak >nul 2>&1
    exit /b 0
)

echo Iniciando Monitor WIN em http://127.0.0.1:8001 ...
start "Monitor WIN" /min python server_win.py
exit /b 0
