@echo off
cd /d %~dp0

if not exist "dist\index.html" (
    echo A instalacao ainda nao foi feita.
    echo Rode primeiro o arquivo "instalar.bat".
    pause
    exit /b 1
)

echo Iniciando Calculadora de Horas Extras...
echo (Feche esta janela para encerrar o programa)
echo.

start "" cmd /c "timeout /t 2 /nobreak >nul & start "" http://localhost:5175"
call npm start
if errorlevel 1 (
    echo.
    echo O programa foi encerrado com erro.
    pause
)
