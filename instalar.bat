@echo off
cd /d %~dp0

echo ============================================
echo  Calculadora de Horas Extras - Instalacao
echo ============================================
echo.

where node >nul 2>nul
if errorlevel 1 (
    echo ERRO: Node.js nao foi encontrado neste computador.
    echo Baixe e instale em https://nodejs.org e rode este arquivo de novo.
    pause
    exit /b 1
)

echo Instalando dependencias, aguarde...
call npm install
if errorlevel 1 (
    echo.
    echo ERRO: falha ao instalar as dependencias.
    pause
    exit /b 1
)

echo.
echo Gerando versao de producao...
call npm run build
if errorlevel 1 (
    echo.
    echo ERRO: falha ao gerar o build.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Instalacao concluida!
echo  Use o arquivo "iniciar.bat" para abrir o programa.
echo ============================================
pause
