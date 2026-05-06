@echo off
title Calculadora de Propinas - Christian Lera
color 0A

echo =================================================
echo    🍽️  CALCULADORA DE PROPINAS  🍽️
echo =================================================
echo    Autor: Christian Lera
echo =================================================
echo.

:: Verificar si Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no esta instalado en el sistema.
    echo.
    echo Por favor, instala Python desde: https://python.org
    echo.
    pause
    exit /b 1
)

:: Mostrar versión de Python
echo [OK] Python detectado:
python --version
echo.

:: Verificar si openpyxl está instalado (opcional)
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo [AVISO] openpyxl no esta instalado.
    echo         La exportacion a Excel usara CSV como alternativa.
    echo         Para instalar: pip install openpyxl
    echo.
) else (
    echo [OK] openpyxl detectado (soportara Excel)
    echo.
)

:: Ejecutar la calculadora
echo Iniciando la calculadora...
echo =================================================
echo.

python CalculadoraDePropinas.py

:: Pausar si ocurre un error o al salir
if errorlevel 1 (
    echo.
    echo =================================================
    echo [ERROR] La calculadora cerro inesperadamente.
    echo =================================================
)

echo.
echo =================================================
echo    ¡Gracias por usar la calculadora!
echo =================================================
pause
