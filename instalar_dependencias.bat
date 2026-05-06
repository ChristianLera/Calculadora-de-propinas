@echo off
title Instalar Dependencias - Calculadora de Propinas
color 0E

echo =================================================
echo    📦 INSTALADOR DE DEPENDENCIAS
echo    Calculadora de Propinas - Christian Lera
echo =================================================
echo.

:: Verificar Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no esta instalado.
    echo.
    echo Por favor, instala Python desde: https://python.org
    echo.
    pause
    exit /b 1
)

echo [OK] Python detectado:
python --version
echo.

:: Verificar si pip está disponible
pip --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] pip no esta disponible.
    echo.
    pause
    exit /b 1
)

echo [OK] pip detectado
echo.

:: Mostrar estado actual de openpyxl
echo Comprobando instalacion actual...
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo [AVISO] openpyxl NO esta instalado
    echo.
    echo Instalando openpyxl...
    echo.
    pip install openpyxl
    echo.
) else (
    echo [OK] openpyxl YA esta instalado
    echo.
    set /p actualizar="¿Quieres actualizarlo? (s/n): "
    if /i "!actualizar!"=="s" (
        echo.
        echo Actualizando openpyxl...
        pip install --upgrade openpyxl
        echo.
    )
)

:: Verificar instalación
echo.
echo Verificando instalacion...
python -c "import openpyxl; print(f'[OK] openpyxl version {openpyxl.__version__} instalada')" >nul 2>&1
if errorlevel 1 (
    echo [ERROR] No se pudo instalar openpyxl
    echo.
    echo La calculadora funcionara igual, pero exportara a CSV en lugar de Excel
) else (
    python -c "import openpyxl; print(f'[OK] openpyxl version {openpyxl.__version__} instalada correctamente')"
)

echo.
echo =================================================
echo    Proceso completado
echo =================================================
echo.
pause