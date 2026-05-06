# ejecutar.ps1 - Script para ejecutar la Calculadora de Propinas
# Autor: Christian Lera

# Configurar la consola
$Host.UI.RawUI.WindowTitle = "Calculadora de Propinas - Christian Lera"
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "   🍽️  CALCULADORA DE PROPINAS  🍽️" -ForegroundColor Yellow
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "   Autor: Christian Lera" -ForegroundColor White
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host ""

# Verificar si Python está instalado
try {
    $pythonVersion = python --version 2>&1
    Write-Host "[OK] Python detectado:" -ForegroundColor Green
    Write-Host "     $pythonVersion" -ForegroundColor White
    Write-Host ""
} catch {
    Write-Host "[ERROR] Python no está instalado en el sistema." -ForegroundColor Red
    Write-Host ""
    Write-Host "Por favor, instala Python desde: https://python.org" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Presiona Enter para salir"
    exit 1
}

# Verificar openpyxl (opcional)
try {
    python -c "import openpyxl" 2>&1 | Out-Null
    Write-Host "[OK] openpyxl detectado (soportará Excel)" -ForegroundColor Green
    Write-Host ""
} catch {
    Write-Host "[AVISO] openpyxl no está instalado." -ForegroundColor Yellow
    Write-Host "        La exportación a Excel usará CSV como alternativa." -ForegroundColor Yellow
    Write-Host "        Para instalar: pip install openpyxl" -ForegroundColor White
    Write-Host ""
}

# Ejecutar la calculadora
Write-Host "Iniciando la calculadora..." -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host ""

try {
    python CalculadoraDePropinas.py
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "=================================================" -ForegroundColor Red
        Write-Host "[ERROR] La calculadora cerró inesperadamente." -ForegroundColor Red
        Write-Host "=================================================" -ForegroundColor Red
    }
} catch {
    Write-Host ""
    Write-Host "[ERROR] No se pudo ejecutar la calculadora:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor White
}

Write-Host ""
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "   ¡Gracias por usar la calculadora!" -ForegroundColor Yellow
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host ""
Read-Host "Presiona Enter para salir"