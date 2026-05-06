# 🍽️ Calculadora de Propinas

**Autor:** Christian Lera

Una calculadora de propinas profesional y completa con interfaz en consola, que incluye historial de cálculos, estadísticas, múltiples monedas y exportación a Excel/CSV.

## ✨ Características

- **Cálculo de propinas** con porcentajes predefinidos (0% a 25%) o personalizados
- **División de cuenta** entre varias personas
- **Múltiples monedas**: USD, EUR, MXN, GBP, COP, ARS
- **Historial** automático de todos los cálculos
- **Estadísticas** (propina promedio, total acumulado, etc.)
- **Exportación** a Excel (.xlsx) o CSV
- **Persistencia** en archivo JSON

## 📋 Requisitos

- Python 3.7 o superior
- Opcional: `openpyxl` para exportar a Excel (si no está instalado, usará CSV automáticamente)

## 🚀 Instalación y Ejecución

### En Windows (Recomendado)
1. Asegúrate de tener Python instalado
2. Ejecuta `ejecutar.bat` (como administrador si es necesario)

### En Windows (PowerShell)
```powershell
.\ejecutar.ps1
```

### En cualquier sistema (Manual)
```bash
# Instalar dependencias (opcional, solo para Excel)
pip install openpyxl

# Ejecutar la calculadora
python CalculadoraDePropinas.py
```

## 🎮 Uso del Programa

Al ejecutar, verás un menú con estas opciones:

| Opción | Descripción |
|--------|-------------|
| 1 | Calcular nueva propina |
| 2 | Ver historial de cálculos |
| 3 | Ver estadísticas |
| 4 | Cambiar moneda |
| 5 | Exportar a Excel/CSV |
| 6 | Salir |

## 📁 Archivos Generados

- `historial_propinas.json` - Guarda automáticamente todos los cálculos
- `historial_propinas.xlsx` - Exportación a Excel (cuando usas la opción 5)
- `historial_propinas.csv` - Exportación a CSV (alternativa)

## 💡 Ejemplo de Uso

```
💰 Monto de la cuenta: 100
⭐ Calidad del servicio:
   1. Pésimo (0%)
   2. Malo (5%)
   3. Regular (10%)
   4. Bueno (15%)
   5. Muy Bueno (18%)
   6. Excelente (20%)
   7. Extraordinario (25%)
   8. Personalizado

Selecciona: 6
👥 ¿Dividir cuenta? (s/n): s
Número de personas: 4

Resultado:
💰 Monto original:      $100.00
⭐ Calidad servicio:    Excelente (20%)
💵 Propina:            $20.00
🌟 Total a pagar:      $120.00
👥 Entre 4 personas:
👤 Cada persona paga:  $30.00
```

## 🛠️ Estructura del Código

```
CalculadoraDePropinas.py
├── CalidadServicio (Enum) - Porcentajes predefinidos
├── Moneda (Enum) - Monedas soportadas
├── CalculoPropina (dataclass) - Modelo de datos
├── GestorHistorial - Guardado/carga/estadísticas/exportación
├── CalculadoraPropina - Lógica de cálculos
└── InterfazUsuario - Menú y entrada de usuario
```

## ⚠️ Solución de Problemas

**Error "openpyxl no está instalado"** → La exportación usará CSV automáticamente o instala con `pip install openpyxl`

**Error al guardar historial** → Verifica que tengas permisos de escritura en la carpeta

**No aparece el menú correctamente** → La consola debe tener soporte UTF-8

## 📝 Notas de Versión

- **Versión:** 1.0
- **Python:** 3.7+
- **Dependencias opcionales:** openpyxl

---

© 2025 Christian Lera - Calculadora de Propinas
