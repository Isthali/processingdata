# Script Unificado de Reportes — Guía de Uso

`unified_report.py` genera los reportes PDF de los distintos ensayos mecánicos
del laboratorio desde un único punto de entrada.

## Requisitos

- **Sistema operativo**: Windows con Microsoft Excel instalado. La conversión
  Excel → PDF usa COM (`pywin32`) y **no funciona en Linux/macOS**.
- **Python 3.9+** y las dependencias de `requirements.txt`
  (`numpy`, `scipy`, `pandas`, `matplotlib`, `openpyxl`, `chardet`, `pypdf`,
  `reportlab`, `pywin32`).

## Entradas esperadas en disco

Antes de correr un subcomando, la carpeta de trabajo `{base-dir}/{infle}/`
debe contener:

1. **Plantilla Excel pre-configurada** con las hojas y celdas donde el reporte
   escribe los resultados. El nombre debe ser:
   ```
   INFLE_{infle}[-{subinfle}]_{standard}_{client}[_{n}].xlsm
   ```
   (los genéricos usan `.xlsx`; el sufijo `_{n}` sólo aparece si hay muestras).

2. **Archivos de datos crudos por muestra**, con la convención de nombre y las
   columnas que espera cada tipo:

   | Subcomando        | Archivo por muestra                         | Columnas                                 |
   |-------------------|---------------------------------------------|------------------------------------------|
   | `cores`           | `{infle}-d{id}.xlsx`                        | Time, Displacement, Load                 |
   | `panels`          | `Losa P{id}.xlsx`                           | Time, Load, Deflection, Displacement     |
   | `panels_residual` | `{infle}-Viga {id}/specimen.dat`            | Time, Displacement, Load, CMOD, Deflection |
   | `beams_residual`  | `{infle}-Viga {id}/specimen.dat`            | Time, Displacement, Load, Deflection, Deflection2 |
   | `tapas`           | `tapas_{id}.xlsx`                           | Time, Load                               |
   | `generic`         | (sin datos crudos — sólo la plantilla)      | —                                        |

Si la plantilla o los archivos de datos no existen, el error aparece al momento
de `add_tests()` o de escribir el Excel, no al parsear los argumentos.

## Subcomandos

| Subcomando        | Ensayo                                             | Estándar por defecto | Estándares disponibles                          | n por defecto | Cliente por defecto |
|-------------------|----------------------------------------------------|----------------------|-------------------------------------------------|---------------|---------------------|
| `cores`           | Compresión axial de testigos                       | `CORES`              | libre                                           | 6             | `PRODIMIN`          |
| `panels`          | Tenacidad de paneles (ASTM C1550, EFNARC, EN 14488-5) | `EFNARC1996`         | `ASTMC1550`, `EFNARC1996`, `EFNARC1999`, `EN14488-5` | 3             | `PRODIMIN`          |
| `panels_residual` | Resistencia residual con CMOD (EN 14651 / EN 14488) | `EN14488`            | `EN14651`, `EN14488`                            | 3             | `PRODIMIN`          |
| `beams_residual`  | Resistencia residual de vigas (ASTM C1609)         | `ASTMC1609`          | `ASTMC1609`                                     | 3             | `PRODIMIN`          |
| `tapas`           | Flexión de tapas de buzón (tránsito)               | `Tapa_Circular_CA`   | libre                                           | 3             | `PRODIMIN`          |
| `generic`         | Conversión Excel → PDF sin procesamiento           | `DM`                 | libre                                           | 0             | `EMPRESA`           |

"Libre" significa que la CLI no restringe el valor; el reporte acepta
cualquier string como identificador del estándar, pero no calculará puntos
característicos por norma si no están mapeados internamente.

## Sintaxis

```bash
python unified_report.py <subcomando> --infle <id> [opciones]
```

### Argumentos comunes

| Flag           | Requerido | Descripción                                                       |
|----------------|-----------|-------------------------------------------------------------------|
| `--infle`      | sí        | Identificador del informe (ej. `336-24`, `111-25`).               |
| `--subinfle`   | no        | Sub-identificador (ej. `S`, `C`). Vacío por defecto.              |
| `--standard`   | no        | Estándar del ensayo. Usa el default del subcomando si se omite.   |
| `--empresa`    | no        | Nombre del cliente.                                               |
| `--base-dir`   | no        | Directorio base. Por defecto `D:/`.                               |
| `--verbose`/`-v` | no      | Activa logging DEBUG.                                             |

### Argumentos de muestras (todos los subcomandos salvo `generic`)

Las opciones `--n` y `--ids` son **mutuamente exclusivas**:

| Flag       | Descripción                                                              |
|------------|--------------------------------------------------------------------------|
| `--n N`    | Número de muestras consecutivas desde `--offset` (default del subcomando si se omite). |
| `--offset` | Valor inicial de ID cuando se usa `--n`. Por defecto 1.                  |
| `--ids`    | Lista explícita de IDs (ej. `--ids 3 4 7`). Anula `--n`/`--offset`.      |

## Ejemplos

### Compresión axial (testigos)

```bash
# Defaults (n=6 muestras desde id=1)
python unified_report.py cores --infle 336-24 --subinfle S

# 8 muestras desde id=3 con carpeta custom
python unified_report.py cores --infle 336-24 --subinfle S \
    --n 8 --offset 3 --base-dir D:/Reportes/Testigos
```

### Tenacidad de paneles

```bash
python unified_report.py panels --infle 111-25 --subinfle C

# Con norma e IDs no consecutivos
python unified_report.py panels --infle 111-25 --subinfle C \
    --standard EFNARC1999 --ids 2 4 5
```

### Resistencia residual con CMOD

```bash
python unified_report.py panels_residual --infle 111-25 --subinfle C \
    --standard EN14488 --n 4 --offset 7
```

### Resistencia residual de vigas (deflexión)

```bash
python unified_report.py beams_residual --infle 222-25 --subinfle B --n 3
```

### Flexión de tapas de buzón

```bash
python unified_report.py tapas --infle 245-25 --subinfle A --empresa FADECO --n 6
```

### Conversión genérica

```bash
python unified_report.py generic --infle 336-24 --subinfle S \
    --empresa EMPRESA_CLIENTE
```

## Archivos de salida

Dentro de `{base-dir}/{infle}/` quedan:

1. La plantilla Excel **modificada** (el script escribe resultados en celdas).
2. `plots.pdf` — gráficos intermedios generados con matplotlib.
3. `INFLE_{infle}[-{subinfle}]_{standard}_{client}[_{n}].pdf` — el reporte
   final con encabezado/pie aplicado.

## Resolución de problemas

Ayuda integrada:

```bash
python unified_report.py --help
python unified_report.py <subcomando> --help
```

Errores comunes:

- **Directorio base no existe**: se crea automáticamente con una advertencia.
- **`--n` negativo**: rechazado (debe ser ≥ 0). Los valores de `--ids` deben ser > 0.
- **Estándar no válido**: los subcomandos con lista cerrada (`panels`,
  `panels_residual`, `beams_residual`) rechazan valores fuera de la lista con
  `--help`. Los demás aceptan cualquier string.
- **Archivo de datos no encontrado**: la traza apunta a la ruta esperada bajo
  `{base-dir}/{infle}/`. Verifica que los nombres coincidan con la convención
  del subcomando (tabla arriba).
- **Excel no convierte a PDF**: requiere Windows con Excel instalado. Si el
  error menciona `win32com`, estás corriendo en Linux/macOS.
