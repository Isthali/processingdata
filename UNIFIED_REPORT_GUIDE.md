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

2. **Archivos de datos crudos por muestra**. Cada tipo tiene una convención de
   nombre y un orden de columnas **por defecto** (tabla abajo); ambos se pueden
   cambiar con `--file-pattern` y `--columns`, o de forma permanente con un
   [`report_config.json`](#archivo-de-configuración-por-informe) en la carpeta
   del informe:

   | Subcomando        | Archivo por muestra                         | Columnas                                 |
   |-------------------|---------------------------------------------|------------------------------------------|
   | `cores`           | `{infle}-d{id}.xlsx`                        | Time, Displacement, Load                 |
   | `cores_local`     | `{infle}-d{id}.xlsx`                        | Time, Load, Displacement, SG1, SG2, SG3 (SG en µε) |
   | `panels`          | `Losa P{id}.xlsx`                           | Time, Load, Deflection, Displacement     |
   | `panels_residual` | `{infle}-Viga {id}/specimen.dat`            | Time, Displacement, Load, CMOD, Deflection |
   | `beams_residual`  | `{infle}-Viga {id}/specimen.dat`            | Time, Displacement, Load, Deflection, Deflection2 |
   | `tapas`           | `tapas_{id}.xlsx`                           | Time, Load                               |
   | `generic`         | (sin datos crudos — sólo la plantilla)      | —                                        |

   Para `cores` y `cores_local` el **diámetro de cada probeta (mm) se lee del
   propio template** en la celda `'Cores'!F{i+16}` (1ª muestra → F16, 2ª → F17,
   etc.). Si la celda está vacía o ≤ 0, el script aborta. La resistencia se
   escribe en columna **N (MPa)** y **O (kgf/cm²)** de la misma fila; la carga
   máxima en la columna **W (kN)**. La fila de la primera muestra (16 por
   defecto) se puede cambiar con `--start-row`.

Si la plantilla o los archivos de datos no existen, el error aparece al momento
de `add_tests()` o de escribir el Excel, no al parsear los argumentos.

## Subcomandos

| Subcomando        | Ensayo                                             | Estándar por defecto | Estándares disponibles                          | n por defecto | Cliente por defecto |
|-------------------|----------------------------------------------------|----------------------|-------------------------------------------------|---------------|---------------------|
| `cores`           | Compresión axial de testigos                       | `CORES`              | libre                                           | 6             | `PRODIMIN`          |
| `cores_local`     | Compresión axial instrumentada con 3 strain gauges (esfuerzo-deformación) | `CORES`              | libre                                           | 3             | `PRODIMIN`          |
| `panels`          | Tenacidad de paneles (ASTM C1550, EFNARC, EN 14488-5) | `EFNARC1996`         | `ASTMC1550`, `EFNARC1996`, `EFNARC1999`, `EN14488-5` | 3             | `PRODIMIN`          |
| `panels_residual` | Resistencia residual con CMOD (EN 14651 / EN 14488) | `EN14488`            | `EN14651`, `EN14488`                            | 3             | `PRODIMIN`          |
| `beams_residual`  | Resistencia residual de vigas (ASTM C1609)         | `ASTMC1609`          | `ASTMC1609`                                     | 3             | `PRODIMIN`          |
| `tapas`           | Flexión de tapas de buzón (tránsito)               | `Tapa_Circular_CA`   | libre                                           | 3             | `PRODIMIN`          |
| `generic`         | Conversión Excel → PDF sin procesamiento           | `GENERIC`                 | libre                                           | 0             | `EMPRESA`           |

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
| `--num-1plot-pag` | Número de página donde inician los gráficos. Controla el rango exportado desde Excel (`pag_f = num_1plot_pag − 1`). Default por subcomando: `panels`/`beams_residual`/`tapas` = 4, `cores`/`cores_local`/`panels_residual` = 5. |
| `--start-row` | Fila de la plantilla Excel correspondiente a la **primera muestra**, para lectura y escritura. Default por subcomando: `cores`/`cores_local` = 16, `panels` = 18, `panels_residual`/`beams_residual` = 19, `tapas` = 26. Las muestras siguientes ocupan filas según el layout del ensayo: fila consecutiva (`cores`, `panels_residual`, `beams_residual`), bloques de 5 filas (`panels`) o de 3 filas (`tapas`). |
| `--file-pattern` | Patrón del nombre del archivo de datos por muestra, relativo a `{base-dir}/{infle}/`. Marcadores: `{id}` (obligatorio), `{infle}`, `{subinfle}`. La extensión decide el lector: `.xlsx`/`.xlsm`/`.xls` se leen como Excel, cualquier otra (`.dat`, `.txt`, `.csv`) como texto delimitado. Ej.: `--file-pattern "Panel M{id}.xlsx"` o `--file-pattern "losa M{id}.xlsx"`. Defaults: ver tabla de archivos arriba (`cores` prueba `{infle}-d{id}/specimen.dat` y luego `{infle}-d{id}.xlsx`). |
| `--columns` | Nombres de las columnas del archivo de datos **en el orden en que aparecen en el archivo** (se asignan posicionalmente). Ej.: `--columns Time Deflection Load Displacement`. Cada ensayo exige ciertas columnas (`Load` siempre; `Deflection`, `CMOD`, `Displacement` o `Time` según el tipo); si falta una, el error lo indica al iniciar. |
| `--config` | Ruta a un JSON de configuración del informe. Si se omite, se busca `report_config.json` en `{base-dir}/{infle}/` automáticamente. Ver sección siguiente. |

## Archivo de configuración por informe

Para no repetir flags en cada corrida, coloca un `report_config.json` dentro de
la carpeta del informe (`{base-dir}/{infle}/`). El script lo carga
automáticamente (o usa `--config <ruta>` para apuntar a otro archivo). Es la
opción recomendada cuando un informe usa nombres de archivo, orden de columnas
o filas de plantilla distintos de los defaults: se configura **una sola vez por
informe** y todas las corridas lo heredan.

```json
{
  "file_pattern": "Panel M{id}.xlsx",
  "columns": ["Time", "Deflection", "Load", "Displacement"],
  "start_row": 22,
  "num_1plot_pag": 5
}
```

Todas las claves son opcionales (`start_row`, `num_1plot_pag`, `file_pattern`,
`columns`); las no reconocidas se ignoran con una advertencia.

**Precedencia**: flag del CLI > `report_config.json` > default del subcomando.
Es decir, un flag explícito siempre gana sobre el archivo.

## Ejemplos

### Compresión axial (testigos)

```bash
# Defaults (n=6 muestras desde id=1)
python unified_report.py cores --infle 336-24 --subinfle S

# 8 muestras desde id=3 con carpeta custom
python unified_report.py cores --infle 336-24 --subinfle S \
    --n 8 --offset 3 --base-dir D:/Reportes/Testigos

# Plantilla con la primera muestra en la fila 20 (en vez de 16)
python unified_report.py cores --infle 336-24 --subinfle S \
    --n 6 --start-row 20
```

### Compresión axial con strain gauges (esfuerzo-deformación)

```bash
# 3 testigos instrumentados con SG1/SG2/SG3 (µε)
python unified_report.py cores_local --infle 336-24 --subinfle S --n 3
```

El archivo de datos por probeta es `{infle}-d{id}.xlsx` con columnas
`Time, Load, Displacement, SG1, SG2, SG3` (las SG en microdeformación). El
informe agrega un segundo gráfico Esfuerzo–Deformación además del usual
Fuerza–Desplazamiento, con esfuerzo en MPa y deformación unitaria (mm/mm).

### Tenacidad de paneles

```bash
python unified_report.py panels --infle 111-25 --subinfle C

# Con norma e IDs no consecutivos
python unified_report.py panels --infle 111-25 --subinfle C \
    --standard EFNARC1999 --ids 2 4 5

# Cambiando la página de inicio de los gráficos (default = 4)
python unified_report.py panels --infle 111-25 --subinfle C \
    --n 3 --num-1plot-pag 6

# Archivos llamados "Panel M1.xlsx", "Panel M2.xlsx", ... con las columnas
# en orden Time, Deflection, Load, Displacement
python unified_report.py panels --infle 111-25 --subinfle C --n 3 \
    --file-pattern "Panel M{id}.xlsx" \
    --columns Time Deflection Load Displacement
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
