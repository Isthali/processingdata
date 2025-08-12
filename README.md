# Procesamiento y Generación de Reportes de Ensayos

Este directorio contiene utilidades para procesar datos de ensayos mecánicos (paneles, vigas, testigos/"cores") y generar reportes en PDF a partir de planillas Excel y archivos de adquisición de datos.

## Estructura principal

- `test_ledi.py`: Clases de modelo y lógica de procesamiento (ensayos, cálculo de índices, generación de gráficos y PDFs).  
- `test_data.py`: Funciones de importación/escritura de datos (texto y Excel).  
- `edit_pdfs.py`: Utilidades para conversión Excel→PDF, fusión y post-procesado (encabezado/pie, orientación).  
- `report_helpers.py`: Lógica común de CLI y ejecución de reportes (parser, rutas, creación de instancia).  
- `panels_report.py`: Script CLI para generar reportes de tenacidad en paneles.  
- `cores_report.py`: Script CLI para generar reportes de compresión axial (testigos).  
- `gen_report.py`: Script genérico para convertir una planilla existente a PDF con formato.  

## Requerimientos

Paquetes Python (ejemplo de instalación en PowerShell):
```powershell
python -m venv .venv
. .\.venv\Scripts\Activate.ps1
pip install numpy scipy pandas matplotlib openpyxl chardet pypdf reportlab pywin32
```
> Nota: `pywin32` y automatización COM (Excel) requieren Windows con Microsoft Excel instalado.

## Uso común de los scripts
Todos los scripts exponen una interfaz CLI homogénea gracias a `report_helpers.parse_common_args`.

Argumentos estándar:
- `--infle`: Identificador principal del informe (obligatorio).
- `--subinfle`: Sufijo opcional.
- `--standard`: Norma o tipo de ensayo.
- `--empresa`: Cliente (alias `--client`).
- `--n`: Número de muestras secuenciales desde 1 (ignorado si se usa `--ids`).
- `--ids`: Lista explícita de IDs (anula `--n`).
- `--base`: Directorio base donde se crea la carpeta `<infle>/` que contendrá salida y buscará datos.

Los scripts añaden parámetros según el contexto (p.ej. opciones de norma en paneles).

## 1. Reporte de Paneles (Tenacidad)
Archivo: `panels_report.py`

Genera reportes de tenacidad en flexión para paneles de acuerdo a normas soportadas: `EFNARC1996`, `EFNARC1999`, `ASTMC1550`.

Ejemplo (2 paneles, IDs 1 y 2):
```powershell
python panels_report.py --infle 090-25 --standard EFNARC1996 --empresa BARCHIP --n 2 --base C:/Datos/Losas
```
Ejemplo (IDs personalizados 3, 5, 8):
```powershell
python panels_report.py --infle 090-25 --standard EFNARC1999 --empresa CLIENTE --ids 3 5 8 --base C:/Datos/Losas
```
Estructura de archivos esperada (uno por panel):
```
<base>/<infle>/
  Losa P1.xlsx
  Losa P2.xlsx
  ...
```
Salida principal: `INFLE_<infle><subinfle>_<standard>_<empresa>[_N].pdf` + archivos intermedios.

## 2. Reporte de Cores (Compresión Axial)
Archivo: `cores_report.py`

Ejemplo (3 testigos secuenciales 1–3):
```powershell
python cores_report.py --infle 061-25 --empresa BCP --n 3 --base C:/Datos/Diamantinas
```
Ejemplo (IDs personalizados):
```powershell
python cores_report.py --infle 061-25 --empresa BCP --ids 2 4 7 --base C:/Datos/Diamantinas
```
Archivos de datos esperados (por ID):
```
<base>/<infle>/
  <infle>-d1/specimen.dat
  <infle>-d2/specimen.dat
  ...
```

## 3. Reporte Genérico (Conversion Excel→PDF)
Archivo: `gen_report.py`

Convierte una planilla Excel ya preparada al PDF final aplicando cabecera/pie y orientación.

Ejemplo:
```powershell
python gen_report.py --infle 336-24 --subinfle -S --standard DM --empresa EXC --base C:/Datos/Diamantinas
```
Se espera encontrar/crear la planilla base dentro de `<base>/<infle>/` siguiendo la convención interna de `Generate_test_report`.

## 4. Helper Común
Archivo: `report_helpers.py`

Funciones clave:
- `parse_common_args(...)`: Construye el parser CLI homogéneo.
- `build_ids(n, ids)`: Retorna la lista final de IDs.
- `prepare_output_dir(base, infle)`: Crea y devuelve ruta `<base>/<infle>/` con barra final.
- `run_report(ClaseReporte, **kwargs)`: Instancia la clase de reporte y ejecuta `make_report_file()`.

## 5. Flujo Interno de Generación (Resumen)
1. Carga de datos (`test_data.py`).
2. Pre-procesamiento (normalización, índices, interpolaciones) en clases de `test_ledi.py`.
3. Escritura de resultados a Excel.
4. Conversión Excel→PDF y post-procesado (`edit_pdfs.py`).
5. Unificación de PDF de datos + figuras + cabeceras.

## 6. Notas y Recomendaciones
- Verifica que Excel no esté abierto con la planilla a convertir (puede generar conflictos COM).
- Si un PDF intermedio queda bloqueado, cerrar procesos Excel (Task Manager) y reintentar.
- Para depuración, puedes ejecutar con un entorno limpio y añadir prints/logging.

## 7. Próximos Mejoras Sugeridas
- Integrar `logging` en lugar de `print`.
- Añadir tests unitarios mínimos (especialmente para `report_helpers` y parsing).
- Validar automáticamente la existencia de cada archivo de datos antes de procesar.
- Agregar opción `--overwrite` y niveles de verbosidad.

---

Cualquier ajuste adicional o automatización CI se puede agregar fácilmente extendiendo este README.
