# Script Unificado de Reportes - Guía de Uso

## Descripción

El script `unified_report.py` combina la funcionalidad de todos los scripts de generación de reportes individuales en una sola herramienta. Permite generar reportes para diferentes tipos de ensayos mecánicos de manera consistente y eficiente.

## Tipos de Ensayo Soportados

### 1. `cores` - Ensayos de Compresión Axial (Testigos)
- **Clase de reporte**: `Axial_compression_test_report`
- **Estándar por defecto**: CORES
- **Directorio base por defecto**: `D:/`
- **Número de muestras por defecto**: 6

### 2. `panels` - Ensayos de Tenacidad de Paneles
- **Clase de reporte**: `Panel_toughness_test_report`
- **Estándar por defecto**: EFNARC1996
- **Estándares disponibles**: EFNARC1996, EFNARC1999, ASTMC1550
- **Directorio base por defecto**: `D:/`
- **Número de muestras por defecto**: 3

### 3. `panels_residual` - Resistencia Residual con CMOD (Paneles/Vigas)
- **Clase de reporte**: `Panel_Beam_residual_strength_test_report`
- **Estándar por defecto**: EN14488 (también EN14651)
- **Medición base**: CMOD como eje x (puntos de referencia en mm de apertura)
- **Directorio base por defecto**: `D:/`
- **Número de muestras por defecto**: 3

### 4. `beams_residual` - Resistencia Residual de Vigas (ASTM C1609)
- **Clase de reporte**: `Beam_residual_strength_test_report`
- **Estándar por defecto**: ASTMC1609
- **Medición base**: Deflexión como eje x (puntos de referencia en mm)
- **Directorio base por defecto**: `D:/`
- **Número de muestras por defecto**: 3

### 5. `tapas` - Ensayos de Flexión de Tapas de Buzón (Tránsito)
- **Clase de reporte**: `Tapa_buzon_flexion_test_report`
- **Estándar por defecto**: NTP339.111
- **Directorio base por defecto**: `D:/`
- **Número de muestras por defecto**: 3

### 6. `generic` - Conversión Genérica Excel → PDF
- **Clase de reporte**: `Generate_test_report`
- **Estándar por defecto**: DM
- **Directorio base por defecto**: `D:/`
- **Sin muestras específicas** (n=0)

## Sintaxis de Uso

```bash
python unified_report.py <tipo_ensayo> [opciones]
```

### Argumentos Comunes

- `--infle`: **(Requerido)** Identificador del informe (ej: 336-24, 111-25)
- `--subinfle`: Sub-identificador del informe (ej: S, C). Por defecto: vacío
- `--standard`: Estándar del ensayo. Cada tipo tiene su valor por defecto
- `--empresa`: Nombre de la empresa cliente. Por defecto según el tipo
- `--base-dir`: Directorio base donde se generará el reporte
- `--verbose`, `-v`: Habilita salida detallada para diagnóstico

### Argumentos Específicos para Ensayos con Muestras

Los tipos `cores`, `panels`, y `panels_residual` admiten:

- `--n`: Número de muestras a procesar
- `--offset`: Valor inicial para los IDs de muestras (por defecto: 1)

## Ejemplos de Uso

### Ensayos de Compresión Axial (Testigos)

```bash
# Configuración básica con valores por defecto
python unified_report.py cores --infle 336-24 --subinfle S

# Configuración personalizada
python unified_report.py cores --infle 336-24 --subinfle S --standard CORES --empresa "MI_EMPRESA" --n 8 --base-dir "D:/Reportes/Testigos"

# Con salida verbose para diagnóstico
python unified_report.py cores --infle 336-24 --subinfle S --verbose
```

### Ensayos de Tenacidad de Paneles

```bash
# Configuración básica
python unified_report.py panels --infle 111-25 --subinfle C

# Con estándar específico
python unified_report.py panels --infle 111-25 --subinfle C --standard EFNARC1999 --n 5

# Configuración completa personalizada
python unified_report.py panels --infle 111-25 --subinfle C --standard ASTMC1550 --empresa "PRODIMIN" --n 6 --offset 3 --base-dir "D:/Reportes/Paneles"
```

### Ensayos de Resistencia Residual (CMOD)

```bash
# Configuración básica
python unified_report.py panels_residual --infle 111-25 --subinfle C

# Configuración personalizada
python unified_report.py panels_residual --infle 111-25 --subinfle C --standard EN14488 --empresa "PRODIMIN" --n 4 --offset 7

### Ensayos de Resistencia Residual de Vigas (ASTM C1609)

```bash
# Configuración básica (deflexión)
python unified_report.py beams_residual --infle 222-25 --subinfle B --standard ASTMC1609 --n 3

# Configuración personalizada
python unified_report.py beams_residual --infle 222-25 --subinfle B --standard ASTMC1609 --empresa "PRODIMIN" --n 4 --offset 2
```

### Ensayos de Flexión de Tapas de Buzón (Tránsito)

```bash
# Configuración básica
python unified_report.py tapas --infle 245-25 --subinfle A --empresa "FADECO" --n 6

# Configuración personalizada
python unified_report.py tapas --infle 222-25 --subinfle A --standard Tapa_Circular_CA --empresa "PRODIMIN" --n 3 --base-dir "D:/Reportes/Tapas"
```

### Conversión Genérica

```bash
# Configuración básica
python unified_report.py generic --infle 336-24 --subinfle S

# Configuración personalizada
python unified_report.py generic --infle 336-24 --subinfle S --standard DM --empresa "EMPRESA_CLIENTE" --base-dir "D:/Reportes/Genericos"
```

## Archivos de Salida

El script generará automáticamente:

1. **Directorio de trabajo**: `{base-dir}/{infle}/`
2. **Archivo Excel**: Con los datos procesados del ensayo
3. **Archivo PDF**: Versión formateada y lista para entrega

## Registro de Actividad (Logging)

- **Modo normal**: Muestra información básica del proceso
- **Modo verbose** (`--verbose`): Muestra detalles técnicos para diagnóstico y resolución de problemas

## Migración desde Scripts Individuales

### Equivalencias de comandos:

**cores_report.py** → `unified_report.py cores`
```bash
# Antes
python cores_report.py --infle 336-24 --subinfle S --standard CORES --empresa PRODIMIN --n 6

# Ahora
python unified_report.py cores --infle 336-24 --subinfle S --standard CORES --empresa PRODIMIN --n 6
```

**panels_report.py** → `unified_report.py panels`
```bash
# Antes
python panels_report.py --infle 111-25 --subinfle C --standard EFNARC1996 --empresa PRODIMIN --n 3

# Ahora
python unified_report.py panels --infle 111-25 --subinfle C --standard EFNARC1996 --empresa PRODIMIN --n 3
```

**gen_report.py** → `unified_report.py generic`
```bash
# Antes
python gen_report.py --infle 336-24 --subinfle S --standard DM --empresa EXC

# Ahora
python unified_report.py generic --infle 336-24 --subinfle S --standard DM --empresa EXC
```

## Notas técnicas sobre el procesamiento (test_ledi)

- Las funciones de obtención de puntos característicos se han generalizado:
	- `Toughness_mechanical_test.get_defl_cps(x_points=None, x_col='Deflection', include_extra_cols=None)`
	- Permiten elegir la columna del eje x (por ejemplo, `x_col='CMOD'`).
	- `include_extra_cols` añade columnas a devolver si existen (p.ej., incluir 'Deflection' cuando se usa `x_col='CMOD'`).
- Los reportes ya fijan el eje correcto por defecto:
	- `Panel_Beam_residual_strength_test_report`: usa `CMOD` como eje x (EN 14651/EN 14488).
	- `Beam_residual_strength_test_report`: usa `Deflection` como eje x (ASTM C1609).
- No es necesario que el usuario final pase `x_col`; el reporte lo configura internamente.

## Ventajas del Script Unificado

1. **Interfaz consistente**: Todos los tipos de ensayo usan la misma sintaxis
2. **Menos archivos**: Un solo script en lugar de cuatro separados
3. **Mantenimiento simplificado**: Cambios centralizados en una sola ubicación
4. **Mejor documentación**: Ayuda integrada para todos los tipos
5. **Validación mejorada**: Verificación consistente de argumentos
6. **Logging unificado**: Sistema de registro consistente para todos los tipos

## Resolución de Problemas

### Ver ayuda detallada:
```bash
python unified_report.py --help
python unified_report.py cores --help
python unified_report.py panels --help
```

### Activar modo diagnóstico:
```bash
python unified_report.py <tipo> --verbose [otros_argumentos]
```

### Errores comunes:
1. **Directorio no existe**: Se creará automáticamente con advertencia
2. **Número de muestras inválido**: Debe ser >= 0
3. **Estándar no válido**: Verificar opciones disponibles con `--help`
