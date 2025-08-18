# Script Unificado de Reportes - Guía de Uso

## Descripción

El script `unified_report.py` combina la funcionalidad de todos los scripts de generación de reportes individuales en una sola herramienta. Permite generar reportes para diferentes tipos de ensayos mecánicos de manera consistente y eficiente.

## Tipos de Ensayo Soportados

### 1. `cores` - Ensayos de Compresión Axial (Testigos)
- **Clase de reporte**: `Axial_compression_test_report`
- **Estándar por defecto**: CORES
- **Directorio base por defecto**: `C:/Users/joela/Documents/MATLAB/Diamantinas`
- **Número de muestras por defecto**: 6

### 2. `panels` - Ensayos de Tenacidad de Paneles
- **Clase de reporte**: `Panel_toughness_test_report`
- **Estándar por defecto**: EFNARC1996
- **Estándares disponibles**: EFNARC1996, EFNARC1999, ASTMC1550
- **Directorio base por defecto**: `C:/Users/joela/Documents/MATLAB/Losas`
- **Número de muestras por defecto**: 3

### 3. `panels_residual` - Ensayos de Resistencia Residual
- **Clase de reporte**: `Panel_Beam_residual_strength_test_report`
- **Estándar por defecto**: EN14488
- **Directorio base por defecto**: `C:/Users/joela/Documents/MATLAB/Losas`
- **Número de muestras por defecto**: 3

### 4. `generic` - Conversión Genérica Excel → PDF
- **Clase de reporte**: `Generate_test_report`
- **Estándar por defecto**: DM
- **Directorio base por defecto**: `C:/Users/joela/Documents/MATLAB/Diamantinas`
- **Sin muestras específicas** (n=0)

## Sintaxis de Uso

```bash
python unified_report.py <tipo_ensayo> [opciones]
```

### Argumentos Comunes

- `--infle`: **(Requerido)** Identificador del informe (ej: 336-24, 111-25)
- `--subinfle`: Sub-identificador del informe (ej: -S, -C). Por defecto: vacío
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
python unified_report.py cores --infle 336-24 --subinfle -S

# Configuración personalizada
python unified_report.py cores --infle 336-24 --subinfle -S --standard CORES --empresa "MI_EMPRESA" --n 8 --base-dir "D:/Reportes/Testigos"

# Con salida verbose para diagnóstico
python unified_report.py cores --infle 336-24 --subinfle -S --verbose
```

### Ensayos de Tenacidad de Paneles

```bash
# Configuración básica
python unified_report.py panels --infle 111-25 --subinfle -C

# Con estándar específico
python unified_report.py panels --infle 111-25 --subinfle -C --standard EFNARC1999 --n 5

# Configuración completa personalizada
python unified_report.py panels --infle 111-25 --subinfle -C --standard ASTMC1550 --empresa "PRODIMIN" --n 6 --offset 3 --base-dir "D:/Reportes/Paneles"
```

### Ensayos de Resistencia Residual

```bash
# Configuración básica
python unified_report.py panels_residual --infle 111-25 --subinfle -C

# Configuración personalizada
python unified_report.py panels_residual --infle 111-25 --subinfle -C --standard EN14488 --empresa "PRODIMIN" --n 4 --offset 7
```

### Conversión Genérica

```bash
# Configuración básica
python unified_report.py generic --infle 336-24 --subinfle -S

# Configuración personalizada
python unified_report.py generic --infle 336-24 --subinfle -S --standard DM --empresa "EMPRESA_CLIENTE" --base-dir "D:/Reportes/Genericos"
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
python cores_report.py --infle 336-24 --subinfle -S --standard CORES --empresa PRODIMIN --n 6

# Ahora
python unified_report.py cores --infle 336-24 --subinfle -S --standard CORES --empresa PRODIMIN --n 6
```

**panels_report.py** → `unified_report.py panels`
```bash
# Antes
python panels_report.py --infle 111-25 --subinfle -C --standard EFNARC1996 --empresa PRODIMIN --n 3

# Ahora
python unified_report.py panels --infle 111-25 --subinfle -C --standard EFNARC1996 --empresa PRODIMIN --n 3
```

**gen_report.py** → `unified_report.py generic`
```bash
# Antes
python gen_report.py --infle 336-24 --subinfle -S --standard DM --empresa EXC

# Ahora
python unified_report.py generic --infle 336-24 --subinfle -S --standard DM --empresa EXC
```

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
