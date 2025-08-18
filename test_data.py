"""Utilidades de importación y escritura de datos para ensayos mecánicos.

Mejoras incluidas:
  - Detección de encoding con fallback robusto.
  - Parámetros configurables (skiprows, dtype).
  - Manejo de errores y validaciones explícitas.
  - Evita errores silenciosos al escribir en hojas inexistentes (opción de crear).
  - Type hints y docstrings para mantenimiento.
  - Logging configurado para debug y tracking.
  - Validaciones de entrada mejoradas.
"""

from __future__ import annotations

import logging
import pandas as pd
import numpy as np
import openpyxl as xl
import chardet
from pathlib import Path
from typing import Iterable, List, Sequence, Tuple, Union, Any, Optional

# Configurar logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# Handler solo si no existe
if not logger.handlers:
    handler = logging.StreamHandler()
    formatter = logging.Formatter('%(name)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)


def detect_encoding(file_path: Union[str, Path], sample_size: int = 10000, default: str = 'utf-8') -> str:
    """Detecta la codificación probable de un archivo binario.

    Args:
        file_path: Ruta al archivo.
        sample_size: Número de bytes a muestrear (debe ser > 0).
        default: Codificación por defecto si chardet no determina.
    Returns:
        Cadena con la codificación detectada o la predeterminada.
    
    Raises:
        FileNotFoundError: Si el archivo no existe.
        ValueError: Si sample_size <= 0.
    """
    if sample_size <= 0:
        raise ValueError("sample_size debe ser mayor que 0")
    
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Archivo no encontrado: {file_path}")
    
    logger.debug(f"Detectando encoding de: {file_path}")
    
    try:
        with file_path.open('rb') as f:
            data = f.read(sample_size)
        result = chardet.detect(data) or {}
        encoding = result.get('encoding') or default
        confidence = result.get('confidence', 0.0)
        
        logger.debug(f"Encoding detectado: {encoding} (confianza: {confidence:.2f})")
        return encoding
    except Exception as e:
        logger.warning(f"Error detectando encoding, usando {default}: {e}")
        return default


def import_data_text(
    file_path: Union[str, Path],
    delimiter: str,
    variable_names: Sequence[str],
    skiprows: int = 15,
    dtype: Any = np.float64,
    dropna: bool = True,
    max_retries: int = 2,
) -> pd.DataFrame:
    """Importa datos desde archivo de texto delimitado.

    Args:
        file_path: Ruta al archivo.
        delimiter: Separador de columnas.
        variable_names: Nombres de columnas a asignar.
        skiprows: Filas a omitir al inicio.
        dtype: Tipo de datos para conversión.
        dropna: Eliminar filas con NaN.
        max_retries: Número máximo de reintentos con diferentes encodings.
    
    Returns:
        DataFrame con los datos importados.
    
    Example:
        >>> df = import_data_text('data.txt', '\t', ['Time', 'Force'], skiprows=10)
    """
    if skiprows < 0:
        raise ValueError("skiprows debe ser >= 0")
    if not variable_names:
        raise ValueError("variable_names no puede estar vacío")
    
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Archivo no encontrado: {file_path}")
    
    logger.info(f"Importando datos de texto: {file_path}")
    
    encodings_to_try = [detect_encoding(file_path), 'utf-8', 'latin-1', 'cp1252']
    
    for attempt, encoding in enumerate(encodings_to_try[:max_retries + 1]):
        try:
            logger.debug(f"Intento {attempt + 1} con encoding: {encoding}")
            df = pd.read_csv(
                filepath_or_buffer=str(file_path),
                sep=delimiter,
                names=list(variable_names),
                dtype=dtype,
                skiprows=skiprows,
                encoding=encoding,
                engine='python',
                on_bad_lines='skip'
            )
            
            logger.info(f"Datos importados exitosamente: {df.shape[0]} filas, {df.shape[1]} columnas")
            break
            
        except (UnicodeDecodeError, UnicodeError) as e:
            if attempt == len(encodings_to_try) - 1:
                logger.error(f"No se pudo decodificar el archivo después de {max_retries + 1} intentos")
                raise RuntimeError(f"Error de encoding después de {max_retries + 1} intentos: {e}")
            logger.warning(f"Error con encoding {encoding}, intentando siguiente: {e}")
            continue
        except Exception as e:
            logger.error(f"Error inesperado importando datos: {e}")
            raise
    
    if dropna:
        initial_rows = len(df)
        df = df.dropna()
        dropped_rows = initial_rows - len(df)
        if dropped_rows > 0:
            logger.info(f"Eliminadas {dropped_rows} filas con valores NaN")
    
    if df.empty:
        logger.warning("DataFrame resultante está vacío")
    
    return df


def import_data_excel(
    file_path: Union[str, Path],
    sheet_idx: Union[int, str],
    variable_names: Sequence[str],
    skiprows: int = 49,
    dtype: Any = np.float64,
    dropna: bool = True,
) -> pd.DataFrame:
    """Importa datos desde una hoja de Excel.

    Args:
        file_path: Ruta al archivo Excel.
        sheet_idx: Índice (int) o nombre (str) de la hoja.
        variable_names: Nombres de columnas a asignar.
        skiprows: Filas a omitir al inicio.
        dtype: Tipo de datos para conversión.
        dropna: Eliminar filas con NaN.
    
    Returns:
        DataFrame con los datos importados.
    
    Example:
        >>> df = import_data_excel('data.xlsx', 0, ['Time', 'Force', 'Disp'])
    """
    if skiprows < 0:
        raise ValueError("skiprows debe ser >= 0")
    if not variable_names:
        raise ValueError("variable_names no puede estar vacío")
    
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Archivo no encontrado: {file_path}")
    
    logger.info(f"Importando datos de Excel: {file_path}, hoja: {sheet_idx}")
    
    try:
        df = pd.read_excel(
            io=str(file_path),
            sheet_name=sheet_idx,
            names=list(variable_names),
            dtype=dtype,
            skiprows=skiprows,
            engine='openpyxl'
        )
        
        logger.info(f"Datos importados exitosamente: {df.shape[0]} filas, {df.shape[1]} columnas")
        
    except ValueError as e:
        if "Worksheet" in str(e):
            raise ValueError(f"La hoja '{sheet_idx}' no existe en el archivo Excel")
        raise
    except Exception as e:
        logger.error(f"Error importando Excel: {e}")
        raise
    
    if dropna:
        initial_rows = len(df)
        df = df.dropna()
        dropped_rows = initial_rows - len(df)
        if dropped_rows > 0:
            logger.info(f"Eliminadas {dropped_rows} filas con valores NaN")
    
    if df.empty:
        logger.warning("DataFrame resultante está vacío")
    
    return df


def get_data_excel(
    file_path: Union[str, Path],
    sheet_idx: Union[int, str, None],
    position: Tuple[int, int] = (1, 1)
) -> Any:
    """Obtiene un valor de una celda de un archivo Excel.

    Args:
        file_path: Ruta al archivo.
        sheet_idx: Índice (0-based) o nombre de la hoja. Si None usa la primera.
        position: (fila, columna) 1-based.
    
    Returns:
        Valor de la celda especificada.
    
    Example:
        >>> value = get_data_excel('data.xlsx', 'Results', (2, 3))
    """
    row, column = position
    if row < 1 or column < 1:
        raise ValueError("Las posiciones deben ser >= 1 (1-based indexing)")
    
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Archivo no encontrado: {file_path}")
    
    logger.debug(f"Leyendo celda ({row}, {column}) de {file_path}, hoja: {sheet_idx}")
    
    wb = xl.load_workbook(filename=str(file_path), keep_vba=True, read_only=True)
    try:
        if sheet_idx is None:
            sheet = wb.worksheets[0]
            logger.debug("Usando primera hoja del workbook")
        elif isinstance(sheet_idx, int):
            if sheet_idx >= len(wb.worksheets):
                raise ValueError(f"Índice de hoja {sheet_idx} fuera de rango (máximo: {len(wb.worksheets) - 1})")
            sheet = wb.worksheets[sheet_idx]
        else:
            if sheet_idx not in wb.sheetnames:
                raise ValueError(f"Hoja '{sheet_idx}' no encontrada. Hojas disponibles: {wb.sheetnames}")
            sheet = wb[sheet_idx]
        
        val = sheet.cell(row, column).value
        logger.debug(f"Valor leído: {val}")
        return val
        
    except Exception as e:
        logger.error(f"Error leyendo celda: {e}")
        raise
    finally:
        wb.close()


def write_data_excel(
    file_path: Union[str, Path],
    sheet_name: str,
    position: Tuple[int, int],
    val: Any,
    create_sheet: bool = True
) -> None:
    """Escribe un valor en una celda de Excel.

    Args:
        file_path: Ruta al archivo existente.
        sheet_name: Nombre de hoja destino.
        position: (fila, columna) 1-based.
        val: Valor a escribir.
        create_sheet: Crear la hoja si no existe.
    
    Example:
        >>> write_data_excel('results.xlsx', 'Data', (1, 1), 42.5)
    """
    row, column = position
    if row < 1 or column < 1:
        raise ValueError("Las posiciones deben ser >= 1 (1-based indexing)")
    if not sheet_name.strip():
        raise ValueError("sheet_name no puede estar vacío")
    
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Archivo no encontrado: {file_path}")
    
    logger.debug(f"Escribiendo valor {val} en ({row}, {column}) de hoja '{sheet_name}'")
    
    wb = xl.load_workbook(filename=str(file_path), keep_vba=True)
    try:
        if sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
        else:
            if create_sheet:
                sheet = wb.create_sheet(title=sheet_name)
                logger.info(f"Creada nueva hoja: {sheet_name}")
            else:
                raise ValueError(f"La hoja '{sheet_name}' no existe y create_sheet=False.")
        
        sheet.cell(row, column, value=val)
        wb.save(str(file_path))
        logger.debug(f"Valor escrito y archivo guardado exitosamente")
        
    except Exception as e:
        logger.error(f"Error escribiendo en Excel: {e}")
        raise
    finally:
        wb.close()


def write_batch_excel(
    file_path: Union[str, Path],
    sheet_name: str,
    data: Sequence[Tuple[int, int, Any]],
    create_sheet: bool = True
) -> None:
    """Escritura por lotes para mejorar rendimiento (abre/cierra una sola vez).
    
    Args:
        file_path: Ruta al archivo Excel.
        sheet_name: Nombre de la hoja destino.
        data: Secuencia de tuplas (fila, columna, valor) para escribir.
        create_sheet: Crear la hoja si no existe.
    
    Example:
        >>> batch_data = [(1, 1, 'Fecha'), (1, 2, 'Fuerza'), (2, 1, '2023-01-01')]
        >>> write_batch_excel('results.xlsx', 'Data', batch_data)
    """
    if not data:
        logger.warning("Lista de datos vacía, no se escribirá nada")
        return
    
    if not sheet_name.strip():
        raise ValueError("sheet_name no puede estar vacío")
    
    # Validar formato de datos
    for i, item in enumerate(data):
        if not isinstance(item, (tuple, list)) or len(item) != 3:
            raise ValueError(f"Elemento {i} debe ser tupla (fila, columna, valor)")
        row, col, _ = item
        if row < 1 or col < 1:
            raise ValueError(f"Elemento {i}: posiciones deben ser >= 1 (1-based)")
    
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Archivo no encontrado: {file_path}")
    
    logger.info(f"Escribiendo {len(data)} valores en lote en hoja '{sheet_name}'")
    
    wb = xl.load_workbook(filename=str(file_path), keep_vba=True)
    try:
        if sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
        else:
            if create_sheet:
                sheet = wb.create_sheet(title=sheet_name)
                logger.info(f"Creada nueva hoja: {sheet_name}")
            else:
                raise ValueError(f"La hoja '{sheet_name}' no existe y create_sheet=False.")
        
        for row, col, value in data:
            sheet.cell(row, col, value=value)
        
        wb.save(str(file_path))
        logger.info(f"Escritura en lote completada exitosamente")
        
    except Exception as e:
        logger.error(f"Error en escritura por lotes: {e}")
        raise
    finally:
        wb.close()


def validate_excel_file(file_path: Union[str, Path]) -> dict:
    """Valida un archivo Excel y retorna información sobre sus hojas.
    
    Args:
        file_path: Ruta al archivo Excel.
    
    Returns:
        Diccionario con información del archivo: hojas, dimensiones, etc.
    
    Example:
        >>> info = validate_excel_file('data.xlsx')
        >>> print(info['sheets'])
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Archivo no encontrado: {file_path}")
    
    logger.info(f"Validando archivo Excel: {file_path}")
    
    wb = xl.load_workbook(filename=str(file_path), read_only=True)
    try:
        info = {
            'file_path': str(file_path),
            'file_size_mb': file_path.stat().st_size / (1024 * 1024),
            'sheets': [],
            'total_sheets': len(wb.worksheets)
        }
        
        for i, sheet in enumerate(wb.worksheets):
            sheet_info = {
                'index': i,
                'name': sheet.title,
                'max_row': sheet.max_row,
                'max_column': sheet.max_column,
                'dimensions': f"{sheet.max_row}x{sheet.max_column}"
            }
            info['sheets'].append(sheet_info)
        
        logger.info(f"Archivo validado: {info['total_sheets']} hojas")
        return info
        
    except Exception as e:
        logger.error(f"Error validando archivo: {e}")
        raise
    finally:
        wb.close()

