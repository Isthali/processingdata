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

# Evitar duplicados: no agregar handler propio; delegar al root.
root_logger = logging.getLogger()
if logger.handlers:
    logger.handlers.clear()
logger.propagate = True

# Opt-in to future pandas behavior to avoid FutureWarning about silent downcasting
try:
    pd.set_option('future.no_silent_downcasting', True)
except Exception:
    pass

__all__ = [
    'detect_encoding', 'import_data_text', 'import_data_excel', 'get_data_excel',
    'write_data_excel', 'write_batch_excel', 'validate_excel_file'
]

_GLOBAL_CONVERSION_WARNING_EMITTED = False
def detect_encoding(file_path: Union[str, Path], sample_size: int = 10000, default: str = 'utf-8') -> str:
    """Detecta la codificación probable de un archivo de texto.

    Args:
        file_path: Ruta al archivo.
        sample_size: Bytes a leer para muestreo (>0).
        default: Codificación a usar si no se detecta.

    Returns:
        Encoding detectado o la codificación por defecto.

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
    coerce_numeric: bool = True,
    min_valid_cols: int = 2,
    debug_sample: bool = False,
    auto_detect_header: bool = True,
    warn_once: bool = True,
    audit_log_dir: Union[str, Path, None] = None,
    audit_head: int = 5,
    audit_tail: int = 5,
    audit_mid: int = 5,
) -> pd.DataFrame:
    """Importa datos de texto con estrategias robustas de limpieza y auditoría.

    Parámetros clave nuevos:
      - auto_detect_header: intenta localizar la línea 'Running Time' y ajusta skiprows.
      - coerce_numeric: fuerza limpieza y conversión numérica columna por columna.
      - min_valid_cols: mínimo de columnas no NaN requerido para conservar una fila.
      - warn_once: evita repetir el warning de conversión fallida.
      - audit_log_dir: si se indica, guarda una muestra audit.csv (head/mid/tail) para trazabilidad.
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

    # Detectar delimitador si se solicita
    if delimiter == 'auto':
        try:
            with file_path.open('r', errors='ignore') as fh:
                raw_lines = []
                for _ in range(400):  # leer líneas suficientes para buscar encabezado y datos
                    try:
                        raw_lines.append(next(fh))
                    except StopIteration:
                        break
        except Exception as e:
            logger.debug(f"No se pudo leer para auto-delimiter: {e}")
            raw_lines = []
        # Intentar detectar encabezado preliminar antes de delimitar
        header_index = None
        for i, line in enumerate(raw_lines):
            lw = line.lower()
            if 'running time' in lw and 'displacement' in lw:
                header_index = i
                break
        data_section = raw_lines[(header_index + 2) if header_index is not None else 0:]
        candidate_delims = ['\t', ',', ';', '|']
        # whitespace flexible se evalúa aparte
        best_delim = '\t'
        best_score = -1
        expected_cols = len(variable_names)
        sample = [ln for ln in data_section if ln.strip()][:50]
        if sample:
            for cand in candidate_delims:
                good = 0
                for ln in sample:
                    parts = ln.strip().split(cand)
                    if len(parts) == expected_cols:
                        good += 1
                score = good / len(sample)
                if score > best_score:
                    best_score = score
                    best_delim = cand
            # Probar whitespace si ninguno alcanza umbral >=0.5
            if best_score < 0.5:
                import re
                good = 0
                for ln in sample:
                    parts = re.split(r'\s+', ln.strip())
                    if len(parts) >= expected_cols:  # puede haber columnas extra de ruido
                        good += 1
                score_ws = good / len(sample)
                if score_ws >= best_score:
                    delimiter = r'\s+'
                else:
                    delimiter = best_delim
            else:
                delimiter = best_delim
            delim_repr = repr(delimiter)
            logger.info(f"Delimitador auto-detectado: {delim_repr} (score={best_score:.2f})")
        else:
            logger.debug("No se pudo muestrear líneas para detectar delimitador; se usará '\t'")
            delimiter = '\t'

    # Auto-detección de encabezado (ajusta skiprows)
    if auto_detect_header:
        try:
            with file_path.open('r', errors='ignore') as fh:
                preview = [next(fh) for _ in range(200)]
        except StopIteration:
            preview = []
        except Exception as e:
            logger.debug(f"No se pudo pre-escANear para encabezado: {e}")
            preview = []
        header_line = None
        for i, line in enumerate(preview):
            low = line.lower()
            if 'running time' in low and ('displacement' in low):
                header_line = i
                break
        if header_line is not None:
            new_skip = header_line + 2  # línea títulos + unidades
            if new_skip != skiprows:
                logger.info(f"Auto-detección encabezado: skiprows {skiprows} -> {new_skip}")
                skiprows = new_skip
        else:
            logger.debug("No se detectó encabezado; se mantiene skiprows proporcionado")

    global _GLOBAL_CONVERSION_WARNING_EMITTED
    warned_conversion = False
    last_error: Optional[Exception] = None
    df: pd.DataFrame | None = None

    for attempt, encoding in enumerate(encodings_to_try[:max_retries + 1]):
        try:
            logger.debug(f"Intento {attempt+1} lectura (encoding={encoding})")
            # Intento directo
            try:
                df = pd.read_csv(
                    str(file_path),
                    sep=delimiter if delimiter != '\\s+' else r'\s+',
                    names=list(variable_names),
                    usecols=range(len(variable_names)),
                    dtype=dtype,
                    skiprows=skiprows,
                    encoding=encoding,
                    engine='python',
                    on_bad_lines='skip'
                )
            except ValueError as ve:
                if coerce_numeric and 'Unable to convert column' in str(ve):
                    col_problem = str(ve).split('column')[-1].strip()
                    if (not warned_conversion) and (not _GLOBAL_CONVERSION_WARNING_EMITTED or not warn_once):
                        logger.warning(f"Fallo conversión directa ({col_problem}). Reintentando con limpieza flexible.")
                        _GLOBAL_CONVERSION_WARNING_EMITTED = True
                    warned_conversion = True
                    df = pd.read_csv(
                        str(file_path),
                        sep=delimiter if delimiter != '\\s+' else r'\s+',
                        names=list(variable_names),
                        usecols=range(len(variable_names)),
                        dtype=str,
                        skiprows=skiprows,
                        encoding=encoding,
                        engine='python',
                        on_bad_lines='skip'
                    )
                    if debug_sample:
                        logger.debug("Muestra cruda:\n" + '\n'.join(df.head(5).astype(str).agg('\t'.join, axis=1)))
                    # Limpieza por columna
                    for col in df.columns:
                        series = (df[col].astype(str)
                                      .str.strip()
                                      .str.replace('\u00a0', ' ', regex=False)
                                      .str.replace(',', '.', regex=False)
                                      .str.replace(r'[^0-9eE+\-\.]+', '', regex=True)
                                 )
                        series = series.replace('', np.nan)
                        df[col] = pd.to_numeric(series, errors='coerce')
                else:
                    raise
            logger.info(f"Datos importados exitosamente: {df.shape[0]} filas, {df.shape[1]} columnas (encoding={encoding})")
            break
        except (UnicodeDecodeError, UnicodeError) as e:
            last_error = e
            if attempt == len(encodings_to_try) - 1:
                logger.error(f"Error de encoding tras {attempt+1} intentos: {e}")
                raise RuntimeError(f"Error de encoding: {e}")
            logger.warning(f"Encoding fallido ({encoding}), probando siguiente")
            continue
        except Exception as e:
            last_error = e
            logger.error(f"Error inesperado importando datos: {e}")
            raise
    if df is None:
        raise RuntimeError(f"No se pudo importar el archivo. Último error: {last_error}")

    # Fallback si todo NaN
    if (df.isna().all(axis=1)).mean() == 1.0:
        logger.warning("Todas las filas NaN tras coerción; intentando autodetección de separador")
        try:
            flex_df = pd.read_csv(
                str(file_path),
                sep=None,
                names=list(variable_names),
                usecols=range(len(variable_names)),
                dtype=str,
                skiprows=skiprows,
                engine='python',
                encoding=encodings_to_try[0],
                on_bad_lines='skip'
            )
            for col in flex_df.columns:
                s = (flex_df[col].astype(str).str.strip().str.replace('\u00a0',' ',regex=False)
                     .str.replace(',', '.', regex=False))
                flex_df[col] = pd.to_numeric(s.replace('', np.nan), errors='coerce')
            if (~flex_df.isna().all(axis=1)).any():
                df = flex_df
                logger.info("Fallback flexible exitoso")
        except Exception as fe:
            logger.error(f"Fallback flexible falló: {fe}")

    # Filtrado de filas
    if dropna:
        initial = len(df)
        df = df.dropna(how='all')
        valid_counts = df.notna().sum(axis=1)
        df = df[valid_counts >= min_valid_cols]
        removed = initial - len(df)
        if removed > 0:
            retention = (len(df) / initial * 100.0) if initial else 0.0
            logger.info(f"Filtrado de filas: eliminadas {removed} (criterio: all-NaN o < {min_valid_cols} válidas) | Retención: {retention:.2f}% ({len(df)}/{initial})")

    # Auditoría
    if audit_log_dir and not df.empty:
        try:
            audit_dir = Path(audit_log_dir)
            audit_dir.mkdir(parents=True, exist_ok=True)
            n = len(df)
            sample_idx = []
            sample_idx.extend(df.index[:audit_head].tolist())
            if n > (audit_head + audit_tail):
                mid_start = max(0, n//2 - audit_mid//2)
                sample_idx.extend(df.index[mid_start: mid_start + audit_mid].tolist())
            sample_idx.extend(df.index[-audit_tail:].tolist())
            sample_idx = sorted(set(sample_idx))
            audit_df = df.loc[sample_idx]
            out_path = audit_dir / f"{file_path.stem}.audit.csv"
            audit_df.to_csv(out_path, index=False)
            logger.debug(f"Auditoría guardada: {out_path}")
        except Exception as ae:
            logger.debug(f"Auditoría no generada: {ae}")

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
    auto_detect_header: bool = True,
    header_search_rows: int = 100,
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

    # Auto detección de fila inicial basada en coincidencia de nombres o patrón numérico
    if auto_detect_header:
        try:
            wb = xl.load_workbook(filename=str(file_path), read_only=True, data_only=True)
            if isinstance(sheet_idx, int):
                ws = wb.worksheets[sheet_idx]
            else:
                ws = wb[sheet_idx]
            target = [v.lower() for v in variable_names]
            header_candidate = None
            for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=header_search_rows, values_only=True), start=1):
                values = [str(v).strip().lower() for v in row if v is not None]
                if not values:
                    continue
                match_count = sum(1 for t in target if t in values)
                if match_count >= max(2, int(0.6 * len(target))):
                    # Asumimos siguiente fila son datos -> skip esa fila
                    header_candidate = r_idx
                    break
            if header_candidate is not None:
                new_skip = header_candidate
                if new_skip != skiprows:
                    logger.info(f"Auto-detección header Excel: skiprows {skiprows} -> {new_skip}")
                    skiprows = new_skip
            wb.close()
        except Exception as e:
            logger.debug(f"No se pudo auto-detectar header Excel: {e}")

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

