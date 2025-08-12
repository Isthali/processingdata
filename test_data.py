"""Utilidades de importación y escritura de datos para ensayos mecánicos.

Mejoras incluidas:
  - Detección de encoding con fallback robusto.
  - Parámetros configurables (skiprows, dtype).
  - Manejo de errores y validaciones explícitas.
  - Evita errores silenciosos al escribir en hojas inexistentes (opción de crear).
  - Type hints y docstrings para mantenimiento.
"""

from __future__ import annotations

import pandas as pd
import numpy as np
import openpyxl as xl
import chardet
from pathlib import Path
from typing import Iterable, List, Sequence, Tuple, Union, Any


def detect_encoding(file_path: Union[str, Path], sample_size: int = 10000, default: str = 'utf-8') -> str:
    """Detecta la codificación probable de un archivo binario.

    Args:
        file_path: Ruta al archivo.
        sample_size: Número de bytes a muestrear.
        default: Codificación por defecto si chardet no determina.
    Returns:
        Cadena con la codificación detectada o la predeterminada.
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Archivo no encontrado: {file_path}")
    with file_path.open('rb') as f:
        data = f.read(sample_size)
    result = chardet.detect(data) or {}
    encoding = result.get('encoding') or default
    return encoding


def import_data_text(
    file_path: Union[str, Path],
    delimiter: str,
    variable_names: Sequence[str],
    skiprows: int = 15,
    dtype: Any = np.float64,
    dropna: bool = True,
) -> pd.DataFrame:
    """Importa datos desde archivo de texto delimitado.

    Args:
        file_path: Ruta al archivo.
        delimiter: Separador de columnas.
        variable_names: Nombres de columnas a asignar.
        skiprows: Filas a omitir al inicio.
        dtype: Tipo de datos para conversión.
        dropna: Eliminar filas con NaN.
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Archivo no encontrado: {file_path}")
    encoding = detect_encoding(file_path)
    try:
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
    except UnicodeDecodeError:
        # Reintento con utf-8
        df = pd.read_csv(
            filepath_or_buffer=str(file_path),
            sep=delimiter,
            names=list(variable_names),
            dtype=dtype,
            skiprows=skiprows,
            encoding='utf-8',
            engine='python',
            on_bad_lines='skip'
        )
    if dropna:
        df = df.dropna()
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

    sheet_idx acepta índice (int) o nombre (str).
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Archivo no encontrado: {file_path}")
    df = pd.read_excel(
        io=str(file_path),
        sheet_name=sheet_idx,
        names=list(variable_names),
        dtype=dtype,
        skiprows=skiprows,
        engine='openpyxl'
    )
    if dropna:
        df = df.dropna()
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
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Archivo no encontrado: {file_path}")
    wb = xl.load_workbook(filename=str(file_path), keep_vba=True, read_only=True)
    try:
        if sheet_idx is None:
            sheet = wb.worksheets[0]
        elif isinstance(sheet_idx, int):
            sheet = wb.worksheets[sheet_idx]
        else:
            sheet = wb[sheet_idx]
        row, column = position
        val = sheet.cell(row, column).value
    finally:
        wb.close()
    return val


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
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Archivo no encontrado: {file_path}")
    wb = xl.load_workbook(filename=str(file_path), keep_vba=True)
    try:
        if sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
        else:
            if create_sheet:
                sheet = wb.create_sheet(title=sheet_name)
            else:
                raise ValueError(f"La hoja '{sheet_name}' no existe y create_sheet=False.")
        row, column = position
        sheet.cell(row, column, value=val)
        wb.save(str(file_path))
    finally:
        wb.close()


def write_batch_excel(
    file_path: Union[str, Path],
    sheet_name: str,
    data: Sequence[Tuple[int, int, Any]],
    create_sheet: bool = True
) -> None:
    """Escritura por lotes para mejorar rendimiento (abre/cierra una sola vez)."""
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Archivo no encontrado: {file_path}")
    wb = xl.load_workbook(filename=str(file_path), keep_vba=True)
    try:
        if sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
        else:
            if create_sheet:
                sheet = wb.create_sheet(title=sheet_name)
            else:
                raise ValueError(f"La hoja '{sheet_name}' no existe y create_sheet=False.")
        for row, col, value in data:
            sheet.cell(row, col, value=value)
        wb.save(str(file_path))
    finally:
        wb.close()

