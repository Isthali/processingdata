"""Funciones utilitarias para conversión y edición de PDFs y hojas Excel.

Mejoras respecto a la versión previa:
  - Manejo robusto de errores y liberación de recursos COM en convert_excel_to_pdf.
  - Soporte de rango de páginas (From/To) en exportación de Excel.
  - Type hints y docstrings para todas las funciones.
  - Validación de rutas con pathlib y mensajes claros.
  - merge_pdfs omite silenciosamente archivos inexistentes con aviso.
  - normalización de orientación usando rotación condicional.
  - apply_header_footer_pdf permite usar 1 o 2 páginas en el PDF de encabezado/pie (portrait/landscape).
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable, List, Sequence, Union

import win32com.client as win32
from pypdf import PdfReader, PdfWriter

# Opcional: mantener imports de reportlab si se amplía para generar overlays dinámicos
from reportlab.pdfgen import canvas  # noqa: F401
from reportlab.lib.pagesizes import A4  # noqa: F401
from reportlab.lib.units import mm  # noqa: F401


def convert_excel_to_pdf(
    excel_path: Union[str, Path],
    pdf_path: Union[str, Path],
    pag_i: int | None = None,
    pag_f: int | None = None,
    visible: bool = False,
    overwrite: bool = True,
) -> None:
    """Convierte un archivo Excel a PDF usando COM (solo Windows).

    Args:
        excel_path: Ruta al archivo origen (.xls/.xlsx/.xlsm).
        pdf_path: Ruta destino del PDF.
        pag_i: Página inicial (1-based) opcional.
        pag_f: Página final (1-based) opcional.
        visible: Mostrar la ventana de Excel.
        overwrite: Si True sobrescribe PDF existente.
    """
    excel_path = Path(excel_path)
    pdf_path = Path(pdf_path)
    if not excel_path.exists():
        raise FileNotFoundError(f"Archivo Excel no encontrado: {excel_path}")
    if pdf_path.exists() and not overwrite:
        raise FileExistsError(f"El PDF destino ya existe: {pdf_path}")

    excel = None
    workbook = None
    try:
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = visible
        workbook = excel.Workbooks.Open(str(excel_path))
        # Constantes: 0 = xlTypePDF
        export_kwargs = {}
        if pag_i is not None:
            export_kwargs['From'] = int(pag_i)
        if pag_f is not None:
            export_kwargs['To'] = int(pag_f)
        workbook.ExportAsFixedFormat(0, str(pdf_path), **export_kwargs)
    except Exception as e:  # pylint: disable=broad-except
        raise RuntimeError(f"Error al convertir Excel a PDF: {e}") from e
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:  # noqa: BLE001
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:  # noqa: BLE001
                pass


def merge_pdfs(pdf_list: Sequence[Union[str, Path]], output_pdf: Union[str, Path]) -> None:
    """Combina varios PDFs en un solo archivo."""
    output_pdf = Path(output_pdf)
    writer = PdfWriter()
    for p in pdf_list:
        p = Path(p)
        if not p.exists():
            print(f"Aviso: se omite PDF inexistente: {p}")
            continue
        try:
            with p.open('rb') as fh:
                reader = PdfReader(fh)
                for page in reader.pages:
                    writer.add_page(page)
        except Exception as e:  # noqa: BLE001
            print(f"Error leyendo {p}: {e}")
    with output_pdf.open('wb') as out_f:
        writer.write(out_f)


def normalize_pdf_orientation(
    input_pdf_path: Union[str, Path],
    output_pdf_path: Union[str, Path],
    desired_orientation: str = 'portrait'
) -> None:
    """Normaliza orientación de todas las páginas a 'portrait' o 'landscape'."""
    input_pdf_path = Path(input_pdf_path)
    output_pdf_path = Path(output_pdf_path)
    if not input_pdf_path.exists():
        raise FileNotFoundError(f"PDF no encontrado: {input_pdf_path}")
    reader = PdfReader(str(input_pdf_path))
    writer = PdfWriter()
    for page in reader.pages:
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)
        if desired_orientation == 'portrait' and w > h:
            page.rotate(-90)
            page.transfer_rotation_to_content()
        elif desired_orientation == 'landscape' and w < h:
            page.rotate(90)
            page.transfer_rotation_to_content()
        writer.add_page(page)
    with output_pdf_path.open('wb') as fh:
        writer.write(fh)


def apply_header_footer_pdf(
    input_pdf_path: Union[str, Path],
    header_footer_pdf_path: Union[str, Path],
    output_pdf_path: Union[str, Path]
) -> None:
    """Aplica overlay (encabezado/pie) a cada página.

    El PDF de header/footer puede contener:
      - 1 página: usada para todos los casos.
      - 2 páginas: primera para portrait, segunda para landscape.
    """
    input_pdf_path = Path(input_pdf_path)
    header_footer_pdf_path = Path(header_footer_pdf_path)
    output_pdf_path = Path(output_pdf_path)
    if not input_pdf_path.exists():
        raise FileNotFoundError(f"PDF base no encontrado: {input_pdf_path}")
    if not header_footer_pdf_path.exists():
        raise FileNotFoundError(f"PDF header/footer no encontrado: {header_footer_pdf_path}")

    base_reader = PdfReader(str(input_pdf_path))
    overlay_reader = PdfReader(str(header_footer_pdf_path))
    pages_overlay = overlay_reader.pages
    portrait_overlay = pages_overlay[0]
    landscape_overlay = pages_overlay[1] if len(pages_overlay) > 1 else pages_overlay[0]
    writer = PdfWriter()
    for page in base_reader.pages:
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)
        if w > h:
            page.merge_page(landscape_overlay, expand=True)
        else:
            page.merge_page(portrait_overlay, expand=True)
        writer.add_page(page)
    with output_pdf_path.open('wb') as fh:
        writer.write(fh)


__all__ = [
    'convert_excel_to_pdf',
    'merge_pdfs',
    'normalize_pdf_orientation',
    'apply_header_footer_pdf'
]
