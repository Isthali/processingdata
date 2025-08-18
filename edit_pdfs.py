"""Funciones utilitarias para conversión y edición de PDFs y hojas Excel.

Mejoras respecto a la versión previa:
  - Manejo robusto de errores y liberación de recursos COM en convert_excel_to_pdf.
  - Soporte de rango de páginas (From/To) en exportación de Excel.
  - Type hints y docstrings para todas las funciones.
  - Validación de rutas con pathlib y mensajes claros.
  - merge_pdfs omite silenciosamente archivos inexistentes con aviso.
  - normalización de orientación usando rotación condicional.
  - apply_header_footer_pdf permite usar 1 o 2 páginas en el PDF de encabezado/pie (portrait/landscape).
  - Sistema de logging integrado.
  - Validaciones mejoradas y manejo de excepciones específicas.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Iterable, List, Sequence, Union, Optional

import win32com.client as win32
from pypdf import PdfReader, PdfWriter

# Opcional: mantener imports de reportlab si se amplía para generar overlays dinámicos
from reportlab.pdfgen import canvas  # noqa: F401
from reportlab.lib.pagesizes import A4  # noqa: F401
from reportlab.lib.units import mm  # noqa: F401

# Configurar logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# Handler solo si no existe
if not logger.handlers:
    handler = logging.StreamHandler()
    formatter = logging.Formatter('%(name)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)


def convert_excel_to_pdf(
    excel_path: Union[str, Path],
    pdf_path: Union[str, Path],
    pag_i: Optional[int] = None,
    pag_f: Optional[int] = None,
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
    
    Raises:
        FileNotFoundError: Si el archivo Excel no existe.
        FileExistsError: Si el PDF existe y overwrite=False.
        RuntimeError: Si hay errores en la conversión COM.
        ValueError: Si los parámetros de páginas son inválidos.
    """
    # Validaciones de entrada
    if pag_i is not None and pag_i < 1:
        raise ValueError("pag_i debe ser >= 1")
    if pag_f is not None and pag_f < 1:
        raise ValueError("pag_f debe ser >= 1")
    if pag_i is not None and pag_f is not None and pag_i > pag_f:
        raise ValueError("pag_i no puede ser mayor que pag_f")
    
    excel_path = Path(excel_path)
    pdf_path = Path(pdf_path)
    
    if not excel_path.exists():
        raise FileNotFoundError(f"Archivo Excel no encontrado: {excel_path}")
    if pdf_path.exists() and not overwrite:
        raise FileExistsError(f"El PDF destino ya existe: {pdf_path}")

    # Crear directorio destino si no existe
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    
    logger.info(f"Convirtiendo Excel a PDF: {excel_path} -> {pdf_path}")
    if pag_i is not None or pag_f is not None:
        logger.info(f"Rango de páginas: {pag_i or 'inicio'} - {pag_f or 'final'}")

    excel = None
    workbook = None
    try:
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = visible
        excel.DisplayAlerts = False  # Evitar diálogos
        
        workbook = excel.Workbooks.Open(str(excel_path))
        
        # Constantes: 0 = xlTypePDF
        export_kwargs = {}
        if pag_i is not None:
            export_kwargs['From'] = int(pag_i)
        if pag_f is not None:
            export_kwargs['To'] = int(pag_f)
        
        workbook.ExportAsFixedFormat(0, str(pdf_path), **export_kwargs)
        logger.info("Conversión Excel->PDF completada exitosamente")
        
    except Exception as e:
        logger.error(f"Error en conversión Excel->PDF: {e}")
        # Intentar limpiar archivo parcial si existe
        if pdf_path.exists():
            try:
                pdf_path.unlink()
                logger.debug("Archivo PDF parcial eliminado")
            except Exception:
                pass
        raise RuntimeError(f"Error al convertir Excel a PDF: {e}") from e
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
                logger.debug("Workbook cerrado")
            except Exception as e:
                logger.warning(f"Error cerrando workbook: {e}")
        if excel is not None:
            try:
                excel.Quit()
                logger.debug("Excel cerrado")
            except Exception as e:
                logger.warning(f"Error cerrando Excel: {e}")


def merge_pdfs(pdf_list: Sequence[Union[str, Path]], output_pdf: Union[str, Path]) -> None:
    """Combina varios PDFs en un solo archivo.
    
    Args:
        pdf_list: Lista de rutas de PDFs a combinar.
        output_pdf: Ruta del PDF resultante.
    
    Raises:
        ValueError: Si la lista está vacía o el archivo de salida no es válido.
        RuntimeError: Si hay errores en la combinación.
    """
    if not pdf_list:
        raise ValueError("La lista de PDFs no puede estar vacía")
    
    output_pdf = Path(output_pdf)
    
    # Crear directorio destino si no existe
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    
    logger.info(f"Combinando {len(pdf_list)} PDFs en: {output_pdf}")
    
    writer = PdfWriter()
    processed_files = 0
    total_pages = 0
    
    for i, p in enumerate(pdf_list):
        p = Path(p)
        if not p.exists():
            logger.warning(f"PDF {i+1}/{len(pdf_list)} no encontrado, omitiendo: {p}")
            continue
        
        try:
            logger.debug(f"Procesando PDF {i+1}/{len(pdf_list)}: {p.name}")
            with p.open('rb') as fh:
                reader = PdfReader(fh)
                pages_in_file = len(reader.pages)
                
                for page_num, page in enumerate(reader.pages):
                    writer.add_page(page)
                
                total_pages += pages_in_file
                processed_files += 1
                logger.debug(f"Agregadas {pages_in_file} páginas de {p.name}")
                
        except Exception as e:
            logger.error(f"Error procesando {p}: {e}")
            continue
    
    if processed_files == 0:
        raise RuntimeError("No se pudo procesar ningún archivo PDF")
    
    try:
        with output_pdf.open('wb') as out_f:
            writer.write(out_f)
        
        logger.info(f"Combinación completada: {processed_files} archivos, {total_pages} páginas totales")
        
    except Exception as e:
        logger.error(f"Error escribiendo PDF combinado: {e}")
        raise RuntimeError(f"Error escribiendo archivo de salida: {e}") from e


def normalize_pdf_orientation(
    input_pdf_path: Union[str, Path],
    output_pdf_path: Union[str, Path],
    desired_orientation: str = 'portrait'
) -> None:
    """Normaliza orientación de todas las páginas a 'portrait' o 'landscape'.
    
    Args:
        input_pdf_path: Ruta al PDF de entrada.
        output_pdf_path: Ruta al PDF de salida.
        desired_orientation: 'portrait' o 'landscape'.
    
    Raises:
        FileNotFoundError: Si el PDF de entrada no existe.
        ValueError: Si la orientación no es válida.
        RuntimeError: Si hay errores en el procesamiento.
    """
    if desired_orientation not in ('portrait', 'landscape'):
        raise ValueError("desired_orientation debe ser 'portrait' o 'landscape'")
    
    input_pdf_path = Path(input_pdf_path)
    output_pdf_path = Path(output_pdf_path)
    
    if not input_pdf_path.exists():
        raise FileNotFoundError(f"PDF no encontrado: {input_pdf_path}")
    
    # Crear directorio destino si no existe
    output_pdf_path.parent.mkdir(parents=True, exist_ok=True)
    
    logger.info(f"Normalizando orientación a {desired_orientation}: {input_pdf_path}")
    
    try:
        reader = PdfReader(str(input_pdf_path))
        writer = PdfWriter()
        
        pages_rotated = 0
        total_pages = len(reader.pages)
        
        for i, page in enumerate(reader.pages):
            w = float(page.mediabox.width)
            h = float(page.mediabox.height)
            
            page_rotated = False
            if desired_orientation == 'portrait' and w > h:
                page.rotate(-90)
                page.transfer_rotation_to_content()
                page_rotated = True
                pages_rotated += 1
            elif desired_orientation == 'landscape' and w < h:
                page.rotate(90)
                page.transfer_rotation_to_content()
                page_rotated = True
                pages_rotated += 1
            
            writer.add_page(page)
            
            if (i + 1) % 10 == 0:  # Log progreso cada 10 páginas
                logger.debug(f"Procesadas {i + 1}/{total_pages} páginas")
        
        with output_pdf_path.open('wb') as fh:
            writer.write(fh)
        
        logger.info(f"Normalización completada: {pages_rotated}/{total_pages} páginas rotadas")
        
    except Exception as e:
        logger.error(f"Error normalizando orientación: {e}")
        raise RuntimeError(f"Error en normalización de orientación: {e}") from e


def apply_header_footer_pdf(
    input_pdf_path: Union[str, Path],
    header_footer_pdf_path: Union[str, Path],
    output_pdf_path: Union[str, Path]
) -> None:
    """Aplica overlay (encabezado/pie) a cada página.

    El PDF de header/footer puede contener:
      - 1 página: usada para todos los casos.
      - 2 páginas: primera para portrait, segunda para landscape.
    
    Args:
        input_pdf_path: PDF base al que aplicar el overlay.
        header_footer_pdf_path: PDF con encabezado/pie.
        output_pdf_path: PDF resultante.
    
    Raises:
        FileNotFoundError: Si algún archivo de entrada no existe.
        RuntimeError: Si hay errores en el procesamiento.
    """
    input_pdf_path = Path(input_pdf_path)
    header_footer_pdf_path = Path(header_footer_pdf_path)
    output_pdf_path = Path(output_pdf_path)
    
    if not input_pdf_path.exists():
        raise FileNotFoundError(f"PDF base no encontrado: {input_pdf_path}")
    if not header_footer_pdf_path.exists():
        raise FileNotFoundError(f"PDF header/footer no encontrado: {header_footer_pdf_path}")

    # Crear directorio destino si no existe
    output_pdf_path.parent.mkdir(parents=True, exist_ok=True)
    
    logger.info(f"Aplicando header/footer: {input_pdf_path}")
    logger.info(f"Overlay desde: {header_footer_pdf_path}")
    
    try:
        base_reader = PdfReader(str(input_pdf_path))
        overlay_reader = PdfReader(str(header_footer_pdf_path))
        
        pages_overlay = overlay_reader.pages
        if len(pages_overlay) == 0:
            raise RuntimeError("El PDF de header/footer no contiene páginas")
        
        portrait_overlay = pages_overlay[0]
        landscape_overlay = pages_overlay[1] if len(pages_overlay) > 1 else pages_overlay[0]
        
        logger.info(f"Overlay configurado: {len(pages_overlay)} página(s) de plantilla")
        
        writer = PdfWriter()
        total_pages = len(base_reader.pages)
        portrait_count = 0
        landscape_count = 0
        
        for i, page in enumerate(base_reader.pages):
            w = float(page.mediabox.width)
            h = float(page.mediabox.height)
            
            if w > h:
                page.merge_page(landscape_overlay, expand=True)
                landscape_count += 1
            else:
                page.merge_page(portrait_overlay, expand=True)
                portrait_count += 1
            
            writer.add_page(page)
            
            if (i + 1) % 10 == 0:  # Log progreso cada 10 páginas
                logger.debug(f"Procesadas {i + 1}/{total_pages} páginas")
        
        with output_pdf_path.open('wb') as fh:
            writer.write(fh)
        
        logger.info(f"Header/footer aplicado: {total_pages} páginas "
                   f"({portrait_count} portrait, {landscape_count} landscape)")
        
    except Exception as e:
        logger.error(f"Error aplicando header/footer: {e}")
        raise RuntimeError(f"Error en aplicación de header/footer: {e}") from e


def get_pdf_info(pdf_path: Union[str, Path]) -> dict:
    """Obtiene información básica de un archivo PDF.
    
    Args:
        pdf_path: Ruta al archivo PDF.
    
    Returns:
        Diccionario con información del PDF (páginas, tamaño, etc.).
    
    Example:
        >>> info = get_pdf_info('document.pdf')
        >>> print(f"Páginas: {info['page_count']}")
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF no encontrado: {pdf_path}")
    
    logger.debug(f"Obteniendo información de PDF: {pdf_path}")
    
    try:
        reader = PdfReader(str(pdf_path))
        
        info = {
            'file_path': str(pdf_path),
            'file_size_mb': pdf_path.stat().st_size / (1024 * 1024),
            'page_count': len(reader.pages),
            'pages_info': []
        }
        
        for i, page in enumerate(reader.pages):
            w = float(page.mediabox.width)
            h = float(page.mediabox.height)
            orientation = 'landscape' if w > h else 'portrait'
            
            page_info = {
                'page_num': i + 1,
                'width': w,
                'height': h,
                'orientation': orientation,
                'rotation': page.rotation if hasattr(page, 'rotation') else 0
            }
            info['pages_info'].append(page_info)
        
        logger.debug(f"PDF info: {info['page_count']} páginas, {info['file_size_mb']:.1f} MB")
        return info
        
    except Exception as e:
        logger.error(f"Error obteniendo info del PDF: {e}")
        raise RuntimeError(f"Error leyendo información del PDF: {e}") from e


__all__ = [
    'convert_excel_to_pdf',
    'merge_pdfs',
    'normalize_pdf_orientation',
    'apply_header_footer_pdf',
    'get_pdf_info'
]
