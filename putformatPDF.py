import win32com.client as win32
from pypdf import PdfReader, PdfWriter, Transformation
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

def convert_excel_to_pdf(excel_path, pdf_path):
    try:
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False  # Opcional: No mostrar Excel durante la operación
        sheets = excel.Workbooks.Open(excel_path)
        sheets.ExportAsFixedFormat(0, pdf_path)  # 0 es el formato PDF
        sheets.Close(False)
        excel.Quit()
        print(f'Archivo guardado como PDF en {pdf_path}')
    except Exception as e:
        print(f'Error al convertir Excel a PDF: {e}')

def merge_pdfs(pdf_list, output_pdf):
    """
    Une múltiples archivos PDF en un solo archivo.

    :param pdf_list: Lista de rutas de archivos PDF a unir.
    :param output_pdf: Ruta del archivo PDF de salida.
    """
    pdf_writer = PdfWriter()

    # Iterar sobre cada archivo PDF y agregar sus páginas al nuevo PDF
    for pdf_path in pdf_list:
        with open(pdf_path, 'rb') as pdf_file:
            pdf_reader = PdfReader(pdf_file)
            for page_num in range(len(pdf_reader.pages)):
                pdf_writer.add_page(pdf_reader.pages[page_num])

    # Guardar el archivo PDF resultante
    with open(output_pdf, 'wb') as output_file:
        pdf_writer.write(output_file)

def normalize_pdf_orientation(input_pdf_path, output_pdf_path, desired_orientation='portrait'):
    """
    Cambia la orientación de todas las páginas de un PDF a retrato ('portrait') o paisaje ('landscape').
    """
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()

    for page_num, page in enumerate(reader.pages):
        # Rotar la página si es necesario para que coincida con la orientación deseada
        if desired_orientation == 'portrait':
            if page.mediabox.width > page.mediabox.height:  # Si está en paisaje
                page.rotate(-90)  # Rotar 90 grados hacia la izquierda
        elif desired_orientation == 'landscape':
            if page.mediabox.width < page.mediabox.height:  # Si está en retrato
                page.rotate(90)  # Rotar 90 grados hacia la derecha

        writer.add_page(page)

    # Guardar el nuevo PDF con las páginas normalizadas en la misma orientación
    with open(output_pdf_path, "wb") as output_pdf:
        writer.write(output_pdf)

def apply_header_footer_pdf(input_pdf_path, header_footer_pdf_path, output_pdf_path):
    """
    Aplica un encabezado y pie de página (de un PDF en orientación portrait) a otro PDF que puede tener diferentes orientaciones.
    
    :param input_pdf_path: Ruta al archivo PDF original.
    :param header_footer_pdf_path: Ruta al archivo PDF que contiene el encabezado y pie de página (en portrait).
    :param output_pdf_path: Ruta de salida para el nuevo PDF.
    """
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()

    # Leer el archivo PDF de encabezado y pie de página
    header_footer_reader = PdfReader(header_footer_pdf_path)
    header_footer_page_port = header_footer_reader.pages[0]  # Asumimos que solo hay una página en el archivo de encabezado/pie
    header_footer_page_land = header_footer_reader.pages[1]  # Asumimos que solo hay una página en el archivo de encabezado/pie

    for page_num, page in enumerate(reader.pages):
        # Obtener el tamaño de la página y orientación
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)

        # Ajustar la rotación del encabezado y pie de página según la orientación de la página principal
        if page_width > page_height:
            # Si la página es landscape (paisaje), rotamos el encabezado y pie de página
            # Aplicar el encabezado y pie de página a la página del archivo original
            page.merge_page(header_footer_page_land)
        else:
            # Si la página es landscape (paisaje), rotamos el encabezado y pie de página
            # Aplicar el encabezado y pie de página a la página del archivo original
            page.merge_page(header_footer_page_port)

        # Agregar la página modificada al nuevo PDF
        writer.add_page(page)

    # Guardar el nuevo archivo PDF con los encabezados y pies de página aplicados
    with open(output_pdf_path, "wb") as output_pdf:
        writer.write(output_pdf)
