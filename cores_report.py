import os
from openpyxl import load_workbook
from matplotlib.backends.backend_pdf import PdfPages
#from putformatPDF import convert_excel_to_pdf, merge_pdfs, normalize_pdf_orientation, apply_header_footer_pdf
from test_ledi import Axial_compression_test_report

# Parámetros iniciales
infle = '274-24'
subinfle = ''
empresa = 'TECCA'
cores_id = [id+1 for id in range(24)]
acred = 'acreditado'

# Directorios
base_dir = f'C:/Users/joela/Documents/MATLAB/Diamantinas/{infle}/'
os.makedirs(base_dir, exist_ok=True)

# Crear el informe en excel
test_report = Axial_compression_test_report(infle=infle, subinfle=subinfle, folder=base_dir, empresa=empresa, samples_id=cores_id)
test_report.add_tests()
test_report.write_report()
test_report.plot_report()

# Convertir a pdf la parte del informe en excel
#excel_path = f'{base_dir}INFLE_{infle}{subinfle}_Cores_{empresa}_{num_cores}.xlsm'
#pdf_path = f'{base_dir}cuadros.pdf'
#convert_excel_to_pdf(excel_path, pdf_path)

# Unir y normalizar la orientacion de los archivos que forman parte del informe
#pdf_files = [f'{base_dir}cuadros.pdf', f'{base_dir}graficos.pdf']  # Lista de archivos PDF
#output_pdf = f'{base_dir}informe_completo.pdf'  # Archivo PDF de salida
#merge_pdfs(pdf_files, output_pdf)
#normalize_pdf_orientation(output_pdf, output_pdf, desired_orientation='portrait')

# Agregar el encabezado y pie de página
#header_footer_pdf_path = f'C:/Users/joela/Documents/PYTHON/formato_{acred}.pdf'           # PDF temporal con el encabezado
#apply_header_footer_pdf(output_pdf, header_footer_pdf_path, output_pdf)

#print("Proceso completado. PDF final generado con éxito.")

# Limpieza
#plt.close('all')
