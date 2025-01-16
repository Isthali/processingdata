import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
from openpyxl import load_workbook
from matplotlib.backends.backend_pdf import PdfPages
from putformatPDF import convert_excel_to_pdf, merge_pdfs, normalize_pdf_orientation, apply_header_footer_pdf
from import_test_data import import_test_data_text

# Parámetros iniciales
infle = '274-24'
subinfle = ''
empresa = 'TECCA'
num_cores = 24
num_1core = 1
num_variables = 3
data_lines = [15, None]  # Equivalente a [15, Inf] en MATLAB
delimiter = '\t'
variable_names = ['t', 'D', 'P']
variable_types = [float, float, float]
acred = 'acreditado'

# Directorios
base_dir = f'C:/Users/joela/Documents/PYTHON/Cores/{infle}/'
os.makedirs(base_dir, exist_ok=True)

# Función para calcular valores de control (basada en ControlValuesCORE.m)
def control_values_core(D, P):
    # Eliminar el desplazamiento inicial y tomar valores absolutos
    D = abs(D - D.iloc[0])
    P = abs(P)
    
    # Carga máxima
    imaxP = P.idxmax()
    
    # Límite de desplazamiento
    ilim = (abs(P[imaxP:] - 0.8 * P[imaxP])).idxmin() + imaxP
    
    D_mod = D[:ilim]
    P_mod = P[:ilim]
    
    CoreControlValues = np.array([[D.iloc[imaxP]], [P.iloc[imaxP]]])
    return CoreControlValues, D_mod, P_mod

# Archivo PDF único donde se guardarán todas las figuras
combined_pdf_path = f'{base_dir}graficos.pdf'

with PdfPages(combined_pdf_path) as pdf:
    for m in range(num_1core, num_1core + num_cores):
        file_path = f'{base_dir}{infle}-d{m}/specimen.dat'
        core = import_test_data_text(file_path, num_variables, data_lines, delimiter, variable_names, variable_types)

        coreControlValuesD, core_D, core_P = control_values_core(core['D'], core['P'])

        plt.figure(figsize=(10, 8))
        plt.plot(core_D, core_P, linewidth=2)
        str1 = f'P_max = {round(coreControlValuesD[1, 0], 2)}'
        plt.axvline(x=coreControlValuesD[0, 0], color='b', linestyle='--', label=str1)
        plt.title(f'Fuerza - Deformación: Testigo {m}')
        plt.xlabel('Deformación Axial [mm]', fontsize=9)
        plt.ylabel('Fuerza Axial [kN]', fontsize=9)
        plt.axis([0, max(core_D) * 1.1, 0, max(core_P) * 1.1])
        plt.grid(True, which='both', linestyle='--', linewidth=0.5)
        plt.figtext(0.1, 0.03, f'INF-LE {infle}{subinfle}', fontsize=9, ha='left')
        plt.figtext(0.5, 0.03, 'Ensayo de Resistencia a la Compresión', fontsize=9, ha='center')
        plt.figtext(0.9, 0.03, f'Pág. {m + 3}', fontsize=9, ha='right')

        if m == num_cores:
            plt.figtext(0.9, 0.01, 'Fin del informe', fontsize=9, ha='right')

        # Agregar la figura actual al archivo PDF
        pdf.savefig()
        plt.close()

        # Guardar valores en un archivo Excel (necesitas openpyxl o xlsxwriter)
        with pd.ExcelWriter(f'{base_dir}INFLE_{infle}{subinfle}_Cores_{empresa}_{num_cores}.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df = pd.DataFrame([coreControlValuesD[1, 0]])
            print(df.shape)
            print(df.values[0,0])
            df.to_excel(writer, sheet_name='Sheet2', startrow=15+m-1, startcol=22, index=False)

# Convertir a pdf la parte del informe en excel
excel_path = f'{base_dir}INFLE_{infle}{subinfle}_Cores_{empresa}_{num_cores}.xlsm'
pdf_path = f'{base_dir}cuadros.pdf'
convert_excel_to_pdf(excel_path, pdf_path)

# Unir y normalizar la orientacion de los archivos que forman parte del informe
pdf_files = [f'{base_dir}cuadros.pdf', f'{base_dir}graficos.pdf']  # Lista de archivos PDF
output_pdf = f'{base_dir}informe_completo.pdf'  # Archivo PDF de salida
merge_pdfs(pdf_files, output_pdf)
normalize_pdf_orientation(output_pdf, output_pdf, desired_orientation='portrait')

# Agregar el encabezado y pie de página
header_footer_pdf_path = f'C:/Users/joela/Documents/PYTHON/formato_{acred}.pdf'           # PDF temporal con el encabezado
apply_header_footer_pdf(output_pdf, header_footer_pdf_path, output_pdf)

print("Proceso completado. PDF final generado con éxito.")

# Limpieza
plt.close('all')
