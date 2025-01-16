import numpy as np
import openpyxl as xl
import pandas as pd
import matplotlib.pyplot as plt
import test_lab as test
import os
from matplotlib.backends.backend_pdf import PdfPages

# Variables
infle = '080-23'
subinfle = ''
empresa = 'HVS'
num_cores = 18
num_1core = 1

base_dir = f'C:/Users/joela/Documents/OneDrive/PYTHON/Compresion_axial/Cores/{infle}/'
os.makedirs(base_dir, exist_ok=True)
grafico_file = f'{base_dir}graficos.pdf'
informe_file = f'{base_dir}INFLE_{infle}{subinfle}_Cores_{empresa}_{num_cores}.xlsm'

with PdfPages(grafico_file) as grafico:
    with xl.load_workbook(informe_file, keep_vba=True) as informe:
        for n in range(num_1core, num_1core + num_cores):
            core = test.Axial_load_test(f'{infle}{subinfle}')
            data_file = f'{base_dir}{infle}-d{n}/specimen.dat'
            core.get_data(data_file=data_file, data_source='csv', column_names=['t', 'D', 'F'])
            core.preprocess_data(withend=True)
            (idx_maxF, maxF, D_maxF) = core.get_max_load()
            fig, ax = plt.subplots()
            core.data_preprocess.plot(x='D', y='F', ax=ax)
            
            fig.set_size_inches(11.7, 8.3)
            ax.set_title(f'Fuerza - Desplazamiento: Testigo {n}')
            ax.set_xlabel('Desplazamiento axial [mm]')
            ax.set_ylabel('Fuerza axial [kN]')
            ax.get_legend().remove()
            ax.grid(visible=True, which='both', linestyle='--')
            ax.minorticks_on()
            ax.annotate(f'Fmax = {np.round(maxF, decimals=2)} kN', (D_maxF, maxF))
            ax.set_position([0.1, 0.15, 0.7, 0.75])
            fig.text(0.05, 0.05, f'INF-LE {infle}{subinfle}', fontsize=8, horizontalalignment='left')
            fig.text(0.5, 0.05, 'Ensayo de resistencia a la compresión', fontsize=8, horizontalalignment='center')
            fig.text(0.85, 0.05, f'Pág. {3 + n}/{3 + num_cores}', fontsize=8, horizontalalignment='right')

            if n == num_cores:
                fig.text(0.85, 0.01, 'Fin del informe', fontsize=8, horizontalalignment='right')    

            grafico.savefig(fig)
            plt.close()

            informe.worksheets[1].cell(row=16-num_1core+n, column=22, value=maxF)
            informe.save()
