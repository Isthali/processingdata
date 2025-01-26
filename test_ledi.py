import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from test_data import import_data_text, import_data_excel, get_data_excel, write_data_excel
from putformatPDF import convert_excel_to_pdf, merge_pdfs, normalize_pdf_orientation, apply_header_footer_pdf

class Mechanical_test:
    def __init__(self):
        self.sample_id = 0
        self.data = pd.DataFrame()
        self.data_process = pd.DataFrame()
        self.data_file = 'data_file'
        self.maxLoad = 0
        self.minLoad = 0
        self.idx = {'minLoad': 0, 'maxLoad': 0}
    
    def get_data(self, data_file, data_source, variable_names):
        self.data_file = data_file
        
        if data_source == 'csv':
            self.data = import_data_text(file_path=self.data_file, delimiter='\t', variable_names=variable_names)
        elif data_source == 'xlsx':
            self.data = import_data_excel(file_path=self.data_file, sheet_idx=0, variable_names=variable_names)
        else:
            print("Error: data_source not recognized.")

        return self.data
            
    def make_positive_data(self):
        for col in self.data.columns:
            self.data[col] = np.abs(self.data[col])

        return self.data
    
    def get_max_load(self):
        if self.data.empty:
            raise ValueError("Error: load data before calculating the maximum load.")
        else:
            self.idx['maxLoad'] = np.argmax(np.abs(self.data['Load'].to_numpy()))
            self.maxLoad = np.max(self.data['Load'].to_numpy())

        return self.maxLoad
    
    def get_min_load(self):
        if self.data.empty:
            raise ValueError("Error: load data before calculating the maximum load.")
        else:
            self.idx['minLoad'] = np.argmin(np.abs(self.data['Load'].to_numpy()))
            self.maxLoad = np.min(self.data['Load'].to_numpy())

        return self.maxLoad

    def plot_data(self, x, y, title, xlabel, ylabel, legend, infle, test_name, num_pag, final_pag=False):
        fig, ax = plt.subplots(figsize=(11.7, 8.3))
        ax.plot(self.data_process[x], self.data_process[y], 'b-', linewidth=2)
        ax.set_title(title, fontsize=10)
        ax.set_xlabel(xlabel, fontsize=9)
        ax.set_ylabel(ylabel, fontsize=9)
        ax.legend(legend, fontsize=9)
        ax.grid(visible=True, which='both', linestyle='--')
        ax.minorticks_on()
        ax.set_position([0.10, 0.15, 0.70, 0.75])
        fig.text(0.05, 0.05, f'INF-LE {infle}', fontsize=8, horizontalalignment='left')
        fig.text(0.5, 0.05, f'LEDI-{test_name}', fontsize=8, horizontalalignment='center')
        fig.text(0.85, 0.05, f'Pág. {num_pag}', fontsize=8, horizontalalignment='right')

        if final_pag:
            fig.text(0.85, 0.01, 'Fin del informe', fontsize=8, horizontalalignment='right')
        
        return fig, ax

class Resistance_mechanical_test(Mechanical_test):
    def __init__(self):
        super().__init__()
        self.idx = {'i': 0, 'f': 0, 'maxLoad': 0}

    def preprocess_data(self):
        super().get_max_load()
        super().make_positive_data()
        imaxP = self.idx['maxLoad']
        self.idx['i'] = np.argmin(np.abs(self.data.loc[:imaxP, 'Load'].to_numpy() - 0.01 * self.maxLoad))
        self.idx['f'] = np.argmin(np.abs(self.data.loc[imaxP:, 'Load'].to_numpy() - 0.8 * self.maxLoad)) + imaxP
        self.data_process = self.data.loc[self.idx['i']:self.idx['f'], :]

        return self.idx

class Axial_compression_test(Resistance_mechanical_test):
    def __init__(self, sample_id=None, data_file= None):
        super().__init__()
        self.sample_id = sample_id
        self.data_file = data_file
        self.area_sec = 0
        self.strength = 0

    def get_sample_id(self):
        return self.sample_id

    def get_area_section(self, length_sec=None, section_type=None):
        if section_type=='circular':
            self.area_sec = np.pi*(length_sec/2)**2
        elif section_type=='square':
            self.area_sec = length_sec**2
        elif section_type=='rectangular':
            self.area_sec = length_sec[0]*length_sec[1]
        else:
            print("Error: no section type selected.")
               
        return self.area_sec
    
    def get_strength(self):
        self.strength = self.maxLoad/self.area_sec

        return self.strength

class Axial_compression_test_report:
    def __init__(self, infle=None, subinfle=None, folder=None, empresa=None, samples_id=[]):
        self.repor_id = {'infle': infle, 'subinfle': subinfle}
        self.folder = folder
        self.empresa = empresa
        self.samples_id = samples_id
        self.tests = []
        self.excel_file = f'{folder}INFLE_{infle}{subinfle}_Cores_{empresa}_{len(self.samples_id)}.xlsm'
        self.plots_file = f'{folder}plots.pdf'
        self.report_file = f'{folder}INFLE_{infle}{subinfle}_Cores_{empresa}_{len(self.samples_id)}.pdf'
    
    def add_tests(self):
        for id in self.samples_id:
            test = Axial_compression_test(sample_id=id, data_file=f'{self.folder}{self.repor_id['infle']}-d{id}/specimen.dat')
            test.get_data(data_file=test.data_file, data_source='csv', variable_names=['Time', 'Displacement', 'Load'])
            test.preprocess_data()
            self.tests.append(test)
    
    def write_report(self):
        for i, test in enumerate(self.tests):
            row = i + 16
            column = 23
            write_data_excel(file_path=self.excel_file, sheet_name='Cores', position=(row, column), val=test.get_max_load())
            #column = 2
            #write_data_excel(file_path=self.report_file, sheet_idx=1, position=(row, column), val=test.get_area_section(cell=(1, 1)))
            #column = 3
            #write_data_excel(file_path=self.report_file, sheet_idx=1, position=(row, column), val=test.get_strength())

    def plot_report(self):
        with PdfPages(f'{self.plots_file}') as pdf_file:
            for i, test in enumerate(self.tests):
                if i == len(self.tests) - 1:
                    fig, ax = test.plot_data(x='Displacement', y='Load', title='Fuerza-Desplazamiento', xlabel='Desplazamiento (mm)', ylabel='Fuerza (kN)', legend=[f'TESTIGO {test.get_sample_id()}'], infle=f'{self.repor_id['infle']}{self.repor_id['subinfle']}', test_name='ENSAYO DE RESISTENCIA A LA COMPRESIÓN', num_pag=i+4, final_pag=True)
                else:
                    fig, ax = test.plot_data(x='Displacement', y='Load', title='Fuerza-Desplazamiento', xlabel='Desplazamiento (mm)', ylabel='Fuerza (kN)', legend=[f'TESTIGO {test.get_sample_id()}'], infle=f'{self.repor_id['infle']}{self.repor_id['subinfle']}', test_name='ENSAYO DE RESISTENCIA A LA COMPRESIÓN', num_pag=i+4, final_pag=False)
                
                pdf_file.savefig(fig)
                plt.close()
    
    def get_report_file(self, acred):
        header_footer_pdf_path = f'C:/Users/joela/Documents/GitHub/processingdata/formato_{acred}.pdf'
        self.add_tests()
        self.write_report()
        convert_excel_to_pdf(excel_path=self.excel_file, pdf_path=self.report_file)
        self.plot_report()
        merge_pdfs(pdf_list=[self.report_file, self.plots_file], output_pdf=self.report_file)
        normalize_pdf_orientation(input_pdf_path=self.report_file, output_pdf_path=self.report_file, desired_orientation='portrait')
        apply_header_footer_pdf(input_pdf_path=self.report_file, header_footer_pdf_path=header_footer_pdf_path, output_pdf_path=self.report_file)
        