import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from test_data import import_test_data_text, import_test_data_excel, get_report_data, write_report_data

class Mechanical_test:
    def __init__(self):
        self.sample_id = 0
        self.data = pd.DataFrame()
        self.data_process = pd.DataFrame()
        self.data_file = 'data_file'
        self.report_file = 'report_file'
    
    def get_data(self, data_file, data_source, variable_names):
        self.data_file = data_file
        
        if data_source == 'csv':
            self.data = import_test_data_text(file_path=self.data_file, delimiter='\t', variable_names=variable_names)
        elif data_source == 'xlsx':
            self.data = import_test_data_excel(file_path=self.data_file, sheet_idx=0, variable_names=variable_names)
        else:
            print('Error: data_source not recognized')

        return self.data
            
    def preprocess_data(self):
        return self.data
    
    def plot_data(self, x, y, title, xlabel, ylabel, legend, test_name, num_pag, final_pag=False):
        fig, ax = plt.subplots(figsize=(11.7, 8.3))
        ax.plot(self.data[x], self.data[y], 'b-', linewidth=2)
        ax.set_title(title, fontsize=10)
        ax.set_xlabel(xlabel, fontsize=9)
        ax.set_ylabel(ylabel, fontsize=9)
        ax.legend(legend, fontsize=9)
        ax.grid(visible=True, which='both', linestyle='--')
        ax.minorticks_on()
        ax.set_position([0.1, 0.15, 0.7, 0.75])
        fig.text(0.05, 0.05, f'INF-LE {self.infle}', fontsize=8, horizontalalignment='left')
        fig.text(0.5, 0.05, f'LEDI-{test_name}', fontsize=8, horizontalalignment='center')
        fig.text(0.85, 0.05, f'Pág. {num_pag}', fontsize=8, horizontalalignment='right')

        if final_pag:
            fig.text(0.85, 0.01, 'Fin del informe', fontsize=8, horizontalalignment='right')
        
        return fig, ax

class Resistance_mechanical_test(Mechanical_test):
    def __init__(self):
        super().__init__(self)
        self.maxF = 0
        self.idx = {'i': 0, 'f': 0, 'maxP': 0}

    def get_positive_data(self):
        for col in self.data.columns:
            self.data[col] = np.abs(self.data[col])

        return self.data

    def get_data(self, data_file, data_source, variable_names):
        self.data = super().get_data(data_file, data_source, variable_names)
        self.get_positive_data()

        return self.data

    def get_max_load(self):
        if self.data.empty:
            raise ValueError("No data loaded. Please load data before calculating the maximum load.")
        else:
            self.idx['maxF'] = np.argmax(self.data['F'])
            self.maxF = np.max(self.data['F'])

        return self.maxF

    def preprocess_data(self):
        imaxP = self.idx['maxF']
        self.idx['i'] = np.argmin(np.abs(self.data.loc[:imaxP, 'F'].to_numpy() - 0.01 * self.maxF))
        self.idx['f'] = np.argmin(np.abs(self.data.loc[imaxP:, 'F'].to_numpy() - 0.8 * self.maxF)) + imaxP

        return self.idx

class Axial_compression_test(Resistance_mechanical_test):
    def __init__(self, sample_id=None, data_file= None, report_file=None):
        super().__init__()
        self.sample_id = sample_id
        self.report_file = report_file
        self.data_file = data_file
        self.area_sec = 0
        self.strength = 0

    def get_area_section(self, cell):
        row, column = cell
        diameter = get_report_data(file_path=self.report_file, sheet_idx=1, position=(row, column))
        self.area_sec = np.pi*(diameter/2)**2

        return self.area_sec
    
    def get_strength(self):
        self.strength = self.maxF/self.area_sec

        return self.strength

class Axial_compression_test_report:
    def __init__(self):
        self.infle = 'infle'
        self.folder= 'folder'
        self.report_file = 'report_file'
        self.report_data = pd.DataFrame()
    
    def add_test(self, test):
        if isinstance(test, Axial_compression_test):
            self.tests.append(test)
        else:
            print("Only Axial_compression_test objects can be added to the report.")

    def get_report_data(self):
        report = []
        for test in self.tests:
            report.append([test.get_max_load() , test.get_area_section(cell=(1, 1)), test.get_strength()])

        self.report_data = pd.DataFrame(report, columns=['Max Load', 'Area Section', 'Strength'])

        return self.report_data
    
    def write_report(self):
        for i, test in enumerate(self.tests):
            row = i + 2
            column = 1
            write_report_data(file_path=self.report_file, sheet_idx=1, position=(row, column), val=test.maxF)
            column = 2
            write_report_data(file_path=self.report_file, sheet_idx=1, position=(row, column), val=test.area_sec)
            column = 3
            write_report_data(file_path=self.report_file, sheet_idx=1, position=(row, column), val=test.strength)