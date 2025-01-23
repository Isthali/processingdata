import numpy as np
import openpyxl as xl
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from test_data import import_test_data_text, import_test_data_excel

class Mechanical_test:
    def __init__(self):
        self.infle = 'inf-le'
        self.data = pd.DataFrame()
        self.data_process = pd.DataFrame()
        self.data_file = 'data_file'
    
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

    def plot_data(self, x, y, title, xlabel, ylabel, legend, savefig=False, figname='figname'):
        plt.figure(figsize=(10, 8))
        plt.plot(self.data[x], self.data[y], 'b-', linewidth=2)
        plt.title(title, fontsize=10)
        plt.xlabel(xlabel, fontsize=9)
        plt.ylabel(ylabel, fontsize=9)
        plt.legend(legend, fontsize=9)
        
        if savefig:
            plt.savefig(figname + '.png')
        
        return plt

class Resistant_test(Mechanical_test):
    def __init__(self, infle):
        super().__init__()
        self.infle = infle
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
        self.idx['maxF'] = np.argmax(self.data['F'])
        self.maxF = np.max(self.data['F'])

        return self.maxF

    def preprocess_data(self):
        imaxP = self.idx['maxF']
        self.idx['i'] = np.argmin(np.abs(self.data.loc[:imaxP, ['F']] - 0.01*self.maxF))
        self.idx['f'] = np.argmin(np.abs(self.data.loc[imaxP:, ['F']] - 0.8*self.maxF)) + imaxP

        return self.idx

class Axial_load_test(Mechanical_test):
    def __init__(self, infle):
        super().__init__()
        self.infle = infle

    def preprocess_data(self, withend=True):
        F = np.abs(np.array(self.data['F']))
        idx_maxF = np.argmax(F)
        maxF = np.amax(F)
        idx_i = np.argmin(np.abs(F[:idx_maxF] - 0.01*maxF))

        if withend:
            idx_f = np.argmin(np.abs(F[idx_maxF:] - 0.8*maxF)) + idx_maxF
        else:
            idx_f = F.shape[0] - 1

        data_array = self.data.iloc[idx_i:idx_f + 1,:].to_numpy()

        n = 0
        for col in self.data.columns:
            data_array[:, n] = np.abs(data_array[:, n])
            
            if col != 'F':
                data_array[:, n] = data_array[:, n] - data_array[0, n]

            n += 1
        
        self.data_preprocess = pd.DataFrame(data_array, columns=self.data.columns)

    def get_max_load(self):
        F = np.abs(np.array(self.data_preprocess['F']))
        D = np.abs(np.array(self.data_preprocess['D']))
        idx_maxF = np.argmax(F)
        maxF = np.amax(F)
        D_maxF = D[idx_maxF]

        return (idx_maxF, maxF, D_maxF)