import numpy as np
import openpyxl as xl
import pandas as pd
import matplotlib.pyplot as plt
import chardet
from matplotlib.backends.backend_pdf import PdfPages

def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        result = chardet.detect(f.read(10000))  # Analiza las primeras 10,000 bytes
        return result['encoding']

class Test_lab:
    def __init__(self):
        self.infle = 'inf-le'
        self.data = pd.DataFrame()
        self.data_preprocess = pd.DataFrame()
        self.data_file = 'data_file'
    
    def get_data(self, data_file, data_source, variable_names):
        self.data_file = data_file
        
        if data_source == 'csv':
            encoding = detect_encoding(self.data_file)
            self.data = pd.read_csv(self.data_file, sep='\t', header=None, names=variable_names, skiprows=14, encoding=encoding)
        elif data_source == 'xlsx':
            self.data = pd.read_csv(self.data_file, sep='\t', header=None, names=variable_names, skiprows=14, encoding=encoding)

    def preprocess_data(self):
        return self.data
    


class Axial_load_test(Test_lab):
    def __init__(self, infle):
        self.infle = infle
        self.data = pd.DataFrame()
        self.data_preprocess = pd.DataFrame()
        self.data_file = 'data_file'

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