import pandas as pd
import chardet

def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        result = chardet.detect(f.read(10000))  # Analiza las primeras 10,000 bytes
        return result['encoding']
    
# Función para importar datos de un archivo de texto
def import_test_data_text(file_path, delimiter, variable_names):
    # Opciones para importar
    encoding = detect_encoding(file_path)  # Detecta la codificación
    
    # Lectura de datos con pandas
    df = pd.read_csv(filepath_or_buffer=file_path, sep=delimiter, names=variable_names, encoding=encoding)
    df = df.dropna()  # Omitir filas con errores de importación o vacías
    return df

# Función para importar datos de un archivo de Excel
def import_test_data_excel(file_path, sheet_idx, variable_names):
    # Opciones para importar
    encoding = detect_encoding(file_path)  # Detecta la codificación
    
    # Lectura de datos con pandas
    df = pd.read_excel(io=file_path, sheet_name=sheet_idx, names=variable_names, encoding=encoding)
    df = df.dropna()  # Omitir filas con errores de importación o vacías
    return df
