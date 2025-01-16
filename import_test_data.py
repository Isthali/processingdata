import pandas as pd
import chardet

def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        result = chardet.detect(f.read(10000))  # Analiza las primeras 10,000 bytes
        return result['encoding']
    
# Función para importar datos de un archivo de texto
def import_test_data_text(file_path, num_variables, data_lines, delimiter, variable_names, variable_types):
    # Opciones para importar
    encoding = detect_encoding(file_path)  # Detecta la codificación
    skip_rows = list(range(data_lines[0] - 1))  # Filas a saltar
    dtype = {variable_names[i]: variable_types[i] for i in range(num_variables)}
    
    # Lectura de datos con pandas
    df = pd.read_csv(file_path, delimiter=delimiter, skiprows=skip_rows, names=variable_names, dtype=dtype, encoding=encoding)
    df = df.dropna()  # Omitir filas con errores de importación o vacías
    return df
