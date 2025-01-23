import pandas as pd
import openpyxl as xl
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

# Función para obtener valores para procesamiento
def get_report_data(file_path, sheet_idx=None, position=(1, 1)):
    # Abrir el archivo Excel
    with xl.load_workbook(filename=file_path, keep_vba=True) as wb:
        # Seleccionar la hoja
        if sheet_idx:
            sheet = wb[sheet_idx]
        else:
            # Si no se proporciona nombre, usar la primera hoja
            sheet = wb.worksheets[0]
        
        # Obtener el valor de la celda especificada
        row, column = position
        val = sheet.cell(row, column).value

    return val

# Función para guardar valor en un archivo Excel
def write_report_data(file_path, sheet_idx, position, val):
    # Abrir el archivo Excel
    with xl.load_workbook(filename=file_path, keep_vba=True) as wb:
        # Seleccionar la hoja
        if sheet_idx in wb.sheetnames:
            sheet = wb[sheet_idx]
        else:
            # Crear una nueva hoja
            sheet = wb.create_sheet(title=sheet_idx)
        
        # Guardar el valor en la celda especificada
        row, column = position
        sheet.cell(row, column, value=val)
        
        # Guardar los cambios
        wb.save(file_path)
