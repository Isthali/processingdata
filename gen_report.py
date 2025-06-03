import os
from test_ledi import Generate_test_report

# Parámetros iniciales
infle = '336-24'
subinfle = '-S'
standar = 'DM'
empresa = 'EXC'

# Directorios
base_dir = f'C:/Users/joela/Documents/MATLAB/Diamantinas/{infle}/'
os.makedirs(base_dir, exist_ok=True)

# Crear el informe en excel
test_report = Generate_test_report(infle=infle, subinfle=subinfle, folder=base_dir, standard=standar, client_id=empresa)
test_report.make_report_file()
