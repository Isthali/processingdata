import os
from test_ledi import Panel_toughness_test_report

# Parámetros iniciales
infle = '010-25'
subinfle = '-B'
standar = 'EFNARC1996'
empresa = 'SIKA'
panels_id = [id+4 for id in range(3)]

# Directorios
base_dir = f'C:/Users/joela/Documents/MATLAB/Losas/{infle}/'
os.makedirs(base_dir, exist_ok=True)

# Crear el informe en excel
test_report = Panel_toughness_test_report(infle=infle, subinfle=subinfle, folder=base_dir, standard=standar, client_id=empresa, samples_id=panels_id)
test_report.make_report_file()
