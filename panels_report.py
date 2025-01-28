import os
from test_ledi import Panel_toughness_test_report

# Parámetros iniciales
infle = '359-24'
subinfle = ''
standar = 'ASTMC1550'
empresa = 'MONOFILAMENTOS'
panels_id = [id+1 for id in range(3)]
acred = 'no_acreditado'

# Directorios
base_dir = f'C:/Users/joela/Documents/MATLAB/Losas/{infle}/'
os.makedirs(base_dir, exist_ok=True)

# Crear el informe en excel
test_report = Panel_toughness_test_report(infle=infle, subinfle=subinfle, folder=base_dir, standard=standar, empresa=empresa, samples_id=panels_id)
test_report.get_report_file(acred)
