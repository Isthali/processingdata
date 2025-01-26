import os
from test_ledi import Axial_compression_test_report

# Parámetros iniciales
infle = '274-24'
subinfle = ''
empresa = 'TECCA'
cores_id = [id+1 for id in range(24)]
acred = 'no_acreditado'

# Directorios
base_dir = f'C:/Users/joela/Documents/MATLAB/Diamantinas/{infle}/'
os.makedirs(base_dir, exist_ok=True)

# Crear el informe en excel
test_report = Axial_compression_test_report(infle=infle, subinfle=subinfle, folder=base_dir, empresa=empresa, samples_id=cores_id)
test_report.get_report_file(acred)
