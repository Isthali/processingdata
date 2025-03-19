import os
from test_ledi import Axial_compression_test_report

# Parámetros iniciales
infle = '292-24'
subinfle = '-B'
standar = 'CORES'
empresa = 'PCM'
cores_id = [id+1 for id in range(6)]

# Directorios
base_dir = f'C:/Users/joela/Documents/MATLAB/Diamantinas/{infle}/'
os.makedirs(base_dir, exist_ok=True)

# Crear el informe en excel
test_report = Axial_compression_test_report(infle=infle, subinfle=subinfle, folder=base_dir, standard=standar, client_id=empresa, samples_id=cores_id)
test_report.make_report_file()
