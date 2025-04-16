import os
from test_ledi import Beam_residual_strength_test_report

# Parámetros iniciales
infle = '063-25'
subinfle = ''
standar = 'EN14651'
empresa = 'CHINA-CIVIL'
panels_id = [id+1 for id in range(3)]

# Directorios
base_dir = f'C:/Users/joela/Documents/MATLAB/Vigas/{infle}/'
os.makedirs(base_dir, exist_ok=True)

# Crear el informe en excel
test_report = Beam_residual_strength_test_report(infle=infle, subinfle=subinfle, folder=base_dir, standard=standar, client_id=empresa, samples_id=panels_id)
test_report.make_report_file()
