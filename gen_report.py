"""Script genérico para convertir una planilla Excel de resultados a PDF con formato.

Este script usa la clase Generate_test_report que asume que el archivo Excel ya
existe y sólo debe aplicarse el pipeline de conversión y formateo.

Ejemplo de uso:
	python gen_report.py --infle 336-24 --subinfle -S --standard DM --empresa EXC \
		--base C:/Users/joela/Documents/MATLAB/Diamantinas
"""

from __future__ import annotations

from test_ledi import Generate_test_report
from report_helpers import (
	parse_common_args,
	prepare_output_dir,
	run_report,
)


def main() -> None:
	# Para este tipo de reporte no se necesitan muestras (samples), se usa n=0.
	args = parse_common_args(
		description='Genera reporte genérico (sólo conversión Excel -> PDF).',
		default_standard='DM',
		standard_choices=None,
		default_client='EMPRESA',
		default_base='C:/Users/joela/Documents/MATLAB/Diamantinas',
		default_n=0,
	)
	folder = prepare_output_dir(args.base_dir, args.infle)
	run_report(
		Generate_test_report,
		infle=args.infle,
		subinfle=args.subinfle,
		folder=folder,
		standard=args.standard,
		client_id=args.empresa,
		samples_id=[],
	)


if __name__ == '__main__':  # pragma: no cover
	main()
