"""Script para generar reportes de ensayos de tenacidad en paneles.

Permite ejecutarse directamente desde la línea de comandos:
	python panels_report.py --infle 090-25 --standard EFNARC1996 --empresa BARCHIP --n 2 \
		--base C:/Users/joela/Documents/MATLAB/Losas

Requiere que los archivos de datos existan con el patrón:
	<base>/<infle>/Losa P<ID>.xlsx

Donde ID es el número de panel en samples_id.
"""

from __future__ import annotations

from test_ledi import Panel_toughness_test_report
from report_helpers import (
	parse_common_args,
	build_ids,
	prepare_output_dir,
	run_report,
)


def main() -> None:
	args = parse_common_args(
		description="Genera reporte de paneles de tenacidad.",
		default_standard='EFNARC1996',
		standard_choices=['EFNARC1996', 'EFNARC1999', 'ASTMC1550'],
		default_client='EMPRESA',
		default_base='C:/Users/joela/Documents/MATLAB/Losas',
		default_n=2,
	)
	panels_id = build_ids(args.n, args.ids)
	folder = prepare_output_dir(args.base_dir, args.infle)
	if not panels_id:
		raise ValueError("No hay IDs de panel proporcionados.")
	run_report(
		Panel_toughness_test_report,
		infle=args.infle,
		subinfle=args.subinfle,
		folder=folder,
		standard=args.standard,
		client_id=args.empresa,
		samples_id=panels_id,
	)


if __name__ == '__main__':  # pragma: no cover
	main()
