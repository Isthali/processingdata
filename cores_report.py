"""Script para generar reportes de ensayos de compresión axial (cores).

Uso CLI:
	python cores_report.py --infle 061-25 --empresa BCP --n 3 \
		--base C:/Users/joela/Documents/MATLAB/Diamantinas

También se puede especificar IDs manualmente:
	python cores_report.py --infle 061-25 --ids 2 5 7
"""

from __future__ import annotations

from test_ledi import Axial_compression_test_report
from report_helpers import (
	parse_common_args,
	build_ids,
	prepare_output_dir,
	run_report,
)


def main() -> None:
	args = parse_common_args(
		description="Genera reporte de compresión axial (cores).",
		default_standard='CORES',
		standard_choices=None,
		default_client='EMPRESA',
		default_base='C:/Users/joela/Documents/MATLAB/Diamantinas',
		default_n=3,
	)
	cores_id = build_ids(args.n, args.ids)
	folder = prepare_output_dir(args.base_dir, args.infle)
	if not cores_id:
		raise ValueError("No se proporcionaron IDs de muestras.")
	run_report(
		Axial_compression_test_report,
		infle=args.infle,
		subinfle=args.subinfle,
		folder=folder,
		standard=args.standard,
		client_id=args.empresa,
		samples_id=cores_id,
	)


if __name__ == '__main__':  # pragma: no cover
	main()
