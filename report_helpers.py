"""Helpers comunes para scripts de generación de reportes."""
from __future__ import annotations

from pathlib import Path
from typing import Any, Type


def prepare_output_dir(base_dir: str | Path, infle: str) -> str:
    """Crea (si no existe) y devuelve la ruta ``base_dir/infle`` con barra final."""
    path = Path(base_dir).expanduser().resolve() / infle
    path.mkdir(parents=True, exist_ok=True)
    # La barra final preserva compatibilidad con concatenaciones existentes en test_ledi.
    return str(path) + '/'


def run_report(report_cls: Type[Any], **kwargs) -> Any:
    """Instancia la clase de reporte y ejecuta ``make_report_file``."""
    report = report_cls(**kwargs)
    report.make_report_file()
    print(f"Reporte generado: {report.report_file}")
    return report


__all__ = [
    'prepare_output_dir',
    'run_report',
]
