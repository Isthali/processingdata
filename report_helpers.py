"""Funciones auxiliares comunes para scripts de generación de reportes.

Objetivo: eliminar duplicación entre panels_report.py y cores_report.py.

Incluye:
  - parse_common_args: construcción parametrizable de CLI.
  - build_ids: genera lista de IDs secuenciales o usa los explícitos.
  - prepare_output_dir: crea carpeta base/infle y devuelve ruta con barra final.
  - run_report: instancia clase de reporte y ejecuta make_report_file.

Todas las rutas se normalizan usando pathlib. Se devuelve siempre un string para
compatibilidad con código existente que concatena cadenas.
"""
from __future__ import annotations

from pathlib import Path
from typing import List, Sequence, Type, Any, Callable, Optional
import argparse


def parse_common_args(
    *,
    description: str,
    default_standard: str,
    standard_choices: Sequence[str] | None,
    default_client: str,
    default_base: str,
    default_n: int,
) -> argparse.Namespace:
    """Genera un parser de argumentos común.

    Args:
        description: Texto de ayuda para el script.
        default_standard: Valor por defecto para --standard.
        standard_choices: Lista de opciones válidas o None (sin restricción).
        default_client: Cliente por defecto.
        default_base: Directorio base por defecto.
        default_n: Número de muestras por defecto.
    """
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('--infle', required=True, help='Identificador INFLE')
    parser.add_argument('--subinfle', default='', help='Sub-identificador INFLE')
    if standard_choices:
        parser.add_argument('--standard', default=default_standard, choices=standard_choices, help='Norma / estándar')
    else:
        parser.add_argument('--standard', default=default_standard, help='Norma / estándar')
    parser.add_argument('--empresa', '--client', dest='empresa', default=default_client, help='Nombre del cliente')
    parser.add_argument('--n', type=int, default=default_n, help='Número de muestras secuenciales desde 1')
    parser.add_argument('--ids', type=int, nargs='*', help='Lista explícita de IDs; anula --n')
    parser.add_argument('--base', dest='base_dir', default=default_base, help='Directorio base para resultados')
    return parser.parse_args()


def build_ids(n: int, explicit_ids: Optional[Sequence[int]]) -> List[int]:
    """Devuelve IDs explícitos o genera secuencia 1..n."""
    return list(explicit_ids) if explicit_ids else list(range(1, n + 1))


def prepare_output_dir(base_dir: str | Path, infle: str) -> str:
    """Crea (si no existe) y devuelve la ruta base/infle con barra final."""
    path = Path(base_dir).expanduser().resolve() / infle
    path.mkdir(parents=True, exist_ok=True)
    # Asegurar barra final para concatenaciones existentes en test_ledi
    return str(path) + '/'


def run_report(report_cls: Type[Any], **kwargs) -> Any:
    """Instancia la clase de reporte y ejecuta make_report_file.

    Retorna la instancia creada.
    """
    report = report_cls(**kwargs)
    report.make_report_file()
    print(f"Reporte generado: {report.report_file}")
    return report

__all__ = [
    'parse_common_args',
    'build_ids',
    'prepare_output_dir',
    'run_report',
]
