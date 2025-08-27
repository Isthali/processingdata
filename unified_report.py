"""Script unificado para generar reportes de ensayos de diferentes tipos.

Este script puede generar reportes para múltiples tipos de ensayos mecánicos:
- cores: Ensayos de compresión axial (testigos de hormigón)
- panels: Ensayos de tenacidad de paneles
- panels_residual: Ensayos de resistencia residual de paneles/vigas
- generic: Conversión genérica Excel -> PDF

Ejemplos de uso:
    # Reporte de testigos de hormigón
    python unified_report.py cores --infle 336-24 --subinfle S --standard CORES --empresa PRODIMIN --n 6

    # Reporte de tenacidad de paneles
    python unified_report.py panels --infle 111-25 --subinfle C --standard EFNARC1996 --empresa PRODIMIN --n 3

    # Reporte de resistencia residual
    python unified_report.py panels_residual --infle 111-25 --subinfle C --standard EN14488 --empresa PRODIMIN --n 3

    # Conversión genérica
    python unified_report.py generic --infle 336-24 --subinfle S --standard DM --empresa EXC
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path
from typing import Dict, List, Optional, Type, Union

from test_ledi import (
    Axial_compression_test_report,
    Panel_toughness_test_report,
    Panel_Beam_residual_strength_test_report,
    Generate_test_report,
    Test_report
)
from report_helpers import prepare_output_dir, run_report


# Configuración de tipos de ensayo
REPORT_CONFIGS = {
    'cores': {
        'report_class': Axial_compression_test_report,
        'description': 'Genera reporte de ensayos de compresión axial (testigos)',
        'default_standard': 'CORES',
        'standard_choices': None,
        'default_client': 'PRODIMIN',
        'default_base': 'C:/Users/joela/Documents/MATLAB/Diamantinas',
        'default_n': 6,
        'requires_samples': True,
    },
    'panels': {
        'report_class': Panel_toughness_test_report,
        'description': 'Genera reporte de ensayos de tenacidad de paneles',
        'default_standard': 'EFNARC1996',
        'standard_choices': ['EFNARC1996', 'EFNARC1999', 'ASTMC1550', 'EN14488-5'],
        'default_client': 'PRODIMIN',
        'default_base': 'C:/Users/joela/Documents/MATLAB/Losas',
        'default_n': 3,
        'requires_samples': True,
    },
    'panels_residual': {
        'report_class': Panel_Beam_residual_strength_test_report,
        'description': 'Genera reporte de ensayos de resistencia residual de paneles/vigas',
        'default_standard': 'EN14488',
        'standard_choices': None,
        'default_client': 'PRODIMIN',
        'default_base': 'C:/Users/joela/Documents/MATLAB/Losas',
        'default_n': 3,
        'requires_samples': True,
    },
    'generic': {
        'report_class': Generate_test_report,
        'description': 'Genera reporte genérico (sólo conversión Excel -> PDF)',
        'default_standard': 'DM',
        'standard_choices': None,
        'default_client': 'EMPRESA',
        'default_base': 'C:/Users/joela/Documents/MATLAB/Diamantinas',
        'default_n': 0,
        'requires_samples': False,
    },
}


def setup_logging(verbose: bool = False) -> None:
    """Configura el sistema de logging."""
    level = logging.DEBUG if verbose else logging.INFO
    format_str = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    
    logging.basicConfig(
        level=level,
        format=format_str,
        handlers=[
            logging.StreamHandler(sys.stdout),
        ]
    )


def build_samples_id(n: int, offset: int = 1) -> List[int]:
    """Construye lista de IDs de muestras consecutivos.
    
    Args:
        n: Número de muestras
        offset: Valor inicial (por defecto 1)
        
    Returns:
        Lista de IDs de muestras
    """
    if n <= 0:
        return []
    return [offset + i for i in range(n)]


def get_samples_id(args: argparse.Namespace) -> List[int]:
    """Obtiene lista de IDs de muestras desde argumentos.
    
    Args:
        args: Argumentos parseados
        
    Returns:
        Lista de IDs de muestras
    """
    config = REPORT_CONFIGS[args.test_type]
    
    if not config['requires_samples']:
        return []
    
    # Si se especificaron IDs directamente
    if hasattr(args, 'ids') and args.ids is not None:
        return args.ids
    
    # Si se especificó número de muestras
    if hasattr(args, 'n') and args.n is not None:
        return build_samples_id(args.n, args.offset)
    
    # Usar valores por defecto
    return build_samples_id(config['default_n'], args.offset)


def parse_arguments() -> argparse.Namespace:
    """Parsea argumentos de línea de comandos."""
    parser = argparse.ArgumentParser(
        description='Script unificado para generar reportes de ensayos mecánicos',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    
    # Subcomandos para tipos de ensayo
    subparsers = parser.add_subparsers(
        dest='test_type',
        help='Tipo de ensayo/reporte a generar',
        required=True
    )
    
    # Crear subparser para cada tipo de ensayo
    for test_type, config in REPORT_CONFIGS.items():
        subparser = subparsers.add_parser(
            test_type,
            help=config['description']
        )
        
        # Argumentos comunes
        subparser.add_argument(
            '--infle',
            required=True,
            help='Identificador del informe (ej: 336-24)'
        )
        
        subparser.add_argument(
            '--subinfle',
            default='',
            help='Sub-identificador del informe (ej: S, C)'
        )
        
        subparser.add_argument(
            '--standard',
            default=config['default_standard'],
            choices=config['standard_choices'],
            help=f'Estándar del ensayo (por defecto: {config["default_standard"]})'
        )
        
        subparser.add_argument(
            '--empresa',
            default=config['default_client'],
            help=f'Nombre de la empresa cliente (por defecto: {config["default_client"]})'
        )
        
        subparser.add_argument(
            '--base-dir',
            default=config['default_base'],
            help=f'Directorio base (por defecto: {config["default_base"]})'
        )
        
        if config['requires_samples']:
            # Grupo mutuamente exclusivo para especificar muestras
            sample_group = subparser.add_mutually_exclusive_group()
            
            sample_group.add_argument(
                '--n',
                type=int,
                default=config['default_n'],
                help=f'Número de muestras consecutivas (por defecto: {config["default_n"]})'
            )
            
            sample_group.add_argument(
                '--ids',
                type=int,
                nargs='+',
                help='IDs específicos de muestras (ej: --ids 3 4 7)'
            )
            
            subparser.add_argument(
                '--offset',
                type=int,
                default=1,
                help='Valor inicial para IDs de muestras cuando se usa --n (por defecto: 1)'
            )
        
        subparser.add_argument(
            '--verbose',
            '-v',
            action='store_true',
            help='Habilita salida verbose'
        )
    
    return parser.parse_args()


def validate_arguments(args: argparse.Namespace) -> None:
    """Valida argumentos de entrada.
    
    Args:
        args: Argumentos parseados
        
    Raises:
        ValueError: Si hay argumentos inválidos
    """
    config = REPORT_CONFIGS[args.test_type]
    
    # Validar directorio base
    base_path = Path(args.base_dir)
    if not base_path.exists():
        logging.warning(f"El directorio base {base_path} no existe, se creará automáticamente")
    
    if config['requires_samples']:
        # Validar IDs específicos si se proporcionaron
        if hasattr(args, 'ids') and args.ids is not None:
            if any(id_val <= 0 for id_val in args.ids):
                raise ValueError(f"Todos los IDs de muestras deben ser > 0, recibidos: {args.ids}")
        
        # Validar número de muestras si se proporcionó
        elif hasattr(args, 'n') and args.n is not None:
            if args.n < 0:
                raise ValueError(f"El número de muestras debe ser >= 0, recibido: {args.n}")
        
        # Validar offset
        if hasattr(args, 'offset') and args.offset < 0:
            raise ValueError(f"El offset debe ser >= 0, recibido: {args.offset}")
    
    logging.info(f"Argumentos validados para tipo de ensayo: {args.test_type}")


def generate_report(args: argparse.Namespace) -> None:
    """Genera el reporte basado en los argumentos.
    
    Args:
        args: Argumentos parseados y validados
    """
    config = REPORT_CONFIGS[args.test_type]
    report_class = config['report_class']
    
    # Preparar directorio de salida
    folder = prepare_output_dir(args.base_dir, args.infle)
    
    # Preparar lista de muestras
    samples_id = get_samples_id(args)
    if config['requires_samples']:
        logging.info(f"Generando reporte para {len(samples_id)} muestras: {samples_id}")
    else:
        logging.info("Generando reporte genérico (sin muestras específicas)")
    
    # Ejecutar generación del reporte
    try:
        run_report(
            report_class,
            infle=args.infle,
            subinfle=args.subinfle,
            folder=folder,
            standard=args.standard,
            client_id=args.empresa,
            samples_id=samples_id,
        )
        logging.info(f"Reporte generado exitosamente en: {folder}")
        
    except Exception as e:
        logging.error(f"Error al generar el reporte: {e}")
        raise


def main() -> None:
    """Función principal del script."""
    try:
        # Parsear argumentos
        args = parse_arguments()
        
        # Configurar logging
        setup_logging(args.verbose)
        
        logging.info(f"Iniciando generación de reporte tipo: {args.test_type}")
        logging.debug(f"Argumentos recibidos: {vars(args)}")
        
        # Validar argumentos
        validate_arguments(args)
        
        # Generar reporte
        generate_report(args)
        
        logging.info("Proceso completado exitosamente")
        
    except KeyboardInterrupt:
        logging.warning("Proceso interrumpido por el usuario")
        sys.exit(1)
        
    except Exception as e:
        logging.error(f"Error en el proceso: {e}")
        if hasattr(args, 'verbose') and args.verbose:
            logging.exception("Detalles del error:")
        sys.exit(1)


if __name__ == '__main__':  # pragma: no cover
    main()
