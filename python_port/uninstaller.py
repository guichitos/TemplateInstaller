"""Desinstalador único basado en la carpeta actual."""
from __future__ import annotations

import argparse
import logging
from pathlib import Path

from . import common


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Desinstalador de plantillas de Office (Python)")
    parser.add_argument(
        "--design-mode",
        action="store_true",
        help="Muestra salida detallada para depuración.",
    )
    return parser.parse_args()


def main(argv: list[str] | None = None) -> int:
    args = parse_args()
    design_mode = args.design_mode or common.DEFAULT_DESIGN_MODE
    common.configure_logging(design_mode)

    base_dir = common.resolve_base_directory(Path.cwd())
    if base_dir == Path.cwd() and common.path_in_appdata(base_dir):
        common.exit_with_error(
            '[ERROR] No se recibió la ruta de las plantillas. Ejecute el desinstalador desde "1. Pin templates..." para que se le pase la carpeta correcta.'
        )

    if design_mode:
        logging.getLogger(__name__).info("[INFO] Desinstalando desde: %s", base_dir)

    common.close_office_apps(design_mode)

    destinations = common.default_destinations()
    common.remove_installed_templates(destinations, design_mode)
    common.delete_custom_copies(base_dir, destinations, design_mode)

    if design_mode:
        logging.getLogger(__name__).info("[FINAL] Desinstalación completada.")
    else:
        print("Ready")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
