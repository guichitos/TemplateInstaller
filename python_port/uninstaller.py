"""Desinstalador único basado en la carpeta actual."""
from __future__ import annotations

import argparse
import logging
from pathlib import Path

# Configuración manual para el modo diseño.
# - Establece en True para forzar modo diseño siempre.
# - Establece en False para desactivarlo siempre.
# - Deja en None para usar la lógica normal basada en entorno.
MANUAL_IS_DESIGN_MODE: bool | None = True

try:
    from . import common
except ImportError:  # pragma: no cover - permite ejecución directa como script
    import sys

    sys.path.append(str(Path(__file__).resolve().parent))
    import common  # type: ignore[no-redef]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Desinstalador de plantillas de Office (Python)")
    return parser.parse_args()


def main(argv: list[str] | None = None) -> int:
    args = parse_args()
    design_mode = _resolve_design_mode()
    common.refresh_design_log_flags(design_mode)
    common.configure_logging(design_mode)
    common.close_office_apps(design_mode)

    base_dir = common.resolve_base_directory(Path.cwd())
    if base_dir == Path.cwd() and common.path_in_appdata(base_dir):
        common.exit_with_error(
            '[ERROR] No se recibió la ruta de las plantillas. Ejecute el desinstalador desde "1. Pin templates..." para que se le pase la carpeta correcta.'
        )

    if design_mode and common.DESIGN_LOG_UNINSTALLER:
        logging.getLogger(__name__).info("[INFO] Desinstalando desde: %s", base_dir)

    destinations = common.default_destinations()
    common.remove_installed_templates(destinations, design_mode)
    common.delete_custom_copies(base_dir, destinations, design_mode)
    common.clear_mru_entries_for_payload(base_dir, destinations, design_mode)

    if design_mode and common.DESIGN_LOG_UNINSTALLER:
        logging.getLogger(__name__).info("[FINAL] Desinstalación completada.")
    else:
        print("Ready")
    return 0


def _resolve_design_mode() -> bool:
    if MANUAL_IS_DESIGN_MODE is not None:
        return bool(MANUAL_IS_DESIGN_MODE)
    return bool(common.DEFAULT_DESIGN_MODE)


if __name__ == "__main__":
    raise SystemExit(main())
