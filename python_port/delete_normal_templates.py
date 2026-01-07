"""Elimina plantillas Normal de Word en una ruta fija."""
from __future__ import annotations

from pathlib import Path

try:
    from . import common
except ImportError:  # pragma: no cover - permite ejecución directa como script
    import sys

    sys.path.append(str(Path(__file__).resolve().parent))
    import common  # type: ignore[no-redef]


def delete_normal_templates() -> None:
    common.remove_normal_templates(design_mode=False, emit=print)


if __name__ == "__main__":
    delete_normal_templates()
