"""Elimina plantillas Normal de Word en una ruta fija."""
from __future__ import annotations

from pathlib import Path

try:
    from . import common
except ImportError:  # pragma: no cover - permite ejecución directa como script
    import sys

    sys.path.append(str(Path(__file__).resolve().parent))
    import common  # type: ignore[no-redef]


TEMPLATE_DIR = common.resolve_template_paths()["ROAMING"]
TARGET_FILES = ("Normal.dotx", "Normal.dotm", "NormalEmail.dotx", "NormalEmail.dotm")


def delete_normal_templates() -> None:
    results: list[str] = []
    if not TEMPLATE_DIR.exists():
        print(f"[ERROR] La carpeta no existe: {TEMPLATE_DIR}")
        return

    for filename in TARGET_FILES:
        target = TEMPLATE_DIR / filename
        if not target.exists():
            results.append(f"[SKIP] No existe: {target}")
            continue
        try:
            target.unlink()
            if target.exists():
                results.append(f"[WARN] Persistió tras borrar: {target}")
            else:
                results.append(f"[OK] Eliminado: {target}")
        except OSError as exc:
            results.append(f"[ERROR] No se pudo eliminar {target} ({exc})")

    if results:
        print("\n".join(results))


if __name__ == "__main__":
    delete_normal_templates()
