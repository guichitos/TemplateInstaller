"""Funciones de utilería comunes para el port en Python."""
from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Iterable, Iterator

from . import constants


LOGGER = logging.getLogger(__name__)


def is_windows() -> bool:
    """Indica si el intérprete corre sobre Windows.

    El instalador interactúa con el registro y con Explorer, por lo que
    las operaciones reales solo son efectivas en Windows. En otros SO los
    pasos críticos se simulan y se dejan trazas de depuración.
    """

    return os.name == "nt"


def normalize_path(path: Path | str | None) -> Path:
    """Normaliza rutas quitando espacios o separadores finales."""

    if path is None:
        return Path()
    text = str(path).strip().rstrip("\\/")
    return Path(text)


def ensure_directory(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def path_in_appdata(path: Path) -> bool:
    try:
        return normalize_path(path).resolve().as_posix().startswith(
            normalize_path(constants.APPDATA_PATH).resolve().as_posix()
        )
    except OSError:
        return False


def iter_template_files(base_dir: Path) -> Iterator[Path]:
    """Itera sobre archivos de plantilla admitidos en el nivel raíz."""

    for extension in sorted(constants.SUPPORTED_TEMPLATE_EXTENSIONS):
        for file in base_dir.glob(f"*{extension}"):
            if file.is_file():
                yield file


def humanize_bool(value: bool) -> str:
    return "true" if value else "false"


def log_platform_warning(action: str) -> None:
    if not is_windows():
        LOGGER.warning(
            "[AVISO] %s se omite porque el sistema no es Windows.", action
        )


def format_authors_list(authors: Iterable[str]) -> str:
    return ";".join(sorted({a.strip() for a in authors if a.strip()}))
