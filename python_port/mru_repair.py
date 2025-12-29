"""Rutinas básicas para reparar listas MRU de plantillas."""
from __future__ import annotations

import logging
from pathlib import Path
from typing import List

from . import constants
from .mru_resolver import MRUPathInfo, detect_mru_path
from .utils import is_windows, normalize_path

LOGGER = logging.getLogger(__name__)

try:  # pragma: no cover
    import winreg  # type: ignore
except ImportError:  # pragma: no cover
    winreg = None  # type: ignore[misc]


def repair_template_mru(design_mode: bool = False) -> List[str]:
    """Limpia entradas MRU apuntando a archivos inexistentes."""

    processed_paths: list[str] = []
    for app in constants.APP_NAMES:
        for auth_mode in (None, "ADAL", "LIVEID"):
            info = detect_mru_path(app, auth_mode)
            processed_paths.append(info.registry_path)
            if not is_windows() or not winreg:
                LOGGER.debug(
                    "[DEBUG] Limpieza MRU simulada para %s (%s)", app, info.registry_path
                )
                continue
            _clear_missing_entries(info, design_mode)
    return processed_paths


def _clear_missing_entries(info: MRUPathInfo, design_mode: bool) -> None:
    key_path = info.registry_path
    if key_path.upper().startswith("HKCU\\"):
        key_path = key_path[5:]

    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            key_path,
            0,
            winreg.KEY_ALL_ACCESS,  # type: ignore[arg-type]
        ) as key:  # pragma: no cover - solo en Windows
            stale_values: list[str] = []
            index = 0
            while True:
                try:
                    name, data, _ = winreg.EnumValue(key, index)
                except OSError:
                    break
                index += 1
                file_path = _extract_file_path_from_mru(data)
                if file_path and not Path(file_path).exists():
                    stale_values.append(name)

            if not stale_values:
                if design_mode:
                    LOGGER.info("[INFO] %s sin cambios.", info.registry_path)
                return

            if design_mode:
                LOGGER.info(
                    "[INFO] Eliminando %d entradas obsoletas en %s.",
                    len(stale_values),
                    info.registry_path,
                )

            for name in stale_values:
                try:
                    winreg.DeleteValue(key, name)
                except OSError:
                    LOGGER.warning(
                        "[WARN] No se pudo borrar %s en %s.", name, info.registry_path
                    )
    except FileNotFoundError:
        if design_mode:
            LOGGER.info("[INFO] Clave MRU ausente: %s", info.registry_path)


def _extract_file_path_from_mru(raw_value: str) -> str:
    try:
        value = raw_value.split("*")[-1]
    except AttributeError:
        return ""
    return str(normalize_path(value))
