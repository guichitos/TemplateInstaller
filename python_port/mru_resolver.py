"""Resolución de rutas MRU y utilidades relacionadas."""
from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from . import constants
from .utils import is_windows, normalize_path

LOGGER = logging.getLogger(__name__)

try:  # pragma: no cover - solo se usa en Windows
    import winreg  # type: ignore
except ImportError:  # pragma: no cover - entornos no Windows
    winreg = None  # type: ignore[misc]


@dataclass
class MRUPathInfo:
    app_name: str
    auth_mode: Optional[str]
    registry_path: str
    source: str


def _resolve_registry_name(app_name: str) -> Optional[str]:
    return constants.APP_REGISTRY_NAMES.get(app_name.upper())


def detect_mru_path(app_name: str, auth_mode: Optional[str] = None) -> MRUPathInfo:
    """Replica la detección de listas MRU por app y modo de autenticación."""

    registry_name = _resolve_registry_name(app_name)
    if not registry_name:
        raise ValueError(f"Aplicación desconocida: {app_name}")

    auth_mode_normalized = auth_mode.upper() if auth_mode else None
    fallback = f"HKCU\\Software\\Microsoft\\Office\\16.0\\{registry_name}\\Recent Templates\\File MRU"

    if not is_windows() or not winreg:
        source = "fallback (no Windows)" if not is_windows() else "winreg no disponible"
        return MRUPathInfo(app_name=app_name, auth_mode=auth_mode_normalized, registry_path=fallback, source=source)

    for version in constants.OFFICE_VERSIONS:
        base = f"Software\\Microsoft\\Office\\{version}\\{registry_name}\\Recent Templates"
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, base) as key:  # type: ignore[arg-type]
                subkey_path = _scan_recent_template_key(key, base, auth_mode_normalized)
                if subkey_path:
                    return MRUPathInfo(
                        app_name=app_name,
                        auth_mode=auth_mode_normalized,
                        registry_path=subkey_path,
                        source=f"registry {version}",
                    )
        except FileNotFoundError:
            continue

    return MRUPathInfo(app_name=app_name, auth_mode=auth_mode_normalized, registry_path=fallback, source="fallback")


def _scan_recent_template_key(key: "winreg.HKEYType", base_path: str, auth_mode: Optional[str]) -> Optional[str]:
    # Buscar primero contenedores ADAL/LIVEID si se pidió un auth_mode específico
    try:
        subkeys = []
        index = 0
        while True:
            try:
                subkeys.append(winreg.EnumKey(key, index))  # type: ignore[arg-type]
                index += 1
            except OSError:
                break
        for subkey in subkeys:
            prefix = subkey.upper()
            if auth_mode and not prefix.startswith(f"{auth_mode.upper()}_"):
                continue
            if prefix.startswith("ADAL_") or prefix.startswith("LIVEID_"):
                return f"HKCU\\{base_path}\\{subkey}\\File MRU"
    except OSError:
        pass

    # Si no se pidió auth_mode o no hubo coincidencias, intentar File MRU directo
    try:
        winreg.OpenKey(key, "File MRU")  # type: ignore[arg-type]
        return f"HKCU\\{base_path}\\File MRU"
    except FileNotFoundError:
        return None


def record_recent_template(app_name: str, file_path: Path, design_mode: bool = False) -> None:
    """Guarda una traza de MRU (se simula en Linux, se respeta en Windows)."""

    file_path = normalize_path(file_path)
    info = detect_mru_path(app_name)
    if not is_windows() or not winreg:
        LOGGER.info(
            "[DEBUG] Registro MRU simulado para %s en %s (no Windows).",
            app_name,
            info.registry_path,
        )
        return

    # Para evitar efectos secundarios, solo registramos una traza de intención.
    LOGGER.info(
        "[INFO] Se debería registrar el archivo %s en %s (modo diseño=%s).",
        file_path,
        info.registry_path,
        design_mode,
    )
