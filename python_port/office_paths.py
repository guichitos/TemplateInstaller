"""Detección de rutas de plantillas para Word, PowerPoint y Excel."""
from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from . import constants
from .utils import ensure_directory, is_windows, normalize_path

LOGGER = logging.getLogger(__name__)

try:  # pragma: no cover
    import winreg  # type: ignore
except ImportError:  # pragma: no cover
    winreg = None  # type: ignore[misc]


@dataclass
class OfficePaths:
    word: Path
    powerpoint: Path
    excel: Path
    document_theme: Path
    default_custom_dir: Path
    appdata_expanded: Path
    documents_path: Path
    statuses: dict[str, str]

    def normalized(self) -> "OfficePaths":
        return OfficePaths(
            word=normalize_path(self.word),
            powerpoint=normalize_path(self.powerpoint),
            excel=normalize_path(self.excel),
            document_theme=normalize_path(self.document_theme),
            default_custom_dir=normalize_path(self.default_custom_dir),
            appdata_expanded=normalize_path(self.appdata_expanded),
            documents_path=normalize_path(self.documents_path),
            statuses=self.statuses,
        )



def detect_office_paths(design_mode: bool = False) -> OfficePaths:
    appdata_expanded = _read_user_shell_folder("AppData") or constants.APPDATA_PATH
    documents_path = _read_user_shell_folder("Personal") or constants.DOCUMENTS_PATH

    document_theme_folder = appdata_expanded / "Microsoft" / "Templates" / "Document Themes"
    default_custom_dir = documents_path / "Custom Templates"

    statuses: dict[str, str] = {}

    document_theme_folder = _ensure_directory_with_status(
        document_theme_folder, statuses, "DOCUMENT_THEME_FOLDER"
    )
    default_custom_dir = _ensure_directory_with_status(
        default_custom_dir, statuses, "DEFAULT_CUSTOM_DIR"
    )

    word_path = _lookup_template_path("Word", ("Options", "Common\\General")) or default_custom_dir
    ppt_path = _lookup_template_path("PowerPoint", ("Options", "Common\\General")) or default_custom_dir
    excel_path = _lookup_template_path("Excel", ("Options", "Common\\General")) or default_custom_dir

    if design_mode:
        LOGGER.debug("[DEBUG] Word templates path: %s", word_path)
        LOGGER.debug("[DEBUG] PowerPoint templates path: %s", ppt_path)
        LOGGER.debug("[DEBUG] Excel templates path: %s", excel_path)
        LOGGER.debug("[DEBUG] Document Themes path: %s", document_theme_folder)

    for path in (word_path, ppt_path, excel_path, document_theme_folder):
        try:
            ensure_directory(path)
        except OSError:
            LOGGER.warning("[WARN] No se pudo crear la carpeta %s", path)

    return OfficePaths(
        word=word_path,
        powerpoint=ppt_path,
        excel=excel_path,
        document_theme=document_theme_folder,
        default_custom_dir=default_custom_dir,
        appdata_expanded=appdata_expanded,
        documents_path=documents_path,
        statuses=statuses,
    ).normalized()


def _ensure_directory_with_status(path: Path, statuses: dict[str, str], label: str) -> Path:
    try:
        ensure_directory(path)
        statuses[label] = "created" if not path.exists() else "exists"
    except OSError:
        statuses[label] = "create_failed"
    return path


def _lookup_template_path(app_registry_name: str, subkeys: tuple[str, str]) -> Optional[Path]:
    if not is_windows() or not winreg:
        return None

    for version in constants.OFFICE_VERSIONS:
        base = f"Software\\Microsoft\\Office\\{version}\\{app_registry_name}"
        for leaf, value_name in (
            (subkeys[0], "PersonalTemplates"),
            (subkeys[1], "UserTemplates"),
        ):
            target = f"{base}\\{leaf}"
            result = _read_registry_value(winreg.HKEY_CURRENT_USER, target, value_name)
            if result:
                return normalize_path(result)
    return None


def _read_registry_value(root: "winreg.HKEYType", key: str, value_name: str) -> Optional[Path]:
    try:
        with winreg.OpenKey(root, key) as registry_key:  # type: ignore[arg-type]
            value, _ = winreg.QueryValueEx(registry_key, value_name)
            return normalize_path(value)
    except FileNotFoundError:
        return None
    except OSError:
        return None


def _read_user_shell_folder(value_name: str) -> Optional[Path]:
    if not is_windows() or not winreg:
        return None
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders",
        ) as key:  # type: ignore[arg-type]
            value, _ = winreg.QueryValueEx(key, value_name)
            return normalize_path(value)
    except FileNotFoundError:
        return None
    except OSError:
        return None
