"""Constantes y configuraciones compartidas para el port en Python."""
from __future__ import annotations

from pathlib import Path
import os

APPDATA_PATH = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
DOCUMENTS_PATH = Path(os.environ.get("USERPROFILE", Path.home())) / "Documents"

DEFAULT_ALLOWED_TEMPLATE_AUTHORS = [
    "www.grada.cc",
    "www.gradaz.com",
]

DEFAULT_DOCUMENT_THEME_DELAY_SECONDS = int(
    os.environ.get("DOCUMENT_THEME_OPEN_DELAY_SECONDS", "0") or 0
)
DEFAULT_DESIGN_MODE = os.environ.get("IsDesignModeEnabled", "false").lower() == "true"

DEFAULT_CUSTOM_OFFICE_TEMPLATE_PATH = Path(
    os.environ.get(
        "CUSTOM_OFFICE_TEMPLATE_PATH",
        DOCUMENTS_PATH / "Custom Office Templates",
    )
)
DEFAULT_CUSTOM_OFFICE_ADDITIONAL_TEMPLATE_PATH = Path(
    os.environ.get(
        "CUSTOM_OFFICE_ADDITIONAL_TEMPLATE_PATH",
        DOCUMENTS_PATH / "Plantillas personalizadas de Office",
    )
)

DEFAULT_ROAMING_TEMPLATE_FOLDER = Path(
    os.environ.get(
        "ROAMING_TEMPLATE_FOLDER_PATH",
        APPDATA_PATH / "Microsoft" / "Templates",
    )
)
DEFAULT_EXCEL_STARTUP_FOLDER = Path(
    os.environ.get(
        "EXCEL_STARTUP_FOLDER_PATH",
        APPDATA_PATH / "Microsoft" / "Excel" / "XLSTART",
    )
)

OFFICE_VERSIONS = ("16.0", "15.0", "14.0", "12.0")
SUPPORTED_TEMPLATE_EXTENSIONS = {".dotx", ".dotm", ".potx", ".potm", ".xltx", ".xltm", ".thmx"}
BASE_TEMPLATE_NAMES = {
    "Normal.dotx",
    "Normal.dotm",
    "NormalEmail.dotx",
    "NormalEmail.dotm",
    "Blank.potx",
    "Blank.potm",
    "Book.xltx",
    "Book.xltm",
    "Sheet.xltx",
    "Sheet.xltm",
}

# Asociación de apps a nombres usados en el registro
APP_REGISTRY_NAMES = {
    "WORD": "Word",
    "POWERPOINT": "PowerPoint",
    "EXCEL": "Excel",
}

# Aplicaciones soportadas para instalación y MRU
APP_NAMES = ("WORD", "POWERPOINT", "EXCEL")
