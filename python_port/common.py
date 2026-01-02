"""Funciones compartidas para instalar/desinstalar plantillas de Office."""
from __future__ import annotations

import logging
import os
import shutil
import sys
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, Iterator, List, Optional
import xml.etree.ElementTree as ET

LOGGER = logging.getLogger(__name__)


# --------------------------------------------------------------------------- #
# Constantes base
# --------------------------------------------------------------------------- #

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
AUTHOR_VALIDATION_ENABLED = os.environ.get("AuthorValidationEnabled", "TRUE").lower() != "false"

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

SUPPORTED_TEMPLATE_EXTENSIONS = {
    ".dotx",
    ".dotm",
    ".potx",
    ".potm",
    ".xltx",
    ".xltm",
    ".thmx",
}

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


# --------------------------------------------------------------------------- #
# Helpers genéricos
# --------------------------------------------------------------------------- #


def normalize_path(path: Path | str | None) -> Path:
    if path is None:
        return Path()
    return Path(str(path).strip().rstrip("\\/"))


def ensure_directory(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def iter_template_files(base_dir: Path) -> Iterator[Path]:
    for ext in SUPPORTED_TEMPLATE_EXTENSIONS:
        yield from base_dir.glob(f"*{ext}")


def resolve_base_directory(base_dir: Path) -> Path:
    """Busca la carpeta que contiene las plantillas dentro de la ruta actual."""
    candidates = [base_dir, base_dir / "payload", base_dir / "templates", base_dir / "extracted"]
    for candidate in candidates:
        if any(candidate.glob("*.dot*")) or any(candidate.glob("*.pot*")) or any(candidate.glob("*.xlt*")):
            return normalize_path(candidate)
    return normalize_path(base_dir)


def path_in_appdata(path: Path) -> bool:
    try:
        return normalize_path(path).resolve().as_posix().startswith(
            normalize_path(APPDATA_PATH).resolve().as_posix()
        )
    except OSError:
        return False


def ensure_parents_and_copy(source: Path, destination: Path) -> None:
    ensure_directory(destination.parent)
    shutil.copy2(source, destination)


# --------------------------------------------------------------------------- #
# Autoría
# --------------------------------------------------------------------------- #


@dataclass
class AuthorCheckResult:
    allowed: bool
    message: str
    authors: List[str]
    error: bool = False

    def as_cli_output(self) -> str:
        return "TRUE" if self.allowed and not self.error else "FALSE"


def check_template_author(
    target: Path,
    allowed_authors: Iterable[str] | None = None,
    validation_enabled: bool = True,
) -> AuthorCheckResult:
    allowed = _normalize_allowed_authors(allowed_authors or DEFAULT_ALLOWED_TEMPLATE_AUTHORS)
    target = normalize_path(target)

    if not target.exists():
        return AuthorCheckResult(
            allowed=False,
            message=f"[ERROR] No se encontró la ruta: \"{target}\"",
            authors=[],
            error=True,
        )

    if target.is_dir():
        authors_found: list[str] = []
        for file in iter_template_files(target):
            author, error = _extract_author(file)
            if error:
                LOGGER.warning(error)
            if author:
                authors_found.append(author)
                LOGGER.info("Archivo: %s - Autor: %s", file.name, author)
            else:
                LOGGER.info("Archivo: %s - Autor: [VACÍO]", file.name)

        message = (
            f"[INFO] Autores listados para la carpeta \"{target}\"."
            if authors_found
            else f"[WARN] No se encontraron plantillas en \"{target}\"."
        )
        return AuthorCheckResult(True, message, authors_found)

    if not validation_enabled:
        return AuthorCheckResult(True, "[INFO] Validación de autores deshabilitada.", [])

    author, error = _extract_author(target)
    if error:
        return AuthorCheckResult(False, error, [], error=True)
    if not author:
        return AuthorCheckResult(False, f"[WARN] El archivo \"{target}\" no tiene autor asignado.", [])

    is_allowed = any(author.lower() == a.lower() for a in allowed)
    message = "[OK] Autor aprobado." if is_allowed else f"[BLOCKED] Autor no permitido para \"{target}\"."
    return AuthorCheckResult(is_allowed, message, [author])


def _normalize_allowed_authors(authors: Iterable[str]) -> list[str]:
    normalized: list[str] = []
    for author in authors:
        cleaned = author.strip()
        if cleaned:
            normalized.append(cleaned)
    return normalized


def _extract_author(template_path: Path) -> tuple[Optional[str], Optional[str]]:
    if not template_path.exists():
        return None, f"[ERROR] No se encontró la ruta: \"{template_path}\""

    try:
        with zipfile.ZipFile(template_path) as zipped:
            try:
                with zipped.open("docProps/core.xml") as core_file:
                    tree = ET.fromstring(core_file.read())
            except KeyError:
                return None, f"[WARN] No se pudo obtener el autor para \"{template_path.name}\" (core.xml ausente)."
    except Exception as exc:  # noqa: BLE001
        return None, f"[ERROR] {template_path.name}: {exc}"

    for candidate in ("{http://purl.org/dc/elements/1.1/}creator", "creator"):
        node = tree.find(candidate)
        if node is not None and node.text:
            return node.text.strip(), None
    return None, f"[WARN] \"{template_path.name}\" sin autor definido."


# --------------------------------------------------------------------------- #
# Instalación / desinstalación
# --------------------------------------------------------------------------- #


@dataclass
class InstallFlags:
    open_word: bool = False
    open_ppt: bool = False
    open_excel: bool = False
    open_document_theme: bool = False
    document_theme_selection: Optional[Path] = None
    custom_selection: Optional[Path] = None
    roaming_selection: Optional[Path] = None
    excel_startup_selection: Optional[Path] = None
    totals: dict = None

    def __post_init__(self) -> None:
        if self.totals is None:
            self.totals = {"files": 0, "errors": 0, "blocked": 0}


def install_template(
    app_label: str,
    filename: str,
    source_root: Path,
    destination_root: Path,
    flags: InstallFlags,
    allowed_authors: Iterable[str],
    validation_enabled: bool,
    design_mode: bool,
) -> None:
    source = normalize_path(source_root / filename)
    destination_root = ensure_directory(normalize_path(destination_root))
    destination = destination_root / filename

    if not source.exists():
        if design_mode:
            LOGGER.warning("[WARNING] Archivo fuente no encontrado: %s", source)
        flags.totals["errors"] += 1
        return

    author_check = check_template_author(source, allowed_authors=allowed_authors, validation_enabled=validation_enabled)
    if not author_check.allowed:
        if design_mode:
            LOGGER.warning(author_check.message)
        flags.totals["blocked"] += 1
        return

    backup_existing(destination, design_mode)
    try:
        ensure_parents_and_copy(source, destination)
        flags.totals["files"] += 1
        if design_mode:
            LOGGER.info("[OK] Copiado %s a %s", filename, destination)
    except OSError as exc:
        flags.totals["errors"] += 1
        LOGGER.error("[ERROR] Falló la copia de %s (%s)", filename, exc)
        return

    if app_label == "WORD":
        flags.open_word = True
        if destination_root == DEFAULT_ROAMING_TEMPLATE_FOLDER:
            flags.roaming_selection = destination
    elif app_label == "POWERPOINT":
        flags.open_ppt = True
        if destination_root == DEFAULT_ROAMING_TEMPLATE_FOLDER:
            flags.roaming_selection = destination
    elif app_label == "EXCEL":
        flags.open_excel = True
        if destination_root == DEFAULT_EXCEL_STARTUP_FOLDER:
            flags.excel_startup_selection = destination

    if destination_root == DEFAULT_CUSTOM_OFFICE_TEMPLATE_PATH:
        flags.custom_selection = destination
    if destination_root == DEFAULT_CUSTOM_OFFICE_ADDITIONAL_TEMPLATE_PATH:
        flags.custom_selection = flags.custom_selection or destination
    if destination_root == DEFAULT_ROAMING_TEMPLATE_FOLDER and filename.lower().endswith(".thmx"):
        flags.open_document_theme = True
        flags.document_theme_selection = destination


def copy_custom_templates(base_dir: Path, destinations: dict[str, Path], flags: InstallFlags, allowed: Iterable[str], validation_enabled: bool, design_mode: bool) -> None:
    for file in iter_template_files(base_dir):
        filename = file.name
        extension = file.suffix.lower()
        if filename in BASE_TEMPLATE_NAMES:
            continue
        destination_root = _destination_for_extension(extension, destinations)
        if destination_root is None:
            if design_mode:
                LOGGER.warning("[WARNING] No hay destino para %s", filename)
            continue

        result = check_template_author(file, allowed_authors=allowed, validation_enabled=validation_enabled)
        if not result.allowed:
            flags.totals["blocked"] += 1
            if design_mode:
                LOGGER.warning(result.message)
            continue

        try:
            ensure_parents_and_copy(file, destination_root / filename)
            flags.totals["files"] += 1
        except OSError as exc:
            flags.totals["errors"] += 1
            LOGGER.error("[ERROR] Falló la copia de %s (%s)", filename, exc)
            continue

        if extension in {".dotx", ".dotm"}:
            flags.open_word = True
        if extension in {".potx", ".potm"}:
            flags.open_ppt = True
        if extension in {".xltx", ".xltm"}:
            flags.open_excel = True
        if destination_root == DEFAULT_ROAMING_TEMPLATE_FOLDER:
            flags.roaming_selection = destination_root / filename
        if destination_root == DEFAULT_EXCEL_STARTUP_FOLDER:
            flags.excel_startup_selection = destination_root / filename
        if extension == ".thmx":
            flags.open_document_theme = True
            flags.document_theme_selection = destination_root / filename
        if destination_root in {DEFAULT_CUSTOM_OFFICE_TEMPLATE_PATH, DEFAULT_CUSTOM_OFFICE_ADDITIONAL_TEMPLATE_PATH}:
            flags.custom_selection = flags.custom_selection or destination_root / filename


def remove_installed_templates(destinations: dict[str, Path], design_mode: bool) -> None:
    targets = {
        destinations["WORD"]: ["Normal.dotx", "Normal.dotm", "NormalEmail.dotx", "NormalEmail.dotm"],
        destinations["POWERPOINT"]: ["Blank.potx", "Blank.potm"],
        destinations["EXCEL"]: ["Book.xltx", "Book.xltm", "Sheet.xltx", "Sheet.xltm"],
        destinations["THEMES"]: [],
    }
    for root, files in targets.items():
        for name in files:
            target = normalize_path(root / name)
            try:
                if target.exists():
                    target.unlink()
                    if design_mode:
                        LOGGER.info("[INFO] Eliminado %s", target)
            except OSError as exc:
                LOGGER.warning("[WARN] No se pudo eliminar %s (%s)", target, exc)


def delete_custom_copies(base_dir: Path, destinations: dict[str, Path], design_mode: bool) -> None:
    for file in iter_template_files(base_dir):
        if file.name in BASE_TEMPLATE_NAMES:
            continue
        for dest in destinations.values():
            candidate = normalize_path(dest / file.name)
            try:
                if candidate.exists():
                    candidate.unlink()
                    if design_mode:
                        LOGGER.info("[INFO] Eliminado %s", candidate)
            except OSError as exc:
                LOGGER.warning("[WARN] No se pudo eliminar %s (%s)", candidate, exc)


def backup_existing(target_file: Path, design_mode: bool) -> None:
    if not target_file.exists():
        return
    backup_dir = target_file.parent / "Backup"
    ensure_directory(backup_dir)
    timestamp = datetime.now().strftime("%Y.%m.%d.%H%M")
    backup_path = backup_dir / f"{timestamp}_{target_file.name}"
    try:
        shutil.copy2(target_file, backup_path)
        if design_mode:
            LOGGER.info("[BACKUP] Copia creada en %s", backup_path)
    except OSError as exc:
        LOGGER.warning("[WARN] No se pudo crear backup de %s (%s)", target_file, exc)


# --------------------------------------------------------------------------- #
# Utilidades plataforma
# --------------------------------------------------------------------------- #


def is_windows() -> bool:
    return os.name == "nt"


def close_office_apps(design_mode: bool) -> None:
    if not is_windows():
        return
    for exe in ("WINWORD.EXE", "POWERPNT.EXE", "EXCEL.EXE"):
        try:
            os.system(f"taskkill /IM {exe} /F >nul 2>&1")
        except OSError:
            if design_mode:
                LOGGER.debug("[DEBUG] No se pudo cerrar %s", exe)


def launch_office_apps(flags: InstallFlags, design_mode: bool) -> None:
    if not is_windows():
        if design_mode:
            LOGGER.info("[WARN] Apertura de aplicaciones omitida: no es Windows.")
        return
    launches = []
    if flags.open_word:
        launches.append(("winword.exe", "Microsoft Word"))
    if flags.open_ppt:
        launches.append(("powerpnt.exe", "Microsoft PowerPoint"))
    if flags.open_excel:
        launches.append(("excel.exe", "Microsoft Excel"))
    for exe, label in launches:
        try:
            if design_mode:
                LOGGER.info("[ACTION] Lanzando %s", label)
            os.startfile(exe)  # type: ignore[arg-type]
        except OSError as exc:
            LOGGER.warning("[WARN] No se pudo iniciar %s (%s)", label, exc)


# --------------------------------------------------------------------------- #
# Utilidades de ruta
# --------------------------------------------------------------------------- #


def default_destinations() -> dict[str, Path]:
    return {
        "WORD": DEFAULT_ROAMING_TEMPLATE_FOLDER,
        "POWERPOINT": DEFAULT_ROAMING_TEMPLATE_FOLDER,
        "EXCEL": DEFAULT_EXCEL_STARTUP_FOLDER,
        "CUSTOM": DEFAULT_CUSTOM_OFFICE_TEMPLATE_PATH,
        "CUSTOM_ALT": DEFAULT_CUSTOM_OFFICE_ADDITIONAL_TEMPLATE_PATH,
        "ROAMING": DEFAULT_ROAMING_TEMPLATE_FOLDER,
        "THEMES": DEFAULT_ROAMING_TEMPLATE_FOLDER / "Document Themes",
    }


def _destination_for_extension(extension: str, destinations: dict[str, Path]) -> Optional[Path]:
    if extension in {".dotx", ".dotm"}:
        return destinations["WORD"]
    if extension in {".potx", ".potm"}:
        return destinations["POWERPOINT"]
    if extension in {".xltx", ".xltm"}:
        return destinations["EXCEL"]
    if extension == ".thmx":
        return destinations["THEMES"]
    return None


def configure_logging(design_mode: bool) -> None:
    level = logging.DEBUG if design_mode else logging.INFO
    logging.basicConfig(level=level, format="%(message)s")


def exit_with_error(message: str) -> None:
    print(message)
    sys.exit(1)
