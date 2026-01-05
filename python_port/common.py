"""Funciones compartidas para instalar/desinstalar plantillas de Office."""
from __future__ import annotations

import logging
import os
import shutil
import subprocess
import sys
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, Iterator, List, Optional
import xml.etree.ElementTree as ET


def normalize_path(path: Path | str | None) -> Path:
    if path is None:
        return Path()
    return Path(str(path).strip().rstrip("\\/"))


try:
    import winreg  # type: ignore[import-not-found]
except Exception:  # pragma: no cover - entornos no Windows
    winreg = None  # type: ignore[assignment]

LOGGER = logging.getLogger(__name__)

# --------------------------------------------------------------------------- #
# Flags manuales de diseño (override de variables de entorno)
# --------------------------------------------------------------------------- #
# Pon en True/False para forzar logs por categoría; deja en None para usar
# la variable de entorno correspondiente o, en su defecto, IsDesignModeEnabled.
MANUAL_DESIGN_LOG_PATHS: bool | None = None
MANUAL_DESIGN_LOG_MRU: bool | None = None
MANUAL_DESIGN_LOG_OPENING: bool | None = None
MANUAL_DESIGN_LOG_AUTHOR: bool | None = None
MANUAL_DESIGN_LOG_COPY_BASE: bool | None = None
MANUAL_DESIGN_LOG_COPY_CUSTOM: bool | None = None
MANUAL_DESIGN_LOG_BACKUP: bool | None = None
MANUAL_DESIGN_LOG_APP_LAUNCH: bool | None = None
MANUAL_DESIGN_LOG_CLOSE_APPS: bool | None = None
MANUAL_DESIGN_LOG_INSTALLER: bool | None = None
MANUAL_DESIGN_LOG_UNINSTALLER: bool | None = None


# --------------------------------------------------------------------------- #
# Constantes base
# --------------------------------------------------------------------------- #

_BASE_PATHS = None


def _read_registry_value(path: str, name: str) -> Optional[str]:
    if winreg is None:
        return None
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
            value, _ = winreg.QueryValueEx(key, name)
            return os.path.expandvars(str(value))
    except OSError:
        return None


def _resolve_appdata_path() -> Path:
    appdata = _read_registry_value(
        r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "AppData"
    )
    if not appdata:
        appdata = os.environ.get("APPDATA")
    return normalize_path(appdata or (Path.home() / "AppData" / "Roaming"))


def _resolve_documents_path() -> Path:
    documents = _read_registry_value(
        r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders", "Personal"
    )
    if not documents:
        documents = os.environ.get("USERPROFILE")
        if documents:
            documents = str(Path(documents) / "Documents")
    return normalize_path(documents or (Path.home() / "Documents"))


def _resolve_custom_template_path(default_custom_dir: Path) -> Path:
    if winreg:
        for version in ("16.0", "15.0", "14.0", "12.0"):
            value = _read_registry_value(
                fr"Software\Microsoft\Office\{version}\Word\Options", "PersonalTemplates"
            )
            if value:
                return normalize_path(value)
        for version in ("16.0", "15.0", "14.0", "12.0"):
            value = _read_registry_value(
                fr"Software\Microsoft\Office\{version}\Common\General", "UserTemplates"
            )
            if value:
                return normalize_path(value)
    return normalize_path(default_custom_dir)


def _resolve_custom_alt_path(custom_primary: Path, default_custom_dir: Path, default_alt_dir: Path) -> Path:
    if winreg:
        for version in ("16.0", "15.0", "14.0", "12.0"):
            value = _read_registry_value(
                fr"Software\Microsoft\Office\{version}\PowerPoint\Options", "PersonalTemplates"
            )
            if value:
                return normalize_path(value)
        for version in ("16.0", "15.0", "14.0", "12.0"):
            value = _read_registry_value(
                fr"Software\Microsoft\Office\{version}\Common\General", "UserTemplates"
            )
            if value:
                return normalize_path(value)
    return normalize_path(custom_primary or default_custom_dir or default_alt_dir)


def _resolve_excel_template_path(custom_primary: Path, default_custom_dir: Path, default_alt_dir: Path) -> Path:
    if winreg:
        for version in ("16.0", "15.0", "14.0", "12.0"):
            value = _read_registry_value(
                fr"Software\Microsoft\Office\{version}\Excel\Options", "PersonalTemplates"
            )
            if value:
                return normalize_path(value)
        for version in ("16.0", "15.0", "14.0", "12.0"):
            value = _read_registry_value(
                fr"Software\Microsoft\Office\{version}\Common\General", "UserTemplates"
            )
            if value:
                return normalize_path(value)
    return normalize_path(custom_primary or default_custom_dir or default_alt_dir)


def _resolve_base_paths() -> dict[str, Path]:
    documents_path = _resolve_documents_path()
    default_custom_dir = documents_path / "Custom Office Templates"
    default_custom_alt_dir = documents_path / "Plantillas personalizadas de Office"
    custom_word = _resolve_custom_template_path(default_custom_dir)
    custom_ppt = _resolve_custom_alt_path(custom_word, default_custom_dir, default_custom_alt_dir)
    custom_excel = _resolve_excel_template_path(custom_word, default_custom_dir, default_custom_alt_dir)
    appdata_path = _resolve_appdata_path()
    return {
        "APPDATA": appdata_path,
        "DOCUMENTS": documents_path,
        "CUSTOM_WORD": custom_word,
        "CUSTOM_PPT": custom_ppt,
        "CUSTOM_EXCEL": custom_excel,
        "CUSTOM_ADDITIONAL": default_custom_alt_dir,
        "THEME": appdata_path / "Microsoft" / "Templates" / "Document Themes",
        "ROAMING": appdata_path / "Microsoft" / "Templates",
        "EXCEL_STARTUP": appdata_path / "Microsoft" / "Excel" / "XLSTART",
    }


_BASE_PATHS = _resolve_base_paths()
APPDATA_PATH = _BASE_PATHS["APPDATA"]
DOCUMENTS_PATH = _BASE_PATHS["DOCUMENTS"]

DEFAULT_ALLOWED_TEMPLATE_AUTHORS = [
    "www.grada.cc",
    "www.gradaz.com",
]

DEFAULT_DOCUMENT_THEME_DELAY_SECONDS = int(
    os.environ.get("DOCUMENT_THEME_OPEN_DELAY_SECONDS", "0") or 0
)
DEFAULT_DESIGN_MODE = os.environ.get("IsDesignModeEnabled", "false").lower() == "true"
AUTHOR_VALIDATION_ENABLED = os.environ.get("AuthorValidationEnabled", "TRUE").lower() != "false"


def _design_flag(env_var: str, manual_override: bool | None) -> bool:
    if manual_override is not None:
        return bool(manual_override)
    raw = os.environ.get(env_var)
    if raw is None:
        return DEFAULT_DESIGN_MODE
    return raw.lower() == "true"


DESIGN_LOG_PATHS = _design_flag("DesignLogPaths", MANUAL_DESIGN_LOG_PATHS)
DESIGN_LOG_MRU = _design_flag("DesignLogMRU", MANUAL_DESIGN_LOG_MRU)
DESIGN_LOG_OPENING = _design_flag("DesignLogOpening", MANUAL_DESIGN_LOG_OPENING)
DESIGN_LOG_AUTHOR = _design_flag("DesignLogAuthor", MANUAL_DESIGN_LOG_AUTHOR)
DESIGN_LOG_COPY_BASE = _design_flag("DesignLogCopyBase", MANUAL_DESIGN_LOG_COPY_BASE)
DESIGN_LOG_COPY_CUSTOM = _design_flag("DesignLogCopyCustom", MANUAL_DESIGN_LOG_COPY_CUSTOM)
DESIGN_LOG_BACKUP = _design_flag("DesignLogBackup", MANUAL_DESIGN_LOG_BACKUP)
DESIGN_LOG_APP_LAUNCH = _design_flag("DesignLogAppLaunch", MANUAL_DESIGN_LOG_APP_LAUNCH)
DESIGN_LOG_CLOSE_APPS = _design_flag("DesignLogCloseApps", MANUAL_DESIGN_LOG_CLOSE_APPS)
DESIGN_LOG_INSTALLER = _design_flag("DesignLogInstaller", MANUAL_DESIGN_LOG_INSTALLER)
DESIGN_LOG_UNINSTALLER = _design_flag("DesignLogUninstaller", MANUAL_DESIGN_LOG_UNINSTALLER)

DEFAULT_CUSTOM_OFFICE_TEMPLATE_PATH = normalize_path(
    os.environ.get("CUSTOM_OFFICE_TEMPLATE_PATH", _BASE_PATHS["CUSTOM_WORD"])
)
DEFAULT_POWERPOINT_TEMPLATE_PATH = normalize_path(
    os.environ.get("POWERPOINT_TEMPLATE_PATH", _BASE_PATHS["CUSTOM_PPT"])
)
DEFAULT_EXCEL_TEMPLATE_PATH = normalize_path(
    os.environ.get("EXCEL_TEMPLATE_PATH", _BASE_PATHS["CUSTOM_EXCEL"])
)
DEFAULT_CUSTOM_OFFICE_ADDITIONAL_TEMPLATE_PATH = normalize_path(
    os.environ.get("CUSTOM_OFFICE_ADDITIONAL_TEMPLATE_PATH", _BASE_PATHS["CUSTOM_ADDITIONAL"])
)
DEFAULT_ROAMING_TEMPLATE_FOLDER = normalize_path(
    os.environ.get("ROAMING_TEMPLATE_FOLDER_PATH", _BASE_PATHS["ROAMING"])
)
DEFAULT_EXCEL_STARTUP_FOLDER = normalize_path(
    os.environ.get("EXCEL_STARTUP_FOLDER_PATH", _BASE_PATHS["EXCEL_STARTUP"])
)
DEFAULT_THEME_FOLDER = normalize_path(_BASE_PATHS["THEME"])

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


def _design_log(enabled: bool, design_mode: bool, level: int, message: str, *args: object) -> None:
    if design_mode and enabled:
        LOGGER.log(level, message, *args)


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
    design_mode: bool = False,
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
            if file.suffix.lower() == ".thmx":
                _design_log(DESIGN_LOG_AUTHOR, design_mode, logging.INFO, "Archivo: %s - Autor: [OMITIDO TEMA]", file.name)
                continue
            author, error = _extract_author(file)
            if error:
                _design_log(DESIGN_LOG_AUTHOR, design_mode, logging.WARNING, error)
            if author:
                authors_found.append(author)
                _design_log(DESIGN_LOG_AUTHOR, design_mode, logging.INFO, "Archivo: %s - Autor: %s", file.name, author)
            else:
                _design_log(DESIGN_LOG_AUTHOR, design_mode, logging.INFO, "Archivo: %s - Autor: [VACÍO]", file.name)

        message = (
            f"[INFO] Autores listados para la carpeta \"{target}\"."
            if authors_found
            else f"[WARN] No se encontraron plantillas en \"{target}\"."
        )
        return AuthorCheckResult(True, message, authors_found)

    if not validation_enabled:
        return AuthorCheckResult(True, "[INFO] Validación de autores deshabilitada.", [])

    if target.suffix.lower() == ".thmx":
        return AuthorCheckResult(True, "[INFO] Validación de autor omitida para temas.", [])

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
    open_theme_folder: bool = False
    open_custom_word_folder: bool = False
    open_custom_ppt_folder: bool = False
    open_custom_excel_folder: bool = False
    open_roaming_folder: bool = False
    open_excel_startup_folder: bool = False
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
    destinations_map: dict[str, Path],
    flags: InstallFlags,
    allowed_authors: Iterable[str],
    validation_enabled: bool,
    design_mode: bool,
) -> None:
    source = normalize_path(source_root / filename)
    destination_root = ensure_directory(normalize_path(destination_root))
    destination = destination_root / filename

    if not source.exists():
        _design_log(DESIGN_LOG_COPY_BASE, design_mode, logging.WARNING, "[WARNING] Archivo fuente no encontrado: %s", source)
        flags.totals["errors"] += 1
        return

    author_check = check_template_author(
        source,
        allowed_authors=allowed_authors,
        validation_enabled=validation_enabled,
        design_mode=design_mode,
    )
    if not author_check.allowed:
        _design_log(DESIGN_LOG_AUTHOR, design_mode, logging.WARNING, author_check.message)
        flags.totals["blocked"] += 1
        return

    backup_existing(destination, design_mode)
    try:
        ensure_parents_and_copy(source, destination)
        flags.totals["files"] += 1
        _design_log(DESIGN_LOG_COPY_BASE, design_mode, logging.INFO, "[OK] Copiado %s a %s", filename, destination)
        _mark_folder_open_flag(destination_root, flags, destinations_map)
        _update_mru_if_applicable(app_label, destination, design_mode)
    except OSError as exc:
        flags.totals["errors"] += 1
        _design_log(DESIGN_LOG_COPY_BASE, design_mode, logging.ERROR, "[ERROR] Falló la copia de %s (%s)", filename, exc)
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
        flags.open_custom_word_folder = True
    if destination_root == DEFAULT_CUSTOM_OFFICE_ADDITIONAL_TEMPLATE_PATH:
        flags.custom_selection = flags.custom_selection or destination
        flags.open_custom_excel_folder = True
    if destination_root == DEFAULT_ROAMING_TEMPLATE_FOLDER and filename.lower().endswith(".thmx"):
        flags.open_document_theme = True
        flags.document_theme_selection = destination


def copy_custom_templates(base_dir: Path, destinations: dict[str, Path], flags: InstallFlags, allowed: Iterable[str], validation_enabled: bool, design_mode: bool) -> None:
    for file in iter_template_files(base_dir):
        filename = file.name
        extension = file.suffix.lower()
        if filename in BASE_TEMPLATE_NAMES:
            continue
        if extension in {".xltx", ".xltm"}:
            destination_root = destinations["EXCEL_CUSTOM"]
        elif extension in {".dotx", ".dotm"}:
            destination_root = destinations["WORD_CUSTOM"]
        elif extension in {".potx", ".potm"}:
            destination_root = destinations["POWERPOINT_CUSTOM"]
        else:
            destination_root = _destination_for_extension(extension, destinations)
        if destination_root is None:
            _design_log(DESIGN_LOG_COPY_CUSTOM, design_mode, logging.WARNING, "[WARNING] No hay destino para %s", filename)
            continue

        result = check_template_author(
            file,
            allowed_authors=allowed,
            validation_enabled=validation_enabled,
            design_mode=design_mode,
        )
        if not result.allowed:
            flags.totals["blocked"] += 1
            _design_log(DESIGN_LOG_AUTHOR, design_mode, logging.WARNING, result.message)
            continue

        try:
            ensure_parents_and_copy(file, destination_root / filename)
            flags.totals["files"] += 1
            _mark_folder_open_flag(destination_root, flags, destinations)
            _design_log(
                DESIGN_LOG_COPY_CUSTOM,
                design_mode,
                logging.INFO,
                "[OK] Copiado %s a %s",
                filename,
                destination_root / filename,
            )
            _update_mru_if_applicable_extension(extension, destination_root / filename, design_mode)
        except OSError as exc:
            flags.totals["errors"] += 1
            _design_log(DESIGN_LOG_COPY_CUSTOM, design_mode, logging.ERROR, "[ERROR] Falló la copia de %s (%s)", filename, exc)
            continue

        if extension in {".dotx", ".dotm"}:
            flags.open_word = True
        if extension in {".potx", ".potm"}:
            flags.open_ppt = True
        if extension in {".xltx", ".xltm"}:
            flags.open_excel = True
        if destination_root == DEFAULT_CUSTOM_OFFICE_TEMPLATE_PATH:
            flags.open_custom_word_folder = True
        if destination_root == DEFAULT_POWERPOINT_TEMPLATE_PATH or destination_root == DEFAULT_CUSTOM_OFFICE_TEMPLATE_PATH:
            flags.open_custom_ppt_folder = True
        if destination_root == DEFAULT_EXCEL_TEMPLATE_PATH or destination_root == DEFAULT_CUSTOM_OFFICE_ADDITIONAL_TEMPLATE_PATH:
            flags.open_custom_excel_folder = True
        if destination_root == DEFAULT_ROAMING_TEMPLATE_FOLDER:
            flags.roaming_selection = destination_root / filename
            flags.open_roaming_folder = True
        if destination_root == DEFAULT_EXCEL_STARTUP_FOLDER:
            flags.excel_startup_selection = destination_root / filename
            flags.open_excel_startup_folder = True
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
                    _design_log(DESIGN_LOG_UNINSTALLER, design_mode, logging.INFO, "[INFO] Eliminado %s", target)
            except OSError as exc:
                _design_log(DESIGN_LOG_UNINSTALLER, design_mode, logging.WARNING, "[WARN] No se pudo eliminar %s (%s)", target, exc)


def delete_custom_copies(base_dir: Path, destinations: dict[str, Path], design_mode: bool) -> None:
    for file in iter_template_files(base_dir):
        if file.name in BASE_TEMPLATE_NAMES:
            continue
        for dest in destinations.values():
            candidate = normalize_path(dest / file.name)
            try:
                if candidate.exists():
                    candidate.unlink()
                    _design_log(DESIGN_LOG_UNINSTALLER, design_mode, logging.INFO, "[INFO] Eliminado %s", candidate)
            except OSError as exc:
                _design_log(DESIGN_LOG_UNINSTALLER, design_mode, logging.WARNING, "[WARN] No se pudo eliminar %s (%s)", candidate, exc)


def backup_existing(target_file: Path, design_mode: bool) -> None:
    if not target_file.exists():
        return
    backup_dir = target_file.parent / "Backup"
    ensure_directory(backup_dir)
    timestamp = datetime.now().strftime("%Y.%m.%d.%H%M")
    backup_path = backup_dir / f"{timestamp}_{target_file.name}"
    try:
        shutil.copy2(target_file, backup_path)
        _design_log(DESIGN_LOG_BACKUP, design_mode, logging.INFO, "[BACKUP] Copia creada en %s", backup_path)
    except OSError as exc:
        _design_log(
            DESIGN_LOG_BACKUP,
            design_mode,
            logging.WARNING,
            "[WARN] No se pudo crear backup de %s (%s)",
            target_file,
            exc,
        )


def open_template_folders(paths: dict[str, Path], design_mode: bool, flags: InstallFlags | None = None) -> None:
    if not is_windows():
        _design_log(DESIGN_LOG_OPENING, design_mode, logging.INFO, "[WARN] Apertura de carpetas omitida: no es Windows.")
        return
    ordered = [
        ("THEME_PATH", "open_theme_folder", paths.get("THEME")),
        ("CUSTOM_WORD_TEMPLATE_PATH", "open_custom_word_folder", paths.get("CUSTOM_WORD")),
        ("CUSTOM_PPT_TEMPLATE_PATH", "open_custom_ppt_folder", paths.get("CUSTOM_PPT")),
        ("CUSTOM_EXCEL_TEMPLATE_PATH", "open_custom_excel_folder", paths.get("CUSTOM_EXCEL")),
        ("ROAMING_TEMPLATE_PATH", "open_roaming_folder", paths.get("ROAMING")),
        ("EXCEL_STARTUP_PATH", "open_excel_startup_folder", paths.get("EXCEL")),
        ("CUSTOM_ADDITIONAL_PATH", "open_custom_excel_folder", paths.get("CUSTOM_ADDITIONAL")),
    ]
    for label, flag_name, target in ordered:
        if target is None:
            continue
        if flags is not None and not getattr(flags, flag_name, False):
            continue
        try:
            ensure_directory(target)
            if not target.exists():
                _design_log(DESIGN_LOG_OPENING, design_mode, logging.WARNING, "[WARN] La carpeta %s no existe tras crearla: %s", label, target)
            _design_log(DESIGN_LOG_OPENING, design_mode, logging.INFO, "[ACTION] Abriendo carpeta %s: %s", label, target)
            try:
                os.startfile(str(target))  # type: ignore[arg-type]
                _design_log(DESIGN_LOG_OPENING, design_mode, logging.INFO, "[OK] startfile lanzado para %s", label)
            except OSError as exc:
                _design_log(DESIGN_LOG_OPENING, design_mode, logging.WARNING, "[WARN] startfile falló para %s (%s); usando explorer.", label, exc)
                try:
                    subprocess.run(["explorer", str(target)], check=False)
                except OSError as exc2:
                    _design_log(DESIGN_LOG_OPENING, design_mode, logging.WARNING, "[WARN] explorer también falló para %s (%s)", label, exc2)
        except OSError as exc:
            _design_log(DESIGN_LOG_OPENING, design_mode, logging.WARNING, "[WARN] No se pudo abrir carpeta %s (%s)", label, exc)


def _mark_folder_open_flag(destination_root: Path, flags: InstallFlags, destinations: dict[str, Path]) -> None:
    if destination_root == destinations.get("THEMES"):
        flags.open_theme_folder = True
    if destination_root in {destinations.get("CUSTOM"), destinations.get("WORD_CUSTOM")}:
        flags.open_custom_word_folder = True
    if destination_root == destinations.get("POWERPOINT_CUSTOM"):
        flags.open_custom_ppt_folder = True
    if destination_root in {destinations.get("EXCEL_CUSTOM"), destinations.get("CUSTOM_ALT")}:
        flags.open_custom_excel_folder = True
    if destination_root == destinations.get("ROAMING"):
        flags.open_roaming_folder = True
    if destination_root == destinations.get("EXCEL"):
        flags.open_excel_startup_folder = True


def _update_mru_if_applicable(app_label: str, destination: Path, design_mode: bool) -> None:
    if not _should_update_mru(destination):
        return
    ext = destination.suffix.lower()
    if ext in {".dotx", ".dotm", ".potx", ".potm", ".xltx", ".xltm"}:
        update_mru_for_template(app_label, destination, design_mode)


def _update_mru_if_applicable_extension(extension: str, destination: Path, design_mode: bool) -> None:
    if not _should_update_mru(destination):
        return
    if extension in {".dotx", ".dotm"}:
        update_mru_for_template("WORD", destination, design_mode)
    if extension in {".potx", ".potm"}:
        update_mru_for_template("POWERPOINT", destination, design_mode)
    if extension in {".xltx", ".xltm"}:
        update_mru_for_template("EXCEL", destination, design_mode)


def _should_update_mru(path: Path) -> bool:
    name = path.name
    ext = path.suffix.lower()
    if name in BASE_TEMPLATE_NAMES:
        return False
    if ext == ".thmx":
        return False
    return True


def open_template_folders(paths: dict[str, Path], design_mode: bool, flags: InstallFlags | None = None) -> None:
    if not is_windows():
        if design_mode:
            LOGGER.info("[WARN] Apertura de carpetas omitida: no es Windows.")
        return
    ordered = [
        ("THEME_PATH", "open_theme_folder", paths.get("THEME")),
        ("CUSTOM_WORD_TEMPLATE_PATH", "open_custom_word_folder", paths.get("CUSTOM_WORD")),
        ("CUSTOM_PPT_TEMPLATE_PATH", "open_custom_ppt_folder", paths.get("CUSTOM_PPT")),
        ("CUSTOM_EXCEL_TEMPLATE_PATH", "open_custom_excel_folder", paths.get("CUSTOM_EXCEL")),
        ("ROAMING_TEMPLATE_PATH", "open_roaming_folder", paths.get("ROAMING")),
        ("EXCEL_STARTUP_PATH", "open_excel_startup_folder", paths.get("EXCEL")),
        ("CUSTOM_ADDITIONAL_PATH", "open_custom_excel_folder", paths.get("CUSTOM_ADDITIONAL")),
    ]
    for label, flag_name, target in ordered:
        if target is None:
            continue
        if flags is not None and not getattr(flags, flag_name, False):
            continue
        try:
            ensure_directory(target)
            if design_mode and DESIGN_LOG_OPENING:
                LOGGER.info("[ACTION] Abriendo carpeta %s: %s", label, target)
                if not target.exists():
                    LOGGER.warning("[WARN] La carpeta %s no existe tras crearla: %s", label, target)
            try:
                os.startfile(str(target))  # type: ignore[arg-type]
                if design_mode and DESIGN_LOG_OPENING:
                    LOGGER.info("[OK] startfile lanzado para %s", label)
            except OSError as exc:
                if design_mode and DESIGN_LOG_OPENING:
                    LOGGER.warning("[WARN] startfile falló para %s (%s); usando explorer.", label, exc)
                try:
                    subprocess.run(["explorer", str(target)], check=False)
                except OSError as exc2:
                    if design_mode and DESIGN_LOG_OPENING:
                        LOGGER.warning("[WARN] explorer también falló para %s (%s)", label, exc2)
        except OSError as exc:
            if design_mode and DESIGN_LOG_OPENING:
                LOGGER.warning("[WARN] No se pudo abrir carpeta %s (%s)", label, exc)


def _mark_folder_open_flag(destination_root: Path, flags: InstallFlags, destinations: dict[str, Path]) -> None:
    if destination_root == destinations.get("THEMES"):
        flags.open_theme_folder = True
    if destination_root in {destinations.get("CUSTOM"), destinations.get("WORD_CUSTOM")}:
        flags.open_custom_word_folder = True
    if destination_root == destinations.get("POWERPOINT_CUSTOM"):
        flags.open_custom_ppt_folder = True
    if destination_root in {destinations.get("EXCEL_CUSTOM"), destinations.get("CUSTOM_ALT")}:
        flags.open_custom_excel_folder = True
    if destination_root == destinations.get("ROAMING"):
        flags.open_roaming_folder = True
    if destination_root == destinations.get("EXCEL"):
        flags.open_excel_startup_folder = True


def _update_mru_if_applicable(app_label: str, destination: Path, design_mode: bool) -> None:
    if not _should_update_mru(destination):
        return
    ext = destination.suffix.lower()
    if ext in {".dotx", ".dotm", ".potx", ".potm", ".xltx", ".xltm"}:
        update_mru_for_template(app_label, destination, design_mode)


def _update_mru_if_applicable_extension(extension: str, destination: Path, design_mode: bool) -> None:
    if not _should_update_mru(destination):
        return
    if extension in {".dotx", ".dotm"}:
        update_mru_for_template("WORD", destination, design_mode)
    if extension in {".potx", ".potm"}:
        update_mru_for_template("POWERPOINT", destination, design_mode)
    if extension in {".xltx", ".xltm"}:
        update_mru_for_template("EXCEL", destination, design_mode)


def _should_update_mru(path: Path) -> bool:
    name = path.name
    ext = path.suffix.lower()
    if name in BASE_TEMPLATE_NAMES:
        return False
    if ext == ".thmx":
        return False
    return True


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
            _design_log(DESIGN_LOG_CLOSE_APPS, design_mode, logging.DEBUG, "[DEBUG] No se pudo cerrar %s", exe)


def launch_office_apps(flags: InstallFlags, design_mode: bool) -> None:
    if not is_windows():
        _design_log(DESIGN_LOG_APP_LAUNCH, design_mode, logging.INFO, "[WARN] Apertura de aplicaciones omitida: no es Windows.")
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
            _design_log(DESIGN_LOG_APP_LAUNCH, design_mode, logging.INFO, "[ACTION] Lanzando %s", label)
            os.startfile(exe)  # type: ignore[arg-type]
        except OSError as exc:
            _design_log(DESIGN_LOG_APP_LAUNCH, design_mode, logging.WARNING, "[WARN] No se pudo iniciar %s (%s)", label, exc)


# --------------------------------------------------------------------------- #
# Utilidades de ruta
# --------------------------------------------------------------------------- #


def default_destinations() -> dict[str, Path]:
    paths = resolve_template_paths()
    return {
        "WORD": paths["ROAMING"],
        "POWERPOINT": paths["ROAMING"],
        "EXCEL": paths["EXCEL"],
        "CUSTOM": paths["CUSTOM_WORD"],
        "CUSTOM_ALT": paths["CUSTOM_ADDITIONAL"],
        "WORD_CUSTOM": paths["CUSTOM_WORD"],
        "POWERPOINT_CUSTOM": paths["CUSTOM_PPT"],
        "EXCEL_CUSTOM": paths["CUSTOM_EXCEL"],
        "ROAMING": paths["ROAMING"],
        "THEMES": paths["THEME"],
    }


def resolve_template_paths() -> dict[str, Path]:
    return {
        "THEME": DEFAULT_THEME_FOLDER,
        "CUSTOM_WORD": DEFAULT_CUSTOM_OFFICE_TEMPLATE_PATH,
        "CUSTOM_PPT": DEFAULT_POWERPOINT_TEMPLATE_PATH or DEFAULT_CUSTOM_OFFICE_TEMPLATE_PATH,
        "CUSTOM_EXCEL": DEFAULT_EXCEL_TEMPLATE_PATH or DEFAULT_CUSTOM_OFFICE_TEMPLATE_PATH,
        "CUSTOM_ADDITIONAL": DEFAULT_CUSTOM_OFFICE_ADDITIONAL_TEMPLATE_PATH,
        "ROAMING": DEFAULT_ROAMING_TEMPLATE_FOLDER,
        "EXCEL": DEFAULT_EXCEL_STARTUP_FOLDER,
    }


def log_template_paths(paths: dict[str, Path], design_mode: bool) -> None:
    if not design_mode or not DESIGN_LOG_PATHS:
        return
    logger = logging.getLogger(__name__)
    logger.info("================= RUTAS CALCULADAS =================")
    logger.info("THEME_PATH                  = %s", paths["THEME"])
    logger.info("CUSTOM_WORD_TEMPLATE_PATH   = %s", paths["CUSTOM_WORD"])
    logger.info("CUSTOM_PPT_TEMPLATE_PATH    = %s", paths["CUSTOM_PPT"])
    logger.info("CUSTOM_EXCEL_TEMPLATE_PATH  = %s", paths["CUSTOM_EXCEL"])
    logger.info("CUSTOM_ADDITIONAL_PATH      = %s", paths["CUSTOM_ADDITIONAL"])
    logger.info("ROAMING_TEMPLATE_PATH       = %s", paths["ROAMING"])
    logger.info("EXCEL_STARTUP_PATH          = %s", paths["EXCEL"])
    logger.info("====================================================")


def log_registry_sources(design_mode: bool) -> None:
    if not design_mode or not DESIGN_LOG_MRU:
        return
    logger = logging.getLogger(__name__)
    word_personal = _read_registry_value(r"Software\Microsoft\Office\\16.0\\Word\\Options", "PersonalTemplates")
    word_user = _read_registry_value(r"Software\Microsoft\Office\\16.0\\Common\\General", "UserTemplates")
    ppt_personal = _read_registry_value(r"Software\Microsoft\Office\\16.0\\PowerPoint\\Options", "PersonalTemplates")
    ppt_user = _read_registry_value(r"Software\Microsoft\Office\\16.0\\Common\\General", "UserTemplates")
    excel_personal = _read_registry_value(r"Software\Microsoft\Office\\16.0\\Excel\\Options", "PersonalTemplates")
    excel_user = _read_registry_value(r"Software\Microsoft\Office\\16.0\\Common\\General", "UserTemplates")
    logger.info("[REG] Word PersonalTemplates: %s", word_personal or "[no valor]")
    logger.info("[REG] Word UserTemplates: %s", word_user or "[no valor]")
    logger.info("[REG] PowerPoint PersonalTemplates: %s", ppt_personal or "[no valor]")
    logger.info("[REG] PowerPoint UserTemplates: %s", ppt_user or "[no valor]")
    logger.info("[REG] Excel PersonalTemplates: %s", excel_personal or "[no valor]")
    logger.info("[REG] Excel UserTemplates: %s", excel_user or "[no valor]")


def update_mru_for_template(app_label: str, file_path: Path, design_mode: bool) -> None:
    if not is_windows() or winreg is None:
        return
    mru_paths = _find_mru_paths(app_label)
    if design_mode and DESIGN_LOG_MRU:
        LOGGER.info("[MRU] Actualizando MRU para %s en rutas: %s", app_label, mru_paths)
    for mru_path in mru_paths:
        try:
            _write_mru_entry(mru_path, file_path, design_mode)
        except OSError as exc:
            if design_mode and DESIGN_LOG_MRU:
                LOGGER.warning("[MRU] No se pudo escribir en %s (%s)", mru_path, exc)


def _find_mru_paths(app_label: str) -> list[str]:
    reg_name = _app_registry_name(app_label)
    if not reg_name:
        return []
    roots: list[str] = []
    versions = ("16.0", "15.0", "14.0", "12.0")
    for version in versions:
        base = fr"Software\Microsoft\Office\{version}\{reg_name}\Recent Templates"
        # Prefer LiveID/ADAL containers si existen
        if winreg:
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, base) as root:
                    sub_count = winreg.QueryInfoKey(root)[0]
                    for idx in range(sub_count):
                        sub = winreg.EnumKey(root, idx)
                        if sub.upper().startswith("ADAL_") or sub.upper().startswith("LIVEID_"):
                            roots.append(f"HKCU\\{base}\\{sub}\\File MRU")
            except OSError:
                pass
        roots.append(f"HKCU\\{base}\\File MRU")
    # Deduplicar manteniendo orden
    seen: set[str] = set()
    ordered: list[str] = []
    for path in roots:
        if path not in seen:
            seen.add(path)
            ordered.append(path)
    return ordered


def _app_registry_name(app_label: str) -> str:
    mapping = {"WORD": "Word", "POWERPOINT": "PowerPoint", "EXCEL": "Excel"}
    return mapping.get(app_label.upper(), "")


def _write_mru_entry(reg_path: str, file_path: Path, design_mode: bool) -> None:
    if winreg is None:
        return
    file_path = normalize_path(file_path)
    full_path = str(file_path)
    basename = file_path.stem
    hive, subkey = reg_path.split("\\", 1)
    hive_obj = winreg.HKEY_CURRENT_USER if hive.upper() == "HKCU" else None
    if hive_obj is None:
        return
    try:
        key = winreg.CreateKeyEx(hive_obj, subkey, 0, winreg.KEY_ALL_ACCESS)
    except OSError:
        return
    with key:
        # Leer entradas existentes
        existing_items: list[tuple[int, str, str]] = []
        index = 0
        try:
            while True:
                name, value, _ = winreg.EnumValue(key, index)
                if name.startswith("Item Metadata "):
                    index += 1
                    continue
                if name.startswith("Item "):
                    try:
                        num = int(name.split(" ", 1)[1])
                    except Exception:
                        num = 0
                    existing_items.append((num, name, value))
                index += 1
        except OSError:
            pass
        # Filtrar duplicados del mismo path
        filtered = []
        for _, name, value in existing_items:
            if isinstance(value, str) and full_path.lower() in value.lower():
                continue
            filtered.append((name, value))
        # Preparar nueva lista con el archivo al frente
        new_entries: list[str] = [full_path] + [val for _, val in filtered]
        # Limitar, p.ej., a 10 entradas
        new_entries = new_entries[:10]
        # Reescribir
        for idx, entry in enumerate(new_entries, start=1):
            item_name = f"Item {idx}"
            meta_name = f"Item Metadata {idx}"
            reg_value = f"[F00000000][T0000000000000000][O00000000]*{entry}"
            meta_value = f"<Metadata><AppSpecific><id>{entry}</id><nm>{basename}</nm><du>{entry}</du></AppSpecific></Metadata>"
            if design_mode and DESIGN_LOG_MRU:
                LOGGER.info("[MRU] Escribiendo %s='%s' en %s", item_name, reg_value, reg_path)
                LOGGER.info("[MRU] Escribiendo %s='%s' en %s", meta_name, meta_value, reg_path)
            winreg.SetValueEx(key, item_name, 0, winreg.REG_SZ, reg_value)
            winreg.SetValueEx(key, meta_name, 0, winreg.REG_SZ, meta_value)
        if design_mode and DESIGN_LOG_MRU:
            LOGGER.info("[MRU] %s actualizado con %s", reg_path, full_path)
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
