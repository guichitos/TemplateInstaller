"""Port en Python del instalador principal de plantillas de Office.

El flujo replica la orquestación de ``Script/1-2. MainInstaller.bat`` y
sus dependencias, pero estructurado en módulos reutilizables.
"""
from __future__ import annotations

import argparse
import logging
import os
import shutil
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable, List, Optional

from . import constants
from . import auth_tools, folder_opener, mru_repair, mru_resolver, office_paths
from .folder_opener import FolderOpenRequest
from .utils import (
    ensure_directory,
    humanize_bool,
    is_windows,
    iter_template_files,
    normalize_path,
    path_in_appdata,
)

LOGGER = logging.getLogger(__name__)


@dataclass
class InstallerConfig:
    design_mode: bool
    allowed_authors: List[str]
    author_validation_enabled: bool
    document_theme_delay_seconds: int
    script_directory: Path
    base_hint: Path
    base_directory: Path


@dataclass
class InstallationContext:
    force_open_word: bool = False
    force_open_ppt: bool = False
    force_open_excel: bool = False
    should_open_document_theme_folder: bool = False
    document_theme_selection_path: Optional[Path] = None
    should_open_custom_template_folder: bool = False
    custom_template_folder_path_to_open: Optional[Path] = None
    custom_template_selection_path: Optional[Path] = None
    custom_template_additional_selection_path: Optional[Path] = None
    should_open_roaming_template_folder: bool = False
    roaming_template_selection_path: Optional[Path] = None
    should_open_excel_startup_folder: bool = False
    excel_startup_selection_path: Optional[Path] = None
    open_requests: List[FolderOpenRequest] = field(default_factory=list)
    totals: dict = field(default_factory=lambda: {"files": 0, "errors": 0, "blocked": 0})


@dataclass
class InstallOutcome:
    app: str
    filename: str
    installed: bool
    destination: Optional[Path] = None
    selected: Optional[Path] = None


# ------------------------------------------------------------
# CLI
# ------------------------------------------------------------


def parse_args(argv: Optional[Iterable[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Instalador universal de plantillas de Office (port Python)",
    )
    parser.add_argument(
        "base_hint",
        nargs="?",
        help="Pista o ruta base con los archivos de plantilla (equivale a %~1).",
    )
    parser.add_argument(
        "--design-mode",
        action="store_true",
        help="Habilita modo diseño (salida detallada).",
    )
    parser.add_argument(
        "--allowed-authors",
        help="Lista separada por ';' de autores permitidos.",
    )
    parser.add_argument(
        "--disable-author-validation",
        action="store_true",
        help="Desactiva la validación de autores.",
    )
    parser.add_argument(
        "--check-author",
        metavar="PATH",
        help="Solo valida el autor de un archivo/carpeta y termina.",
    )
    parser.add_argument(
        "--document-theme-delay",
        type=int,
        help="Segundos de espera antes de abrir aplicaciones tras manejar temas de documento.",
    )
    return parser.parse_args(argv)


# ------------------------------------------------------------
# Lógica principal
# ------------------------------------------------------------


def main(argv: Optional[Iterable[str]] = None) -> int:
    args = parse_args(argv)

    design_mode = args.design_mode or constants.DEFAULT_DESIGN_MODE
    log_level = logging.DEBUG if design_mode else logging.INFO
    logging.basicConfig(level=log_level, format="%(message)s")

    script_directory = Path(__file__).resolve().parent
    base_hint = normalize_path(
        args.base_hint
        or os.environ.get("PIN_LAUNCHER_DIR")
        or script_directory
    )

    if base_hint == script_directory and path_in_appdata(script_directory):
        print(
            "[ERROR] No se recibió la ruta de las plantillas. Ejecute el instalador "
            'desde "1. Pin templates..." para que se le pase la carpeta correcta.'
        )
        return 1

    base_directory = resolve_base_directory(base_hint)
    allowed_authors = _resolve_allowed_authors(args.allowed_authors)
    author_validation_enabled = not args.disable_author_validation
    if os.environ.get("AuthorValidationEnabled", "TRUE").lower() == "false":
        author_validation_enabled = False

    document_theme_delay = (
        args.document_theme_delay
        if args.document_theme_delay is not None
        else constants.DEFAULT_DOCUMENT_THEME_DELAY_SECONDS
    )

    config = InstallerConfig(
        design_mode=design_mode,
        allowed_authors=allowed_authors,
        author_validation_enabled=author_validation_enabled,
        document_theme_delay_seconds=document_theme_delay,
        script_directory=script_directory,
        base_hint=base_hint,
        base_directory=base_directory,
    )

    if args.check_author:
        result = auth_tools.check_template_author(
            Path(args.check_author),
            allowed_authors=allowed_authors,
            validation_enabled=author_validation_enabled,
            design_mode=design_mode,
        )
        print(result.as_cli_output())
        if design_mode:
            LOGGER.info(result.message)
        return 0 if result.allowed else 1

    _print_intro(config)
    check_environment(design_mode)
    close_office_apps(design_mode)
    mru_repair.repair_template_mru(design_mode)

    context = InstallationContext()

    base_results = install_base_templates(config, context)
    detected_paths = office_paths.detect_office_paths(design_mode)
    copy_custom_templates(config, detected_paths, context)

    handle_folder_open_requests(detected_paths, context, design_mode)

    if context.should_open_document_theme_folder:
        if design_mode:
            LOGGER.info(
                "[INFO] Esperando %s segundos antes de abrir aplicaciones...",
                config.document_theme_delay_seconds,
            )
        time.sleep(max(config.document_theme_delay_seconds, 0))

    launch_office_apps(context, design_mode)

    if design_mode:
        LOGGER.info(
            "[FINAL] Instalación completada. Archivos copiados=%s, errores=%s, bloqueados=%s.",
            context.totals["files"],
            context.totals["errors"],
            context.totals["blocked"],
        )
    else:
        print("Ready")
    return 0


# ------------------------------------------------------------
# Helpers de alto nivel
# ------------------------------------------------------------


def _print_intro(config: InstallerConfig) -> None:
    if config.design_mode:
        LOGGER.info("[DEBUG] Design mode habilitado=%s", humanize_bool(config.design_mode))
        LOGGER.info("[INFO] Script ejecutado desde: %s", config.script_directory)
        LOGGER.info("[INFO] Base directory: %s", config.base_directory)
    else:
        print(
            "Installing custom templates and applying them as the new Microsoft Office defaults..."
        )


def resolve_base_directory(base_hint: Path) -> Path:
    candidates = [base_hint]
    for child in ("payload", "templates", "extracted"):
        candidates.append(base_hint / child)

    for candidate in candidates:
        if any(iter_template_files(candidate)):
            return normalize_path(candidate)
    return normalize_path(base_hint)


def check_environment(design_mode: bool) -> None:
    if design_mode:
        LOGGER.info("[DEBUG] Environment check starting...")
        LOGGER.info("[DEBUG] Sistema operativo: %s", os.name)
        LOGGER.info("[DEBUG] winreg disponible: %s", bool(mru_resolver.winreg))
        LOGGER.info("[DEBUG] Environment check completed.")


def close_office_apps(design_mode: bool) -> None:
    if design_mode:
        LOGGER.info("[DEBUG] Cerrando aplicaciones de Office si están abiertas...")
    if not is_windows():
        LOGGER.debug("[DEBUG] Cierre de apps omitido: no es Windows.")
        return

    for exe in ("WINWORD.EXE", "POWERPNT.EXE", "EXCEL.EXE"):
        try:
            os.system(f"taskkill /IM {exe} /F >nul 2>&1")
        except OSError:
            LOGGER.debug("[DEBUG] No se pudo cerrar %s", exe)


# ------------------------------------------------------------
# Instalación de plantillas
# ------------------------------------------------------------


def install_base_templates(config: InstallerConfig, context: InstallationContext) -> List[InstallOutcome]:
    outcomes: list[InstallOutcome] = []
    base_targets = [
        ("WORD", "Normal.dotx", constants.DEFAULT_ROAMING_TEMPLATE_FOLDER),
        ("WORD", "Normal.dotm", constants.DEFAULT_ROAMING_TEMPLATE_FOLDER),
        ("WORD", "NormalEmail.dotx", constants.DEFAULT_ROAMING_TEMPLATE_FOLDER),
        ("WORD", "NormalEmail.dotm", constants.DEFAULT_ROAMING_TEMPLATE_FOLDER),
        ("POWERPOINT", "Blank.potx", constants.DEFAULT_ROAMING_TEMPLATE_FOLDER),
        ("POWERPOINT", "Blank.potm", constants.DEFAULT_ROAMING_TEMPLATE_FOLDER),
        ("EXCEL", "Book.xltx", constants.DEFAULT_EXCEL_STARTUP_FOLDER),
        ("EXCEL", "Book.xltm", constants.DEFAULT_EXCEL_STARTUP_FOLDER),
        ("EXCEL", "Sheet.xltx", constants.DEFAULT_EXCEL_STARTUP_FOLDER),
        ("EXCEL", "Sheet.xltm", constants.DEFAULT_EXCEL_STARTUP_FOLDER),
    ]

    for app, filename, destination in base_targets:
        outcome = _install_single_template(
            app,
            filename,
            config.base_directory,
            destination,
            config,
            context,
        )
        outcomes.append(outcome)
    return outcomes


def _install_single_template(
    app: str,
    filename: str,
    base_directory: Path,
    destination_directory: Path,
    config: InstallerConfig,
    context: InstallationContext,
) -> InstallOutcome:
    source = normalize_path(base_directory / filename)
    destination_directory = ensure_directory(normalize_path(destination_directory))
    destination = destination_directory / filename

    if not source.exists():
        if config.design_mode:
            LOGGER.warning("[WARNING] Archivo fuente no encontrado: %s", source)
        context.totals["errors"] += 1
        return InstallOutcome(app=app, filename=filename, installed=False)

    author_check = auth_tools.check_template_author(
        source,
        allowed_authors=config.allowed_authors,
        validation_enabled=config.author_validation_enabled,
        design_mode=config.design_mode,
    )
    if not author_check.allowed:
        if config.design_mode:
            LOGGER.warning(author_check.message)
        context.totals["blocked"] += 1
        return InstallOutcome(app=app, filename=filename, installed=False)

    backup_existing_template(destination, config.design_mode)
    shutil.copy2(source, destination)

    if destination.exists():
        context.totals["files"] += 1
        if config.design_mode:
            LOGGER.info("[OK] Instalado %s en %s", filename, destination)

        if app == "WORD":
            context.force_open_word = True
            _capture_selection_for_roaming(destination, context)
        elif app == "POWERPOINT":
            context.force_open_ppt = True
            _capture_selection_for_roaming(destination, context)
        elif app == "EXCEL":
            context.force_open_excel = True
            _capture_selection_for_excel_startup(destination, context)

        return InstallOutcome(app=app, filename=filename, installed=True, destination=destination, selected=destination)

    context.totals["errors"] += 1
    return InstallOutcome(app=app, filename=filename, installed=False)


def backup_existing_template(target_file: Path, design_mode: bool) -> None:
    if not target_file.exists():
        return

    backup_dir = target_file.parent / "Backup"
    try:
        ensure_directory(backup_dir)
    except OSError:
        LOGGER.warning("[WARN] No se pudo crear la carpeta de backup: %s", backup_dir)
        return

    timestamp = datetime.now().strftime("%Y.%m.%d.%H%M")
    backup_path = backup_dir / f"{timestamp}_{target_file.name}"
    try:
        shutil.copy2(target_file, backup_path)
        if design_mode:
            LOGGER.info("[BACKUP] Copia creada en %s", backup_path)
    except OSError as exc:
        LOGGER.warning("[WARN] No se pudo crear backup de %s (%s)", target_file, exc)


def _capture_selection_for_roaming(destination: Path, context: InstallationContext) -> None:
    roaming_compare = normalize_path(constants.DEFAULT_ROAMING_TEMPLATE_FOLDER)
    dest_parent = normalize_path(destination.parent)
    if dest_parent == roaming_compare:
        context.should_open_roaming_template_folder = True
        if not context.roaming_template_selection_path:
            context.roaming_template_selection_path = destination
    else:
        context.open_requests.append(
            FolderOpenRequest(enabled=True, path=dest_parent, select=destination)
        )


def _capture_selection_for_excel_startup(destination: Path, context: InstallationContext) -> None:
    excel_startup_compare = normalize_path(constants.DEFAULT_EXCEL_STARTUP_FOLDER)
    dest_parent = normalize_path(destination.parent)
    if dest_parent == excel_startup_compare:
        context.should_open_excel_startup_folder = True
        if not context.excel_startup_selection_path:
            context.excel_startup_selection_path = destination
    else:
        context.open_requests.append(
            FolderOpenRequest(enabled=True, path=dest_parent, select=destination)
        )


# ------------------------------------------------------------
# Copia de plantillas personalizadas (CopyAll)
# ------------------------------------------------------------


def copy_custom_templates(
    config: InstallerConfig,
    detected_paths: office_paths.OfficePaths,
    context: InstallationContext,
) -> None:
    base_dir = config.base_directory
    word_path = normalize_path(detected_paths.word)
    ppt_path = normalize_path(detected_paths.powerpoint)
    excel_path = normalize_path(detected_paths.excel)
    document_theme_path = normalize_path(detected_paths.document_theme)

    custom_template_path = normalize_path(constants.DEFAULT_CUSTOM_OFFICE_TEMPLATE_PATH)
    custom_additional_path = normalize_path(constants.DEFAULT_CUSTOM_OFFICE_ADDITIONAL_TEMPLATE_PATH)
    roaming_template_folder = normalize_path(constants.DEFAULT_ROAMING_TEMPLATE_FOLDER)
    excel_startup_folder = normalize_path(constants.DEFAULT_EXCEL_STARTUP_FOLDER)

    for file in iter_template_files(base_dir):
        filename = file.name
        extension = file.suffix.lower()
        if filename in constants.BASE_TEMPLATE_NAMES:
            if config.design_mode:
                LOGGER.info("[INFO] Archivo base omitido en CopyAll: %s", filename)
            continue

        destination = _destination_for_extension(extension, word_path, ppt_path, excel_path, document_theme_path)
        if not destination:
            if config.design_mode:
                LOGGER.warning("[WARNING] No hay destino para %s", filename)
            continue

        result = auth_tools.check_template_author(
            file,
            allowed_authors=config.allowed_authors,
            validation_enabled=config.author_validation_enabled,
            design_mode=config.design_mode,
        )
        if not result.allowed:
            context.totals["blocked"] += 1
            if config.design_mode:
                LOGGER.warning(result.message)
            continue

        try:
            ensure_directory(destination)
            shutil.copy2(file, destination / filename)
            context.totals["files"] += 1
            if config.design_mode:
                LOGGER.info("[OK] Copiado: %s -> %s", filename, destination)
        except OSError as exc:
            context.totals["errors"] += 1
            LOGGER.warning("[ERROR] Falló la copia de %s (%s)", filename, exc)
            continue

        _update_flags_after_copy(
            destination,
            filename,
            extension,
            context,
            custom_template_path,
            custom_additional_path,
            roaming_template_folder,
            excel_startup_folder,
            document_theme_path,
        )

        _register_mru(extension, destination / filename, config.design_mode)



def _destination_for_extension(
    extension: str,
    word_path: Path,
    ppt_path: Path,
    excel_path: Path,
    document_theme_path: Path,
) -> Optional[Path]:
    if extension in {".dotx", ".dotm"}:
        return word_path
    if extension in {".potx", ".potm"}:
        return ppt_path
    if extension in {".xltx", ".xltm"}:
        return excel_path
    if extension == ".thmx":
        return document_theme_path
    return None


def _update_flags_after_copy(
    destination: Path,
    filename: str,
    extension: str,
    context: InstallationContext,
    custom_template_path: Path,
    custom_additional_path: Path,
    roaming_template_folder: Path,
    excel_startup_folder: Path,
    document_theme_path: Path,
) -> None:
    destination = normalize_path(destination)
    installed_path = destination / filename

    if extension in {".dotx", ".dotm"}:
        context.force_open_word = True
        if destination == custom_template_path:
            context.should_open_custom_template_folder = True
            context.custom_template_folder_path_to_open = destination
            if not context.custom_template_selection_path:
                context.custom_template_selection_path = installed_path
        elif destination == custom_additional_path:
            context.should_open_custom_template_folder = True
            context.custom_template_folder_path_to_open = destination
            if not context.custom_template_additional_selection_path:
                context.custom_template_additional_selection_path = installed_path
        elif destination == roaming_template_folder:
            context.should_open_roaming_template_folder = True
            if not context.roaming_template_selection_path:
                context.roaming_template_selection_path = installed_path
        else:
            context.open_requests.append(
                FolderOpenRequest(enabled=True, path=destination, select=installed_path)
            )

    elif extension in {".potx", ".potm"}:
        context.force_open_ppt = True
        if destination in {custom_template_path, custom_additional_path, roaming_template_folder}:
            context.should_open_custom_template_folder = True
            context.custom_template_folder_path_to_open = destination
            if not context.custom_template_selection_path:
                context.custom_template_selection_path = installed_path
        else:
            context.open_requests.append(
                FolderOpenRequest(enabled=True, path=destination, select=installed_path)
            )

    elif extension in {".xltx", ".xltm"}:
        context.force_open_excel = True
        if destination == custom_template_path:
            context.should_open_custom_template_folder = True
            context.custom_template_folder_path_to_open = destination
            if not context.custom_template_selection_path:
                context.custom_template_selection_path = installed_path
        elif destination == custom_additional_path:
            context.should_open_custom_template_folder = True
            context.custom_template_folder_path_to_open = destination
            if not context.custom_template_additional_selection_path:
                context.custom_template_additional_selection_path = installed_path
        elif destination == roaming_template_folder:
            context.should_open_roaming_template_folder = True
            if not context.roaming_template_selection_path:
                context.roaming_template_selection_path = installed_path
        elif destination == excel_startup_folder:
            context.should_open_excel_startup_folder = True
            if not context.excel_startup_selection_path:
                context.excel_startup_selection_path = installed_path
        else:
            context.open_requests.append(
                FolderOpenRequest(enabled=True, path=destination, select=installed_path)
            )

    elif extension == ".thmx":
        context.should_open_document_theme_folder = True
        if not context.document_theme_selection_path:
            context.document_theme_selection_path = installed_path



def _register_mru(extension: str, target: Path, design_mode: bool) -> None:
    if extension in {".dotx", ".dotm"}:
        mru_resolver.record_recent_template("WORD", target, design_mode)
    elif extension in {".potx", ".potm"}:
        mru_resolver.record_recent_template("POWERPOINT", target, design_mode)
    elif extension in {".xltx", ".xltm"}:
        mru_resolver.record_recent_template("EXCEL", target, design_mode)


# ------------------------------------------------------------
# Apertura de carpetas y aplicaciones
# ------------------------------------------------------------


def handle_folder_open_requests(
    detected_paths: office_paths.OfficePaths,
    context: InstallationContext,
    design_mode: bool,
) -> None:
    requests: list[FolderOpenRequest] = list(context.open_requests)

    if context.should_open_document_theme_folder:
        requests.append(
            FolderOpenRequest(
                enabled=True,
                path=detected_paths.document_theme,
                select=context.document_theme_selection_path,
            )
        )

    if context.should_open_custom_template_folder and context.custom_template_folder_path_to_open:
        requests.append(
            FolderOpenRequest(
                enabled=True,
                path=context.custom_template_folder_path_to_open,
                select=(context.custom_template_selection_path or context.custom_template_additional_selection_path),
            )
        )

    if context.should_open_roaming_template_folder:
        requests.append(
            FolderOpenRequest(
                enabled=True,
                path=constants.DEFAULT_ROAMING_TEMPLATE_FOLDER,
                select=context.roaming_template_selection_path,
            )
        )

    if context.should_open_excel_startup_folder:
        requests.append(
            FolderOpenRequest(
                enabled=True,
                path=constants.DEFAULT_EXCEL_STARTUP_FOLDER,
                select=context.excel_startup_selection_path,
            )
        )

    folder_opener.open_template_folders(design_mode, requests)


def launch_office_apps(context: InstallationContext, design_mode: bool) -> None:
    if not any((context.force_open_word, context.force_open_ppt, context.force_open_excel)):
        if design_mode:
            LOGGER.info("[INFO] No es necesario abrir aplicaciones de Office.")
        return

    targets = []
    if context.force_open_word:
        targets.append(("winword.exe", "Microsoft Word"))
    if context.force_open_ppt:
        targets.append(("powerpnt.exe", "Microsoft PowerPoint"))
    if context.force_open_excel:
        targets.append(("excel.exe", "Microsoft Excel"))

    if not is_windows():
        LOGGER.info("[WARN] Apertura de aplicaciones omitida: no es Windows.")
        return

    for executable, friendly in targets:
        try:
            if design_mode:
                LOGGER.info("[ACTION] Lanzando %s", friendly)
            os.startfile(executable)  # type: ignore[arg-type]
        except OSError as exc:
            LOGGER.warning("[WARN] No se pudo iniciar %s (%s)", friendly, exc)


# ------------------------------------------------------------
# Utilidades auxiliares
# ------------------------------------------------------------


def _resolve_allowed_authors(cli_value: Optional[str]) -> List[str]:
    env_value = os.environ.get("AllowedTemplateAuthors")
    raw = cli_value or env_value
    if not raw:
        return constants.DEFAULT_ALLOWED_TEMPLATE_AUTHORS
    return [author.strip() for author in raw.split(";") if author.strip()]


if __name__ == "__main__":
    sys.exit(main())
