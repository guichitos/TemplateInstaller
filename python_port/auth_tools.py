"""Herramientas de validación de autores y utilidades relacionadas."""
from __future__ import annotations

import logging
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional
import xml.etree.ElementTree as ET

from . import constants
from .utils import format_authors_list, iter_template_files, normalize_path

LOGGER = logging.getLogger(__name__)


@dataclass
class AuthorCheckResult:
    allowed: bool
    message: str
    authors: List[str]
    error: bool = False

    def as_cli_output(self) -> str:
        return "TRUE" if self.allowed and not self.error else "FALSE"


def _extract_author(template_path: Path) -> tuple[Optional[str], Optional[str]]:
    if not template_path.exists():
        return None, f"[ERROR] No se encontró la ruta: \"{template_path}\""

    try:
        with zipfile.ZipFile(template_path) as zipped:
            try:
                with zipped.open("docProps/core.xml") as core_file:
                    tree = ET.fromstring(core_file.read())
            except KeyError:
                return None, "[WARN] No se pudo obtener el autor (core.xml ausente)."
    except Exception as exc:  # noqa: BLE001 - queremos presentar el mensaje tal cual
        return None, f"[ERROR] {exc}"

    creator = None
    for candidate in ("{http://purl.org/dc/elements/1.1/}creator", "creator"):
        node = tree.find(candidate)
        if node is not None and node.text:
            creator = node.text.strip()
            break

    if not creator:
        return None, "[WARN] Archivo sin autor definido."

    return creator, None


def _normalize_allowed_authors(authors: Iterable[str]) -> list[str]:
    normalized = []
    for author in authors:
        cleaned = author.strip()
        if cleaned:
            normalized.append(cleaned)
    return normalized


def check_template_author(
    target: Path,
    allowed_authors: Iterable[str] | None = None,
    validation_enabled: bool = True,
    design_mode: bool = False,
) -> AuthorCheckResult:
    """Valida los autores de un archivo o carpeta de plantillas."""

    allowed = _normalize_allowed_authors(
        allowed_authors or constants.DEFAULT_ALLOWED_TEMPLATE_AUTHORS
    )
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
        return AuthorCheckResult(
            allowed=True,
            message=message,
            authors=authors_found,
        )

    if not validation_enabled:
        return AuthorCheckResult(
            allowed=True,
            message="[INFO] Validación de autores deshabilitada; se permite el archivo.",
            authors=[],
        )

    author, error = _extract_author(target)
    if error:
        return AuthorCheckResult(
            allowed=False,
            message=error,
            authors=[],
            error=True,
        )

    if not author:
        return AuthorCheckResult(
            allowed=False,
            message=f"[WARN] El archivo \"{target}\" no tiene autor asignado.",
            authors=[],
        )

    is_allowed = any(author.strip().lower() == a.lower() for a in allowed)
    message = (
        "[OK] Autor aprobado."
        if is_allowed
        else f"[BLOCKED] Autor no permitido para \"{target}\"."
    )

    if design_mode:
        LOGGER.debug("Autor encontrado: %s", author)
        LOGGER.debug("Lista de autores permitidos: %s", format_authors_list(allowed))

    return AuthorCheckResult(allowed=is_allowed, message=message, authors=[author])
