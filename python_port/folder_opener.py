"""Lógica para abrir carpetas de plantillas de forma controlada."""
from __future__ import annotations

import logging
import os
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from .utils import is_windows, log_platform_warning, normalize_path

LOGGER = logging.getLogger(__name__)


@dataclass
class FolderOpenRequest:
    enabled: bool
    path: Path | None
    select: Path | None = None

    def normalized(self) -> "FolderOpenRequest":
        return FolderOpenRequest(
            enabled=self.enabled,
            path=normalize_path(self.path) if self.path else None,
            select=normalize_path(self.select) if self.select else None,
        )



def open_template_folders(design_mode: bool, requests: Iterable[FolderOpenRequest]) -> None:
    normalized_requests = [req.normalized() for req in requests]
    for request in normalized_requests:
        _open_if_enabled(request, design_mode)


def _open_if_enabled(request: FolderOpenRequest, design_mode: bool) -> None:
    if not request.enabled:
        if design_mode:
            LOGGER.debug("[DEBUG] Apertura omitida para %s", request)
        return

    if not request.path or not request.path.exists():
        if design_mode:
            LOGGER.debug("[DEBUG] Carpeta no disponible: %s", request.path)
        return

    if not is_windows():
        log_platform_warning(f"Abrir carpeta {request.path}")
        return

    try:
        if request.select and request.select.exists():
            if design_mode:
                LOGGER.info(
                    "[ACTION] Abriendo %s y seleccionando %s", request.path, request.select
                )
            subprocess.Popen(
                ["explorer", f"/select,{request.select}"],
                shell=False,
            )
        else:
            if design_mode:
                LOGGER.info("[ACTION] Abriendo %s", request.path)
            os.startfile(request.path)  # type: ignore[arg-type]
    except OSError as exc:  # pragma: no cover - depende del SO
        LOGGER.warning("[WARN] No se pudo abrir %s (%s)", request.path, exc)
