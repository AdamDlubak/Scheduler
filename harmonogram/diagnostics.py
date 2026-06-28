#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Moduł diagnostyki — zbieranie logów, błędów, stacktrace i danych wejściowych.
"""

from __future__ import annotations

import logging
import shutil
import traceback
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Optional

# Konfiguracja
DIAGNOSTICS_DIR = Path(__file__).parent.parent / "diagnostics"
DIAGNOSTICS_DIR.mkdir(exist_ok=True)

# Logger
logger = logging.getLogger("harmonogram_diagnostics")
logger.setLevel(logging.DEBUG)

# Unikaj duplikowania handlery
if not logger.handlers:
    handler = logging.StreamHandler()
    handler.setLevel(logging.DEBUG)
    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)


def create_diagnostic_report(
    error: Exception,
    input_file_path: Optional[str] = None,
    context: Optional[dict] = None,
) -> Path:
    """
    Tworzy raport diagnostyczny — folder z logami, stacktrace, i kopią pliku wejściowego.
    
    Args:
        error: Obiekt wyjątku
        input_file_path: Ścieżka do pliku wejściowego (Excel)
        context: Dodatkowy kontekst (np. nazwa użytkownika, parametry)
    
    Returns:
        Ścieżka do folderu raportu diagnostycznego
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_dir = DIAGNOSTICS_DIR / f"error_{timestamp}"
    report_dir.mkdir(exist_ok=True)

    # 1. Log stacktrace
    log_file = report_dir / "error.log"
    with open(log_file, "w", encoding="utf-8") as f:
        f.write("=== RAPORT DIAGNOSTYCZNY ===\n")
        f.write(f"Timestamp: {datetime.now().isoformat()}\n")
        f.write(f"Error Type: {type(error).__name__}\n")
        f.write(f"Error Message: {str(error)}\n\n")
        f.write("=== STACKTRACE ===\n")
        f.write(traceback.format_exc())
        f.write("\n\n=== KONTEKST ===\n")
        if context:
            for key, value in context.items():
                f.write(f"{key}: {value}\n")
        else:
            f.write("(brak dodatkowego kontekstu)\n")

    # 2. Kopiuj plik wejściowy
    if input_file_path and Path(input_file_path).exists():
        input_copy = report_dir / "input_file.xlsx"
        shutil.copy(input_file_path, input_copy)
        logger.debug(f"Skopiowany plik wejściowy: {input_copy}")

    # 3. Stwórz ZIP
    zip_path = report_dir.parent / f"diagnostic_{timestamp}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for file in report_dir.glob("*"):
            zf.write(file, arcname=file.name)

    logger.info(f"Raport diagnostyczny: {zip_path}")
    return report_dir


def log_info(message: str) -> None:
    """Log informacyjny"""
    logger.info(message)


def log_warning(message: str) -> None:
    """Log ostrzeżenia"""
    logger.warning(message)


def log_error(message: str) -> None:
    """Log błędu"""
    logger.error(message)


def get_diagnostics_folder() -> Path:
    """Zwraca ścieżkę do folderu diagnostyki"""
    return DIAGNOSTICS_DIR

