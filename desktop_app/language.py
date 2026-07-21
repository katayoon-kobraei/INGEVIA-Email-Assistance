from __future__ import annotations

import sys

from PySide6.QtCore import QProcess, QSettings
from PySide6.QtWidgets import QApplication

ORGANIZATION = "INGEVIA"
APPLICATION = "Email Assistant"
DEFAULT_LANGUAGE = "en"


def get_language() -> str:
    """Return the saved UI language code: ``en`` or ``es``."""
    value = QSettings(ORGANIZATION, APPLICATION).value(
        "language", DEFAULT_LANGUAGE, type=str
    )
    return value if value in {"en", "es"} else DEFAULT_LANGUAGE


def set_language(code: str) -> None:
    """Persist the selected UI language."""
    if code not in {"en", "es"}:
        raise ValueError(f"Unsupported language: {code}")
    settings = QSettings(ORGANIZATION, APPLICATION)
    settings.setValue("language", code)
    settings.sync()


def restart_application() -> None:
    """Restart the current desktop application with the saved language."""
    QProcess.startDetached(sys.executable, sys.argv)
    app = QApplication.instance()
    if app is not None:
        app.quit()
