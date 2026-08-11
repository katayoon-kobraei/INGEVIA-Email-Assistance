from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from desktop_app.language import get_language
from desktop_app.styles import APP_STYLESHEET


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("INGEVIA Email Assistant")
    app.setOrganizationName("INGEVIA")
    icon_path = Path(__file__).resolve().parent.parent / "assets" / "app.ico"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    app.setStyleSheet(APP_STYLESHEET)

    if get_language() == "es":
        from desktop_app.main_window import MainWindow
    else:
        from desktop_app.main_window_en import MainWindow

    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
