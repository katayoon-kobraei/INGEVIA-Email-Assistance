from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication

from desktop_app.language import get_language
from desktop_app.styles import APP_STYLESHEET


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("INGEVIA Email Assistant")
    app.setOrganizationName("INGEVIA")
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
