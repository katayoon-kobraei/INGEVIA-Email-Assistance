"""
prompt_editor.py

"Edit AI prompts" dialog, opened from the Configuration/Settings page.
Lists every file in data_service.PROMPTS_DIR (the same folder each
classification agent reads its prompt from -- see src/config.py's
PROMPTS_DIR and src/classification/*_agent.py) and lets the user read
and edit them directly, with an automatic timestamped backup on every
save (see data_service.write_prompt_file). Changes take effect on the
very next email the pipeline processes -- no restart needed, since
every agent just reads its .md file fresh at import time, and the
scheduled task starts a brand-new Python process each run.

Shared between both language windows (main_window.py / main_window_en.py)
rather than duplicated like those two files -- only the label text
differs by language, not the behavior, so parametrizing one dialog is
simpler than maintaining two copies of file I/O and dirty-state logic.

Where this goes: desktop_app/prompt_editor.py.
"""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
)

from desktop_app import data_service

_LABELS = {
    "es": {
        "title": "Editar prompts de IA",
        "subtitle": (
            "Estos archivos controlan exactamente cómo la IA clasifica y archiva cada "
            "correo. Los cambios se aplican de inmediato, en el próximo correo procesado -- "
            "no es necesario reiniciar la aplicación."
        ),
        "unsaved_title": "Cambios sin guardar",
        "unsaved_body": 'Tiene cambios sin guardar en "{name}". ¿Qué desea hacer?',
        "save_and_continue": "Guardar y continuar",
        "discard": "Descartar cambios",
        "cancel": "Cancelar",
        "save": "Guardar",
        "revert": "Deshacer cambios",
        "close": "Cerrar",
        "status_saved": "Guardado -- copia de seguridad creada en _backups.",
        "status_unsaved": "Cambios sin guardar",
        "status_error": "No se pudo guardar: {error}",
        "no_files": "No se encontraron archivos de prompts en:\n{path}",
        "select_prompt": "Seleccione un archivo de la izquierda para editarlo.",
        "viewer_note": (
            "⚠ Este ordenador está en modo VISOR. Editar aquí solo cambia la copia local de "
            "este ordenador -- no afecta a la clasificación real, que usa la copia del "
            "ordenador de procesamiento."
        ),
    },
    "en": {
        "title": "Edit AI prompts",
        "subtitle": (
            "These files control exactly how the AI classifies and files every email. "
            "Changes take effect immediately, on the next email processed -- no need to "
            "restart the application."
        ),
        "unsaved_title": "Unsaved changes",
        "unsaved_body": 'You have unsaved changes in "{name}". What would you like to do?',
        "save_and_continue": "Save and continue",
        "discard": "Discard changes",
        "cancel": "Cancel",
        "save": "Save",
        "revert": "Revert changes",
        "close": "Close",
        "status_saved": "Saved -- backup created in _backups.",
        "status_unsaved": "Unsaved changes",
        "status_error": "Could not save: {error}",
        "no_files": "No prompt files were found in:\n{path}",
        "select_prompt": "Select a file on the left to edit it.",
        "viewer_note": (
            "⚠ This computer is in VIEWER mode. Editing here only changes this computer's "
            "local copy -- it does not affect real classification, which uses the "
            "processing computer's copy."
        ),
    },
}


class PromptEditorDialog(QDialog):
    def __init__(self, parent=None, language: str = "es") -> None:
        super().__init__(parent)
        self.labels = _LABELS.get(language, _LABELS["es"])
        self.setWindowTitle(self.labels["title"])
        self.resize(980, 640)
        self._current_path: Path | None = None
        self._loaded_content: str = ""

        outer = QVBoxLayout(self)
        outer.setContentsMargins(20, 18, 20, 18)
        outer.setSpacing(10)

        title = QLabel(self.labels["title"])
        title.setObjectName("PageTitle")
        subtitle = QLabel(self.labels["subtitle"])
        subtitle.setObjectName("PageSubtitle")
        subtitle.setWordWrap(True)
        outer.addWidget(title)
        outer.addWidget(subtitle)

        if data_service.IS_VIEWER:
            # Prompt files are local to each install, not shared through
            # OUTPUT_ROOT/ARCHIVE_ROOT like the archived emails are -- so
            # editing them here has no effect on the processing computer.
            viewer_note = QLabel(self.labels["viewer_note"])
            viewer_note.setObjectName("MutedText")
            viewer_note.setWordWrap(True)
            outer.addWidget(viewer_note)

        body = QHBoxLayout()
        body.setSpacing(14)

        self.file_list = QListWidget()
        self.file_list.setFixedWidth(260)
        self.file_list.currentItemChanged.connect(self._on_selection_changed)
        body.addWidget(self.file_list)

        right = QVBoxLayout()
        self.editor = QPlainTextEdit()
        self.editor.setPlaceholderText(self.labels["select_prompt"])
        self.editor.setEnabled(False)
        editor_font = self.editor.font()
        editor_font.setFamily("Consolas")
        editor_font.setPointSize(10)
        self.editor.setFont(editor_font)
        self.editor.textChanged.connect(self._on_text_changed)
        right.addWidget(self.editor, 1)

        self.status_label = QLabel("")
        self.status_label.setObjectName("MutedText")
        right.addWidget(self.status_label)

        buttons = QHBoxLayout()
        self.revert_button = QPushButton(self.labels["revert"])
        self.revert_button.setObjectName("SecondaryButton")
        self.revert_button.clicked.connect(self._revert)
        self.revert_button.setEnabled(False)
        self.save_button = QPushButton(self.labels["save"])
        self.save_button.setObjectName("PrimaryButton")
        self.save_button.clicked.connect(self._save)
        self.save_button.setEnabled(False)
        close_button = QPushButton(self.labels["close"])
        close_button.setObjectName("SecondaryButton")
        close_button.clicked.connect(self.close)
        buttons.addWidget(self.revert_button)
        buttons.addStretch()
        buttons.addWidget(close_button)
        buttons.addWidget(self.save_button)
        right.addLayout(buttons)

        body.addLayout(right, 1)
        outer.addLayout(body, 1)

        self._populate_file_list()

    def _populate_file_list(self) -> None:
        files = data_service.list_prompt_files()
        if not files:
            self.status_label.setText(self.labels["no_files"].format(path=data_service.PROMPTS_DIR))
            return
        for path in files:
            item = QListWidgetItem(path.name)
            item.setData(Qt.UserRole, str(path))
            self.file_list.addItem(item)
        self.file_list.setCurrentRow(0)

    def _dirty(self) -> bool:
        return self._current_path is not None and self.editor.toPlainText() != self._loaded_content

    def _on_text_changed(self) -> None:
        dirty = self._dirty()
        self.save_button.setEnabled(dirty)
        self.revert_button.setEnabled(dirty)
        if dirty:
            self.status_label.setText(self.labels["status_unsaved"])
        elif self._current_path is not None:
            self.status_label.setText("")

    def _on_selection_changed(self, current: QListWidgetItem | None, previous: QListWidgetItem | None) -> None:
        if previous is not None and self._dirty():
            choice = self._prompt_unsaved(previous.text())
            if choice == "cancel":
                self.file_list.blockSignals(True)
                self.file_list.setCurrentItem(previous)
                self.file_list.blockSignals(False)
                return
            if choice == "save":
                self._save(silent=True)

        if current is None:
            return
        path = Path(current.data(Qt.UserRole))
        try:
            content = data_service.read_prompt_file(path)
        except OSError as exc:
            QMessageBox.warning(self, self.labels["title"], str(exc))
            return
        self._current_path = path
        self._loaded_content = content
        self.editor.blockSignals(True)
        self.editor.setPlainText(content)
        self.editor.blockSignals(False)
        self.editor.setEnabled(True)
        self.save_button.setEnabled(False)
        self.revert_button.setEnabled(False)
        self.status_label.setText("")

    def _prompt_unsaved(self, name: str) -> str:
        box = QMessageBox(self)
        box.setWindowTitle(self.labels["unsaved_title"])
        box.setText(self.labels["unsaved_body"].format(name=name))
        save_btn = box.addButton(self.labels["save_and_continue"], QMessageBox.AcceptRole)
        discard_btn = box.addButton(self.labels["discard"], QMessageBox.DestructiveRole)
        box.addButton(self.labels["cancel"], QMessageBox.RejectRole)
        box.exec()
        clicked = box.clickedButton()
        if clicked is save_btn:
            return "save"
        if clicked is discard_btn:
            return "discard"
        return "cancel"

    def _revert(self) -> None:
        self.editor.blockSignals(True)
        self.editor.setPlainText(self._loaded_content)
        self.editor.blockSignals(False)
        self.save_button.setEnabled(False)
        self.revert_button.setEnabled(False)
        self.status_label.setText("")

    def _save(self, silent: bool = False) -> None:
        if self._current_path is None:
            return
        content = self.editor.toPlainText()
        ok, error = data_service.write_prompt_file(self._current_path, content)
        if ok:
            self._loaded_content = content
            self.save_button.setEnabled(False)
            self.revert_button.setEnabled(False)
            self.status_label.setText(self.labels["status_saved"])
        elif not silent:
            QMessageBox.warning(self, self.labels["title"], self.labels["status_error"].format(error=error))

    def closeEvent(self, event) -> None:
        if self._dirty():
            choice = self._prompt_unsaved(self._current_path.name)
            if choice == "cancel":
                event.ignore()
                return
            if choice == "save":
                self._save(silent=True)
        event.accept()