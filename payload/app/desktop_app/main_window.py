from __future__ import annotations

import io
import os
import re
import subprocess
import sys
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from typing import Any

from PySide6.QtCore import (
    QEasingCurve,
    QParallelAnimationGroup,
    QPoint,
    QPropertyAnimation,
    Qt,
    QThread,
    QTimer,
    Signal,
)
from PySide6.QtGui import QColor, QDesktopServices, QFont, QIcon
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QButtonGroup,
    QComboBox,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QSpinBox,
    QStackedWidget,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from desktop_app import data_service
from desktop_app.language import get_language, restart_application, set_language


DISPLAY_ROW_LIMIT = 500


class PipelineWorker(QThread):
    finished_with_result = Signal(bool, str)

    def run(self) -> None:
        app_root = Path(__file__).resolve().parent.parent
        launcher = app_root / "launch_pipeline.pyw"
        executable = Path(sys.executable)
        if executable.name.casefold() == "pythonw.exe":
            console_python = executable.with_name("python.exe")
            if console_python.is_file():
                executable = console_python

        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0
        try:
            completed = subprocess.run(
                [str(executable), str(launcher), "--interactive"],
                cwd=str(app_root),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=creationflags,
                check=False,
            )
            output = "\n".join(
                part.strip()
                for part in (completed.stdout, completed.stderr)
                if part and part.strip()
            ).strip()
            if not output:
                output = "Proceso completado correctamente." if completed.returncode == 0 else f"Processing failed with exit code {completed.returncode}."
            self.finished_with_result.emit(completed.returncode == 0, output)
        except Exception as exc:
            self.finished_with_result.emit(False, str(exc))


class OutlookFlagWorker(QThread):
    loaded = Signal(object)

    def run(self) -> None:
        self.loaded.emit(data_service.load_outlook_flagged())


class LocalDataWorker(QThread):
    loaded = Signal(object)
    failed = Signal(str)

    def run(self) -> None:
        try:
            self.loaded.emit(data_service.load_ui_snapshot())
        except Exception as exc:
            self.failed.emit(str(exc))


class SchedulerStatusWorker(QThread):
    loaded = Signal(object)

    def run(self) -> None:
        self.loaded.emit(data_service.get_scheduler_status())


class MetricCard(QFrame):
    """Tarjeta de métrica con panel informativo deslizante al pasar el cursor."""

    def __init__(
        self,
        icon: str,
        title: str,
        hint: str = "",
        tone: str = "blue",
        details: str = "",
    ) -> None:
        super().__init__()

        self.setObjectName("MetricCard")
        self.setProperty("tone", tone)
        self.setFixedHeight(148)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setCursor(Qt.PointingHandCursor)

        # The frame itself has no layout. Both pages are direct children and
        # are moved horizontally by QPropertyAnimation.
        self._showing_details = False

        self.front_page = QWidget(self)
        self.front_page.setObjectName("MetricSlidePage")

        front_layout = QVBoxLayout(self.front_page)
        front_layout.setContentsMargins(18, 16, 18, 16)
        front_layout.setSpacing(7)

        top_layout = QHBoxLayout()
        self.icon = QLabel(icon)
        self.icon.setObjectName("MetricIcon")
        self.icon.setProperty("tone", tone)
        self.icon.setFixedSize(38, 38)
        self.icon.setAlignment(Qt.AlignCenter)
        top_layout.addWidget(self.icon)
        top_layout.addStretch()
        front_layout.addLayout(top_layout)

        self.value = QLabel("—")
        self.value.setObjectName("MetricValue")
        self.value.setProperty("tone", tone)

        self.title = QLabel(title)
        self.title.setObjectName("MetricTitle")

        self.hint = QLabel(hint)
        self.hint.setObjectName("MetricHint")
        self.hint.setWordWrap(True)

        front_layout.addWidget(self.value)
        front_layout.addWidget(self.title)
        front_layout.addWidget(self.hint)

        self.details_page = QWidget(self)
        self.details_page.setObjectName("MetricSlidePage")

        details_layout = QVBoxLayout(self.details_page)
        details_layout.setContentsMargins(18, 15, 18, 13)
        details_layout.setSpacing(6)

        details_header = QHBoxLayout()
        details_icon = QLabel(icon)
        details_icon.setObjectName("MetricIcon")
        details_icon.setProperty("tone", tone)
        details_icon.setFixedSize(34, 34)
        details_icon.setAlignment(Qt.AlignCenter)

        self.details_title = QLabel(title)
        self.details_title.setObjectName("MetricDetailTitle")
        self.details_title.setWordWrap(True)

        details_header.addWidget(details_icon)
        details_header.addWidget(self.details_title, 1)
        details_layout.addLayout(details_header)

        self.details_text = QLabel(details or hint)
        self.details_text.setObjectName("MetricDetailText")
        self.details_text.setWordWrap(True)
        details_layout.addWidget(self.details_text)

        self.details_status = QLabel(hint)
        self.details_status.setObjectName("MetricDetailStatus")
        self.details_status.setWordWrap(True)
        details_layout.addWidget(self.details_status)
        details_layout.addStretch()

        footer = QLabel("Retire el cursor para volver al contador")
        footer.setObjectName("MetricDetailFooter")
        details_layout.addWidget(footer)

        self.front_animation = QPropertyAnimation(self.front_page, b"pos", self)
        self.details_animation = QPropertyAnimation(self.details_page, b"pos", self)

        for animation in (self.front_animation, self.details_animation):
            animation.setDuration(260)
            animation.setEasingCurve(QEasingCurve.OutCubic)

        self.animation_group = QParallelAnimationGroup(self)
        self.animation_group.addAnimation(self.front_animation)
        self.animation_group.addAnimation(self.details_animation)

    def resizeEvent(self, event) -> None:
        """Mantiene ambas páginas alineadas al redimensionar la ventana."""
        super().resizeEvent(event)

        width = self.width()
        height = self.height()
        self.front_page.resize(width, height)
        self.details_page.resize(width, height)

        if self._showing_details:
            self.front_page.move(-width, 0)
            self.details_page.move(0, 0)
        else:
            self.front_page.move(0, 0)
            self.details_page.move(width, 0)

    def _slide(self, show_details: bool) -> None:
        """Anima el cambio entre el contador y el panel informativo."""
        if self._showing_details == show_details:
            return

        self._showing_details = show_details
        width = max(self.width(), 1)

        self.animation_group.stop()
        self.front_animation.setStartValue(self.front_page.pos())
        self.details_animation.setStartValue(self.details_page.pos())

        if show_details:
            self.front_animation.setEndValue(QPoint(-width, 0))
            self.details_animation.setEndValue(QPoint(0, 0))
        else:
            self.front_animation.setEndValue(QPoint(0, 0))
            self.details_animation.setEndValue(QPoint(width, 0))

        self.animation_group.start()

    def enterEvent(self, event) -> None:
        self._slide(True)
        super().enterEvent(event)

    def leaveEvent(self, event) -> None:
        self._slide(False)
        super().leaveEvent(event)

    def set_value(self, value: int | str, hint: str | None = None) -> None:
        """Actualiza el contador y, cuando procede, el estado del panel."""
        self.value.setText(str(value))
        if hint is not None:
            self.hint.setText(hint)
            self.details_status.setText(f"Estado actual: {hint}")


class DetailDrawer(QFrame):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("Drawer")
        self.setMaximumWidth(0)
        self.setMinimumWidth(0)
        self._record: dict[str, Any] | None = None

        self.animation = QPropertyAnimation(self, b"maximumWidth", self)
        self.animation.setDuration(220)
        self.animation.setEasingCurve(QEasingCurve.OutCubic)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(9)

        header = QHBoxLayout()
        title = QLabel("Detalle del correo")
        title.setObjectName("DrawerTitle")
        close_button = QPushButton("✕")
        close_button.setObjectName("SecondaryButton")
        close_button.setFixedWidth(42)
        close_button.clicked.connect(self.close_drawer)
        header.addWidget(title)
        header.addStretch()
        header.addWidget(close_button)
        layout.addLayout(header)

        self.values: dict[str, QLabel] = {}
        for key, caption in (
            ("Date", "FECHA"),
            ("Direction", "DIRECCIÓN"),
            ("Project Folder", "PROYECTO"),
            ("Address Folder", "UBICACIÓN / DIRECCIÓN"),
            ("Contact", "CONTACTO"),
            ("Topic", "TEMA"),
            ("Subject", "ASUNTO"),
            ("Attachments", "ADJUNTOS"),
            ("Priority", "PRIORIDAD"),
            ("_pending", "PENDIENTE DE RESPUESTA"),
            ("Summary", "RESUMEN"),
            ("_status", "ESTADO"),
        ):
            label = QLabel(caption)
            label.setObjectName("DrawerLabel")
            value = QLabel("—")
            value.setObjectName("DrawerValue")
            value.setWordWrap(True)
            value.setTextInteractionFlags(Qt.TextSelectableByMouse)
            layout.addWidget(label)
            layout.addWidget(value)
            self.values[key] = value

        layout.addStretch()
        self.open_button = QPushButton("Abrir carpeta")
        self.open_button.setObjectName("PrimaryButton")
        self.open_button.clicked.connect(self._open_folder)
        self.copy_button = QPushButton("Copiar ruta")
        self.copy_button.setObjectName("SecondaryButton")
        self.copy_button.clicked.connect(self._copy_path)
        layout.addWidget(self.open_button)
        layout.addWidget(self.copy_button)

    def show_record(self, record: dict[str, Any]) -> None:
        self._record = record
        for key, label in self.values.items():
            value = record.get(key, "")
            if key == "Priority":
                value = priority_label(value)
            elif key == "_pending":
                value = "Sí" if value else "No"
            label.setText(str(value or "—"))
        self.animation.stop()
        self.animation.setStartValue(self.maximumWidth())
        self.animation.setEndValue(370)
        self.animation.start()

    def close_drawer(self) -> None:
        self.animation.stop()
        self.animation.setStartValue(self.maximumWidth())
        self.animation.setEndValue(0)
        self.animation.start()

    def _open_folder(self) -> None:
        if not self._record:
            return
        ok, message = data_service.open_path(self._record.get("Folder Path", ""))
        if not ok:
            QMessageBox.warning(self, "No se pudo abrir", message)

    def _copy_path(self) -> None:
        if not self._record:
            return
        QApplication.clipboard().setText(str(self._record.get("Folder Path", "")))


class DashboardPage(QWidget):
    process_requested = Signal()
    refresh_requested = Signal()
    open_output_requested = Signal()

    def __init__(self) -> None:
        super().__init__()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 24)
        outer.setSpacing(18)

        hero = QFrame()
        hero.setObjectName("HeroPanel")
        hero_layout = QHBoxLayout(hero)
        hero_layout.setContentsMargins(22, 18, 22, 18)
        hero_text = QVBoxLayout()
        title = QLabel("Gestión automática del correo")
        title.setObjectName("SectionTitle")
        subtitle = QLabel(
            "Clasifica correos de Outlook, guarda adjuntos y organiza la documentación por proyecto."
        )
        subtitle.setObjectName("MutedText")
        subtitle.setWordWrap(True)
        hero_text.addWidget(title)
        hero_text.addWidget(subtitle)
        hero_layout.addLayout(hero_text, 1)

        self.process_button = QPushButton("▶  Procesar nuevos correos")
        self.process_button.setObjectName("ProcessButton")
        self.process_button.clicked.connect(self.process_requested)
        if data_service.IS_VIEWER:
            title.setText("Vista compartida del correo")
            subtitle.setText(
                "Consulta los correos del buzón configurado que procesa el ordenador principal."
            )
            self.process_button.hide()
        self.refresh_button = QPushButton("↻  Actualizar")
        self.refresh_button.setObjectName("SecondaryButton")
        self.refresh_button.clicked.connect(self.refresh_requested)
        output_button = QPushButton("▣  Abrir archivo")
        output_button.setObjectName("SecondaryButton")
        output_button.clicked.connect(self.open_output_requested)
        hero_layout.addWidget(output_button)
        hero_layout.addWidget(self.refresh_button)
        hero_layout.addWidget(self.process_button)
        outer.addWidget(hero)

        cards = QHBoxLayout()
        cards.setSpacing(14)
        self.processed_card = MetricCard(
            "✓",
            "Correos procesados",
            "Total de correos registrados por el sistema",
            "green",
            details=(
                "Muestra el número de correos únicos que ya han sido gestionados "
                "por el asistente. No se procesarán de nuevo en futuras ejecuciones."
            ),
        )
        self.flagged_card = MetricCard(
            "⚑",
            "Marcados en Outlook",
            "Correos identificados por el asistente en Outlook",
            "orange",
            details=(
                "Muestra los correos que contienen la categoría configurada por el "
                "asistente. Puede consultarlos desde Marcados en Outlook."
            ),
        )
        self.filed_card = MetricCard(
            "▤",
            "Correos archivados",
            "Correos guardados correctamente en carpetas",
            "blue",
            details=(
                "Muestra los correos clasificados y guardados en la carpeta del "
                "proyecto correspondiente, junto con sus archivos adjuntos."
            ),
        )
        self.review_card = MetricCard(
            "!",
            "Pendientes de revisión",
            "Correos sin una clasificación segura",
            "red",
            details=(
                "Muestra los correos que el sistema no pudo clasificar con suficiente "
                "seguridad. Estos mensajes deben revisarse manualmente."
            ),
        )
        self.pending_card = MetricCard(
            "↩",
            "Pendientes de respuesta",
            "Correos que requieren una respuesta de la empresa",
            "purple",
            details=(
                "Muestra los correos entrantes relevantes que el agente de IA ha "
                "identificado como pendientes de una respuesta escrita."
            ),
        )
        cards.addWidget(self.processed_card)
        cards.addWidget(self.flagged_card)
        cards.addWidget(self.filed_card)
        cards.addWidget(self.review_card)
        cards.addWidget(self.pending_card)
        outer.addLayout(cards)

        content = QHBoxLayout()
        content.setSpacing(16)

        activity_panel = QFrame()
        activity_panel.setObjectName("Panel")
        activity_layout = QVBoxLayout(activity_panel)
        activity_layout.setContentsMargins(18, 16, 18, 16)
        activity_title = QLabel("Actividad reciente")
        activity_title.setObjectName("SectionTitle")
        activity_layout.addWidget(activity_title)
        self.recent_table = QTableWidget(0, 4)
        self.recent_table.setHorizontalHeaderLabels(
            ["Fecha", "Guardado en", "Asunto", "Prioridad"]
        )
        configure_table(self.recent_table)
        recent_header = self.recent_table.horizontalHeader()
        recent_header.setSectionResizeMode(0, QHeaderView.Interactive)
        self.recent_table.setColumnWidth(0, 105)
        recent_header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        recent_header.setSectionResizeMode(2, QHeaderView.Stretch)
        recent_header.setSectionResizeMode(3, QHeaderView.Interactive)
        self.recent_table.setColumnWidth(3, 95)
        activity_layout.addWidget(self.recent_table)
        content.addWidget(activity_panel, 3)

        attention_panel = QFrame()
        attention_panel.setObjectName("AttentionPanel")
        attention_panel.setMinimumWidth(280)
        attention_layout = QVBoxLayout(attention_panel)
        attention_layout.setContentsMargins(18, 16, 18, 16)
        attention_title = QLabel("Atención necesaria")
        attention_title.setObjectName("SectionTitle")
        self.review_summary = QLabel("Cargando...")
        self.review_summary.setWordWrap(True)
        self.review_summary.setObjectName("MutedText")
        self.quarantine_summary = QLabel("")
        self.quarantine_summary.setWordWrap(True)
        self.quarantine_summary.setObjectName("MutedText")
        self.priority_summary = QLabel("")
        self.priority_summary.setWordWrap(True)
        self.priority_summary.setObjectName("MutedText")
        self.pending_summary = QLabel("")
        self.pending_summary.setWordWrap(True)
        self.pending_summary.setObjectName("MutedText")
        self.outlook_summary = QLabel("Consultando Outlook...")
        self.outlook_summary.setWordWrap(True)
        self.outlook_summary.setObjectName("MutedText")
        attention_layout.addWidget(attention_title)
        attention_layout.addSpacing(8)
        attention_layout.addWidget(self.review_summary)
        attention_layout.addWidget(self.quarantine_summary)
        attention_layout.addWidget(self.priority_summary)
        attention_layout.addWidget(self.pending_summary)
        attention_layout.addWidget(self.outlook_summary)
        attention_layout.addStretch()
        content.addWidget(attention_panel, 1)
        outer.addLayout(content, 1)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.hide()
        outer.addWidget(self.progress)

    def set_processing(self, active: bool) -> None:
        if data_service.IS_VIEWER:
            return
        self.process_button.setDisabled(active)
        self.refresh_button.setDisabled(active)
        self.process_button.setText("Procesando..." if active else "▶  Procesar nuevos correos")
        self.progress.setVisible(active)

    def update_local_data(
        self,
        rows: list[dict[str, Any]],
        attachments: list[dict[str, Any]],
        processed: int,
        activity_rows: list[dict[str, Any]] | None = None,
        pending_rows: list[dict[str, Any]] | None = None,
    ) -> None:
        review = [row for row in rows if row.get("Project Folder") == "UNSORTED" or row.get("_status") == "AI_REVIEW"]
        quarantined = [row for row in attachments if row.get("Estado") == "quarantined"]
        self.processed_card.set_value(processed)
        self.filed_card.set_value(len([row for row in rows if row.get("_status") not in {"PENDING_AI", "AI_REVIEW"}]))
        self.review_card.set_value(len(review))
        pending_total = len(pending_rows or [])
        high_priority = [row for row in rows if int(row.get("_priority") or 0) >= 4]
        critical_priority = [row for row in rows if int(row.get("_priority") or 0) == 5]
        self.pending_card.set_value(pending_total)
        self.review_summary.setText(
            "✓ No hay correos pendientes de clasificación."
            if not review
            else f"⚠ {len(review)} correo(s) necesitan revisión manual de clasificación/IA."
        )
        self.quarantine_summary.setText(
            "✓ No hay adjuntos bloqueados."
            if not quarantined
            else f"⚠ {len(quarantined)} adjunto(s) fueron bloqueados por seguridad."
        )
        self.priority_summary.setText(
            "✓ No hay correos de prioridad alta o crítica."
            if not high_priority
            else f"⚠ {len(high_priority)} correo(s) tienen prioridad 4 o 5; {len(critical_priority)} son críticos."
        )
        self.pending_summary.setText(
            "✓ No hay correos pendientes de respuesta."
            if pending_total == 0
            else f"↩ {pending_total} correo(s) necesitan una respuesta de la empresa."
        )

        source_rows = activity_rows or rows
        status_labels = {
            "ARCHIVADO": "Archivado",
            "OMITIDO": "Omitido",
            "REVISAR": "Revisar",
            "ERROR": "Error",
            "PENDIENTE": "Pendiente de respuesta",
            "PENDING_AI": "Pendiente IA",
            "AI_REVIEW": "Revisar IA",
            "PROCESADO": "Procesado",
        }
        self.recent_table.setRowCount(0)
        for record in source_rows[:12]:
            row = self.recent_table.rowCount()
            self.recent_table.insertRow(row)
            saved_path = (
                record.get("Folder Path")
                or record.get("Path")
                or record.get("Ruta")
                or ""
            )
            values = [
                record.get("Date", ""),
                saved_path or "—",
                record.get("Subject", ""),
                priority_label(record.get("_priority")),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if column == 1:
                    item.setToolTip(str(saved_path))
                if column == 2:
                    item.setToolTip(str(record.get("Subject", "")))
                apply_priority_style(item, record.get("_priority"), emphasize=column == 3)
                self.recent_table.setItem(row, column, item)

    def update_flagged(self, summary: data_service.OutlookFlagSummary) -> None:
        if summary.error:
            self.flagged_card.set_value("—", "No se pudo consultar Outlook")
            self.outlook_summary.setText(f"⚠ Outlook: {summary.error}")
            return
        hint = "Recuento estimado" if summary.estimated else "Recuento directo desde Outlook"
        self.flagged_card.set_value(summary.total, hint)
        suffix = " (estimado)" if summary.estimated else ""
        self.outlook_summary.setText(
            f"✓ {summary.total} correo(s) tienen la categoría «{data_service.PROCESSED_CATEGORY_NAME}»{suffix}."
        )


class EmailsPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.all_rows: list[dict[str, Any]] = []

        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 24)
        outer.setSpacing(14)

        heading = QLabel("Correos archivados")
        heading.setObjectName("PageTitle")
        subtitle = QLabel(
            "Busca, filtra y consulta el resumen de cada correo procesado."
        )
        subtitle.setObjectName("PageSubtitle")
        priority_legend = QLabel(
            "Prioridad: 5 crítica (rojo) · 4 alta (naranja) · 3 normal (amarillo)"
        )
        priority_legend.setObjectName("MutedText")
        outer.addWidget(heading)
        outer.addWidget(subtitle)
        outer.addWidget(priority_legend)

        filters = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText(
            "Buscar por asunto, resumen, contacto, remitente o proyecto..."
        )
        self.project_filter = QComboBox()
        self.project_filter.addItem("Todos los proyectos")
        self.status_filter = QComboBox()
        self.status_filter.addItems(
            ["Todos los estados", "Procesado", "Pendiente IA", "Revisar", "Revisar IA"]
        )
        filters.addWidget(self.search, 2)
        filters.addWidget(self.project_filter, 1)
        filters.addWidget(self.status_filter, 1)
        outer.addLayout(filters)

        # Resumen de una línea generado por el mismo flujo que alimenta
        # el archivo Excel "Informe de Emails.xlsx".
        summary_panel = QFrame()
        summary_panel.setObjectName("EmailSummaryPanel")
        summary_layout = QVBoxLayout(summary_panel)
        summary_layout.setContentsMargins(16, 12, 16, 12)
        summary_layout.setSpacing(4)

        summary_title = QLabel("Resumen del correo seleccionado")
        summary_title.setObjectName("EmailSummaryTitle")
        self.email_summary = QLabel(
            "Seleccione un correo para ver su resumen de una línea."
        )
        self.email_summary.setObjectName("EmailSummaryText")
        self.email_summary.setWordWrap(True)
        self.email_summary.setTextInteractionFlags(Qt.TextSelectableByMouse)
        summary_layout.addWidget(summary_title)
        summary_layout.addWidget(self.email_summary)
        outer.addWidget(summary_panel)

        self.table = QTableWidget(0, 8)
        self.table.setHorizontalHeaderLabels(
            [
                "Fecha",
                "Guardado en",
                "Contacto",
                "Remitente",
                "Asunto",
                "Resumen",
                "Adjuntos",
                "Prioridad",
            ]
        )
        configure_table(self.table)
        self.table.setWordWrap(False)

        # Todas las columnas de Correos son ajustables por el usuario.
        # Arrastra el separador de la cabecera para ampliar o reducir cualquier
        # columna. Doble clic en un separador autoajusta esa columna al contenido.
        header = self.table.horizontalHeader()
        header.setMinimumSectionSize(24)
        header.setStretchLastSection(False)
        header.setCascadingSectionResizes(False)
        for column in range(self.table.columnCount()):
            header.setSectionResizeMode(column, QHeaderView.Interactive)

        initial_widths = {
            0: 95,   # Fecha
            1: 520,  # Guardado en
            2: 125,  # Contacto
            3: 190,  # Remitente
            4: 320,  # Asunto
            5: 420,  # Resumen
            6: 90,   # Adjuntos
            7: 95,   # Prioridad
        }
        for column, width in initial_widths.items():
            self.table.setColumnWidth(column, width)

        header.sectionHandleDoubleClicked.connect(
            lambda logical_index: self.table.resizeColumnToContents(logical_index)
        )
        self.table.setTextElideMode(Qt.ElideRight)
        # La tabla ahora utiliza todo el ancho disponible. El panel lateral de
        # detalles se ha eliminado; la selección solo actualiza el resumen superior.
        outer.addWidget(self.table, 1)

        self.search.textChanged.connect(self.apply_filters)
        self.project_filter.currentTextChanged.connect(self.apply_filters)
        self.status_filter.currentTextChanged.connect(self.apply_filters)
        self.table.itemSelectionChanged.connect(self._selection_changed)
        self.table.cellDoubleClicked.connect(self._open_selected)

    def set_rows(self, rows: list[dict[str, Any]]) -> None:
        self.all_rows = rows
        current = self.project_filter.currentText()
        projects = sorted(
            {
                str(row.get("Project Folder", ""))
                for row in rows
                if row.get("Project Folder")
            }
        )
        self.project_filter.blockSignals(True)
        self.project_filter.clear()
        self.project_filter.addItem("Todos los proyectos")
        self.project_filter.addItems(projects)
        if current in projects:
            self.project_filter.setCurrentText(current)
        self.project_filter.blockSignals(False)
        self.apply_filters()

    def apply_filters(self) -> None:
        query = self.search.text().strip().lower()
        project = self.project_filter.currentText()
        status = self.status_filter.currentText()

        filtered: list[dict[str, Any]] = []
        for record in self.all_rows:
            haystack = " ".join(
                str(record.get(key, ""))
                for key in (
                    "Subject",
                    "Summary",
                    "Contact",
                    "SenderName",
                    "SenderEmail",
                    "SenderDisplay",
                    "Project Folder",
                    "Folder Path",
                    "Topic",
                    "Sender/Recipient",
                    "Priority",
                )
            ).lower()
            if query and query not in haystack:
                continue
            if (
                project != "Todos los proyectos"
                and record.get("Project Folder") != project
            ):
                continue
            if record.get("Direction") == "SALIENTE":
                continue
            if (
                status == "Procesado"
                and record.get("_status") != "PROCESADO"
            ):
                continue
            if status == "Pendiente IA" and record.get("_status") != "PENDING_AI":
                continue
            if status == "Revisar" and record.get("_status") != "REVISAR":
                continue
            if status == "Revisar IA" and record.get("_status") != "AI_REVIEW":
                continue
            filtered.append(record)

        visible = filtered[:DISPLAY_ROW_LIMIT]
        self.table.setUpdatesEnabled(False)
        self.table.setRowCount(len(visible))
        for row, record in enumerate(visible):
            summary = str(record.get("Summary") or "").strip()
            saved_path = str(record.get("Folder Path") or "").strip()
            values = [
                record.get("Date", ""),
                saved_path,
                record.get("Contact", ""),
                record.get("SenderDisplay") or data_service.sender_display(record),
                record.get("Subject", ""),
                summary or "Sin resumen disponible",
                record.get("Attachments", "0"),
                priority_label(record.get("_priority")),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if column == 0:
                    item.setData(Qt.UserRole, record)
                if column == 1:
                    item.setToolTip(saved_path)
                if column == 2:
                    item.setToolTip(str(record.get("Contact", "")))
                if column == 3:
                    item.setToolTip(str(record.get("SenderDisplay") or data_service.sender_display(record)))
                if column == 4:
                    item.setToolTip(str(record.get("Subject", "")))
                if column == 5:
                    item.setToolTip(summary or "Sin resumen disponible")
                apply_priority_style(
                    item,
                    record.get("_priority"),
                    emphasize=column == 7,
                )
                self.table.setItem(row, column, item)
        self.table.setUpdatesEnabled(True)

        if visible:
            # Do not auto-select the first row: Windows selection colors hide
            # the red/orange/yellow priority background until selection moves.
            self.table.clearSelection()
        else:
            self.email_summary.setText(
                "No hay correos que coincidan con los filtros seleccionados."
            )

    def _selected_record(self) -> dict[str, Any] | None:
        row = self.table.currentRow()
        if row < 0:
            return None
        item = self.table.item(row, 0)
        return item.data(Qt.UserRole) if item else None

    def _selection_changed(self) -> None:
        record = self._selected_record()
        if not record:
            self.email_summary.setText(
                "Seleccione un correo para ver su resumen de una línea."
            )
            return

        summary = str(record.get("Summary") or "").strip()
        self.email_summary.setText(
            summary
            or "No hay un resumen disponible para este correo."
        )

    def _open_selected(self, _row: int, _column: int) -> None:
        record = self._selected_record()
        if not record:
            return
        ok, message = data_service.open_email_record(record)
        if not ok:
            QMessageBox.warning(self, "No se pudo abrir", message)


class FlaggedPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 24)
        outer.setSpacing(14)

        heading_line = QHBoxLayout()
        text = QVBoxLayout()
        title = QLabel("Correos marcados en Outlook")
        title.setObjectName("PageTitle")
        subtitle = QLabel(
            f"Correos con la categoría «{data_service.PROCESSED_CATEGORY_NAME}» y marca de seguimiento."
        )
        subtitle.setObjectName("PageSubtitle")
        text.addWidget(title)
        text.addWidget(subtitle)
        heading_line.addLayout(text)
        heading_line.addStretch()
        self.count_label = QLabel("Consultando Outlook...")
        self.count_label.setObjectName("MutedText")
        heading_line.addWidget(self.count_label)
        outer.addLayout(heading_line)

        self.search = QLineEdit()
        self.search.setPlaceholderText("Buscar dentro de los correos marcados...")
        outer.addWidget(self.search)

        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["Fecha", "Dirección", "Contacto", "Asunto", "Categoría"])
        configure_table(self.table)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        outer.addWidget(self.table, 1)
        self._rows: list[dict[str, Any]] = []
        self.search.textChanged.connect(self._render)

    def set_summary(self, summary: data_service.OutlookFlagSummary) -> None:
        if summary.error:
            self.count_label.setText(f"No disponible: {summary.error}")
            self._rows = []
        else:
            suffix = " (estimado)" if summary.estimated else ""
            self.count_label.setText(f"{summary.total} marcados{suffix}; mostrando hasta {len(summary.rows)}")
            self._rows = summary.rows
        self._render()

    def _render(self) -> None:
        query = self.search.text().strip().lower()
        self.table.setRowCount(0)
        for record in self._rows:
            haystack = " ".join(str(value) for value in record.values()).lower()
            if query and query not in haystack:
                continue
            row = self.table.rowCount()
            self.table.insertRow(row)
            for column, key in enumerate(("Fecha", "Dirección", "Contacto", "Asunto", "Categoría")):
                self.table.setItem(row, column, QTableWidgetItem(str(record.get(key, ""))))


class PendingPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self._rows: list[dict[str, Any]] = []

        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 24)
        outer.setSpacing(14)

        heading = QLabel("Pendientes de respuesta")
        heading.setObjectName("PageTitle")
        subtitle = QLabel(
            "Correos entrantes que el sistema ha identificado como pendientes de una respuesta escrita."
        )
        subtitle.setObjectName("PageSubtitle")
        subtitle.setWordWrap(True)
        outer.addWidget(heading)
        outer.addWidget(subtitle)

        controls = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("Buscar por remitente o asunto...")
        self.search.textChanged.connect(self._render)
        open_list = QPushButton("Abrir lista CSV")
        open_list.setObjectName("SecondaryButton")
        open_list.clicked.connect(self._open_list)
        open_output = QPushButton("Abrir carpeta de salida")
        open_output.setObjectName("PrimaryButton")
        open_output.clicked.connect(self._open_output)
        controls.addWidget(self.search, 1)
        controls.addWidget(open_list)
        controls.addWidget(open_output)
        outer.addLayout(controls)

        self.summary = QLabel("0 correos pendientes")
        self.summary.setObjectName("MutedText")
        outer.addWidget(self.summary)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Fecha", "Remitente", "Asunto", "Prioridad"])
        configure_table(self.table)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        outer.addWidget(self.table, 1)

    def set_rows(self, rows: list[dict[str, Any]]) -> None:
        self._rows = rows
        self._render()

    def _render(self) -> None:
        query = self.search.text().strip().lower()
        filtered = []
        for record in self._rows:
            text = " ".join(
                str(record.get(key, ""))
                for key in ("Fecha", "Remitente", "Asunto")
            ).lower()
            if query and query not in text:
                continue
            filtered.append(record)

        visible = filtered[:DISPLAY_ROW_LIMIT]
        self.table.setUpdatesEnabled(False)
        self.table.setRowCount(len(visible))
        for row, record in enumerate(visible):
            values = [
                record.get("Fecha", ""),
                record.get("Remitente", ""),
                record.get("Asunto", ""),
                priority_label(record.get("_priority")),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                apply_priority_style(item, record.get("_priority"), emphasize=column == 3)
                self.table.setItem(row, column, item)
        self.table.setUpdatesEnabled(True)
        suffix = "" if len(filtered) <= DISPLAY_ROW_LIMIT else f"; mostrando {len(visible)}"
        self.summary.setText(f"{len(filtered)} correo(s) pendiente(s) de respuesta{suffix}")

    def _open_list(self) -> None:
        ok, message = data_service.open_path(data_service.PENDING_LIST_PATH)
        if not ok:
            QMessageBox.warning(self, "No se pudo abrir", message)

    def _open_output(self) -> None:
        ok, message = data_service.open_path(data_service.OUTPUT_ROOT)
        if not ok:
            QMessageBox.warning(self, "No se pudo abrir", message)


class ReportPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self._rows: list[dict[str, Any]] = []

        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 24)
        outer.setSpacing(14)

        heading = QLabel("Informe de correos")
        heading.setObjectName("PageTitle")
        subtitle = QLabel(
            "Resumen legible de los correos archivados, generado por el backend después de cada ejecución."
        )
        subtitle.setObjectName("PageSubtitle")
        subtitle.setWordWrap(True)
        outer.addWidget(heading)
        outer.addWidget(subtitle)

        controls = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("Buscar por remitente, asunto o resumen...")
        self.search.textChanged.connect(self._render)
        open_excel = QPushButton("Abrir informe Excel")
        open_excel.setObjectName("PrimaryButton")
        open_excel.clicked.connect(self._open_excel)
        open_csv = QPushButton("Abrir registro CSV")
        open_csv.setObjectName("SecondaryButton")
        open_csv.clicked.connect(self._open_csv)
        controls.addWidget(self.search, 1)
        controls.addWidget(open_csv)
        controls.addWidget(open_excel)
        outer.addLayout(controls)

        self.summary = QLabel("0 entradas en el informe")
        self.summary.setObjectName("MutedText")
        outer.addWidget(self.summary)

        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(
            ["Fecha", "Remitente", "Email", "Asunto", "Prioridad", "Resumen"]
        )
        configure_table(self.table)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self.table.cellDoubleClicked.connect(self._open_selected)
        outer.addWidget(self.table, 1)

    def set_rows(self, rows: list[dict[str, Any]]) -> None:
        self._rows = rows
        self._render()

    def _render(self) -> None:
        query = self.search.text().strip().lower()
        filtered = []
        for record in self._rows:
            text = " ".join(
                str(record.get(key, ""))
                for key in ("NombreRemitente", "EmailRemitente", "Asunto", "Resumen")
            ).lower()
            if query and query not in text:
                continue
            filtered.append(record)

        visible = filtered[:DISPLAY_ROW_LIMIT]
        self.table.setUpdatesEnabled(False)
        self.table.setRowCount(len(visible))
        for row, record in enumerate(visible):
            values = [
                record.get("FechaHora", ""),
                record.get("NombreRemitente", ""),
                record.get("EmailRemitente", ""),
                record.get("Asunto", ""),
                priority_label(record.get("_priority")),
                record.get("Resumen", ""),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if column == 0:
                    item.setData(Qt.UserRole, record)
                apply_priority_style(item, record.get("_priority"), emphasize=column == 4)
                self.table.setItem(row, column, item)
        self.table.setUpdatesEnabled(True)
        suffix = "" if len(filtered) <= DISPLAY_ROW_LIMIT else f"; mostrando {len(visible)}"
        self.summary.setText(f"{len(filtered)} entrada(s) en el informe{suffix}")

    def _selected(self) -> dict[str, Any] | None:
        row = self.table.currentRow()
        if row < 0:
            return None
        item = self.table.item(row, 0)
        return item.data(Qt.UserRole) if item else None

    def _open_selected(self, _row: int, _column: int) -> None:
        record = self._selected()
        if not record:
            return
        ok, message = data_service.open_path(record.get("Ruta", ""))
        if not ok:
            QMessageBox.warning(self, "No se pudo abrir", message)

    def _open_excel(self) -> None:
        ok, message = data_service.open_path(data_service.REPORT_XLSX_PATH)
        if not ok:
            QMessageBox.warning(self, "No se pudo abrir", message)

    def _open_csv(self) -> None:
        ok, message = data_service.open_path(data_service.REPORT_LOG_PATH)
        if not ok:
            QMessageBox.warning(self, "No se pudo abrir", message)


class ExcelReportPage(QWidget):
    """Página dedicada al informe Excel generado en la carpeta de salida."""

    def __init__(self) -> None:
        super().__init__()

        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 24)
        outer.setSpacing(14)

        heading = QLabel("Informe Excel")
        heading.setObjectName("PageTitle")
        subtitle = QLabel(
            "Acceso directo al archivo «Informe de Emails.xlsx» generado en la carpeta de salida."
        )
        subtitle.setObjectName("PageSubtitle")
        subtitle.setWordWrap(True)
        outer.addWidget(heading)
        outer.addWidget(subtitle)

        panel = QFrame()
        panel.setObjectName("Panel")
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(22, 20, 22, 20)
        panel_layout.setSpacing(12)

        header = QHBoxLayout()
        icon = QLabel("XLSX")
        icon.setObjectName("MetricIcon")
        icon.setProperty("tone", "green")
        icon.setFixedSize(58, 42)
        icon.setAlignment(Qt.AlignCenter)

        title_block = QVBoxLayout()
        self.file_name = QLabel(data_service.REPORT_XLSX_PATH.name)
        self.file_name.setObjectName("SectionTitle")
        self.status_label = QLabel("Consultando el archivo...")
        self.status_label.setObjectName("WarningPill")
        self.status_label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        title_block.addWidget(self.file_name)
        title_block.addWidget(self.status_label, 0, Qt.AlignLeft)

        header.addWidget(icon)
        header.addLayout(title_block, 1)
        panel_layout.addLayout(header)

        self.path_value = add_setting_row(
            panel_layout,
            "Ubicación",
            str(data_service.REPORT_XLSX_PATH),
        )
        self.size_value = add_setting_row(panel_layout, "Tamaño", "—")
        self.modified_value = add_setting_row(panel_layout, "Última actualización", "—")
        self.entries_value = add_setting_row(panel_layout, "Correos incluidos", "0")
        outer.addWidget(panel)

        actions = QHBoxLayout()
        self.open_button = QPushButton("Abrir informe Excel")
        self.open_button.setObjectName("PrimaryButton")
        self.open_button.clicked.connect(self._open_excel)

        open_folder = QPushButton("Abrir carpeta de salida")
        open_folder.setObjectName("SecondaryButton")
        open_folder.clicked.connect(self._open_output)

        refresh = QPushButton("↻  Actualizar estado")
        refresh.setObjectName("SecondaryButton")
        refresh.clicked.connect(self.refresh_status)

        actions.addWidget(self.open_button)
        actions.addWidget(open_folder)
        actions.addWidget(refresh)
        actions.addStretch()
        outer.addLayout(actions)

        note = QLabel(
            "El informe se vuelve a generar después del procesamiento. Si todavía no existe, "
            "procese al menos un correo relevante y actualice esta página."
        )
        note.setObjectName("MutedText")
        note.setWordWrap(True)
        outer.addWidget(note)
        outer.addStretch()

    def refresh_status(self) -> None:
        self.apply_info(data_service.get_report_excel_info())

    def apply_info(self, info: dict[str, Any]) -> None:
        self.path_value.setText(str(info.get("path", "")))
        self.size_value.setText(str(info.get("size_text", "—")))
        self.modified_value.setText(str(info.get("modified", "—")))
        self.entries_value.setText(str(info.get("entries", 0)))

        exists = bool(info.get("exists"))
        self.open_button.setEnabled(exists)
        if exists:
            self.status_label.setText("Disponible")
            self.status_label.setObjectName("SuccessPill")
        else:
            self.status_label.setText("Todavía no generado")
            self.status_label.setObjectName("WarningPill")
        self.status_label.style().unpolish(self.status_label)
        self.status_label.style().polish(self.status_label)

    def _open_excel(self) -> None:
        ok, message = data_service.open_path(data_service.REPORT_XLSX_PATH)
        if not ok:
            QMessageBox.warning(self, "No se pudo abrir", message)

    def _open_output(self) -> None:
        ok, message = data_service.open_path(data_service.OUTPUT_ROOT)
        if not ok:
            QMessageBox.warning(self, "No se pudo abrir", message)


class AttachmentsPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.rows: list[dict[str, Any]] = []
        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 24)
        outer.setSpacing(14)

        title = QLabel("Adjuntos")
        title.setObjectName("PageTitle")
        subtitle = QLabel("Consulta los archivos guardados y los adjuntos bloqueados por seguridad.")
        subtitle.setObjectName("PageSubtitle")
        outer.addWidget(title)
        outer.addWidget(subtitle)

        filters = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("Buscar archivo, asunto o proyecto...")
        self.status_filter = QComboBox()
        self.status_filter.addItems(["Todos", "Guardados", "Bloqueados"])
        filters.addWidget(self.search, 1)
        filters.addWidget(self.status_filter)
        outer.addLayout(filters)

        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["Fecha", "Proyecto", "Asunto", "Archivo", "Estado", "Motivo"])
        configure_table(self.table)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        outer.addWidget(self.table, 1)

        buttons = QHBoxLayout()
        buttons.addStretch()
        open_folder = QPushButton("Abrir carpeta")
        open_folder.setObjectName("SecondaryButton")
        open_folder.clicked.connect(self._open_folder)
        open_file = QPushButton("Abrir archivo")
        open_file.setObjectName("PrimaryButton")
        open_file.clicked.connect(self._open_file)
        buttons.addWidget(open_folder)
        buttons.addWidget(open_file)
        outer.addLayout(buttons)

        self.search.textChanged.connect(self._render)
        self.status_filter.currentTextChanged.connect(self._render)
        self.table.cellDoubleClicked.connect(lambda _r, _c: self._open_file())

    def set_rows(self, rows: list[dict[str, Any]]) -> None:
        self.rows = rows
        self._render()

    def _render(self) -> None:
        query = self.search.text().strip().lower()
        status = self.status_filter.currentText()
        filtered: list[dict[str, Any]] = []
        for record in self.rows:
            haystack = " ".join(str(record.get(key, "")) for key in ("Proyecto", "Asunto", "Archivo", "Motivo")).lower()
            if query and query not in haystack:
                continue
            is_quarantined = record.get("Estado") == "quarantined"
            if status == "Guardados" and is_quarantined:
                continue
            if status == "Bloqueados" and not is_quarantined:
                continue
            filtered.append(record)
        visible = filtered[:DISPLAY_ROW_LIMIT]
        self.table.setUpdatesEnabled(False)
        self.table.setRowCount(len(visible))
        for row, record in enumerate(visible):
            is_quarantined = record.get("Estado") == "quarantined"
            values = [
                str(record.get("Fecha", ""))[:16].replace("T", " "),
                record.get("Proyecto", ""),
                record.get("Asunto", ""),
                record.get("Archivo", ""),
                "Bloqueado" if is_quarantined else "Guardado",
                record.get("Motivo", ""),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if column == 0:
                    item.setData(Qt.UserRole, record)
                if column == 4:
                    item.setForeground(QColor("#b42318" if is_quarantined else "#087443"))
                self.table.setItem(row, column, item)
        self.table.setUpdatesEnabled(True)

    def _selected(self) -> dict[str, Any] | None:
        row = self.table.currentRow()
        if row < 0:
            return None
        item = self.table.item(row, 0)
        return item.data(Qt.UserRole) if item else None

    def _open_folder(self) -> None:
        record = self._selected()
        if not record:
            return
        ok, message = data_service.open_path(record.get("Carpeta", ""))
        if not ok:
            QMessageBox.warning(self, "No se pudo abrir", message)

    def _open_file(self) -> None:
        record = self._selected()
        if not record:
            return
        if record.get("Estado") == "quarantined":
            QMessageBox.information(
                self,
                "Archivo bloqueado",
                "Este adjunto fue puesto en cuarentena y no se abre desde la aplicación.",
            )
            return
        ok, message = data_service.open_path(record.get("Ruta", ""))
        if not ok:
            QMessageBox.warning(self, "No se pudo abrir", message)


class SettingsPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 24)
        outer.setSpacing(14)

        title = QLabel("Configuración y estado")
        title.setObjectName("PageTitle")
        subtitle = QLabel("Información local de Outlook, almacenamiento, Gemini y automatización.")
        subtitle.setObjectName("PageSubtitle")
        outer.addWidget(title)
        outer.addWidget(subtitle)

        panel = QFrame()
        panel.setObjectName("Panel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(12)

        mode_text = "Ordenador de procesamiento" if data_service.IS_PROCESSOR else "Ordenador visor"
        self.mode_value = add_setting_row(layout, "Modo de instalación", mode_text)
        self.shared_value = add_setting_row(layout, "Carpeta de datos compartida", str(data_service.SHARED_DATA_PATH))
        self.output_value = add_setting_row(layout, "Carpeta de salida", str(data_service.OUTPUT_ROOT))
        self.archive_value = add_setting_row(layout, "Archivo de correos procesados", str(data_service.PROJECT_YEAR_ROOT))
        self.category_value = add_setting_row(layout, "Categoría de Outlook", data_service.PROCESSED_CATEGORY_NAME)
        gemini_status = "Configurada" if os.environ.get("GEMINI_API_KEY") else "No configurada"
        self.gemini_value = add_setting_row(layout, "Clave de Gemini", gemini_status)
        flag_status = "Activado" if data_service.FLAG_PROCESSED_EMAILS else "Desactivado"
        self.flag_value = add_setting_row(layout, "Marcado automático en Outlook", flag_status)
        pending_status = "Activado" if data_service.CHECK_PENDING_RESPONSES else "Desactivado"
        self.pending_value = add_setting_row(layout, "Detección de respuestas pendientes", pending_status)
        self.pending_folder_value = add_setting_row(layout, "Carpeta de pendientes en Outlook", data_service.PENDING_FOLDER_NAME)
        self.junk_value = add_setting_row(
            layout,
            "Correos no relevantes",
            "Se dejan sin cambios en la bandeja de entrada (comportamiento actual)",
        )
        self.billing_value = add_setting_row(layout, "Carpeta de facturación", str(data_service.BILLING_OUTPUT_ROOT))
        self.admin_value = add_setting_row(layout, "Correo de administración", data_service.ADMINISTRACION_EMAIL)
        self.boss_value = add_setting_row(layout, "Correo del responsable", data_service.BOSS_EMAIL)
        self.target_mailbox_value = add_setting_row(layout, "Buzón Outlook procesado", data_service.TARGET_MAILBOX)
        description_path = str(data_service.DESCRIPTIONS_XLSX_PATH) if data_service.DESCRIPTIONS_XLSX_PATH else "No configurado"
        self.description_value = add_setting_row(layout, "Archivo de descripciones", description_path)
        self.description_limit_value = add_setting_row(layout, "Límite por descripción", f"{data_service.DESCRIPTION_MAX_CHARS} caracteres")
        self.ignore_domains_value = add_setting_row(layout, "Dominios internos ignorados", data_service.IGNORE_DOMAINS)
        self.ignore_senders_value = add_setting_row(layout, "Remitentes ignorados", data_service.IGNORE_SENDERS)
        self.plenergy_value = add_setting_row(layout, "Dominios Plenergy / Plainco", data_service.PLENERGY_SENDER_DOMAINS)
        self.report_value = add_setting_row(layout, "Informe generado", str(data_service.REPORT_XLSX_PATH))
        self.scheduler_value = add_setting_row(layout, "Tarea programada", "Consultando...")
        outer.addWidget(panel)

        actions = QHBoxLayout()
        open_output = QPushButton("Abrir carpeta de salida")
        open_output.setObjectName("PrimaryButton")
        open_output.clicked.connect(self._open_output)
        open_env = QPushButton("Abrir archivo .env")
        open_env.setObjectName("SecondaryButton")
        open_env.clicked.connect(self._open_env)
        edit_prompts = QPushButton("Editar prompts de IA")
        edit_prompts.setObjectName("SecondaryButton")
        edit_prompts.clicked.connect(self._edit_prompts)
        actions.addWidget(open_output)
        actions.addWidget(open_env)
        actions.addWidget(edit_prompts)
        actions.addStretch()
        outer.addLayout(actions)
        outer.addStretch()

    def set_scheduler_status(self, status: dict[str, Any] | None) -> None:
        if status and status.get("Mode") == "VIEWER":
            self.scheduler_value.setText("Desactivada en modo visor")
            return
        if not status:
            self.scheduler_value.setText("No encontrada o no disponible")
            return
        result = status.get("LastTaskResult")
        result_text = "OK" if result == 0 else f"Código {result}"
        self.scheduler_value.setText(
            f"Última ejecución: {status.get('LastRunTime', '—')} | Resultado: {result_text} | "
            f"Próxima: {status.get('NextRunTime', '—')}"
        )

    def _open_output(self) -> None:
        ok, message = data_service.open_path(data_service.OUTPUT_ROOT)
        if not ok:
            QMessageBox.warning(self, "No se pudo abrir", message)

    def _open_env(self) -> None:
        ok, message = data_service.open_path(data_service.ENV_PATH)
        if not ok:
            QMessageBox.warning(self, "No se pudo abrir", message)

    def _edit_prompts(self) -> None:
        from desktop_app.prompt_editor import PromptEditorDialog

        dialog = PromptEditorDialog(self, language="es")
        dialog.exec()


class NewFoldersPage(QWidget):
    """Página para dar de alta un proyecto/empresa o subcarpeta nueva en
    Folder_Data.xlsx (DESCRIPTIONS_XLSX_PATH), el archivo de referencia
    que project_descriptions.py lee para dar contexto a la IA en cada
    clasificación. Los cambios se aplican de inmediato, sin reiniciar
    la aplicación."""

    def __init__(self) -> None:
        super().__init__()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 24)
        outer.setSpacing(14)

        title = QLabel("Carpetas nuevas")
        title.setObjectName("PageTitle")
        subtitle = QLabel(
            "Añada una fila cuando se cree un proyecto, empresa o subcarpeta nueva, para que la "
            "IA la reconozca en el próximo correo procesado."
        )
        subtitle.setObjectName("PageSubtitle")
        subtitle.setWordWrap(True)
        outer.addWidget(title)
        outer.addWidget(subtitle)

        file_panel = QFrame()
        file_panel.setObjectName("Panel")
        file_layout = QVBoxLayout(file_panel)
        file_layout.setContentsMargins(22, 20, 22, 20)
        file_layout.setSpacing(12)

        header = QHBoxLayout()
        icon = QLabel("XLSX")
        icon.setObjectName("MetricIcon")
        icon.setProperty("tone", "green")
        icon.setFixedSize(58, 42)
        icon.setAlignment(Qt.AlignCenter)
        title_block = QVBoxLayout()
        file_name = data_service.DESCRIPTIONS_XLSX_PATH.name if data_service.DESCRIPTIONS_XLSX_PATH else "No configurado"
        file_label = QLabel(file_name)
        file_label.setObjectName("SectionTitle")
        title_block.addWidget(file_label)
        header.addWidget(icon)
        header.addLayout(title_block, 1)
        file_layout.addLayout(header)

        path_text = str(data_service.DESCRIPTIONS_XLSX_PATH) if data_service.DESCRIPTIONS_XLSX_PATH else "No configurado (DESCRIPTIONS_XLSX_PATH)"
        add_setting_row(file_layout, "Ubicación", path_text)
        outer.addWidget(file_panel)

        open_actions = QHBoxLayout()
        open_button = QPushButton("Abrir archivo Excel")
        open_button.setObjectName("SecondaryButton")
        open_button.setEnabled(bool(data_service.DESCRIPTIONS_XLSX_PATH))
        open_button.clicked.connect(self._open_excel)
        open_actions.addWidget(open_button)
        open_actions.addStretch()
        outer.addLayout(open_actions)

        form_panel = QFrame()
        form_panel.setObjectName("Panel")
        form_outer = QVBoxLayout(form_panel)
        form_outer.setContentsMargins(22, 20, 22, 20)
        form_outer.setSpacing(12)

        form_title = QLabel("Nueva fila")
        form_title.setObjectName("SectionTitle")
        form_outer.addWidget(form_title)

        form = QFormLayout()
        form.setSpacing(10)
        form.setLabelAlignment(Qt.AlignLeft)

        self.project_id_input = QLineEdit()
        self.project_id_input.setPlaceholderText("26-002-02")
        form.addRow("ID proyecto:", self.project_id_input)

        self.company_input = QLineEdit()
        self.company_input.setPlaceholderText("CONSUM")
        form.addRow("Empresa \\ Proyecto:", self.company_input)

        self.address_input = QLineEdit()
        self.address_input.setPlaceholderText("C DEL FORN VICENTE ROCA CERVERA, 33 XIRIVELLA")
        form.addRow("Dirección:", self.address_input)

        self.year_input = QSpinBox()
        self.year_input.setRange(2000, 2100)
        self.year_input.setValue(data_service.CURRENT_YEAR)
        self.year_input.setGroupSeparatorShown(False)
        form.addRow("Año:", self.year_input)

        self.company_description_input = QPlainTextEdit()
        self.company_description_input.setPlaceholderText("Descripción de la empresa/proyecto (opcional pero recomendado)")
        self.company_description_input.setFixedHeight(70)
        form.addRow("Empresa \\ Proyecto Descripción:", self.company_description_input)

        self.subfolder_description_input = QPlainTextEdit()
        self.subfolder_description_input.setPlaceholderText("Descripción de la subcarpeta, si la hay (opcional)")
        self.subfolder_description_input.setFixedHeight(70)
        form.addRow("Subfolder Descripción:", self.subfolder_description_input)

        form_outer.addLayout(form)
        outer.addWidget(form_panel)

        self.status_label = QLabel("")
        self.status_label.setObjectName("MutedText")
        self.status_label.setWordWrap(True)
        outer.addWidget(self.status_label)

        buttons = QHBoxLayout()
        clear_button = QPushButton("Limpiar campos")
        clear_button.setObjectName("SecondaryButton")
        clear_button.clicked.connect(self._clear)
        save_button = QPushButton("Guardar fila")
        save_button.setObjectName("PrimaryButton")
        save_button.clicked.connect(self._save)
        buttons.addWidget(clear_button)
        buttons.addStretch()
        buttons.addWidget(save_button)
        outer.addLayout(buttons)
        outer.addStretch()

    def _open_excel(self) -> None:
        ok, message = data_service.open_path(data_service.DESCRIPTIONS_XLSX_PATH)
        if not ok:
            QMessageBox.warning(self, "No se pudo abrir", message)

    def _clear(self) -> None:
        self.project_id_input.clear()
        self.company_input.clear()
        self.address_input.clear()
        self.year_input.setValue(data_service.CURRENT_YEAR)
        self.company_description_input.clear()
        self.subfolder_description_input.clear()
        self.status_label.setText("")

    def _save(self) -> None:
        project_id = self.project_id_input.text().strip()
        company = self.company_input.text().strip()
        address = self.address_input.text().strip()
        year = self.year_input.value()
        company_description = self.company_description_input.toPlainText().strip()
        subfolder_description = self.subfolder_description_input.toPlainText().strip()

        if not project_id or not company or not address:
            QMessageBox.warning(
                self,
                "Faltan datos",
                "Complete al menos ID proyecto, Empresa \\ Proyecto y Dirección antes de guardar.",
            )
            return

        if not re.match(r"^\d{2}-\d+(-\d+)?$", project_id):
            confirm = QMessageBox.question(
                self,
                "Formato de ID inusual",
                f'«{project_id}» no tiene el formato habitual (ej. 26-002-02). ¿Guardar de todas formas?',
            )
            if confirm != QMessageBox.Yes:
                return

        ok, message = data_service.save_folder_data_row(
            project_id, company, address, year, company_description, subfolder_description
        )
        if not ok:
            QMessageBox.warning(self, "No se pudo guardar", message)
            self.status_label.setText("")
            return

        self.status_label.setText(message.split(" / ")[0])
        QMessageBox.information(self, "Guardado", message.split(" / ")[0])
        self._clear()


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        mode_suffix = " - Visor" if data_service.IS_VIEWER else " - Procesamiento"
        self.setWindowTitle(f"INGEVIA Email Assistant{mode_suffix}")
        self.resize(1420, 860)
        self.setMinimumSize(1120, 700)
        self.pipeline_worker: PipelineWorker | None = None
        self.outlook_worker: OutlookFlagWorker | None = None
        self.local_worker: LocalDataWorker | None = None
        self.scheduler_worker: SchedulerStatusWorker | None = None
        icon_path = Path(__file__).resolve().parent.parent / "assets" / "app.ico"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        root = QWidget()
        root.setObjectName("AppRoot")
        self.setCentralWidget(root)
        root_layout = QHBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        sidebar = self._build_sidebar()
        root_layout.addWidget(sidebar)

        main = QVBoxLayout()
        main.setContentsMargins(0, 0, 0, 0)
        main.setSpacing(0)
        main.addWidget(self._build_topbar())

        self.stack = QStackedWidget()
        self.dashboard = DashboardPage()
        self.emails = EmailsPage()
        self.flagged = FlaggedPage()
        self.pending_page = PendingPage()
        self.report_page = ReportPage()
        self.excel_page = ExcelReportPage()
        self.attachments = AttachmentsPage()
        self.settings_page = SettingsPage()
        self.new_folders_page = NewFoldersPage()
        for page in (
            self.dashboard,
            self.emails,
            self.flagged,
            self.pending_page,
            self.report_page,
            self.excel_page,
            self.attachments,
            self.settings_page,
            self.new_folders_page,
        ):
            self.stack.addWidget(page)
        main.addWidget(self.stack, 1)
        root_layout.addLayout(main, 1)

        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("Listo")

        if data_service.IS_PROCESSOR:
            self.dashboard.process_requested.connect(self.run_pipeline)
        self.dashboard.refresh_requested.connect(self.refresh_all)
        self.dashboard.open_output_requested.connect(self.open_output)
        QTimer.singleShot(120, self.refresh_all)
        # Actualiza automáticamente solo los datos compartidos/locales cada 2
        # minutos para reflejar cuándo la cola IA mueve un correo a su carpeta
        # final. Este temporizador NO consulta Outlook.
        self.local_auto_timer = QTimer(self)
        self.local_auto_timer.setInterval(120000)
        self.local_auto_timer.timeout.connect(self.refresh_local_only)
        self.local_auto_timer.start()

    def _build_sidebar(self) -> QFrame:
        sidebar = QFrame()
        sidebar.setObjectName("Sidebar")
        sidebar.setFixedWidth(238)
        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(16, 22, 16, 18)
        layout.setSpacing(8)

        brand = QLabel("INGEVIA")
        brand.setObjectName("BrandTitle")
        brand_subtitle = QLabel("Email Assistant")
        brand_subtitle.setObjectName("BrandSubtitle")
        layout.addWidget(brand)
        layout.addWidget(brand_subtitle)
        layout.addSpacing(22)

        group = QButtonGroup(self)
        group.setExclusive(True)
        buttons = [
            ("▦  Panel principal", 0),
            ("✉  Correos", 1),
            ("⚑  Marcados en Outlook", 2),
            ("↩  Pendientes de respuesta", 3),
            ("▤  Informe de correos", 4),
            ("▥  Informe Excel", 5),
            ("▣  Adjuntos", 6),
            ("⚙  Configuración", 7),
            ("🗂  Carpetas nuevas", 8),
        ]
        for text, index in buttons:
            button = QPushButton(text)
            button.setObjectName("NavButton")
            button.setCheckable(True)
            button.clicked.connect(lambda checked=False, i=index: self._navigate(i))
            group.addButton(button)
            layout.addWidget(button)
            if index == 0:
                button.setChecked(True)
        layout.addStretch()
        status_text = "●  Visor · datos compartidos" if data_service.IS_VIEWER else "●  Ordenador de procesamiento"
        status = QLabel(status_text)
        status.setObjectName("SidebarStatus")
        status.setToolTip(
            "Este equipo solo consulta los datos compartidos."
            if data_service.IS_VIEWER
            else f"Este equipo procesa el buzón {data_service.TARGET_MAILBOX}."
        )
        layout.addWidget(status)
        version = QLabel("Desktop UI 1.26.9")
        version.setObjectName("SidebarStatus")
        layout.addWidget(version)
        return sidebar

    def _build_topbar(self) -> QFrame:
        topbar = QFrame()
        topbar.setObjectName("TopBar")
        topbar.setFixedHeight(78)
        layout = QHBoxLayout(topbar)
        layout.setContentsMargins(26, 12, 26, 12)
        title_block = QVBoxLayout()
        self.top_title = QLabel("Panel principal")
        self.top_title.setObjectName("PageTitle")
        self.top_subtitle = QLabel("Resumen del procesamiento y del estado de Outlook")
        self.top_subtitle.setObjectName("PageSubtitle")
        title_block.addWidget(self.top_title)
        title_block.addWidget(self.top_subtitle)
        layout.addLayout(title_block)
        layout.addStretch()

        self.language_combo = QComboBox()
        self.language_combo.setObjectName("LanguageCombo")
        self.language_combo.addItem("English", "en")
        self.language_combo.addItem("Español", "es")
        current_language = get_language()
        selected_index = self.language_combo.findData(current_language)
        if selected_index >= 0:
            self.language_combo.setCurrentIndex(selected_index)
        self.language_combo.setToolTip("Cambiar idioma de la aplicación")
        self.language_combo.currentIndexChanged.connect(self._change_language)
        layout.addWidget(self.language_combo)

        initial_status = "Datos compartidos: consultando..." if data_service.IS_VIEWER else "Outlook: consultando..."
        self.outlook_pill = QLabel(initial_status)
        self.outlook_pill.setObjectName("WarningPill")
        layout.addWidget(self.outlook_pill)
        self.top_refresh_button = QPushButton("↻")
        self.top_refresh_button.setObjectName("SecondaryButton")
        self.top_refresh_button.setFixedWidth(44)
        self.top_refresh_button.setToolTip("Actualizar todos los datos")
        self.top_refresh_button.clicked.connect(self.refresh_all)
        layout.addWidget(self.top_refresh_button)
        return topbar

    def _change_language(self) -> None:
        code = self.language_combo.currentData()
        if not code or code == get_language():
            return
        set_language(str(code))
        QMessageBox.information(
            self,
            "Idioma actualizado",
            "La aplicación se reiniciará para aplicar el nuevo idioma.",
        )
        restart_application()

    def _navigate(self, index: int) -> None:
        titles = [
            ("Panel principal", "Resumen del procesamiento y del estado de Outlook"),
            ("Correos", "Archivo de correos organizados por proyecto"),
            ("Marcados en Outlook", "Correos identificados como procesados por el asistente"),
            ("Pendientes de respuesta", "Mensajes que necesitan una respuesta escrita de la empresa"),
            ("Informe de correos", "Resumen generado de los correos archivados"),
            ("Informe Excel", "Archivo Excel generado en la carpeta de salida"),
            ("Adjuntos", "Archivos guardados y bloqueados por seguridad"),
            ("Configuración", "Estado de la aplicación local y sus rutas"),
            ("Carpetas nuevas", "Añadir una fila nueva al archivo de referencia Folder_Data.xlsx"),
        ]
        self.stack.setCurrentIndex(index)
        self.top_title.setText(titles[index][0])
        self.top_subtitle.setText(titles[index][1])
        if index == 7:
            self.refresh_scheduler_status()

    def refresh_all(self) -> None:
        if self.local_worker and self.local_worker.isRunning():
            self.statusBar().showMessage("La actualización ya está en curso", 2500)
            return
        self.statusBar().showMessage("Cargando datos locales en segundo plano...")
        self.dashboard.refresh_button.setDisabled(True)
        self.top_refresh_button.setDisabled(True)
        self.local_worker = LocalDataWorker(self)
        self.local_worker.loaded.connect(self._local_data_loaded)
        self.local_worker.failed.connect(self._local_data_failed)
        self.local_worker.finished.connect(self._local_refresh_finished)
        self.local_worker.start()
        self.refresh_outlook_flags()

    def refresh_local_only(self) -> None:
        if self.local_worker and self.local_worker.isRunning():
            return
        self.local_worker = LocalDataWorker(self)
        self.local_worker.loaded.connect(self._local_data_loaded)
        self.local_worker.failed.connect(self._local_data_failed)
        self.local_worker.finished.connect(self._local_refresh_finished)
        self.local_worker.start()

    def _local_data_loaded(self, snapshot: dict[str, Any]) -> None:
        rows = snapshot.get("rows", [])
        activity_rows = snapshot.get("activity_rows", [])
        attachments = snapshot.get("attachments", [])
        pending_rows = snapshot.get("pending_rows", [])
        report_rows = snapshot.get("report_rows", [])
        processed = int(snapshot.get("processed", 0) or 0)
        self.dashboard.update_local_data(rows, attachments, processed, activity_rows, pending_rows)
        self.emails.set_rows(rows)
        self.pending_page.set_rows(pending_rows)
        self.report_page.set_rows(report_rows)
        self.excel_page.apply_info(snapshot.get("excel_info", {}))
        self.attachments.set_rows(attachments)
        self.statusBar().showMessage("Datos locales actualizados", 4000)

    def _local_data_failed(self, message: str) -> None:
        QMessageBox.critical(self, "Error de actualización", message)
        self.statusBar().showMessage("Error al actualizar los datos locales", 4000)

    def _local_refresh_finished(self) -> None:
        self.dashboard.refresh_button.setDisabled(False)
        self.top_refresh_button.setDisabled(False)

    def refresh_scheduler_status(self) -> None:
        if data_service.IS_VIEWER:
            self.settings_page.set_scheduler_status({"Mode": "VIEWER"})
            return
        if self.scheduler_worker and self.scheduler_worker.isRunning():
            return
        self.settings_page.scheduler_value.setText("Consultando...")
        self.scheduler_worker = SchedulerStatusWorker(self)
        self.scheduler_worker.loaded.connect(self.settings_page.set_scheduler_status)
        self.scheduler_worker.start()

    def refresh_outlook_flags(self) -> None:
        if self.outlook_worker and self.outlook_worker.isRunning():
            return
        self.outlook_pill.setText(
            "Datos compartidos: consultando..." if data_service.IS_VIEWER else "Outlook: consultando..."
        )
        self.outlook_pill.setObjectName("WarningPill")
        self.outlook_pill.style().unpolish(self.outlook_pill)
        self.outlook_pill.style().polish(self.outlook_pill)
        self.outlook_worker = OutlookFlagWorker(self)
        self.outlook_worker.loaded.connect(self._outlook_loaded)
        self.outlook_worker.start()

    def _outlook_loaded(self, summary: data_service.OutlookFlagSummary) -> None:
        self.dashboard.update_flagged(summary)
        self.flagged.set_summary(summary)
        if summary.error:
            self.outlook_pill.setText(
                "Datos compartidos: no disponibles" if data_service.IS_VIEWER else "Outlook: no disponible"
            )
            self.outlook_pill.setObjectName("ErrorPill")
        else:
            self.outlook_pill.setText(
                f"Datos compartidos: conectados · {summary.total} correos"
                if data_service.IS_VIEWER
                else f"Outlook: conectado · {summary.total} marcados"
            )
            self.outlook_pill.setObjectName("SuccessPill")
        self.outlook_pill.style().unpolish(self.outlook_pill)
        self.outlook_pill.style().polish(self.outlook_pill)

    def run_pipeline(self) -> None:
        if data_service.IS_VIEWER:
            QMessageBox.information(
                self,
                "Modo visor",
                "Este equipo solo muestra los datos compartidos. El procesamiento se ejecuta en el ordenador principal.",
            )
            return
        answer = QMessageBox.question(
            self,
            "Procesar correos",
            "Se revisarán únicamente los correos recientes de la bandeja de entrada.\n\n¿Deseas continuar?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if answer != QMessageBox.Yes:
            return
        if self.pipeline_worker and self.pipeline_worker.isRunning():
            return
        self.dashboard.set_processing(True)
        self.statusBar().showMessage("Procesando correos en segundo plano...")
        self.pipeline_worker = PipelineWorker(self)
        self.pipeline_worker.finished_with_result.connect(self._pipeline_finished)
        self.pipeline_worker.start()

    def _pipeline_finished(self, success: bool, message: str) -> None:
        self.dashboard.set_processing(False)
        if success:
            QMessageBox.information(self, "Proceso completado", message[-2500:])
            self.statusBar().showMessage("Procesamiento completado", 5000)
            self.refresh_all()
        else:
            QMessageBox.critical(self, "Error durante el procesamiento", message[-3500:])
            self.statusBar().showMessage("El procesamiento terminó con errores", 5000)

    def open_output(self) -> None:
        ok, message = data_service.open_path(data_service.OUTPUT_ROOT)
        if not ok:
            QMessageBox.warning(self, "No se pudo abrir", message)


def configure_table(table: QTableWidget) -> None:
    table.setAlternatingRowColors(True)
    table.setSelectionBehavior(QAbstractItemView.SelectRows)
    table.setSelectionMode(QAbstractItemView.SingleSelection)
    table.setEditTriggers(QAbstractItemView.NoEditTriggers)
    table.setShowGrid(False)
    table.verticalHeader().setVisible(False)
    table.horizontalHeader().setStretchLastSection(False)
    table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)


def direction_label(value: str) -> str:
    if value == "ENTRANTE":
        return "Entrante"
    return value


def priority_label(value: Any) -> str:
    try:
        priority = int(value)
    except (TypeError, ValueError):
        return "—"
    return {
        5: "5 - Crítica",
        4: "4 - Alta",
        3: "3 - Normal",
        2: "2 - Baja",
        1: "1 - Muy baja",
    }.get(priority, str(priority))


def apply_priority_style(
    item: QTableWidgetItem,
    value: Any,
    *,
    emphasize: bool = False,
) -> None:
    try:
        priority = int(value)
    except (TypeError, ValueError):
        return
    colors = {
        # Required office mapping: 5 red, 4 orange, 3 yellow.
        # The full row gets a light tint; the Priority cell gets the
        # stronger solid color so the level is immediately visible.
        5: ("#f8d7da", "#d32f2f", "#7f1d1d", "#ffffff"),
        4: ("#ffe0b2", "#f57c00", "#7c2d12", "#ffffff"),
        3: ("#fff3cd", "#fbc02d", "#5f4600", "#1f2937"),
    }.get(priority)
    if not colors:
        return
    background, strong_background, foreground, strong_foreground = colors
    item.setBackground(QColor(strong_background if emphasize else background))
    item.setForeground(QColor(strong_foreground if emphasize else foreground))
    if emphasize:
        font = item.font()
        font.setBold(True)
        item.setFont(font)


def add_setting_row(layout: QVBoxLayout, title: str, value: str) -> QLabel:
    container = QFrame()
    row = QHBoxLayout(container)
    row.setContentsMargins(0, 4, 0, 4)
    title_label = QLabel(title)
    title_label.setObjectName("MetricTitle")
    value_label = QLabel(value)
    value_label.setObjectName("MutedText")
    value_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
    value_label.setWordWrap(True)
    row.addWidget(title_label)
    row.addStretch()
    row.addWidget(value_label, 1)
    layout.addWidget(container)
    return value_label