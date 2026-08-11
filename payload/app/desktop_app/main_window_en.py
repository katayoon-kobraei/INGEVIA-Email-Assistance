from __future__ import annotations

import io
import os
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
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSizePolicy,
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
        stream = io.StringIO()
        pythoncom = None
        try:
            if os.name == "nt":
                import pythoncom as _pythoncom

                pythoncom = _pythoncom
                pythoncom.CoInitialize()
            with redirect_stdout(stream), redirect_stderr(stream):
                from src.pipeline import run as run_pipeline

                run_pipeline()
            output = stream.getvalue().strip() or "Process completed successfully."
            self.finished_with_result.emit(True, output)
        except Exception as exc:
            log = stream.getvalue().strip()
            message = f"{exc}"
            if log:
                message = f"{message}\n\nDetails:\n{log}"
            self.finished_with_result.emit(False, message)
        finally:
            if pythoncom is not None:
                pythoncom.CoUninitialize()


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
    """Dashboard metric with a sliding information panel on hover."""

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
        self._showing_details = False

        self.front_page = QWidget(self)
        self.front_page.setObjectName("MetricSlidePage")
        front_layout = QVBoxLayout(self.front_page)
        front_layout.setContentsMargins(18, 16, 18, 16)
        front_layout.setSpacing(7)
        top = QHBoxLayout()
        self.icon = QLabel(icon)
        self.icon.setObjectName("MetricIcon")
        self.icon.setProperty("tone", tone)
        self.icon.setFixedSize(38, 38)
        self.icon.setAlignment(Qt.AlignCenter)
        top.addWidget(self.icon)
        top.addStretch()
        front_layout.addLayout(top)
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
        header = QHBoxLayout()
        details_icon = QLabel(icon)
        details_icon.setObjectName("MetricIcon")
        details_icon.setProperty("tone", tone)
        details_icon.setFixedSize(34, 34)
        details_icon.setAlignment(Qt.AlignCenter)
        details_title = QLabel(title)
        details_title.setObjectName("MetricDetailTitle")
        details_title.setWordWrap(True)
        header.addWidget(details_icon)
        header.addWidget(details_title, 1)
        details_layout.addLayout(header)
        self.details_text = QLabel(details or hint)
        self.details_text.setObjectName("MetricDetailText")
        self.details_text.setWordWrap(True)
        details_layout.addWidget(self.details_text)
        self.details_status = QLabel(hint)
        self.details_status.setObjectName("MetricDetailStatus")
        self.details_status.setWordWrap(True)
        details_layout.addWidget(self.details_status)
        details_layout.addStretch()
        footer = QLabel("Move the pointer away to return to the counter")
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
        super().resizeEvent(event)
        width, height = self.width(), self.height()
        self.front_page.resize(width, height)
        self.details_page.resize(width, height)
        if self._showing_details:
            self.front_page.move(-width, 0)
            self.details_page.move(0, 0)
        else:
            self.front_page.move(0, 0)
            self.details_page.move(width, 0)

    def _slide(self, show_details: bool) -> None:
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
        self.value.setText(str(value))
        if hint is not None:
            self.hint.setText(hint)
            self.details_status.setText(f"Current status: {hint}")


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
        title = QLabel("Email details")
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
            ("Date", "DATE"),
            ("Direction", "DIRECTION"),
            ("Project Folder", "PROJECT"),
            ("Address Folder", "LOCATION / ADDRESS"),
            ("Contact", "CONTACT"),
            ("Topic", "TOPIC"),
            ("Subject", "SUBJECT"),
            ("Attachments", "ATTACHMENTS"),
            ("Priority", "PRIORITY"),
            ("_pending", "AWAITING RESPONSE"),
            ("Summary", "SUMMARY"),
            ("_status", "STATUS"),
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
        self.open_button = QPushButton("Open folder")
        self.open_button.setObjectName("PrimaryButton")
        self.open_button.clicked.connect(self._open_folder)
        self.copy_button = QPushButton("Copy path")
        self.copy_button.setObjectName("SecondaryButton")
        self.copy_button.clicked.connect(self._copy_path)
        layout.addWidget(self.open_button)
        layout.addWidget(self.copy_button)

    def show_record(self, record: dict[str, Any]) -> None:
        self._record = record
        for key, label in self.values.items():
            value = record.get(key, "") or "—"
            if key == "Direction":
                value = direction_label(str(value))
            elif key == "_status":
                value = "Review" if value == "REVISAR" else "Processed"
            label.setText(str(value))
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
            QMessageBox.warning(self, "Could not open", message)

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
        title = QLabel("Automatic email management")
        title.setObjectName("SectionTitle")
        subtitle = QLabel(
            "Classify Outlook emails, save attachments, and organize documentation by project."
        )
        subtitle.setObjectName("MutedText")
        subtitle.setWordWrap(True)
        hero_text.addWidget(title)
        hero_text.addWidget(subtitle)
        hero_layout.addLayout(hero_text, 1)

        self.process_button = QPushButton("▶  Process new emails")
        self.process_button.setObjectName("ProcessButton")
        self.process_button.clicked.connect(self.process_requested)
        if data_service.IS_VIEWER:
            title.setText("Shared email view")
            subtitle.setText(
                "View the configured mailbox emails processed by the main computer."
            )
            self.process_button.hide()
        self.refresh_button = QPushButton("↻  Refresh")
        self.refresh_button.setObjectName("SecondaryButton")
        self.refresh_button.clicked.connect(self.refresh_requested)
        output_button = QPushButton("▣  Open files")
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
            "Processed emails",
            "Total unique emails registered by the system",
            "green",
            details=(
                "Counts messages already handled by the assistant. They will not "
                "be processed again during future runs."
            ),
        )
        self.flagged_card = MetricCard(
            "⚑",
            "Flagged in Outlook",
            "Messages stamped with the assistant category",
            "orange",
            details=(
                "Shows Outlook messages carrying the configured assistant category. "
                "Open the Flagged in Outlook page to inspect them."
            ),
        )
        self.filed_card = MetricCard(
            "▤",
            "Filed emails",
            "Messages saved in project folders",
            "blue",
            details=(
                "Counts messages classified and stored in the corresponding project "
                "folder together with their accepted attachments."
            ),
        )
        self.review_card = MetricCard(
            "!",
            "Need review",
            "Messages without a confident classification",
            "red",
            details=(
                "Shows messages that could not be assigned confidently and should "
                "be reviewed manually."
            ),
        )
        self.pending_card = MetricCard(
            "↩",
            "Awaiting response",
            "Incoming messages that need a company reply",
            "purple",
            details=(
                "Shows relevant incoming messages that the AI agent identified as "
                "still waiting for a written response from the company."
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
        activity_title = QLabel("Recent activity")
        activity_title.setObjectName("SectionTitle")
        activity_layout.addWidget(activity_title)
        self.recent_table = QTableWidget(0, 6)
        self.recent_table.setHorizontalHeaderLabels(
            ["Date", "Direction", "Project", "Subject", "Priority", "Status"]
        )
        configure_table(self.recent_table)
        self.recent_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        activity_layout.addWidget(self.recent_table)
        content.addWidget(activity_panel, 3)

        attention_panel = QFrame()
        attention_panel.setObjectName("AttentionPanel")
        attention_panel.setMinimumWidth(280)
        attention_layout = QVBoxLayout(attention_panel)
        attention_layout.setContentsMargins(18, 16, 18, 16)
        attention_title = QLabel("Attention required")
        attention_title.setObjectName("SectionTitle")
        self.review_summary = QLabel("Loading...")
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
        self.outlook_summary = QLabel("Checking Outlook...")
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
        self.process_button.setText("Processing..." if active else "▶  Process new emails")
        self.progress.setVisible(active)

    def update_local_data(
        self,
        rows: list[dict[str, Any]],
        attachments: list[dict[str, Any]],
        processed: int,
        activity_rows: list[dict[str, Any]] | None = None,
        pending_rows: list[dict[str, Any]] | None = None,
    ) -> None:
        review = [row for row in rows if row.get("Project Folder") == "UNSORTED"]
        quarantined = [row for row in attachments if row.get("Status") == "quarantined"]
        self.processed_card.set_value(processed)
        self.filed_card.set_value(len(rows))
        self.review_card.set_value(len(review))
        pending_total = len(pending_rows or [])
        high_priority = [row for row in rows if int(row.get("_priority") or 0) >= 4]
        critical_priority = [row for row in rows if int(row.get("_priority") or 0) == 5]
        self.pending_card.set_value(pending_total)
        self.review_summary.setText(
            "✓ No emails require classification review."
            if not review
            else f"⚠ {len(review)} email(s) are in UNSORTED and need manual review."
        )
        self.quarantine_summary.setText(
            "✓ No attachments are blocked."
            if not quarantined
            else f"⚠ {len(quarantined)} attachment(s) were blocked for security."
        )
        self.priority_summary.setText(
            "✓ No high or critical priority emails."
            if not high_priority
            else f"⚠ {len(high_priority)} email(s) have priority 4 or 5; {len(critical_priority)} are critical."
        )
        self.pending_summary.setText(
            "✓ No emails are awaiting a response."
            if pending_total == 0
            else f"↩ {pending_total} email(s) need a company response."
        )

        source_rows = activity_rows or rows
        status_labels = {
            "ARCHIVADO": "Filed",
            "OMITIDO": "Skipped",
            "REVISAR": "Review",
            "ERROR": "Error",
            "PENDIENTE": "Awaiting response",
            "PROCESADO": "Processed",
        }
        self.recent_table.setRowCount(0)
        for record in source_rows[:12]:
            row = self.recent_table.rowCount()
            self.recent_table.insertRow(row)
            raw_status = str(record.get("Status") or record.get("_status") or "PROCESADO")
            status_text = status_labels.get(raw_status, raw_status.title())
            values = [
                record.get("Date", ""),
                direction_label(record.get("Direction", "")),
                record.get("Project Folder", "") or "—",
                record.get("Subject", ""),
                priority_label(record.get("_priority")),
                status_text,
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                apply_priority_style(item, record.get("_priority"), emphasize=column == 4)
                if column == 5:
                    color = {
                        "ERROR": "#b42318",
                        "REVISAR": "#b26a00",
                        "OMITIDO": "#52667a",
                        "PENDIENTE": "#6d28d9",
                    }.get(raw_status, "#087443")
                    item.setForeground(QColor(color))
                self.recent_table.setItem(row, column, item)

    def update_flagged(self, summary: data_service.OutlookFlagSummary) -> None:
        if summary.error:
            self.flagged_card.set_value("—", "Could not query Outlook")
            self.outlook_summary.setText(f"⚠ Outlook: {summary.error}")
            return
        hint = "Estimated count" if summary.estimated else "Direct count from Outlook"
        self.flagged_card.set_value(summary.total, hint)
        suffix = " (estimated)" if summary.estimated else ""
        self.outlook_summary.setText(
            f"✓ {summary.total} email(s) have the category “{data_service.PROCESSED_CATEGORY_NAME}”{suffix}."
        )


class EmailsPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.all_rows: list[dict[str, Any]] = []

        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 24)
        outer.setSpacing(14)

        heading = QLabel("Filed emails")
        heading.setObjectName("PageTitle")
        subtitle = QLabel(
            "Search, filter, and review the summary of every processed email."
        )
        subtitle.setObjectName("PageSubtitle")
        priority_legend = QLabel(
            "Priority: 5 critical (red) · 4 high (orange) · 3 normal (yellow)"
        )
        priority_legend.setObjectName("MutedText")
        outer.addWidget(heading)
        outer.addWidget(subtitle)
        outer.addWidget(priority_legend)

        filters = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText(
            "Search by subject, summary, contact, sender, or project..."
        )
        self.project_filter = QComboBox()
        self.project_filter.addItem("All projects")
        self.direction_filter = QComboBox()
        self.direction_filter.addItems(
            ["All directions", "Incoming"]
        )
        self.status_filter = QComboBox()
        self.status_filter.addItems(
            ["All statuses", "Processed", "Review"]
        )
        filters.addWidget(self.search, 2)
        filters.addWidget(self.project_filter, 1)
        filters.addWidget(self.direction_filter, 1)
        filters.addWidget(self.status_filter, 1)
        outer.addLayout(filters)

        # One-line summary generated by the same backend flow that feeds
        # the Excel file "Informe de Emails.xlsx".
        summary_panel = QFrame()
        summary_panel.setObjectName("EmailSummaryPanel")
        summary_layout = QVBoxLayout(summary_panel)
        summary_layout.setContentsMargins(16, 12, 16, 12)
        summary_layout.setSpacing(4)

        summary_title = QLabel("Selected email summary")
        summary_title.setObjectName("EmailSummaryTitle")
        self.email_summary = QLabel(
            "Select an email to view its one-line summary."
        )
        self.email_summary.setObjectName("EmailSummaryText")
        self.email_summary.setWordWrap(True)
        self.email_summary.setTextInteractionFlags(Qt.TextSelectableByMouse)
        summary_layout.addWidget(summary_title)
        summary_layout.addWidget(self.email_summary)
        outer.addWidget(summary_panel)

        self.table = QTableWidget(0, 10)
        self.table.setHorizontalHeaderLabels(
            [
                "Date",
                "Direction",
                "Project",
                "Contact",
                "Sender",
                "Subject",
                "Summary",
                "Attachments",
                "Priority",
                "Status",
            ]
        )
        configure_table(self.table)
        self.table.setWordWrap(False)
        self.table.setColumnWidth(4, 280)
        self.table.horizontalHeader().setSectionResizeMode(
            5, QHeaderView.Stretch
        )
        self.table.horizontalHeader().setSectionResizeMode(
            6, QHeaderView.Stretch
        )
        # The table now uses the full page width. Email details are intentionally
        # omitted; selecting a row only updates the one-line summary above.
        outer.addWidget(self.table, 1)

        self.search.textChanged.connect(self.apply_filters)
        self.project_filter.currentTextChanged.connect(self.apply_filters)
        self.direction_filter.currentTextChanged.connect(self.apply_filters)
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
        self.project_filter.addItem("All projects")
        self.project_filter.addItems(projects)
        if current in projects:
            self.project_filter.setCurrentText(current)
        self.project_filter.blockSignals(False)
        self.apply_filters()

    def apply_filters(self) -> None:
        query = self.search.text().strip().lower()
        project = self.project_filter.currentText()
        direction = self.direction_filter.currentText()
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
                    "Topic",
                    "Sender/Recipient",
                    "Priority",
                )
            ).lower()
            if query and query not in haystack:
                continue
            if project != "All projects" and record.get("Project Folder") != project:
                continue
            if direction == "Incoming" and record.get("Direction") != "ENTRANTE":
                continue
            if status == "Processed" and record.get("_status") != "PROCESADO":
                continue
            if status == "Review" and record.get("_status") != "REVISAR":
                continue
            filtered.append(record)

        visible = filtered[:DISPLAY_ROW_LIMIT]
        self.table.setUpdatesEnabled(False)
        self.table.setRowCount(len(visible))
        for row, record in enumerate(visible):
            summary = str(record.get("Summary") or "").strip()
            values = [
                record.get("Date", ""),
                direction_label(record.get("Direction", "")),
                record.get("Project Folder", ""),
                record.get("Contact", ""),
                record.get("SenderDisplay") or data_service.sender_display(record),
                record.get("Subject", ""),
                summary or "No summary available",
                record.get("Attachments", "0"),
                priority_label(record.get("_priority")),
                "Review" if record.get("_status") == "REVISAR" else "Processed",
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if column == 0:
                    item.setData(Qt.UserRole, record)
                if column == 4:
                    item.setToolTip(str(record.get("SenderDisplay") or data_service.sender_display(record)))
                if column == 6:
                    item.setToolTip(summary or "No summary available")
                apply_priority_style(item, record.get("_priority"), emphasize=column == 8)
                if column == 9:
                    item.setForeground(QColor("#b26a00" if record.get("_status") == "REVISAR" else "#087443"))
                self.table.setItem(row, column, item)
        self.table.setUpdatesEnabled(True)

        if visible:
            # Do not auto-select the first row: Windows selection colors hide
            # the red/orange/yellow priority background until selection moves.
            self.table.clearSelection()
        else:
            self.email_summary.setText("No emails match the selected filters.")

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
                "Select an email to view its one-line summary."
            )
            return

        summary = str(record.get("Summary") or "").strip()
        self.email_summary.setText(
            summary
            or "No summary is available for this email."
        )

    def _open_selected(self, _row: int, _column: int) -> None:
        record = self._selected_record()
        if not record:
            return
        ok, message = data_service.open_path(record.get("Folder Path", ""))
        if not ok:
            QMessageBox.warning(self, "Could not open", message)


class FlaggedPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 24)
        outer.setSpacing(14)

        heading_line = QHBoxLayout()
        text = QVBoxLayout()
        title = QLabel("Emails flagged in Outlook")
        title.setObjectName("PageTitle")
        subtitle = QLabel(
            f"Emails with category “{data_service.PROCESSED_CATEGORY_NAME}” and follow-up flag."
        )
        subtitle.setObjectName("PageSubtitle")
        text.addWidget(title)
        text.addWidget(subtitle)
        heading_line.addLayout(text)
        heading_line.addStretch()
        self.count_label = QLabel("Checking Outlook...")
        self.count_label.setObjectName("MutedText")
        heading_line.addWidget(self.count_label)
        outer.addLayout(heading_line)

        self.search = QLineEdit()
        self.search.setPlaceholderText("Search flagged emails...")
        outer.addWidget(self.search)

        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["Date", "Direction", "Contact", "Subject", "Category"])
        configure_table(self.table)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        outer.addWidget(self.table, 1)
        self._rows: list[dict[str, Any]] = []
        self.search.textChanged.connect(self._render)

    def set_summary(self, summary: data_service.OutlookFlagSummary) -> None:
        if summary.error:
            self.count_label.setText(f"Unavailable: {summary.error}")
            self._rows = []
        else:
            suffix = " (estimated)" if summary.estimated else ""
            self.count_label.setText(f"{summary.total} flagged{suffix}; showing up to {len(summary.rows)}")
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
            for column, key in enumerate(("Date", "Direction", "Contact", "Subject", "Category")):
                self.table.setItem(row, column, QTableWidgetItem(str(record.get(key, ""))))


class PendingPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self._rows: list[dict[str, Any]] = []
        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 24)
        outer.setSpacing(14)
        heading = QLabel("Awaiting response")
        heading.setObjectName("PageTitle")
        subtitle = QLabel(
            "Incoming messages identified as still waiting for a written company response."
        )
        subtitle.setObjectName("PageSubtitle")
        subtitle.setWordWrap(True)
        outer.addWidget(heading)
        outer.addWidget(subtitle)
        controls = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search by sender or subject...")
        self.search.textChanged.connect(self._render)
        open_list = QPushButton("Open CSV list")
        open_list.setObjectName("SecondaryButton")
        open_list.clicked.connect(self._open_list)
        open_output = QPushButton("Open output folder")
        open_output.setObjectName("PrimaryButton")
        open_output.clicked.connect(self._open_output)
        controls.addWidget(self.search, 1)
        controls.addWidget(open_list)
        controls.addWidget(open_output)
        outer.addLayout(controls)
        self.summary = QLabel("0 pending emails")
        self.summary.setObjectName("MutedText")
        outer.addWidget(self.summary)
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Date", "Sender", "Subject", "Priority"])
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
                for key in ("Date", "Sender", "Subject")
            ).lower()
            if query and query not in text:
                continue
            filtered.append(record)

        visible = filtered[:DISPLAY_ROW_LIMIT]
        self.table.setUpdatesEnabled(False)
        self.table.setRowCount(len(visible))
        for row, record in enumerate(visible):
            values = [
                record.get("Date", ""),
                record.get("Sender", ""),
                record.get("Subject", ""),
                priority_label(record.get("_priority")),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                apply_priority_style(item, record.get("_priority"), emphasize=column == 3)
                self.table.setItem(row, column, item)
        self.table.setUpdatesEnabled(True)
        suffix = "" if len(filtered) <= DISPLAY_ROW_LIMIT else f"; showing {len(visible)}"
        self.summary.setText(f"{len(filtered)} email(s) awaiting response{suffix}")

    def _open_list(self) -> None:
        ok, message = data_service.open_path(data_service.PENDING_LIST_PATH)
        if not ok:
            QMessageBox.warning(self, "Could not open", message)

    def _open_output(self) -> None:
        ok, message = data_service.open_path(data_service.OUTPUT_ROOT)
        if not ok:
            QMessageBox.warning(self, "Could not open", message)


class ReportPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self._rows: list[dict[str, Any]] = []

        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 24)
        outer.setSpacing(14)

        heading = QLabel("Email report")
        heading.setObjectName("PageTitle")
        subtitle = QLabel(
            "Human-readable summary of filed emails generated by the backend after each run."
        )
        subtitle.setObjectName("PageSubtitle")
        subtitle.setWordWrap(True)
        outer.addWidget(heading)
        outer.addWidget(subtitle)

        controls = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search by sender, subject, or summary...")
        self.search.textChanged.connect(self._render)
        open_excel = QPushButton("Open Excel report")
        open_excel.setObjectName("PrimaryButton")
        open_excel.clicked.connect(self._open_excel)
        open_csv = QPushButton("Open CSV log")
        open_csv.setObjectName("SecondaryButton")
        open_csv.clicked.connect(self._open_csv)
        controls.addWidget(self.search, 1)
        controls.addWidget(open_csv)
        controls.addWidget(open_excel)
        outer.addLayout(controls)

        self.summary = QLabel("0 report entries")
        self.summary.setObjectName("MutedText")
        outer.addWidget(self.summary)

        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(
            ["Date", "Sender", "Email", "Subject", "Priority", "Summary"]
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
                for key in ("SenderName", "SenderEmail", "Subject", "Summary")
            ).lower()
            if query and query not in text:
                continue
            filtered.append(record)

        visible = filtered[:DISPLAY_ROW_LIMIT]
        self.table.setUpdatesEnabled(False)
        self.table.setRowCount(len(visible))
        for row, record in enumerate(visible):
            values = [
                record.get("DateTime", ""),
                record.get("SenderName", ""),
                record.get("SenderEmail", ""),
                record.get("Subject", ""),
                priority_label(record.get("_priority")),
                record.get("Summary", ""),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if column == 0:
                    item.setData(Qt.UserRole, record)
                apply_priority_style(item, record.get("_priority"), emphasize=column == 4)
                self.table.setItem(row, column, item)
        self.table.setUpdatesEnabled(True)
        suffix = "" if len(filtered) <= DISPLAY_ROW_LIMIT else f"; showing {len(visible)}"
        self.summary.setText(f"{len(filtered)} report entries{suffix}")

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
        ok, message = data_service.open_path(record.get("Path", ""))
        if not ok:
            QMessageBox.warning(self, "Could not open", message)

    def _open_excel(self) -> None:
        ok, message = data_service.open_path(data_service.REPORT_XLSX_PATH)
        if not ok:
            QMessageBox.warning(self, "Could not open", message)

    def _open_csv(self) -> None:
        ok, message = data_service.open_path(data_service.REPORT_LOG_PATH)
        if not ok:
            QMessageBox.warning(self, "Could not open", message)


class ExcelReportPage(QWidget):
    """Dedicated page for the generated Excel report in the output folder."""

    def __init__(self) -> None:
        super().__init__()

        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 24)
        outer.setSpacing(14)

        heading = QLabel("Excel report")
        heading.setObjectName("PageTitle")
        subtitle = QLabel(
            "Direct access to the generated “Informe de Emails.xlsx” file in the output folder."
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
        self.status_label = QLabel("Checking the file...")
        self.status_label.setObjectName("WarningPill")
        self.status_label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        title_block.addWidget(self.file_name)
        title_block.addWidget(self.status_label, 0, Qt.AlignLeft)

        header.addWidget(icon)
        header.addLayout(title_block, 1)
        panel_layout.addLayout(header)

        self.path_value = add_setting_row(
            panel_layout,
            "Location",
            str(data_service.REPORT_XLSX_PATH),
        )
        self.size_value = add_setting_row(panel_layout, "Size", "—")
        self.modified_value = add_setting_row(panel_layout, "Last updated", "—")
        self.entries_value = add_setting_row(panel_layout, "Emails included", "0")
        outer.addWidget(panel)

        actions = QHBoxLayout()
        self.open_button = QPushButton("Open Excel report")
        self.open_button.setObjectName("PrimaryButton")
        self.open_button.clicked.connect(self._open_excel)

        open_folder = QPushButton("Open output folder")
        open_folder.setObjectName("SecondaryButton")
        open_folder.clicked.connect(self._open_output)

        refresh = QPushButton("↻  Refresh status")
        refresh.setObjectName("SecondaryButton")
        refresh.clicked.connect(self.refresh_status)

        actions.addWidget(self.open_button)
        actions.addWidget(open_folder)
        actions.addWidget(refresh)
        actions.addStretch()
        outer.addLayout(actions)

        note = QLabel(
            "The report is regenerated after processing. If it does not exist yet, "
            "process at least one relevant email and refresh this page."
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
            self.status_label.setText("Available")
            self.status_label.setObjectName("SuccessPill")
        else:
            self.status_label.setText("Not generated yet")
            self.status_label.setObjectName("WarningPill")
        self.status_label.style().unpolish(self.status_label)
        self.status_label.style().polish(self.status_label)

    def _open_excel(self) -> None:
        ok, message = data_service.open_path(data_service.REPORT_XLSX_PATH)
        if not ok:
            QMessageBox.warning(self, "Could not open", message)

    def _open_output(self) -> None:
        ok, message = data_service.open_path(data_service.OUTPUT_ROOT)
        if not ok:
            QMessageBox.warning(self, "Could not open", message)


class AttachmentsPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.rows: list[dict[str, Any]] = []
        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 24)
        outer.setSpacing(14)

        title = QLabel("Attachments")
        title.setObjectName("PageTitle")
        subtitle = QLabel("Review saved files and attachments blocked for security.")
        subtitle.setObjectName("PageSubtitle")
        outer.addWidget(title)
        outer.addWidget(subtitle)

        filters = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search by file, subject, or project...")
        self.status_filter = QComboBox()
        self.status_filter.addItems(["All", "Saved", "Blocked"])
        filters.addWidget(self.search, 1)
        filters.addWidget(self.status_filter)
        outer.addLayout(filters)

        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["Date", "Project", "Subject", "File", "Status", "Reason"])
        configure_table(self.table)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        outer.addWidget(self.table, 1)

        buttons = QHBoxLayout()
        buttons.addStretch()
        open_folder = QPushButton("Open folder")
        open_folder.setObjectName("SecondaryButton")
        open_folder.clicked.connect(self._open_folder)
        open_file = QPushButton("Open file")
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
            haystack = " ".join(
                str(record.get(key, ""))
                for key in ("Project", "Subject", "File", "Reason")
            ).lower()
            if query and query not in haystack:
                continue
            is_quarantined = record.get("Status") == "quarantined"
            if status == "Saved" and is_quarantined:
                continue
            if status == "Blocked" and not is_quarantined:
                continue
            filtered.append(record)

        visible = filtered[:DISPLAY_ROW_LIMIT]
        self.table.setUpdatesEnabled(False)
        self.table.setRowCount(len(visible))
        for row, record in enumerate(visible):
            is_quarantined = record.get("Status") == "quarantined"
            values = [
                str(record.get("Date", ""))[:16].replace("T", " "),
                record.get("Project", ""),
                record.get("Subject", ""),
                record.get("File", ""),
                "Blocked" if is_quarantined else "Saved",
                record.get("Reason", ""),
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
        ok, message = data_service.open_path(record.get("Folder", ""))
        if not ok:
            QMessageBox.warning(self, "Could not open", message)

    def _open_file(self) -> None:
        record = self._selected()
        if not record:
            return
        if record.get("Status") == "quarantined":
            QMessageBox.information(
                self,
                "Blocked file",
                "This attachment was quarantined and cannot be opened from the application.",
            )
            return
        ok, message = data_service.open_path(record.get("Path", ""))
        if not ok:
            QMessageBox.warning(self, "Could not open", message)


class SettingsPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 24)
        outer.setSpacing(14)

        title = QLabel("Settings and status")
        title.setObjectName("PageTitle")
        subtitle = QLabel("Local Outlook, storage, Gemini, and automation information.")
        subtitle.setObjectName("PageSubtitle")
        outer.addWidget(title)
        outer.addWidget(subtitle)

        panel = QFrame()
        panel.setObjectName("Panel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(12)

        mode_text = "Processing computer" if data_service.IS_PROCESSOR else "Viewer computer"
        self.mode_value = add_setting_row(layout, "Installation mode", mode_text)
        self.shared_value = add_setting_row(layout, "Shared data folder", str(data_service.SHARED_DATA_PATH))
        self.output_value = add_setting_row(layout, "Output folder", str(data_service.OUTPUT_ROOT))
        self.archive_value = add_setting_row(layout, "Processed email archive", str(data_service.PROJECT_YEAR_ROOT))
        self.category_value = add_setting_row(layout, "Outlook category", data_service.PROCESSED_CATEGORY_NAME)
        gemini_status = "Configured" if os.environ.get("GEMINI_API_KEY") else "Not configured"
        self.gemini_value = add_setting_row(layout, "Gemini key", gemini_status)
        flag_status = "Enabled" if data_service.FLAG_PROCESSED_EMAILS else "Disabled"
        self.flag_value = add_setting_row(layout, "Automatic Outlook flagging", flag_status)
        pending_status = "Enabled" if data_service.CHECK_PENDING_RESPONSES else "Disabled"
        self.pending_value = add_setting_row(layout, "Pending-response detection", pending_status)
        self.pending_folder_value = add_setting_row(layout, "Pending Outlook folder", data_service.PENDING_FOLDER_NAME)
        self.junk_value = add_setting_row(
            layout,
            "Irrelevant emails",
            "Left unchanged in the Inbox (current pipeline behavior)",
        )
        self.billing_value = add_setting_row(layout, "Billing output folder", str(data_service.BILLING_OUTPUT_ROOT))
        self.admin_value = add_setting_row(layout, "Administration email", data_service.ADMINISTRACION_EMAIL)
        self.boss_value = add_setting_row(layout, "Responsible person's email", data_service.BOSS_EMAIL)
        self.target_mailbox_value = add_setting_row(layout, "Processed Outlook mailbox", data_service.TARGET_MAILBOX)
        description_path = str(data_service.DESCRIPTIONS_XLSX_PATH) if data_service.DESCRIPTIONS_XLSX_PATH else "Not configured"
        self.description_value = add_setting_row(layout, "Descriptions spreadsheet", description_path)
        self.description_limit_value = add_setting_row(layout, "Description limit", f"{data_service.DESCRIPTION_MAX_CHARS} characters")
        self.ignore_domains_value = add_setting_row(layout, "Ignored internal domains", data_service.IGNORE_DOMAINS)
        self.ignore_senders_value = add_setting_row(layout, "Ignored senders", data_service.IGNORE_SENDERS)
        self.plenergy_value = add_setting_row(layout, "Plenergy / Plainco domains", data_service.PLENERGY_SENDER_DOMAINS)
        self.report_value = add_setting_row(layout, "Generated report", str(data_service.REPORT_XLSX_PATH))
        self.scheduler_value = add_setting_row(layout, "Scheduled task", "Checking...")
        outer.addWidget(panel)

        actions = QHBoxLayout()
        open_output = QPushButton("Open output folder")
        open_output.setObjectName("PrimaryButton")
        open_output.clicked.connect(self._open_output)
        open_env = QPushButton("Open .env file")
        open_env.setObjectName("SecondaryButton")
        open_env.clicked.connect(self._open_env)
        actions.addWidget(open_output)
        actions.addWidget(open_env)
        actions.addStretch()
        outer.addLayout(actions)
        outer.addStretch()

    def set_scheduler_status(self, status: dict[str, Any] | None) -> None:
        if status and status.get("Mode") == "VIEWER":
            self.scheduler_value.setText("Disabled in viewer mode")
            return
        if not status:
            self.scheduler_value.setText("Not found or unavailable")
            return
        result = status.get("LastTaskResult")
        result_text = "OK" if result == 0 else f"Code {result}"
        self.scheduler_value.setText(
            f"Last run: {status.get('LastRunTime', '—')} | Result: {result_text} | "
            f"Next: {status.get('NextRunTime', '—')}"
        )

    def _open_output(self) -> None:
        ok, message = data_service.open_path(data_service.OUTPUT_ROOT)
        if not ok:
            QMessageBox.warning(self, "Could not open", message)

    def _open_env(self) -> None:
        ok, message = data_service.open_path(data_service.ENV_PATH)
        if not ok:
            QMessageBox.warning(self, "Could not open", message)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        mode_suffix = " - Viewer" if data_service.IS_VIEWER else " - Processing"
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
        for page in (
            self.dashboard,
            self.emails,
            self.flagged,
            self.pending_page,
            self.report_page,
            self.excel_page,
            self.attachments,
            self.settings_page,
        ):
            self.stack.addWidget(page)
        main.addWidget(self.stack, 1)
        root_layout.addLayout(main, 1)

        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("Ready")

        if data_service.IS_PROCESSOR:
            self.dashboard.process_requested.connect(self.run_pipeline)
        self.dashboard.refresh_requested.connect(self.refresh_all)
        self.dashboard.open_output_requested.connect(self.open_output)
        QTimer.singleShot(120, self.refresh_all)

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
            ("▦  Dashboard", 0),
            ("✉  Emails", 1),
            ("⚑  Flagged in Outlook", 2),
            ("↩  Awaiting response", 3),
            ("▤  Email report", 4),
            ("▥  Excel report", 5),
            ("▣  Attachments", 6),
            ("⚙  Settings", 7),
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
        status_text = "●  Viewer · shared data" if data_service.IS_VIEWER else "●  Processing computer"
        status = QLabel(status_text)
        status.setObjectName("SidebarStatus")
        status.setToolTip(
            "This computer only reads the shared data."
            if data_service.IS_VIEWER
            else f"This computer processes mailbox {data_service.TARGET_MAILBOX}."
        )
        layout.addWidget(status)
        version = QLabel("Desktop UI 1.10")
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
        self.top_title = QLabel("Dashboard")
        self.top_title.setObjectName("PageTitle")
        self.top_subtitle = QLabel("Processing and Outlook status overview")
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
        self.language_combo.setToolTip("Change application language")
        self.language_combo.currentIndexChanged.connect(self._change_language)
        layout.addWidget(self.language_combo)

        initial_status = "Shared data: checking..." if data_service.IS_VIEWER else "Outlook: checking..."
        self.outlook_pill = QLabel(initial_status)
        self.outlook_pill.setObjectName("WarningPill")
        layout.addWidget(self.outlook_pill)
        self.top_refresh_button = QPushButton("↻")
        self.top_refresh_button.setObjectName("SecondaryButton")
        self.top_refresh_button.setFixedWidth(44)
        self.top_refresh_button.setToolTip("Refresh all data")
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
            "Language updated",
            "The application will restart to apply the new language.",
        )
        restart_application()

    def _navigate(self, index: int) -> None:
        titles = [
            ("Dashboard", "Processing and Outlook status overview"),
            ("Emails", "Processed email archive organized by project"),
            ("Flagged in Outlook", "Emails identified as processed by the assistant"),
            ("Awaiting response", "Incoming messages that still need a company reply"),
            ("Email report", "Generated summary of filed correspondence"),
            ("Excel report", "Generated Excel file in the output folder"),
            ("Attachments", "Files saved and blocked for security"),
            ("Settings", "Local application status and paths"),
        ]
        self.stack.setCurrentIndex(index)
        self.top_title.setText(titles[index][0])
        self.top_subtitle.setText(titles[index][1])
        if index == 7:
            self.refresh_scheduler_status()

    def refresh_all(self) -> None:
        if self.local_worker and self.local_worker.isRunning():
            self.statusBar().showMessage("A refresh is already running", 2500)
            return
        self.statusBar().showMessage("Loading local data in the background...")
        self.dashboard.refresh_button.setDisabled(True)
        self.top_refresh_button.setDisabled(True)
        self.local_worker = LocalDataWorker(self)
        self.local_worker.loaded.connect(self._local_data_loaded)
        self.local_worker.failed.connect(self._local_data_failed)
        self.local_worker.finished.connect(self._local_refresh_finished)
        self.local_worker.start()
        self.refresh_outlook_flags()

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
        self.statusBar().showMessage("Local data refreshed", 4000)

    def _local_data_failed(self, message: str) -> None:
        QMessageBox.critical(self, "Refresh error", message)
        self.statusBar().showMessage("Error refreshing local data", 4000)

    def _local_refresh_finished(self) -> None:
        self.dashboard.refresh_button.setDisabled(False)
        self.top_refresh_button.setDisabled(False)

    def refresh_scheduler_status(self) -> None:
        if data_service.IS_VIEWER:
            self.settings_page.set_scheduler_status({"Mode": "VIEWER"})
            return
        if self.scheduler_worker and self.scheduler_worker.isRunning():
            return
        self.settings_page.scheduler_value.setText("Checking...")
        self.scheduler_worker = SchedulerStatusWorker(self)
        self.scheduler_worker.loaded.connect(self.settings_page.set_scheduler_status)
        self.scheduler_worker.start()

    def refresh_outlook_flags(self) -> None:
        if self.outlook_worker and self.outlook_worker.isRunning():
            return
        self.outlook_pill.setText(
            "Shared data: checking..." if data_service.IS_VIEWER else "Outlook: checking..."
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
                "Shared data: unavailable" if data_service.IS_VIEWER else "Outlook: unavailable"
            )
            self.outlook_pill.setObjectName("ErrorPill")
        else:
            self.outlook_pill.setText(
                f"Shared data: connected · {summary.total} emails"
                if data_service.IS_VIEWER
                else f"Outlook: connected · {summary.total} flagged"
            )
            self.outlook_pill.setObjectName("SuccessPill")
        self.outlook_pill.style().unpolish(self.outlook_pill)
        self.outlook_pill.style().polish(self.outlook_pill)

    def run_pipeline(self) -> None:
        if data_service.IS_VIEWER:
            QMessageBox.information(
                self,
                "Viewer mode",
                "This computer only displays shared data. Processing runs on the main computer.",
            )
            return
        answer = QMessageBox.question(
            self,
            "Process emails",
            "Recent emails from Inbox and Sent Items will be checked.\n\nDo you want to continue?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if answer != QMessageBox.Yes:
            return
        if self.pipeline_worker and self.pipeline_worker.isRunning():
            return
        self.dashboard.set_processing(True)
        self.statusBar().showMessage("Processing emails in the background...")
        self.pipeline_worker = PipelineWorker(self)
        self.pipeline_worker.finished_with_result.connect(self._pipeline_finished)
        self.pipeline_worker.start()

    def _pipeline_finished(self, success: bool, message: str) -> None:
        self.dashboard.set_processing(False)
        if success:
            QMessageBox.information(self, "Process completed", message[-2500:])
            self.statusBar().showMessage("Processing completed", 5000)
            self.refresh_all()
        else:
            QMessageBox.critical(self, "Processing error", message[-3500:])
            self.statusBar().showMessage("Processing finished with errors", 5000)

    def open_output(self) -> None:
        ok, message = data_service.open_path(data_service.OUTPUT_ROOT)
        if not ok:
            QMessageBox.warning(self, "Could not open", message)


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
    # Outgoing (SALIENTE) mail is never processed or shown anymore -- see
    # get_recent_emails() in src/ingestion/outlook_local.py and
    # data_service._is_saliente(). Only ENTRANTE is expected here now.
    if value == "ENTRANTE":
        return "Incoming"
    return value


def priority_label(value: Any) -> str:
    try:
        priority = int(value)
    except (TypeError, ValueError):
        return "—"
    return {
        5: "5 - Critical",
        4: "4 - High",
        3: "3 - Normal",
        2: "2 - Low",
        1: "1 - Very low",
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