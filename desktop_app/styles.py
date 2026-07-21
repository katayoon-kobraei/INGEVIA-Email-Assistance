APP_STYLESHEET = r"""
* {
    font-family: "Segoe UI";
    font-size: 13px;
    color: #172033;
}
QMainWindow, QWidget#AppRoot {
    background: #f4f7fb;
}

/* Sidebar */
QFrame#Sidebar {
    background: qlineargradient(
        x1: 0, y1: 0, x2: 0, y2: 1,
        stop: 0 #102a43,
        stop: 0.55 #173b67,
        stop: 1 #0b2239
    );
    border: none;
}
QLabel#BrandTitle {
    color: white;
    font-size: 19px;
    font-weight: 800;
    letter-spacing: 1px;
}
QLabel#BrandSubtitle {
    color: #b8d6f2;
    font-size: 11px;
}
QPushButton#NavButton {
    color: #d8e7f5;
    background: transparent;
    border: none;
    border-left: 4px solid transparent;
    border-radius: 8px;
    padding: 11px 14px;
    text-align: left;
    font-weight: 650;
}
QPushButton#NavButton:hover {
    background: rgba(255, 255, 255, 0.10);
    color: white;
}
QPushButton#NavButton:checked {
    background: #2f80ed;
    color: white;
    border-left: 4px solid #7dd3fc;
}
QLabel#SidebarStatus {
    color: #b8d6f2;
    padding: 6px 12px;
}

/* Header and page titles */
QFrame#TopBar {
    background: white;
    border-bottom: 1px solid #e2e8f0;
}
QLabel#PageTitle {
    font-size: 23px;
    font-weight: 800;
    color: #102a43;
}
QLabel#PageSubtitle {
    color: #62748a;
}
QLabel#SectionTitle {
    font-size: 16px;
    font-weight: 750;
    color: #183b56;
}
QLabel#MutedText { color: #718096; }

/* Buttons */
QPushButton#PrimaryButton {
    background: #2f80ed;
    color: white;
    border: none;
    border-radius: 9px;
    padding: 10px 16px;
    font-weight: 750;
}
QPushButton#PrimaryButton:hover { background: #1f6fd1; }
QPushButton#PrimaryButton:pressed { background: #1859a8; }
QPushButton#PrimaryButton:disabled { background: #a8bdd7; }

QPushButton#ProcessButton {
    background: #16a36a;
    color: white;
    border: none;
    border-radius: 9px;
    padding: 11px 18px;
    font-weight: 800;
}
QPushButton#ProcessButton:hover { background: #118455; }
QPushButton#ProcessButton:pressed { background: #0d6b46; }
QPushButton#ProcessButton:disabled { background: #9ccfba; }

QPushButton#SecondaryButton {
    background: white;
    color: #253b53;
    border: 1px solid #cbd7e5;
    border-radius: 9px;
    padding: 9px 14px;
    font-weight: 650;
}
QPushButton#SecondaryButton:hover {
    background: #edf5ff;
    border-color: #8dbcf3;
    color: #1f6fd1;
}
QPushButton#DangerButton {
    background: #fff1f2;
    color: #b42318;
    border: 1px solid #fecdd3;
    border-radius: 9px;
    padding: 9px 14px;
    font-weight: 650;
}

/* Panels */
QFrame#Panel {
    background: white;
    border: 1px solid #e0e7ef;
    border-radius: 13px;
}
QFrame#HeroPanel {
    background: qlineargradient(
        x1: 0, y1: 0, x2: 1, y2: 0,
        stop: 0 #eaf4ff,
        stop: 0.55 #f5faff,
        stop: 1 #eefbf5
    );
    border: 1px solid #cfe2f6;
    border-radius: 14px;
}
QFrame#AttentionPanel {
    background: #fffaf0;
    border: 1px solid #f6d99b;
    border-radius: 13px;
}

/* Metric cards */
QFrame#MetricCard {
    background: white;
    border: 1px solid #e0e7ef;
    border-radius: 13px;
}
QFrame#MetricCard[tone="green"] {
    background: #ecfdf3;
    border: 1px solid #a7f3d0;
    border-left: 5px solid #16a34a;
}
QFrame#MetricCard[tone="orange"] {
    background: #fff8e8;
    border: 1px solid #f7d58a;
    border-left: 5px solid #f59e0b;
}
QFrame#MetricCard[tone="blue"] {
    background: #edf5ff;
    border: 1px solid #b9d7fb;
    border-left: 5px solid #2f80ed;
}
QFrame#MetricCard[tone="red"] {
    background: #fff1f2;
    border: 1px solid #fecdd3;
    border-left: 5px solid #e5484d;
}
QFrame#MetricCard[tone="purple"] {
    background: #f5f0ff;
    border: 1px solid #d8c8ff;
    border-left: 5px solid #7c3aed;
}
QLabel#MetricIcon {
    border-radius: 9px;
    font-size: 18px;
    font-weight: 800;
    padding: 6px;
}
QLabel#MetricIcon[tone="green"] { background: #d1fae5; color: #15803d; }
QLabel#MetricIcon[tone="orange"] { background: #ffedc2; color: #b45309; }
QLabel#MetricIcon[tone="blue"] { background: #dbeafe; color: #1d4ed8; }
QLabel#MetricIcon[tone="red"] { background: #ffe4e6; color: #be123c; }
QLabel#MetricIcon[tone="purple"] { background: #ede9fe; color: #6d28d9; }
QLabel#MetricValue {
    font-size: 30px;
    font-weight: 800;
    color: #102a43;
}
QLabel#MetricValue[tone="green"] { color: #16794a; }
QLabel#MetricValue[tone="orange"] { color: #b76100; }
QLabel#MetricValue[tone="blue"] { color: #1f6fd1; }
QLabel#MetricValue[tone="red"] { color: #c9363e; }
QLabel#MetricValue[tone="purple"] { color: #6d3cc7; }
QLabel#MetricTitle {
    color: #42566f;
    font-weight: 700;
}
QLabel#MetricHint {
    color: #74869b;
    font-size: 11px;
}

/* Status pills */
QLabel#SuccessPill {
    color: #087443;
    background: #e8f7ef;
    border: 1px solid #b9e8cf;
    border-radius: 10px;
    padding: 4px 9px;
    font-weight: 700;
}
QLabel#WarningPill {
    color: #9a5300;
    background: #fff2d8;
    border: 1px solid #f5d496;
    border-radius: 10px;
    padding: 4px 9px;
    font-weight: 700;
}
QLabel#ErrorPill {
    color: #b42318;
    background: #ffebe9;
    border: 1px solid #fecaca;
    border-radius: 10px;
    padding: 4px 9px;
    font-weight: 700;
}

/* Inputs */
QLineEdit, QComboBox {
    background: white;
    border: 1px solid #cfd9e6;
    border-radius: 9px;
    padding: 8px 10px;
    min-height: 20px;
    selection-background-color: #cfe4ff;
}
QLineEdit:hover, QComboBox:hover { border-color: #8dbcf3; }
QLineEdit:focus, QComboBox:focus {
    border: 2px solid #2f80ed;
    padding: 7px 9px;
}

/* Tables */
QTableWidget {
    background: white;
    border: 1px solid #e0e7ef;
    border-radius: 11px;
    gridline-color: #edf1f6;
    selection-background-color: #dcebff;
    selection-color: #102a43;
    alternate-background-color: #f8fbff;
}
QHeaderView::section {
    background: #eaf2fb;
    color: #34506c;
    border: none;
    border-bottom: 1px solid #cfdceb;
    padding: 9px;
    font-weight: 750;
}
QTableWidget::item { padding: 7px; }
QTableWidget::item:hover { background: #f0f7ff; }

/* Drawer */
QFrame#Drawer {
    background: white;
    border: 1px solid #bfd3e8;
    border-radius: 13px;
}
QLabel#DrawerTitle {
    font-size: 17px;
    font-weight: 800;
    color: #183b56;
}
QLabel#DrawerLabel {
    color: #6f8195;
    font-size: 11px;
    font-weight: 700;
}
QLabel#DrawerValue {
    color: #1c273b;
    font-size: 13px;
}

/* Progress and scrollbars */
QProgressBar {
    background: #dbe7f2;
    border: none;
    border-radius: 4px;
    min-height: 8px;
    max-height: 8px;
}
QProgressBar::chunk {
    background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0, stop: 0 #2f80ed, stop: 1 #16a36a);
    border-radius: 4px;
}
QScrollBar:vertical {
    width: 10px;
    background: transparent;
}
QScrollBar::handle:vertical {
    background: #b8c6d8;
    border-radius: 5px;
    min-height: 30px;
}
QScrollBar::handle:vertical:hover { background: #8ea6c2; }
QStatusBar {
    background: white;
    color: #68758c;
    border-top: 1px solid #e1e7f0;
}
QToolTip {
    background: #17324d;
    color: white;
    border: none;
    padding: 6px;
}
"""

# The language selector uses the same visual treatment as the other inputs.
