"""Единая плоская тема оформления клиента.

Цвета статусов остаются в schedule_client.status_mapping; здесь только
нейтральная палитра интерфейса, отступы и состояния элементов.
"""

from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QApplication, QLabel, QWidget

ACCENT = "#2563EB"
ACCENT_HOVER = "#1D4ED8"
TEXT = "#1E293B"
TEXT_MUTED = "#64748B"
SURFACE = "#FFFFFF"
SURFACE_ALT = "#F8FAFC"
BORDER = "#E2E8F0"
BORDER_STRONG = "#CBD5E1"


APP_STYLE = """
* {
    font-family: "Segoe UI", "Inter", "Arial";
}
QMainWindow, QDialog {
    background: #F4F6FA;
}
QWidget#appRoot {
    background: #F4F6FA;
}
QWidget {
    color: #1E293B;
    font-size: 10pt;
}
QLabel {
    background: transparent;
}
QLabel#appTitle {
    color: #0F172A;
    font-size: 14pt;
    font-weight: 700;
}
QLabel#appSubtitle {
    color: #64748B;
    font-size: 8.5pt;
}
QLabel#sectionCaption {
    color: #94A3B8;
    font-size: 8pt;
    font-weight: 700;
    letter-spacing: 1px;
}
QLabel#mutedLabel {
    color: #64748B;
    font-size: 8.5pt;
}
QLabel#legendLabel {
    color: #475569;
    font-size: 8.5pt;
}
QLabel#hintLabel {
    color: #64748B;
    font-style: italic;
}
QLabel#statusPill {
    border-radius: 8px;
    padding: 3px 8px;
    font-size: 8.5pt;
    font-weight: 600;
}
QLabel#statusPill[state="connected"],
QLabel#statusPill[state="saved"] {
    background: #DCFCE7;
    color: #15803D;
    border: 1px solid #BBF7D0;
}
QLabel#statusPill[state="offline"],
QLabel#statusPill[state="error"] {
    background: #FEE2E2;
    color: #B91C1C;
    border: 1px solid #FECACA;
}
QLabel#statusPill[state="unsaved"] {
    background: #FEF3C7;
    color: #B45309;
    border: 1px solid #FDE68A;
}
QLabel#statusPill[state="local"] {
    background: #EFF6FF;
    color: #1D4ED8;
    border: 1px solid #BFDBFE;
}
QFrame#appHeader {
    background: transparent;
    border: none;
}
QFrame#toolbarCard {
    background: transparent;
    border: none;
    border-bottom: 1px solid #E2E8F0;
    border-radius: 0;
}
QFrame#legendCard {
    background: transparent;
    border: none;
}
QWidget#monthCard {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 10px;
}
QPushButton {
    background: #FFFFFF;
    color: #334155;
    border: 1px solid #E2E8F0;
    border-radius: 7px;
    padding: 0 12px;
    min-height: 30px;
    font-size: 9.5pt;
    font-weight: 600;
}
QPushButton:hover {
    background: #F8FAFC;
    border-color: #CBD5E1;
}
QPushButton:pressed {
    background: #F1F5F9;
}
QPushButton:disabled {
    background: #F8FAFC;
    color: #94A3B8;
    border-color: #E2E8F0;
}
QPushButton[variant="primary"] {
    background: #2563EB;
    color: #FFFFFF;
    border: 1px solid #2563EB;
}
QPushButton[variant="primary"]:hover {
    background: #1D4ED8;
    border-color: #1D4ED8;
}
QPushButton[variant="primary"]:pressed {
    background: #1E40AF;
    border-color: #1E40AF;
}
QPushButton[variant="success"] {
    background: #16A34A;
    color: #FFFFFF;
    border: 1px solid #16A34A;
}
QPushButton[variant="success"]:hover {
    background: #15803D;
    border-color: #15803D;
}
QPushButton[variant="success"]:pressed {
    background: #166534;
    border-color: #166534;
}
QPushButton[variant="ghost"] {
    background: transparent;
    color: #475569;
    border: 1px solid transparent;
}
QPushButton[variant="ghost"]:hover {
    background: #F1F5F9;
}
QPushButton[variant="icon"] {
    background: #F1F5F9;
    color: #475569;
    border: 1px solid transparent;
    padding: 0;
}
QPushButton[variant="icon"]:hover {
    background: #E2E8F0;
}
QPushButton[variant="danger"] {
    background: #FEF2F2;
    color: #B91C1C;
    border: 1px solid #FECACA;
}
QPushButton[variant="danger"]:hover {
    background: #FEE2E2;
    border-color: #FCA5A5;
}
QPushButton[variant="danger"]:pressed {
    background: #FECACA;
}
QPushButton[variant="dangerIcon"] {
    background: #FEF2F2;
    color: #B91C1C;
    border: 1px solid #FECACA;
    padding: 0;
}
QPushButton[variant="dangerIcon"]:hover {
    background: #FEE2E2;
    border-color: #FCA5A5;
}
QPushButton[variant="dangerIcon"]:pressed {
    background: #FECACA;
}
QToolButton#appMenuButton,
QToolButton#cardMenuButton {
    background: transparent;
    color: #64748B;
    border: 1px solid transparent;
    border-radius: 6px;
    padding: 0;
    font-size: 14pt;
    font-weight: 600;
}
QToolButton#appMenuButton:hover,
QToolButton#cardMenuButton:hover {
    background: #F1F5F9;
    color: #334155;
    border-color: #E2E8F0;
}
QToolButton#appMenuButton::menu-indicator,
QToolButton#cardMenuButton::menu-indicator {
    image: none;
    width: 0;
}
QDialogButtonBox QPushButton {
    min-width: 92px;
}
QDialogButtonBox QPushButton:default {
    background: #2563EB;
    color: #FFFFFF;
    border-color: #2563EB;
}
QDialogButtonBox QPushButton:default:hover {
    background: #1D4ED8;
    border-color: #1D4ED8;
}
QLineEdit, QComboBox, QTimeEdit, QSpinBox, QDateEdit {
    background: #FFFFFF;
    color: #1E293B;
    border: 1px solid #E2E8F0;
    border-radius: 7px;
    padding: 0 8px;
    min-height: 30px;
    selection-background-color: #DBEAFE;
}
QLineEdit:focus, QComboBox:focus, QTimeEdit:focus, QSpinBox:focus, QDateEdit:focus {
    border: 1px solid #2563EB;
}
QLineEdit:disabled, QComboBox:disabled, QTimeEdit:disabled {
    background: #F8FAFC;
    color: #94A3B8;
}
QComboBox::drop-down {
    border: none;
    width: 22px;
}
QComboBox#periodCombo {
    min-height: 26px;
    padding: 0 6px;
    font-size: 9pt;
}
QComboBox#periodCombo::drop-down {
    width: 20px;
}
QComboBox QAbstractItemView {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 8px;
    padding: 4px;
    selection-background-color: #EFF6FF;
    selection-color: #1E293B;
    outline: 0;
}
QTableWidget#scheduleTable {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 10px;
    gridline-color: #EEF2F7;
    alternate-background-color: #FAFBFD;
    selection-background-color: transparent;
    selection-color: #0F172A;
}
QTableWidget#scheduleTable::item {
    padding: 0 2px;
}
QTableWidget#scheduleTable::item:selected {
    background: transparent;
    border: 2px solid #2563EB;
}
QHeaderView::section {
    background: #F8FAFC;
    color: #475569;
    padding: 2px 1px;
    border: none;
    border-right: 1px solid #EEF2F7;
    border-bottom: 1px solid #E2E8F0;
    font-weight: 600;
}
QHeaderView::section:vertical {
    padding: 0 8px;
    text-align: left;
}
QTableCornerButton::section {
    background: #F8FAFC;
    border: none;
    border-right: 1px solid #EEF2F7;
    border-bottom: 1px solid #E2E8F0;
}
QMenuBar {
    background: #FFFFFF;
    color: #475569;
    border-bottom: 1px solid #E2E8F0;
    padding: 2px 8px;
}
QMenuBar::item {
    background: transparent;
    padding: 6px 10px;
    border-radius: 6px;
}
QMenuBar::item:selected {
    background: #EFF6FF;
    color: #1D4ED8;
}
QMenu {
    background: #FFFFFF;
    color: #334155;
    border: 1px solid #E2E8F0;
    border-radius: 8px;
    padding: 6px;
}
QMenu::item {
    padding: 6px 22px 6px 12px;
    border-radius: 6px;
}
QMenu::item:selected {
    background: #EFF6FF;
    color: #1D4ED8;
}
QMenu::separator {
    height: 1px;
    background: #EEF2F7;
    margin: 5px 8px;
}
QStatusBar {
    background: transparent;
    color: #94A3B8;
    border-top: none;
    font-size: 8pt;
}
QStatusBar::item {
    border: none;
}
QSplitter::handle:vertical {
    background: transparent;
}
QSplitter::handle:vertical:hover {
    background: #CBD5E1;
    border-radius: 4px;
}
QScrollBar:vertical {
    background: transparent;
    width: 10px;
    margin: 2px;
}
QScrollBar::handle:vertical {
    background: #CBD5E1;
    min-height: 28px;
    border-radius: 5px;
}
QScrollBar::handle:vertical:hover {
    background: #94A3B8;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}
QScrollBar:horizontal {
    background: transparent;
    height: 10px;
    margin: 2px;
}
QScrollBar::handle:horizontal {
    background: #CBD5E1;
    min-width: 28px;
    border-radius: 5px;
}
QScrollBar::handle:horizontal:hover {
    background: #94A3B8;
}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0;
}
QTabWidget::pane {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 10px;
    top: -1px;
}
QTabBar::tab {
    background: transparent;
    color: #64748B;
    padding: 8px 16px;
    margin-right: 4px;
    border-radius: 6px;
}
QTabBar::tab:selected {
    background: #EFF6FF;
    color: #1D4ED8;
    font-weight: 600;
}
QTabBar::tab:hover:!selected {
    background: #F1F5F9;
}
QListWidget {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 8px;
    padding: 4px;
    outline: 0;
}
QListWidget::item {
    padding: 8px 10px;
    border-radius: 6px;
}
QListWidget::item:selected {
    background: #EFF6FF;
    color: #1D4ED8;
}
QListWidget::item:hover:!selected {
    background: #F8FAFC;
}
QGroupBox {
    border: 1px solid #E2E8F0;
    border-radius: 10px;
    margin-top: 14px;
    padding-top: 8px;
    background: #FFFFFF;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 6px;
    color: #475569;
    font-weight: 600;
}
QRadioButton::indicator, QCheckBox::indicator {
    width: 16px;
    height: 16px;
    border: 1px solid #CBD5E1;
    background: #FFFFFF;
}
QRadioButton::indicator {
    border-radius: 8px;
}
QRadioButton::indicator:checked {
    border: 5px solid #2563EB;
}
QCheckBox::indicator {
    border-radius: 4px;
}
QCheckBox::indicator:checked {
    background: #2563EB;
    border-color: #2563EB;
}
QToolTip {
    background: #0F172A;
    color: #FFFFFF;
    border: none;
    border-radius: 6px;
    padding: 6px 8px;
}
QScrollArea {
    border: none;
    background: transparent;
}
"""


def repolish(widget: QWidget) -> None:
    """Пересчитать стиль после смены динамического свойства."""
    style = widget.style()
    if style is not None:
        style.unpolish(widget)
        style.polish(widget)
    widget.update()


def set_pill_state(label: QLabel, state: str) -> None:
    """Состояние цветной плашки статуса: connected/saved/offline/unsaved."""
    if label.property("state") == state:
        return
    label.setProperty("state", state)
    repolish(label)


def set_variant(widget: QWidget, variant: str) -> None:
    """Вариант оформления кнопки: primary/success/ghost/icon/dangerIcon."""
    widget.setProperty("variant", variant)
    repolish(widget)


def apply_theme(app: QApplication | None = None) -> None:
    """Применить Fusion, шрифт и общую плоскую тему."""
    app = app or QApplication.instance()
    if app is None:
        return
    app.setStyle("Fusion")
    font = QFont("Segoe UI", 10)
    font.setStyleStrategy(QFont.PreferAntialias)
    app.setFont(font)
    app.setStyleSheet(APP_STYLE)
