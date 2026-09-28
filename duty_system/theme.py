"""应用浅色主题与运行时生成的复选框图标。"""

from __future__ import annotations

import tempfile
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QImage, QPainter, QPen

STYLE = """
* { font-size: 13px; }

QMainWindow, QWidget#appRoot, QTabWidget::pane {
    background: #f3f6fb;
}
QWidget#leftPanel, QScrollArea#leftScroll {
    background: #f8fafc;
}
QScrollArea#leftScroll > QWidget > QWidget {
    background: #f8fafc;
}
QFrame#topBar {
    background: #ffffff;
    border-bottom: 1px solid #dde5f0;
}
QFrame#summaryCard, QFrame#metricCard, QFrame#emptyState, QFrame#legendBar, Card {
    background: #ffffff;
    border: 1px solid #dde5f0;
    border-radius: 14px;
}
QSplitter::handle { background: transparent; }
QSplitter::handle:horizontal { width: 14px; }

QLabel { color: #132238; background: transparent; }
QLabel#brandMark {
    color: #ffffff;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                stop:0 #3b82f6, stop:1 #1d4ed8);
    border-radius: 11px;
    font-size: 16px;
    font-weight: 600;
}
QLabel#appTitle { font-size: 19px; font-weight: 600; }
QLabel#appSubtitle { font-size: 11px; color: #66758a; }
QLabel#pageTitle { font-size: 18px; font-weight: 600; color: #132238; }
QLabel#legendText { color: #66758a; font-size: 11px; }
QLabel#summary { color: #526177; font-size: 12px; }
QLabel#secondary { color: #66758a; font-size: 12px; }
QLabel#metricLabel { color: #66758a; font-size: 11px; }
QLabel#metricValue { color: #132238; font-size: 22px; font-weight: 600; }
QLabel#metricHint { color: #059669; font-size: 10px; }
QLabel#statusPill {
    color: #475569;
    background: #eef2f7;
    border: 1px solid #dbe3ee;
    border-radius: 12px;
    padding: 5px 10px;
    font-size: 11px;
    font-weight: 500;
}

QGroupBox {
    color: #132238;
    background: #ffffff;
    border: 1px solid #dde5f0;
    border-radius: 14px;
    margin-top: 16px;
    padding: 13px;
    font-weight: 600;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 14px;
    top: 0;
    padding: 0 6px;
    color: #132238;
    background: #ffffff;
}

QPushButton {
    color: #132238;
    background: #ffffff;
    border: 1px solid #d9e2ee;
    border-radius: 9px;
    padding: 7px 13px;
    font-weight: 500;
}
QPushButton:hover {
    background: #f6f9fe;
    border-color: #b9cceb;
}
QPushButton:pressed { background: #eaf1fd; }
QPushButton:disabled {
    color: #9aa8ba;
    background: #f3f6fa;
    border-color: #e3e9f2;
}
QPushButton#primary {
    color: #ffffff;
    background: #2563eb;
    border: 1px solid #2563eb;
    font-weight: 600;
    padding: 9px 16px;
}
QPushButton#primary:hover { background: #1d4ed8; border-color: #1d4ed8; }
QPushButton#primary:pressed { background: #1e40af; }
QPushButton#primary:disabled {
    color: #ffffff;
    background: #9bb8f5;
    border-color: #9bb8f5;
}
QPushButton#danger {
    color: #dc2626;
    background: #fff7f7;
    border: 1px solid #fecaca;
}
QPushButton#danger:hover { background: #fee2e2; }

QSpinBox, QComboBox, QDateEdit, QLineEdit {
    color: #132238;
    background: #ffffff;
    border: 1px solid #d9e2ee;
    border-radius: 8px;
    padding: 5px 9px;
    selection-background-color: #2563eb;
    selection-color: #ffffff;
}
QSpinBox:focus, QComboBox:focus, QDateEdit:focus, QLineEdit:focus {
    border: 1px solid #2563eb;
}
QSpinBox::up-button, QSpinBox::down-button,
QDateEdit::up-button, QDateEdit::down-button {
    width: 0;
    border: none;
    background: none;
}
QComboBox::drop-down { border: none; width: 24px; }
QDateEdit::drop-down { border: none; width: 26px; }
QComboBox QAbstractItemView {
    color: #132238;
    background: #ffffff;
    border: 1px solid #d9e2ee;
    border-radius: 8px;
    selection-background-color: #2563eb;
    selection-color: #ffffff;
    outline: none;
    padding: 3px;
}

QDialog { background: #f3f6fb; }
QCheckBox { color: #132238; background: transparent; spacing: 7px; }
QCheckBox::indicator {
    width: 16px;
    height: 16px;
    background: #ffffff;
    border: 1px solid #c7d3e2;
    border-radius: 4px;
}
QCheckBox::indicator:hover { border-color: #2563eb; }
QCheckBox::indicator:checked {
    background: #2563eb;
    border-color: #2563eb;
    image: url(__CHECK__);
}

QListWidget {
    background: #ffffff;
    border: 1px solid #dde5f0;
    border-radius: 12px;
    outline: none;
    padding: 6px;
}
QListWidget::item {
    padding: 8px 9px;
    margin: 1px 2px;
    border-radius: 8px;
}
QListWidget::item:hover { background: #f3f6fb; }
QListWidget::item:selected { background: #dbeafe; color: #1d4ed8; }

QTabWidget::pane { border: none; background: transparent; }
QTabBar {
    background: #e8eef7;
    border-radius: 10px;
    padding: 3px;
}
QTabBar::tab {
    color: #5a6a80;
    background: transparent;
    border: none;
    border-radius: 8px;
    padding: 7px 14px;
    margin: 1px;
    font-weight: 500;
}
QTabBar::tab:hover { color: #132238; }
QTabBar::tab:selected {
    color: #1d4ed8;
    background: #ffffff;
    font-weight: 600;
}

QTableWidget {
    color: #132238;
    background: #ffffff;
    border: none;
    gridline-color: transparent;
    alternate-background-color: #f7faff;
    selection-background-color: #dbeafe;
    selection-color: #1d4ed8;
    outline: none;
}
QTableWidget::item { padding: 5px 9px; border: none; }
QHeaderView::section {
    color: #66758a;
    background: #f8fafc;
    border: none;
    border-bottom: 1px solid #dde5f0;
    padding: 8px 10px;
    font-weight: 600;
    font-size: 12px;
}
QTableCornerButton::section { background: #f8fafc; border: none; }

QScrollBar:vertical { background: transparent; width: 10px; margin: 2px; }
QScrollBar::handle:vertical { background: #c7d3e2; border-radius: 5px; min-height: 28px; }
QScrollBar::handle:vertical:hover { background: #9fb0c6; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }
QScrollBar:horizontal { background: transparent; height: 10px; margin: 2px; }
QScrollBar::handle:horizontal { background: #c7d3e2; border-radius: 5px; min-width: 28px; }
QScrollBar::handle:horizontal:hover { background: #9fb0c6; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal { background: transparent; }

QStatusBar { color: #66758a; background: transparent; font-size: 12px; }
QStatusBar::item { border: none; }

QToolTip {
    color: #ffffff;
    background: #132238;
    border: none;
    border-radius: 7px;
    padding: 7px 10px;
    font-size: 12px;
}
"""


def _write_check_icon() -> str:
    """生成复选框选中态的白色对勾图标，供 QSS image 引用。"""
    image = QImage(64, 64, QImage.Format_ARGB32)
    image.fill(Qt.transparent)
    painter = QPainter(image)
    painter.setRenderHint(QPainter.Antialiasing)
    pen = QPen(QColor("#ffffff"), 8)
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    painter.setPen(pen)
    painter.drawLine(16, 34, 27, 45)
    painter.drawLine(27, 45, 48, 18)
    painter.end()
    path = Path(tempfile.gettempdir()) / "duty_check_icon.png"
    image.save(str(path))
    return path.as_posix()


def build_style() -> str:
    """装配全局样式表，并注入运行时生成的对勾图标路径。"""
    return STYLE.replace("__CHECK__", _write_check_icon())
