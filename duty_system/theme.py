"""应用浅色主题与运行时生成的复选框图标。"""

from __future__ import annotations

import tempfile
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QImage, QPainter, QPen

STYLE = """
* { font-size: 13px; }

QMainWindow, QTabWidget::pane { background: #f5f5f7; }
QSplitter::handle { background: transparent; }
QSplitter::handle:horizontal { width: 16px; }

QLabel { color: #1d1d1f; background: transparent; }
QLabel#appTitle { font-size: 20px; font-weight: 700; }
QLabel#appSubtitle { font-size: 12px; color: #86868b; }
QLabel#summary { color: #3c3c43; }
QLabel#secondary { color: #86868b; font-size: 12px; }

Card {
    background: #ffffff;
    border: 1px solid #e8e8ed;
    border-radius: 10px;
}

QGroupBox {
    font-weight: 600; color: #1d1d1f;
    background: #ffffff;
    border: 1px solid #e8e8ed; border-radius: 10px;
    margin-top: 16px; padding: 12px;
}
QGroupBox::title {
    subcontrol-origin: margin; subcontrol-position: top left;
    left: 14px; top: 0px; padding: 0 5px;
    background: #ffffff; color: #1d1d1f;
}

QPushButton {
    background: #ffffff; color: #1d1d1f;
    border: 1px solid #d2d2d7; border-radius: 7px;
    padding: 6px 14px; font-weight: 500;
}
QPushButton:hover { background: #f7f7f9; }
QPushButton:pressed { background: #e8e8ed; }
QPushButton:disabled { color: #aeaeb2; background: #f7f7f9; border-color: #e5e5ea; }

QPushButton#primary { background: #007aff; color: #ffffff; border: none; font-weight: 600; }
QPushButton#primary:hover { background: #0071e3; }
QPushButton#primary:pressed { background: #0062cc; }
QPushButton#primary:disabled { background: #99c7ff; color: #ffffff; }

QPushButton#danger { background: transparent; color: #ff3b30; border: 1px solid #ffb3ae; }
QPushButton#danger:hover { background: #fff0ee; }
QPushButton#danger:pressed { background: #ffe1dd; }

QSpinBox, QComboBox, QDateEdit {
    background: #ffffff; color: #1d1d1f;
    border: 1px solid #d2d2d7; border-radius: 6px;
    padding: 3px 8px;
    selection-background-color: #007aff; selection-color: #ffffff;
}
QSpinBox:focus, QComboBox:focus, QDateEdit:focus { border: 1px solid #007aff; }
QSpinBox::up-button, QSpinBox::down-button, QDateEdit::up-button, QDateEdit::down-button { width: 0; border: none; background: none; }
QComboBox::drop-down { border: none; width: 24px; }
QDateEdit::drop-down { border: none; width: 26px; }
QComboBox QAbstractItemView {
    background: #ffffff; border: 1px solid #d2d2d7; border-radius: 8px;
    selection-background-color: #007aff; selection-color: #ffffff;
    outline: none; padding: 2px;
}

QDialog { background: #f5f5f7; }
QLineEdit {
    background: #ffffff; color: #1d1d1f;
    border: 1px solid #d2d2d7; border-radius: 6px; padding: 4px 8px;
    selection-background-color: #007aff; selection-color: #ffffff;
}
QLineEdit:focus { border: 1px solid #007aff; }

QCheckBox { color: #1d1d1f; background: transparent; spacing: 7px; }
QCheckBox::indicator {
    width: 16px; height: 16px;
    border: 1px solid #c7c7cc; border-radius: 4px; background: #ffffff;
}
QCheckBox::indicator:hover { border-color: #007aff; }
QCheckBox::indicator:checked { background: #007aff; border-color: #007aff; image: url(__CHECK__); }

QListWidget {
    background: #ffffff; border: 1px solid #e8e8ed; border-radius: 10px;
    outline: none; padding: 5px;
}
QListWidget::item { padding: 6px 8px; border-radius: 7px; margin: 1px 2px; }
QListWidget::item:hover { background: #f2f2f7; }
QListWidget::item:selected { background: #007aff; color: #ffffff; }

QTabWidget::pane { border: none; background: transparent; }
QTabBar { background: #e9e9eb; border-radius: 9px; padding: 2px; }
QTabBar::tab {
    background: transparent; color: #3c3c43;
    padding: 5px 14px; margin: 1px;
    border: none; border-radius: 7px; font-weight: 500;
}
QTabBar::tab:hover { color: #1d1d1f; }
QTabBar::tab:selected { background: #ffffff; color: #1d1d1f; font-weight: 600; }

QTableWidget {
    background: transparent; border: none;
    gridline-color: transparent;
    alternate-background-color: #f7f7f9;
    selection-background-color: #007aff; selection-color: #ffffff;
    outline: none;
}
QTableWidget::item { padding: 4px 8px; border: none; }
QHeaderView::section {
    background: transparent; color: #86868b;
    border: none; border-bottom: 1px solid #e5e5ea;
    padding: 7px 10px;
    font-weight: 600; font-size: 12px;
}
QTableCornerButton::section { background: transparent; border: none; }

QScrollBar:vertical { background: transparent; width: 10px; margin: 2px; }
QScrollBar::handle:vertical { background: #c7c7cc; border-radius: 4px; min-height: 28px; }
QScrollBar::handle:vertical:hover { background: #aeaeb2; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }
QScrollBar:horizontal { background: transparent; height: 10px; margin: 2px; }
QScrollBar::handle:horizontal { background: #c7c7cc; border-radius: 4px; min-width: 28px; }
QScrollBar::handle:horizontal:hover { background: #aeaeb2; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal { background: transparent; }

QStatusBar { background: transparent; color: #86868b; font-size: 12px; }
QStatusBar::item { border: none; }

QToolTip {
    background: #1d1d1f; color: #ffffff;
    border: none; border-radius: 6px; padding: 6px 10px; font-size: 12px;
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
