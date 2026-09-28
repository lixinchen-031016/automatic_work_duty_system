"""小型 Qt 界面复用组件与表格填充工具。"""

from __future__ import annotations

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialogButtonBox,
    QFrame,
    QHeaderView,
    QLabel,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .parser import BLOCK_LABELS, BLOCK_SESSIONS


class Card(QWidget):
    """白色圆角卡片容器（类苹果表面样式）"""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setAttribute(Qt.WA_StyledBackground, True)


def card(widget: QWidget) -> Card:
    """把控件包进白色圆角卡片。"""
    container = Card()
    layout = QVBoxLayout(container)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.addWidget(widget)
    return container


def fill_table(table: QTableWidget, df: pd.DataFrame) -> None:
    """把 DataFrame 填入只读 QTableWidget，尺寸不变时复用单元格。"""
    rows, cols = len(df), len(df.columns)
    if table.rowCount() != rows or table.columnCount() != cols:
        table.clearContents()
        table.setRowCount(rows)
        table.setColumnCount(cols)
    table.setHorizontalHeaderLabels([str(column) for column in df.columns])
    table.setShowGrid(False)
    table.setAlternatingRowColors(True)
    table.verticalHeader().setVisible(False)
    table.verticalHeader().setDefaultSectionSize(30)

    values = [
        tuple(str(value) for value in row)
        for row in df.itertuples(index=False, name=None)
    ]
    for row_index, row_values in enumerate(values):
        for column_index, text in enumerate(row_values):
            item = table.item(row_index, column_index)
            if item is None:
                item = QTableWidgetItem()
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                table.setItem(row_index, column_index, item)
            if item.text() != text:
                item.setText(text)

    header = table.horizontalHeader()
    if cols and header.sectionResizeMode(0) != QHeaderView.Interactive:
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setStretchLastSection(True)
    if cols:
        table.resizeColumnsToContents()


def special_sessions_label(session_list: list[int]) -> str:
    """把节次列表展示成用户登记时选择的时段块。"""
    sessions = set(session_list)
    labels = [
        BLOCK_LABELS[block].split(" ")[0]
        for block in sorted(BLOCK_SESSIONS)
        if set(BLOCK_SESSIONS[block]) <= sessions
    ]
    return "、".join(labels) or "自定义节次"


def special_weeks_label(week_start: int, week_end: int) -> str:
    """把连续周次范围格式化为界面标签。"""
    if week_start == week_end:
        return f"第{week_start}周"
    return f"第{week_start}–{week_end}周"


def dialog_header(title: str, subtitle: str = "") -> QFrame:
    """创建统一弹窗标题区。"""
    header = QFrame()
    header.setObjectName("dialogHeader")
    layout = QVBoxLayout(header)
    layout.setContentsMargins(0, 0, 0, 10)
    layout.setSpacing(4)
    title_label = QLabel(title)
    title_label.setObjectName("dialogTitle")
    layout.addWidget(title_label)
    if subtitle:
        subtitle_label = QLabel(subtitle)
        subtitle_label.setObjectName("secondary")
        subtitle_label.setWordWrap(True)
        layout.addWidget(subtitle_label)
    return header


def dialog_buttons(
    accept_text: str = "确定",
    cancel_text: str = "取消",
) -> QDialogButtonBox:
    """创建支持 Enter 确认、Esc 取消的标准弹窗按钮区。"""
    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    accept = buttons.button(QDialogButtonBox.Ok)
    cancel = buttons.button(QDialogButtonBox.Cancel)
    accept.setText(accept_text)
    cancel.setText(cancel_text)
    accept.setDefault(True)
    accept.setAutoDefault(True)
    cancel.setAutoDefault(True)
    return buttons
