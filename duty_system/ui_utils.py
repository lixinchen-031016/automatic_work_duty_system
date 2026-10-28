"""小型 Qt 界面复用组件与表格填充工具。"""

from __future__ import annotations

import pandas as pd
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QDialogButtonBox,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QToolButton,
    QVBoxLayout,
    QWidget,
    QWidgetAction,
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


class _CheckList(QListWidget):
    """点整行即可切换勾选项。

    Qt 默认只有在点中文字左侧的小方框时才会切换勾选，鼠标稍有偏差就没反应。
    这里先让基类处理（它会负责选中/焦点/小方框切换），再比对勾选状态：
    状态没变说明这次点在了文字或行内空白上，由我们补一次切换。
    """

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802 - Qt 命名
        if event.button() != Qt.LeftButton:
            super().mouseReleaseEvent(event)
            return
        item = self.itemAt(event.position().toPoint())
        if item is None or not item.flags() & Qt.ItemIsUserCheckable:
            super().mouseReleaseEvent(event)
            return
        before = item.checkState()
        super().mouseReleaseEvent(event)
        if item.checkState() == before:
            item.setCheckState(
                Qt.Unchecked if before == Qt.Checked else Qt.Checked)


class ChecklistMenuButton(QToolButton):
    """下拉多选按钮：勾选若干项，按钮文字实时汇总当前选中范围。

    为什么不用 QComboBox：下拉框只能单选，而且每次点击都会立即收起。
    这里把一张带勾选框的清单塞进 QMenu（QWidgetAction 承载），点勾选框只改变
    勾选状态、弹层保持打开，可以连续勾选多个人；点弹层外任意位置收起。

    按钮文字规则（见 _refresh_text）：
    - 一项没选：``人员：未选择``
    - 全部选中：``人员：全部（12）``
    - 部分选中：``人员：张三、李四 等 3/12``（最多列举 inline_names 个名字）

    ``selectionChanged`` 只在勾选状态真正变化时触发；重建清单（set_items）
    属于「同步数据」而不是「用户操作」，不会触发，避免刷新时递归刷新。
    """

    selectionChanged = Signal()

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        prefix: str = "人员",
        all_text: str = "全部",
        none_text: str = "未选择",
        hint: str = "",
        inline_names: int = 2,
        menu_width: int = 264,
        max_visible: int = 10,
    ) -> None:
        super().__init__(parent)
        self._prefix = prefix
        self._all_text = all_text
        self._none_text = none_text
        self._hint = hint
        self._inline_names = max(1, inline_names)
        self._max_visible = max(1, max_visible)
        self._items: list[tuple[object, str]] = []
        self._updating = False  # 重建清单期间屏蔽勾选信号

        self.setPopupMode(QToolButton.InstantPopup)
        self.setToolButtonStyle(Qt.ToolButtonTextOnly)
        self._menu = QMenu(self)

        panel = QWidget()
        box = QVBoxLayout(panel)
        box.setContentsMargins(8, 8, 8, 8)
        box.setSpacing(6)
        quick = QHBoxLayout()
        quick.setSpacing(6)
        btn_all = QPushButton("全选")
        btn_none = QPushButton("全不选")
        btn_all.clicked.connect(lambda: self.set_all_checked(True))
        btn_none.clicked.connect(lambda: self.set_all_checked(False))
        quick.addWidget(btn_all)
        quick.addWidget(btn_none)
        quick.addStretch(1)
        box.addLayout(quick)

        self._list = _CheckList()
        self._list.setObjectName("checklistPopup")
        self._list.setUniformItemSizes(True)
        self._list.setFixedWidth(menu_width - 16)
        self._list.itemChanged.connect(self._on_item_changed)
        box.addWidget(self._list)

        action = QWidgetAction(self._menu)
        action.setDefaultWidget(panel)
        self._menu.addAction(action)
        self.setMenu(self._menu)
        self._refresh_text()

    # ---------- 数据同步 ----------

    def set_items(
        self,
        items: list[tuple[object, str]],
        *,
        default_checked: bool = True,
    ) -> None:
        """同步清单项；已存在的项保留勾选状态，新增项用 default_checked。

        清单内容完全一致时直接返回：用户在弹层里勾选会触发页面刷新，刷新又
        回来同步清单，若无条件重建会打断正在进行的勾选（列表跳动、弹层收窄）。
        """
        items = [(value, label) for value, label in items]
        if items == self._items:
            return
        previous = self._states()
        self._items = items
        self._updating = True
        try:
            self._list.clear()
            for value, label in items:
                item = QListWidgetItem(label)
                item.setData(Qt.UserRole, value)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                checked = previous.get(value, default_checked)
                item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
                self._list.addItem(item)
        finally:
            self._updating = False
        self._fit_height()
        self._refresh_text()

    def _fit_height(self) -> None:
        """按可见项数定高，超过 max_visible 后交给滚动条。"""
        count = self._list.count()
        row_height = self._list.sizeHintForRow(0) if count else 0
        row_height = row_height if row_height > 0 else 26
        visible = min(max(count, 1), self._max_visible)
        self._list.setFixedHeight(row_height * visible + 8)

    def _states(self) -> dict[object, bool]:
        """当前勾选状态：项值 -> 是否勾选。"""
        states: dict[object, bool] = {}
        for row in range(self._list.count()):
            item = self._list.item(row)
            states[item.data(Qt.UserRole)] = item.checkState() == Qt.Checked
        return states

    # ---------- 勾选状态读写 ----------

    @property
    def total_count(self) -> int:
        return len(self._items)

    def checked_values(self) -> set:
        """已勾选项的取值集合（顺序不敏感，便于直接当筛选条件用）。"""
        return {value for value, checked in self._states().items() if checked}

    def set_checked_values(self, values) -> None:
        """按给定取值批量勾选，未列出的项一律取消。"""
        targets = set(values)
        self._updating = True
        try:
            for row in range(self._list.count()):
                item = self._list.item(row)
                item.setCheckState(
                    Qt.Checked if item.data(Qt.UserRole) in targets
                    else Qt.Unchecked)
        finally:
            self._updating = False
        self._refresh_text()
        self.selectionChanged.emit()

    def set_all_checked(self, checked: bool) -> None:
        """全选 / 全不选（弹层里的快捷按钮）。"""
        state = Qt.Checked if checked else Qt.Unchecked
        if all(item.checkState() == state
               for item in (self._list.item(row)
                            for row in range(self._list.count()))):
            return
        self._updating = True
        try:
            for row in range(self._list.count()):
                self._list.item(row).setCheckState(state)
        finally:
            self._updating = False
        self._refresh_text()
        self.selectionChanged.emit()

    # ---------- 内部 ----------

    def _on_item_changed(self, _item: QListWidgetItem) -> None:
        if self._updating:
            return
        self._refresh_text()
        self.selectionChanged.emit()

    def _refresh_text(self) -> None:
        """按勾选情况生成按钮文字与提示；三种状态各有固定写法。"""
        states = self._states()
        total = len(self._items)
        names = [label for value, label in self._items if states.get(value)]
        count = len(names)
        if total == 0:
            summary = "—"
        elif count == 0:
            summary = self._none_text
        elif count == total:
            summary = f"{self._all_text}（{total}）"
        else:
            summary = f"{'、'.join(names[: self._inline_names])} 等 {count}/{total}"
        self.setText(f"{self._prefix}：{summary}")

        if not total:
            detail = "没有可筛选的项"
        elif count == total:
            detail = f"已选全部 {total} 项"
        else:
            detail = f"已选 {count}/{total}：{'、'.join(names) or '无'}"
        self.setToolTip(f"{self._hint}\n{detail}" if self._hint else detail)


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
