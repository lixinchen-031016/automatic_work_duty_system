"""花名册管理页：增删改查、筛选与匹配查找。

界面分三层：
- 顶部：匹配查找（关键词 + 查找范围 + 匹配方式）与筛选（工作室 / 职位 / 数据完整性）；
- 中部：花名册表格，命中的字段高亮，双击即可查看并修改；
- 底部：新增 / 编辑 / 删除 / 复制 / 撤销 / 重做 / 导入 / 导出。

写操作统一走 Database 的花名册接口；数据库层会在每次变动后把未手工锁定的
成员按新花名册重建工作室归属，因此界面只需在写入后刷新自己的缓存。
"""

from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import replace

from PySide6.QtCore import QItemSelectionModel, Qt
from PySide6.QtGui import QBrush, QColor, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QFormLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .database import Database
from .roster import (
    POSITION_LEADER,
    POSITION_MEMBER,
    ROSTER_MATCH_MODES,
    ROSTER_SEARCH_SCOPES,
    ROSTER_STATUS_FILTERS,
    TARGET_AVAILABILITY_STUDIOS,
    UNKNOWN_STUDIO,
    RosterEntry,
    filter_roster_entries,
    match_roster_entry,
    position_options_for_studio,
    roster_duplicate_keys,
    roster_identity_key,
    roster_leader_conflicts,
    validate_roster_entries,
)
from .ui_utils import dialog_buttons, dialog_header

# 表格列与对应的 RosterEntry 字段（"来源"为展示列，不参与匹配查找）
_COLUMNS: tuple[tuple[str, str], ...] = (
    ("工作室", "studio"),
    ("职位", "position"),
    ("姓名", "name"),
    ("学号", "student_id"),
    ("手机号", "phone"),
    ("学院+专业", "college_major"),
    ("来源", "source_file"),
)
_SOURCE_COLUMN = len(_COLUMNS) - 1
_HIGHLIGHT_COLOR = QColor("#fff3bf")
_MATCH_COLOR = QColor("#1d4ed8")
_WARNING_COLOR = QColor("#b45309")


def _source_text(entry: RosterEntry) -> str:
    """来源列文本：导入文件位置，或标记为手工新增。"""
    if not (entry.source_file or "").strip():
        return "手工新增"
    if entry.row_number:
        return f"{entry.source_file} 第{entry.row_number}行"
    return entry.source_file


class RosterEntryDialog(QDialog):
    """新增 / 编辑单条花名册记录的表单。"""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        entry: RosterEntry | None = None,
        studios: Sequence[str] = (),
        title: str = "新增花名册记录",
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumWidth(480)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(10)
        layout.addWidget(dialog_header(
            title,
            "带 * 的字段必填；保存后会同步更新未手工锁定成员的工作室归属。"))

        form = QFormLayout()
        form.setSpacing(10)
        form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)

        self.studio_combo = QComboBox()
        self.studio_combo.setEditable(True)
        self.studio_combo.setInsertPolicy(QComboBox.NoInsert)
        options = sorted({
            *((studio or "").strip() for studio in studios),
            *TARGET_AVAILABILITY_STUDIOS,
            UNKNOWN_STUDIO,
        } - {""})
        self.studio_combo.addItems(options)

        self.position_combo = QComboBox()
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("必填，例如：张三")
        self.student_id_edit = QLineEdit()
        self.student_id_edit.setPlaceholderText("可留空；补齐后便于与课表成员匹配")
        self.phone_edit = QLineEdit()
        self.college_major_edit = QLineEdit()
        self.college_major_edit.setPlaceholderText("例如：计算机学院软件工程")

        form.addRow("工作室 *", self.studio_combo)
        form.addRow("职位", self.position_combo)
        form.addRow("姓名 *", self.name_edit)
        form.addRow("学号", self.student_id_edit)
        form.addRow("手机号", self.phone_edit)
        form.addRow("学院+专业", self.college_major_edit)
        layout.addLayout(form)

        self.studio_combo.currentTextChanged.connect(self._sync_positions)
        if entry is not None:
            self.studio_combo.setCurrentText(
                (entry.studio or "").strip() or UNKNOWN_STUDIO)
            self._sync_positions(self.studio_combo.currentText())
            self.position_combo.setCurrentIndex(max(
                self.position_combo.findData((entry.position or "").strip()), 0))
            self.name_edit.setText(entry.name or "")
            self.student_id_edit.setText(entry.student_id or "")
            self.phone_edit.setText(entry.phone or "")
            self.college_major_edit.setText(entry.college_major or "")
        else:
            self._sync_positions(self.studio_combo.currentText())
            self.position_combo.setCurrentIndex(max(
                self.position_combo.findData(POSITION_MEMBER), 0))

        buttons = dialog_buttons("保存")
        buttons.accepted.connect(self._try_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.name_edit.setFocus()

    def _sync_positions(self, studio: str) -> None:
        """按工作室刷新职位可选项（运营工作室额外支持高级顾问）。"""
        current = (self.position_combo.currentData() or "").strip()
        self.position_combo.blockSignals(True)
        self.position_combo.clear()
        for position in position_options_for_studio(studio):
            self.position_combo.addItem(position or "（暂不指定）", position)
        self.position_combo.setCurrentIndex(max(
            self.position_combo.findData(current), 0))
        self.position_combo.blockSignals(False)

    def _try_accept(self) -> None:
        if not self.name_edit.text().strip():
            QMessageBox.information(self, "提示", "请填写姓名。")
            self.name_edit.setFocus()
            return
        self.accept()

    def entry(self) -> RosterEntry:
        """把表单内容转换为花名册记录（entry_id 为 0，由数据库分配）。"""
        return RosterEntry(
            studio=self.studio_combo.currentText().strip() or UNKNOWN_STUDIO,
            name=self.name_edit.text().strip(),
            student_id=self.student_id_edit.text().strip(),
            position=(self.position_combo.currentData() or "").strip(),
            phone=self.phone_edit.text().strip(),
            college_major=self.college_major_edit.text().strip(),
        )


class RosterEditorDialog(QDialog):
    """花名册管理窗口：增删改查 + 筛选 + 匹配查找。"""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        db: Database,
        on_changed: Callable[[str], None] | None = None,
        import_callback: Callable[[Callable[[], None]], None] | None = None,
        export_callback: Callable[[], None] | None = None,
    ) -> None:
        super().__init__(parent)
        self.db = db
        self._on_changed = on_changed
        self._import_callback = import_callback
        self._export_callback = export_callback
        self.entries: list[RosterEntry] = db.list_roster_entries()
        self._visible: list[RosterEntry] = []
        self._undo_stack: list[list[RosterEntry]] = []
        self._redo_stack: list[list[RosterEntry]] = []
        self._undo_limit = 50
        self._last_message = ""
        self.setWindowTitle("花名册管理")
        self.setMinimumSize(1080, 640)
        self._build_ui()
        self.refresh()

    # ---------- 界面搭建 ----------

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(10)
        layout.addWidget(dialog_header(
            "花名册管理",
            "支持新增、修改、删除与查询；可按姓名 / 学号 / 手机号等字段匹配查找并筛选。"
            "修改后未手工锁定的成员会自动按花名册更新工作室归属。"))

        layout.addLayout(self._build_search_row())
        layout.addLayout(self._build_filter_row())

        self.table = QTableWidget()
        self.table.setColumnCount(len(_COLUMNS))
        self.table.setHorizontalHeaderLabels([label for label, _field in _COLUMNS])
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(30)
        header = self.table.horizontalHeader()
        for index in range(len(_COLUMNS)):
            header.setSectionResizeMode(index, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.Stretch)
        self.table.doubleClicked.connect(lambda _index: self.edit_selected())
        self.table.itemSelectionChanged.connect(self._update_buttons)
        layout.addWidget(self.table, stretch=1)

        layout.addLayout(self._build_action_row())

        self.status_label = QLabel()
        self.status_label.setObjectName("secondary")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        QShortcut(QKeySequence.StandardKey.Find, self,
                  activated=lambda: self.search_edit.setFocus())
        # Delete 只在表格获得焦点时生效，避免在其他输入控件里误删记录
        delete_shortcut = QShortcut(QKeySequence.StandardKey.Delete, self.table)
        delete_shortcut.setContext(Qt.WidgetShortcut)
        delete_shortcut.activated.connect(self.delete_selected)

        self.search_edit.setAccessibleName("花名册匹配查找")
        self.scope_combo.setAccessibleName("查找范围")
        self.mode_combo.setAccessibleName("匹配方式")
        self.studio_filter.setAccessibleName("工作室筛选")
        self.position_filter.setAccessibleName("职位筛选")
        self.status_filter.setAccessibleName("数据状态筛选")
        self.table.setAccessibleName("花名册表格")

    def _build_search_row(self) -> QHBoxLayout:
        row = QHBoxLayout()
        row.setSpacing(8)
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText(
            "匹配查找：姓名 / 学号 / 手机号 / 学院专业 / 工作室 / 职位")
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.setToolTip(
            "输入关键词即时筛选并高亮命中的字段；按 Enter 定位到第一条命中记录。\n"
            "学号 / 手机号中的空格与全角数字会自动归一化后比对。")
        self.search_edit.textChanged.connect(self.refresh)
        self.search_edit.returnPressed.connect(self.focus_first_match)
        row.addWidget(self.search_edit, stretch=1)

        self.scope_combo = QComboBox()
        self.scope_combo.setToolTip("限定关键词在哪个字段中匹配")
        for value, label in ROSTER_SEARCH_SCOPES:
            self.scope_combo.addItem(label, value)
        self.scope_combo.currentIndexChanged.connect(self.refresh)
        row.addWidget(self.scope_combo)

        self.mode_combo = QComboBox()
        self.mode_combo.setToolTip(
            "包含匹配：字段含关键词即可\n前缀匹配：从字段开头比对\n精确匹配：整个字段与关键词相同")
        for value, label in ROSTER_MATCH_MODES:
            self.mode_combo.addItem(label, value)
        self.mode_combo.currentIndexChanged.connect(self.refresh)
        row.addWidget(self.mode_combo)

        self.btn_clear_search = QPushButton("清除筛选")
        self.btn_clear_search.setToolTip("清空关键词、查找范围与全部筛选条件")
        self.btn_clear_search.clicked.connect(self.clear_filters)
        row.addWidget(self.btn_clear_search)
        return row

    def _build_filter_row(self) -> QHBoxLayout:
        row = QHBoxLayout()
        row.setSpacing(8)
        row.addWidget(QLabel("工作室"))
        self.studio_filter = QComboBox()
        self.studio_filter.setToolTip("按工作室筛选")
        self.studio_filter.currentIndexChanged.connect(self.refresh)
        row.addWidget(self.studio_filter)

        row.addWidget(QLabel("职位"))
        self.position_filter = QComboBox()
        self.position_filter.setToolTip("按职位筛选")
        self.position_filter.currentIndexChanged.connect(self.refresh)
        row.addWidget(self.position_filter)

        row.addWidget(QLabel("数据状态"))
        self.status_filter = QComboBox()
        self.status_filter.setToolTip(
            "筛选缺学号 / 缺手机号 / 同工作室重名 / 部长重复的记录")
        for value, label in ROSTER_STATUS_FILTERS:
            self.status_filter.addItem(label, value)
        self.status_filter.currentIndexChanged.connect(self.refresh)
        row.addWidget(self.status_filter)
        row.addStretch(1)
        return row

    def _build_action_row(self) -> QHBoxLayout:
        row = QHBoxLayout()
        row.setSpacing(8)
        self.btn_add = QPushButton("新增…")
        self.btn_add.setObjectName("primary")
        self.btn_add.setToolTip("新增一条花名册记录")
        self.btn_add.clicked.connect(self.add_entry)
        self.btn_edit = QPushButton("编辑…")
        self.btn_edit.setToolTip("修改选中的一条记录（也可双击表格行）")
        self.btn_edit.clicked.connect(self.edit_selected)
        self.btn_delete = QPushButton("删除")
        self.btn_delete.setObjectName("danger")
        self.btn_delete.setToolTip("删除选中记录（支持多选与 Delete 键）")
        self.btn_delete.clicked.connect(self.delete_selected)
        self.btn_duplicate = QPushButton("复制…")
        self.btn_duplicate.setToolTip("按选中记录复制新条目，适合补充兼任工作室")
        self.btn_duplicate.clicked.connect(self.duplicate_selected)
        self.btn_undo = QPushButton("撤销")
        self.btn_undo.setToolTip("撤回最近一次花名册修改")
        self.btn_undo.clicked.connect(self.undo_change)
        self.btn_redo = QPushButton("重做")
        self.btn_redo.setToolTip("恢复被撤销的花名册修改")
        self.btn_redo.clicked.connect(self.redo_change)
        for button in (
            self.btn_add, self.btn_edit, self.btn_delete, self.btn_duplicate,
            self.btn_undo, self.btn_redo,
        ):
            row.addWidget(button)
        row.addStretch(1)

        self.btn_import = QPushButton("导入花名册…")
        self.btn_import.setToolTip("调用「文件 → 导入花名册…」的冲突对照流程")
        self.btn_import.clicked.connect(self.request_import)
        self.btn_import.setEnabled(self._import_callback is not None)
        row.addWidget(self.btn_import)

        self.btn_export = QPushButton("导出花名册…")
        self.btn_export.setToolTip("导出完整花名册 Excel（含课表状态与课程明细）")
        self.btn_export.clicked.connect(self.request_export)
        self.btn_export.setEnabled(self._export_callback is not None)
        row.addWidget(self.btn_export)

        self.btn_close = QPushButton("关闭")
        self.btn_close.clicked.connect(self.accept)
        row.addWidget(self.btn_close)
        return row

    # ---------- 读取与展示 ----------

    def _reload(self) -> None:
        self.entries = self.db.list_roster_entries()

    def _selected_entries(self) -> list[RosterEntry]:
        """当前选中的记录；用单元格里保存的 id 定位以兼容表格排序。"""
        by_id = {entry.entry_id: entry for entry in self.entries}
        ids: list[int] = []
        model = self.table.selectionModel()
        for index in (model.selectedRows() if model is not None else []):
            item = self.table.item(index.row(), 0)
            if item is None:
                continue
            entry_id = item.data(Qt.UserRole)
            if entry_id is not None and int(entry_id) not in ids:
                ids.append(int(entry_id))
        return [by_id[entry_id] for entry_id in ids if entry_id in by_id]

    def _query(self) -> dict[str, str]:
        return {
            "query": self.search_edit.text(),
            "scope": self.scope_combo.currentData() or "all",
            "mode": self.mode_combo.currentData() or "contains",
            "studio": self.studio_filter.currentData() or "",
            "position": self.position_filter.currentData() or "",
            "status": self.status_filter.currentData() or "",
        }

    def _sync_filter_options(self) -> None:
        """刷新筛选下拉的可选项，并尽量保留用户当前选择。"""
        studios = sorted({
            (entry.studio or "").strip() for entry in self.entries
            if (entry.studio or "").strip()
        })
        positions = sorted({
            (entry.position or "").strip() for entry in self.entries
            if (entry.position or "").strip()
        })
        for combo, values, all_label in (
            (self.studio_filter, studios, "全部工作室"),
            (self.position_filter, positions, "全部职位"),
        ):
            current = combo.currentData() or ""
            combo.blockSignals(True)
            combo.clear()
            combo.addItem(all_label, "")
            for value in values:
                combo.addItem(value, value)
            combo.setCurrentIndex(max(combo.findData(current), 0))
            combo.blockSignals(False)

    def refresh(self, *_args) -> None:
        """按当前查找与筛选条件重建表格。"""
        keep_selected = {entry.entry_id for entry in self._selected_entries()}
        # 先同步下拉可选项：某个工作室被删光后，筛选项会自动回到「全部」，
        # 之后再读取条件，保证下拉显示与表格内容一致。
        self._sync_filter_options()
        query = self._query()
        self._visible = filter_roster_entries(self.entries, **query)
        self._fill_table(query)
        self._select_ids(keep_selected)
        self._update_status(query)
        self._update_buttons()

    def _select_ids(self, ids: set[int]) -> None:
        """在重建后的表格里恢复选中状态（兼容表格排序后的行序）。"""
        model = self.table.selectionModel()
        if not ids or model is None:
            return
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item is None:
                continue
            if int(item.data(Qt.UserRole) or 0) in ids:
                # 直接改 item 选中标记不会同步选择模型，这里显式选中整行
                model.select(
                    self.table.model().index(row, 0),
                    QItemSelectionModel.Select | QItemSelectionModel.Rows)

    def _highlight_columns(self, entry: RosterEntry, query: dict[str, str]) -> set[int]:
        """返回命中的列号；查找范围为「全部字段」时逐列判断。"""
        keyword = query["query"]
        if not keyword.strip():
            return set()
        columns: set[int] = set()
        for column, (_label, field) in enumerate(_COLUMNS):
            if column == _SOURCE_COLUMN:
                continue
            if query["scope"] not in ("all", field):
                continue
            if match_roster_entry(entry, keyword, scope=field, mode=query["mode"]):
                columns.add(column)
        return columns

    def _problem_text(
        self,
        entry: RosterEntry,
        duplicate_keys: set[tuple[str, str]],
        leader_conflicts: dict[str, int],
    ) -> str:
        problems: list[str] = []
        if not (entry.name or "").strip():
            problems.append("姓名为空，保存前必须补全")
        if roster_identity_key(entry) in duplicate_keys:
            problems.append("同一工作室下存在重名记录，请确认是兼任还是重复登记")
        studio = (entry.studio or UNKNOWN_STUDIO)
        if (entry.position or "").strip() == POSITION_LEADER and studio in leader_conflicts:
            problems.append(f"{studio} 指定了多名部长")
        return "；".join(problems)

    def _fill_table(self, query: dict[str, str]) -> None:
        duplicate_keys = roster_duplicate_keys(self.entries)
        leader_conflicts = roster_leader_conflicts(self.entries)
        self.table.setSortingEnabled(False)
        self.table.clearContents()
        self.table.setRowCount(len(self._visible))
        for row, entry in enumerate(self._visible):
            highlighted = self._highlight_columns(entry, query)
            problem_text = self._problem_text(entry, duplicate_keys, leader_conflicts)
            for column, (_label, field) in enumerate(_COLUMNS):
                text = (
                    _source_text(entry) if column == _SOURCE_COLUMN
                    else str(getattr(entry, field, "") or "")
                )
                item = QTableWidgetItem(text)
                if column in highlighted:
                    # 当前主题里 QSS 指定了 QTableWidget::item，模型设置的背景色
                    # 不会绘制；因此用加粗 + 前景色强调，背景色作为主题变化后的兜底。
                    item.setBackground(QBrush(_HIGHLIGHT_COLOR))
                    font = item.font()
                    font.setBold(True)
                    item.setFont(font)
                    item.setForeground(QBrush(_MATCH_COLOR))
                if column == 0:
                    item.setData(Qt.UserRole, entry.entry_id)
                    if problem_text:
                        if column not in highlighted:
                            item.setForeground(QBrush(_WARNING_COLOR))
                        item.setToolTip(problem_text)
                        item.setText(f"⚠ {text}")
                self.table.setItem(row, column, item)
        self.table.setSortingEnabled(True)

    def _update_status(self, query: dict[str, str]) -> None:
        parts = [f"共 {len(self.entries)} 条记录", f"显示 {len(self._visible)} 条"]
        keyword = (query["query"] or "").strip()
        if keyword:
            parts.append(f"命中「{keyword}」{len(self._visible)} 条")
        selected = len(self._selected_entries())
        if selected:
            parts.append(f"已选 {selected} 条")
        problems = validate_roster_entries(self.entries)
        hint = "；".join(problems) if problems else "数据检查通过"
        text = " · ".join(parts) + f"　|　{hint}"
        self.status_label.setText(
            f"{self._last_message}　|　{text}" if self._last_message else text)

    def _update_buttons(self, *_args) -> None:
        selected = len(self._selected_entries())
        self.btn_edit.setEnabled(selected == 1)
        self.btn_delete.setEnabled(selected > 0)
        self.btn_duplicate.setEnabled(selected > 0)
        self.btn_undo.setEnabled(bool(self._undo_stack))
        self.btn_redo.setEnabled(bool(self._redo_stack))

    # ---------- 匹配查找 ----------

    def focus_first_match(self) -> None:
        """按 Enter 时选中并滚动到第一条命中的记录（兼容表格排序后的行序）。"""
        if not self._visible:
            return
        target = self._visible[0].entry_id
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item is None or int(item.data(Qt.UserRole) or 0) != target:
                continue
            self.table.setCurrentCell(row, 0)
            self.table.scrollToItem(item)
            return

    def clear_filters(self) -> None:
        self.search_edit.clear()
        for combo in (
            self.scope_combo, self.mode_combo, self.studio_filter,
            self.position_filter, self.status_filter,
        ):
            combo.blockSignals(True)
            combo.setCurrentIndex(0)
            combo.blockSignals(False)
        self.refresh()

    # ---------- 增删改 ----------

    def _push_undo(self) -> None:
        self._undo_stack.append(list(self.entries))
        del self._undo_stack[:-self._undo_limit]
        self._redo_stack.clear()

    def _confirm_proposed(self, proposed: list[RosterEntry]) -> bool:
        """保存前校验：返回 True 表示可以继续写入。"""
        problems = validate_roster_entries(proposed)
        if not problems:
            return True
        if any("姓名为空" in problem for problem in problems):
            QMessageBox.warning(
                self, "无法保存",
                "花名册中存在姓名为空的记录，请补全后再保存：\n\n"
                + "\n".join(problems))
            return False
        answer = QMessageBox.question(
            self, "花名册校验提醒",
            "发现以下需要确认的问题：\n\n" + "\n".join(problems)
            + "\n\n是否仍然保存？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        return answer == QMessageBox.Yes

    def _notify(self, message: str) -> None:
        self._last_message = message
        self._update_status(self._query())
        if self._on_changed is not None:
            self._on_changed(message)

    def _commit(self, mutate: Callable[[], object], message: str) -> bool:
        """执行写操作；失败时回滚撤销栈并提示。"""
        self._push_undo()
        try:
            mutate()
        except ValueError as exc:
            self._undo_stack.pop()
            QMessageBox.warning(self, "无法保存", str(exc))
            return False
        self._reload()
        self.refresh()
        self._notify(message)
        return True

    def add_entry(self) -> None:
        dialog = RosterEntryDialog(
            self, studios=[entry.studio for entry in self.entries],
            title="新增花名册记录")
        if dialog.exec() != QDialog.Accepted:
            return
        entry = dialog.entry()
        if not self._confirm_proposed([*self.entries, entry]):
            return
        created: list[int] = []
        if self._commit(
                lambda: created.append(self.db.add_roster_entry(entry)),
                f"已新增花名册记录：{entry.name}（{entry.studio}）。"):
            self._select_ids(set(created))

    def edit_selected(self) -> None:
        selected = self._selected_entries()
        if len(selected) != 1:
            QMessageBox.information(self, "提示", "请先选择一条要修改的记录。")
            return
        original = selected[0]
        dialog = RosterEntryDialog(
            self, entry=original,
            studios=[entry.studio for entry in self.entries],
            title="修改花名册记录")
        if dialog.exec() != QDialog.Accepted:
            return
        updated = dialog.entry()
        proposed = [
            updated if item.entry_id == original.entry_id else item
            for item in self.entries
        ]
        if not self._confirm_proposed(proposed):
            return
        self._commit(
            lambda: self.db.update_roster_entry(original.entry_id, updated),
            f"已更新花名册记录：{updated.name}（{updated.studio}）。")

    def duplicate_selected(self) -> None:
        selected = self._selected_entries()
        if not selected:
            QMessageBox.information(self, "提示", "请先选择要复制的记录。")
            return
        copies = [
            replace(entry, entry_id=0, source_file="", row_number=0)
            for entry in selected
        ]

        copies_ids: list[int] = []

        def mutate() -> None:
            for copy in copies:
                copies_ids.append(self.db.add_roster_entry(copy))

        if self._commit(
                mutate,
                f"已复制 {len(copies)} 条花名册记录，请修改姓名与学号后再保存。"):
            self._select_ids(set(copies_ids))

    def delete_selected(self) -> None:
        selected = self._selected_entries()
        if not selected:
            QMessageBox.information(self, "提示", "请先选择要删除的记录。")
            return
        names = "、".join(entry.name or "（未命名）" for entry in selected[:5])
        if len(selected) > 5:
            names += f" 等 {len(selected)} 条"
        answer = QMessageBox.question(
            self, "确认删除",
            f"确定删除花名册记录：{names}？\n"
            "未手工锁定成员的工作室归属会随之更新。",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if answer != QMessageBox.Yes:
            return
        ids = [entry.entry_id for entry in selected]
        self._commit(
            lambda: self.db.delete_roster_entries(ids),
            f"已删除 {len(ids)} 条花名册记录。")

    def undo_change(self) -> None:
        if not self._undo_stack:
            return
        previous = self._undo_stack.pop()
        self._redo_stack.append(list(self.entries))
        self.db.restore_roster_entries(previous)
        self._reload()
        self.refresh()
        self._notify("已撤销上次花名册修改。")

    def redo_change(self) -> None:
        if not self._redo_stack:
            return
        following = self._redo_stack.pop()
        self._undo_stack.append(list(self.entries))
        self.db.restore_roster_entries(following)
        self._reload()
        self.refresh()
        self._notify("已重做花名册修改。")

    # ---------- 导入 / 导出 ----------

    def request_import(self) -> None:
        if self._import_callback is not None:
            self._import_callback(self.reload_after_import)

    def request_export(self) -> None:
        if self._export_callback is not None:
            self._export_callback()

    def reload_after_import(self) -> None:
        """导入流程结束后刷新表格；导入会重写整份花名册，故清空撤销栈。"""
        self._undo_stack.clear()
        self._redo_stack.clear()
        self._reload()
        self.refresh()
