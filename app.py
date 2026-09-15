"""自动值班排班系统 — PySide6 桌面应用

功能：上传成员个人课表(.xls/.xlsx) -> 解析入库(SQLite) -> 按空闲时段生成排班表
     （避免课程冲突、每人每天只值一次、均衡分配）-> 界面展示与导出(Excel/CSV)
     空闲甘特图：按周查看各成员忙闲、高亮全员空闲时段，方便安排任务

运行：python app.py
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor, QImage, QPainter, QPen
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QComboBox, QFileDialog, QFormLayout,
    QGroupBox, QHBoxLayout, QHeaderView, QLabel, QListWidget, QListWidgetItem,
    QMainWindow, QMessageBox, QPushButton, QSpinBox, QSplitter, QTabWidget,
    QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

from duty_system.database import Database
from duty_system.exporter import (
    build_detail_df, build_pivot_df, build_stats_df, export_csv, export_excel,
)
from duty_system.gantt import AvailabilityMatrix, build_availability, export_gantt_excel, slot_header
from duty_system.parser import BLOCK_LABELS, WEEKDAY_LABELS, parse_schedule_file
from duty_system.scheduler import ScheduleConfig, ScheduleResult, generate_schedule

DB_PATH = Path(__file__).parent / "duty_system.db"

# 类苹果设计语言（macOS 浅色模式）：
#   背景 #f5f5f7 / 卡片白色圆角 / 系统蓝 #007aff / 文字 #1d1d1f·#86868b
#   分隔线 #e5e5ea / 控件边框 #d2d2d7 / 8pt 间距网格 / 分段控件式 Tab
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

QSpinBox, QComboBox {
    background: #ffffff; color: #1d1d1f;
    border: 1px solid #d2d2d7; border-radius: 6px;
    padding: 3px 8px;
    selection-background-color: #007aff; selection-color: #ffffff;
}
QSpinBox:focus, QComboBox:focus { border: 1px solid #007aff; }
QSpinBox::up-button, QSpinBox::down-button { width: 0; border: none; background: none; }
QComboBox::drop-down { border: none; width: 24px; }
QComboBox QAbstractItemView {
    background: #ffffff; border: 1px solid #d2d2d7; border-radius: 8px;
    selection-background-color: #007aff; selection-color: #ffffff;
    outline: none; padding: 2px;
}

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
    """生成复选框选中态的白色对勾图标，供 QSS image 引用"""
    img = QImage(64, 64, QImage.Format_ARGB32)
    img.fill(Qt.transparent)
    painter = QPainter(img)
    painter.setRenderHint(QPainter.Antialiasing)
    pen = QPen(QColor("#ffffff"), 8)
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    painter.setPen(pen)
    painter.drawLine(16, 34, 27, 45)
    painter.drawLine(27, 45, 48, 18)
    painter.end()
    path = Path(tempfile.gettempdir()) / "duty_check_icon.png"
    img.save(str(path))
    return path.as_posix()


def build_style() -> str:
    """装配全局样式表（注入运行时生成的对勾图标路径）"""
    return STYLE.replace("__CHECK__", _write_check_icon())


class Card(QWidget):
    """白色圆角卡片容器（类苹果表面样式）"""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setAttribute(Qt.WA_StyledBackground, True)


def _card(widget: QWidget) -> Card:
    """把控件包进白色圆角卡片"""
    card = Card()
    lay = QVBoxLayout(card)
    lay.setContentsMargins(0, 0, 0, 0)
    lay.addWidget(widget)
    return card


def fill_table(table: QTableWidget, df: pd.DataFrame) -> None:
    """把 DataFrame 填入只读 QTableWidget（苹果风格：无网格线、交替行、隐藏行号）"""
    table.clearContents()
    table.setColumnCount(len(df.columns))
    table.setRowCount(len(df))
    table.setHorizontalHeaderLabels([str(c) for c in df.columns])
    table.setShowGrid(False)
    table.setAlternatingRowColors(True)
    table.verticalHeader().setVisible(False)
    table.verticalHeader().setDefaultSectionSize(30)
    for r in range(len(df)):
        for c in range(len(df.columns)):
            item = QTableWidgetItem(str(df.iat[r, c]))
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            table.setItem(r, c, item)
    header = table.horizontalHeader()
    header.setSectionResizeMode(QHeaderView.ResizeToContents)
    header.setStretchLastSection(True)


class MainWindow(QMainWindow):
    def __init__(self, db_path: str | Path = DB_PATH):
        super().__init__()
        self.db = Database(db_path)
        self.result: ScheduleResult | None = None
        self.gantt_matrix: AvailabilityMatrix | None = None
        self.setWindowTitle("自动值班排班系统")
        self.resize(1280, 800)
        self._build_ui()
        self.refresh_members()
        self.statusBar().showMessage("就绪。请上传成员课表。")

    # ---------- 界面构建 ----------

    def _build_ui(self) -> None:
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self._build_left_panel())
        splitter.addWidget(self._build_right_panel())
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([340, 940])
        self.setCentralWidget(splitter)

    def _build_left_panel(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(16, 16, 8, 16)
        layout.setSpacing(12)

        title = QLabel("自动值班排班系统")
        title.setObjectName("appTitle")
        layout.addWidget(title)
        subtitle = QLabel("解析课表 · 智能排班 · 空闲甘特")
        subtitle.setObjectName("appSubtitle")
        layout.addWidget(subtitle)

        btn_upload = QPushButton("上传成员课表…")
        btn_upload.setObjectName("primary")
        btn_upload.setMinimumHeight(36)
        btn_upload.clicked.connect(self.upload_files)
        layout.addWidget(btn_upload)

        grp_members = QGroupBox("成员")
        members_layout = QVBoxLayout(grp_members)
        members_layout.setContentsMargins(8, 4, 8, 8)
        members_layout.setSpacing(8)
        self.member_list = QListWidget()
        self.member_list.setSelectionMode(QAbstractItemView.SingleSelection)
        self.member_list.setMinimumHeight(120)
        members_layout.addWidget(self.member_list)
        btn_remove = QPushButton("删除选中成员")
        btn_remove.setObjectName("danger")
        btn_remove.clicked.connect(self.remove_selected_member)
        members_layout.addWidget(btn_remove)
        layout.addWidget(grp_members, stretch=1)

        grp_cfg = QGroupBox("排班参数")
        form = QFormLayout(grp_cfg)
        form.setContentsMargins(8, 4, 8, 8)
        form.setSpacing(10)
        week_row = QHBoxLayout()
        self.week_from = QSpinBox()
        self.week_from.setRange(1, 25)
        self.week_from.setValue(1)
        self.week_to = QSpinBox()
        self.week_to.setRange(1, 25)
        self.week_to.setValue(18)
        week_row.addWidget(self.week_from)
        week_row.addWidget(QLabel("至"))
        week_row.addWidget(self.week_to)
        form.addRow("值班周范围", week_row)

        self.weekday_checks: dict[int, QCheckBox] = {}
        wd_row = QHBoxLayout()
        for d in (1, 2, 3, 4, 5, 6, 7):
            cb = QCheckBox(WEEKDAY_LABELS[d])
            cb.setChecked(d <= 5)
            cb.stateChanged.connect(self._on_gantt_filter_changed)
            self.weekday_checks[d] = cb
            wd_row.addWidget(cb)
        form.addRow("值班星期", wd_row)

        self.block_checks: dict[int, QCheckBox] = {}
        blk_row = QHBoxLayout()
        for b in (1, 2, 3, 4, 5):
            cb = QCheckBox(BLOCK_LABELS[b].split(" ")[0])
            cb.setChecked(True)
            cb.setToolTip(BLOCK_LABELS[b])
            cb.stateChanged.connect(self._on_gantt_filter_changed)
            self.block_checks[b] = cb
            blk_row.addWidget(cb)
        form.addRow("值班时段", blk_row)

        num_row1 = QHBoxLayout()
        self.per_slot = QSpinBox()
        self.per_slot.setRange(1, 5)
        self.max_week = QSpinBox()
        self.max_week.setRange(1, 10)
        self.max_week.setValue(3)
        self.max_week.setToolTip("每人每周最多值班次数")
        num_row1.addWidget(QLabel("每时段人数"))
        num_row1.addWidget(self.per_slot)
        num_row1.addWidget(QLabel("每周上限"))
        num_row1.addWidget(self.max_week)
        form.addRow("", num_row1)

        num_row2 = QHBoxLayout()
        self.max_day = QSpinBox()
        self.max_day.setRange(1, 5)
        self.max_day.setValue(1)
        self.max_day.setToolTip("每人每天最多值班次数")
        self.seed = QSpinBox()
        self.seed.setRange(0, 9999)
        self.seed.setValue(42)
        self.seed.setToolTip("平手时决定分给谁，固定种子可复现")
        num_row2.addWidget(QLabel("每天上限"))
        num_row2.addWidget(self.max_day)
        num_row2.addWidget(QLabel("随机种子"))
        num_row2.addWidget(self.seed)
        form.addRow("", num_row2)
        layout.addWidget(grp_cfg)

        self.btn_generate = QPushButton("生成排班表")
        self.btn_generate.setObjectName("primary")
        self.btn_generate.setMinimumHeight(42)
        self.btn_generate.clicked.connect(self.generate)
        layout.addWidget(self.btn_generate)
        return panel

    def _build_right_panel(self) -> QWidget:
        self.tabs = QTabWidget()

        # Tab1 排班总表
        tab1 = QWidget()
        v1 = QVBoxLayout(tab1)
        v1.setContentsMargins(16, 14, 16, 16)
        v1.setSpacing(12)
        self.summary_label = QLabel("尚未生成排班表。设置左侧参数后点击「生成排班表」。")
        self.summary_label.setObjectName("summary")
        self.summary_label.setWordWrap(True)
        v1.addWidget(self.summary_label)
        btn_row = QHBoxLayout()
        self.btn_export_xlsx = QPushButton("导出 Excel")
        self.btn_export_xlsx.clicked.connect(self.export_xlsx)
        self.btn_export_csv = QPushButton("导出 CSV")
        self.btn_export_csv.clicked.connect(self.export_csv)
        self.btn_export_xlsx.setEnabled(False)
        self.btn_export_csv.setEnabled(False)
        btn_row.addWidget(self.btn_export_xlsx)
        btn_row.addWidget(self.btn_export_csv)
        btn_row.addStretch()
        v1.addLayout(btn_row)
        self.pivot_table = QTableWidget()
        v1.addWidget(_card(self.pivot_table), stretch=1)
        self.tabs.addTab(tab1, "值班排班表")

        # Tab2 值班明细
        tab2 = QWidget()
        v2 = QVBoxLayout(tab2)
        v2.setContentsMargins(16, 14, 16, 16)
        self.detail_table = QTableWidget()
        v2.addWidget(_card(self.detail_table))
        self.tabs.addTab(tab2, "值班明细")

        # Tab3 成员课表
        tab3 = QWidget()
        v3 = QVBoxLayout(tab3)
        v3.setContentsMargins(16, 14, 16, 16)
        v3.setSpacing(12)
        sel_row = QHBoxLayout()
        sel_row.addWidget(QLabel("选择成员"))
        self.member_combo = QComboBox()
        self.member_combo.setMinimumWidth(220)
        self.member_combo.currentIndexChanged.connect(self.show_member_courses)
        sel_row.addWidget(self.member_combo)
        sel_row.addStretch()
        v3.addLayout(sel_row)
        self.member_info = QLabel("")
        self.member_info.setObjectName("secondary")
        v3.addWidget(self.member_info)
        self.course_table = QTableWidget()
        v3.addWidget(_card(self.course_table), stretch=1)
        self.tabs.addTab(tab3, "成员课表")

        # Tab4 值班统计
        tab4 = QWidget()
        v4 = QVBoxLayout(tab4)
        v4.setContentsMargins(16, 14, 16, 16)
        v4.setSpacing(12)
        stats_title = QLabel("各成员值班次数统计（均衡性参考）")
        stats_title.setObjectName("secondary")
        v4.addWidget(stats_title)
        self.stats_table = QTableWidget()
        v4.addWidget(_card(self.stats_table), stretch=2)
        self.gap_title = QLabel("")
        self.gap_title.setObjectName("secondary")
        v4.addWidget(self.gap_title)
        self.gap_table = QTableWidget()
        v4.addWidget(_card(self.gap_table), stretch=1)
        self.tabs.addTab(tab4, "值班统计")

        # Tab5 空闲甘特图：成员空闲时段一览，方便安排任务
        tab5 = QWidget()
        v5 = QVBoxLayout(tab5)
        v5.setContentsMargins(16, 14, 16, 16)
        v5.setSpacing(12)
        ctl = QHBoxLayout()
        ctl.addWidget(QLabel("查看周次"))
        self.gantt_week = QSpinBox()
        self.gantt_week.setRange(1, 25)
        self.gantt_week.setValue(1)
        self.gantt_week.setToolTip("甘特图按周查看（课程随周次变化）")
        self.gantt_week.valueChanged.connect(self.refresh_gantt)
        ctl.addWidget(self.gantt_week)
        legend = QLabel(
            '<span style="background:#c9f2cf;">&nbsp;&nbsp;&nbsp;&nbsp;</span> 空闲&nbsp;&nbsp;'
            '<span style="background:#f2f2f7;">&nbsp;&nbsp;&nbsp;&nbsp;</span> 有课&nbsp;&nbsp;'
            '<span style="background:#ffd9a8;">&nbsp;&nbsp;&nbsp;&nbsp;</span> 全员空闲')
        legend.setObjectName("secondary")
        ctl.addWidget(legend)
        ctl.addStretch()
        self.btn_export_gantt = QPushButton("导出甘特图 Excel")
        self.btn_export_gantt.clicked.connect(self.export_gantt)
        self.btn_export_gantt.setEnabled(False)
        ctl.addWidget(self.btn_export_gantt)
        v5.addLayout(ctl)
        self.gantt_hint = QLabel("")
        self.gantt_hint.setObjectName("summary")
        self.gantt_hint.setWordWrap(True)
        v5.addWidget(self.gantt_hint)
        self.gantt_table = QTableWidget()
        v5.addWidget(_card(self.gantt_table), stretch=1)
        self.tabs.addTab(tab5, "空闲甘特图")

        self.tabs.currentChanged.connect(self._on_tab_changed)
        return self.tabs

    # ---------- 数据刷新 ----------

    def refresh_members(self) -> None:
        members = self.db.list_members()
        self.member_list.clear()
        self.member_combo.blockSignals(True)
        self.member_combo.clear()
        for m in members:
            self.member_list.addItem(f"{m.name}（{m.course_count} 门课）")
            item = self.member_list.item(self.member_list.count() - 1)
            item.setData(Qt.UserRole, m.id)
            self.member_combo.addItem(f"{m.name}（{m.class_name or '未知班级'}）", m.id)
        self.member_combo.blockSignals(False)
        if self.member_combo.count():
            self.show_member_courses(0)
        else:
            self.member_info.setText("")
            fill_table(self.course_table, pd.DataFrame(columns=["星期", "课程", "教师", "周次", "节次", "地点"]))
        self.refresh_gantt()

    def refresh_schedule_tabs(self) -> None:
        result = self.result
        if result is None:
            return
        pivot = build_pivot_df(result.assignments)
        display = pd.concat(
            [pd.DataFrame(pivot.index.tolist(), columns=["周次", "星期"]),
             pivot.reset_index(drop=True)], axis=1)
        fill_table(self.pivot_table, display)
        fill_table(self.detail_table, build_detail_df(result.assignments))
        fill_table(self.stats_table, build_stats_df(result.member_stats))

        n = len(result.assignments)
        self.summary_label.setText(
            f"排班完成：共 {n} 人次安排 | 参与成员 {len(result.member_stats)} 人 | "
            f"人均 {n / max(len(result.member_stats), 1):.1f} 次 | "
            f"总次数极差 {result.balanced_spread}（越小越均衡）| "
            f"无人可用时段 {len(result.gaps)} 个（成员有课、已达每周/每天上限或当天已值过）")
        self.btn_export_xlsx.setEnabled(bool(result.assignments))
        self.btn_export_csv.setEnabled(bool(result.assignments))

        if result.gaps:
            self.gap_title.setText(f"无人可值时段（{len(result.gaps)} 个）")
            gap_df = pd.DataFrame([{
                "周次": f"第{w}周", "星期": WEEKDAY_LABELS[d], "时段": BLOCK_LABELS[b],
            } for w, d, b in result.gaps])
            fill_table(self.gap_table, gap_df)
        else:
            self.gap_title.setText("所有值班时段均已安排到位，无缺口。")
            fill_table(self.gap_table, pd.DataFrame(columns=["周次", "星期", "时段"]))

    def show_member_courses(self, index: int) -> None:
        if index < 0:
            return
        member_id = self.member_combo.itemData(index)
        m = self.db.get_member(member_id)
        if m is None:
            return
        self.member_info.setText(
            f"学号 {m.student_id or '—'} | {m.term or '—'} | {m.major or '—'} | "
            f"{m.department or '—'} | 来源：{m.file_name or '—'}")
        courses = sorted(self.db.get_courses(member_id),
                         key=lambda c: (c.weekday, min(c.session_list), c.course_name))
        df = pd.DataFrame([{
            "星期": WEEKDAY_LABELS[c.weekday], "课程": c.course_name,
            "教师": c.teacher or "—", "周次": c.weeks_text,
            "节次": f"{c.sessions_text}节", "地点": c.location or "—",
        } for c in courses])
        fill_table(self.course_table, df if not df.empty else
                   pd.DataFrame(columns=["星期", "课程", "教师", "周次", "节次", "地点"]))

    # ---------- 空闲甘特图 ----------

    def _on_tab_changed(self, index: int) -> None:
        if self.tabs.tabText(index) == "空闲甘特图":
            self.refresh_gantt()

    def _on_gantt_filter_changed(self) -> None:
        if self.tabs.tabText(self.tabs.currentIndex()) == "空闲甘特图":
            self.refresh_gantt()

    def refresh_gantt(self) -> None:
        members = self.db.list_members()
        weekdays = [d for d, cb in self.weekday_checks.items() if cb.isChecked()]
        blocks = [b for b, cb in self.block_checks.items() if cb.isChecked()]
        if not members or not weekdays or not blocks:
            self.gantt_matrix = None
            self.gantt_table.clearContents()
            self.gantt_table.setRowCount(0)
            self.gantt_table.setColumnCount(0)
            self.gantt_hint.setText("请先上传成员课表，并至少勾选一个值班星期和值班时段（甘特图跟随左侧筛选）。")
            self.btn_export_gantt.setEnabled(False)
            return

        matrix = build_availability(
            members, self.db.get_courses(),
            self.gantt_week.value(), weekdays, blocks)
        self.gantt_matrix = matrix
        self._fill_gantt_table(matrix)

        all_free = matrix.all_free_slots
        if all_free:
            text = "、".join(
                f"{WEEKDAY_LABELS[d]} {BLOCK_LABELS[b].split(' ')[0]}" for d, b in all_free)
            self.gantt_hint.setText(
                f"第{matrix.week}周全员空闲时段共 {len(all_free)} 个（橙色高亮列）：{text}。"
                "适合安排需要全员参加的任务。")
        else:
            self.gantt_hint.setText(
                f"第{matrix.week}周没有全员空闲的时段；汇总行为各时段空闲人数，"
                "选择空闲人数最多的时段最容易凑齐人。")
        self.btn_export_gantt.setEnabled(True)

    def _fill_gantt_table(self, m: AvailabilityMatrix) -> None:
        """行=成员（末行为汇总），列=星期x时段；空闲绿色、有课灰色、全员空闲橙色"""
        t = self.gantt_table
        FREE, BUSY, ALL_FREE = QColor("#c9f2cf"), QColor("#f2f2f7"), QColor("#ffd9a8")
        t.clearContents()
        t.setRowCount(m.member_count + 1)
        t.setColumnCount(len(m.slots))
        t.setVerticalHeaderLabels(m.member_names + ["空闲人数"])
        t.setHorizontalHeaderLabels([slot_header(d, b) for d, b in m.slots])

        for r in range(m.member_count):
            for i, (d, b) in enumerate(m.slots):
                item = QTableWidgetItem("")
                item.setFlags(Qt.ItemIsEnabled)
                if m.free[r][i]:
                    item.setBackground(FREE)
                    item.setToolTip(f"第{m.week}周 {WEEKDAY_LABELS[d]} {BLOCK_LABELS[b]}：空闲")
                else:
                    names = m.busy_courses.get((r, d, b), [])
                    item.setBackground(BUSY)
                    item.setText("课")
                    item.setForeground(QBrush(QColor("#aeaeb2")))
                    item.setTextAlignment(Qt.AlignCenter)
                    item.setToolTip(
                        f"第{m.week}周 {WEEKDAY_LABELS[d]} {BLOCK_LABELS[b]}：\n"
                        + "\n".join(names))
                t.setItem(r, i, item)

        r = m.member_count
        for i, (d, b) in enumerate(m.slots):
            all_free = m.free_counts[i] == m.member_count
            item = QTableWidgetItem(f"{m.free_counts[i]}/{m.member_count}")
            item.setFlags(Qt.ItemIsEnabled)
            item.setTextAlignment(Qt.AlignCenter)
            if all_free:
                item.setBackground(ALL_FREE)
                item.setForeground(QBrush(QColor("#b25e00")))
            t.setItem(r, i, item)
            head = t.horizontalHeaderItem(i)
            head.setForeground(QBrush(QColor("#ff9500" if all_free else "#86868b")))

        t.verticalHeader().setDefaultSectionSize(28)
        header = t.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setStretchLastSection(False)
        header.setMinimumHeight(44)

    def export_gantt(self) -> None:
        m = self.gantt_matrix
        if m is None:
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "导出甘特图 Excel", f"空闲甘特图_第{m.week}周.xlsx", "Excel 文件 (*.xlsx)")
        if not path:
            return
        Path(path).write_bytes(export_gantt_excel(m))
        self.statusBar().showMessage(f"已导出甘特图 Excel：{path}")

    # ---------- 动作 ----------

    def upload_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择成员课表文件", str(Path.home()),
            "课表文件 (*.xls *.xlsx);;所有文件 (*)")
        if not files:
            return
        ok_count, errors = 0, []
        for f in files:
            try:
                schedule = parse_schedule_file(Path(f).read_bytes(), Path(f).name)
            except Exception as e:
                errors.append(f"{Path(f).name}：{e}")
                continue
            if not schedule.courses and not schedule.name:
                errors.append(f"{Path(f).name}：未解析到课程信息，请确认是教务系统导出的个人课表")
                continue
            self.db.upsert_member(schedule)
            ok_count += 1
        self.refresh_members()
        if errors:
            QMessageBox.warning(self, "部分文件导入失败", "\n".join(errors))
        self.statusBar().showMessage(f"已导入 {ok_count} 份课表，当前共 {len(self.db.list_members())} 名成员。")

    def remove_selected_member(self) -> None:
        item = self.member_list.currentItem()
        if item is None:
            QMessageBox.information(self, "提示", "请先在成员列表中选择要删除的成员。")
            return
        member_id = item.data(Qt.UserRole)
        m = self.db.get_member(member_id)
        if m is None:
            return
        if QMessageBox.question(
                self, "确认删除",
                f"确定删除成员「{m.name}」？其课程与排班记录将一并移除。") == QMessageBox.Yes:
            self.db.delete_member(member_id)
            self.refresh_members()
            self.statusBar().showMessage(f"已删除成员 {m.name}。")

    def generate(self) -> None:
        members = self.db.list_members()
        if not members:
            QMessageBox.information(self, "提示", "请先上传成员课表。")
            return
        weekdays = [d for d, cb in self.weekday_checks.items() if cb.isChecked()]
        blocks = [b for b, cb in self.block_checks.items() if cb.isChecked()]
        if not weekdays or not blocks:
            QMessageBox.warning(self, "提示", "请至少选择一个值班星期和值班时段。")
            return
        if self.week_from.value() > self.week_to.value():
            QMessageBox.warning(self, "提示", "起始周不能大于结束周。")
            return
        config = ScheduleConfig(
            weeks=range(self.week_from.value(), self.week_to.value() + 1),
            weekdays=sorted(weekdays), blocks=sorted(blocks),
            per_slot=self.per_slot.value(),
            max_per_week=self.max_week.value(),
            max_per_day=self.max_day.value(),
            seed=self.seed.value(),
        )
        result = generate_schedule(members, self.db.get_courses(), config)
        self.db.clear_assignments()
        self.db.save_assignments(result.assignments)
        self.result = result
        self.refresh_schedule_tabs()
        self.tabs.setCurrentIndex(0)
        self.statusBar().showMessage(
            f"排班完成：{len(result.assignments)} 人次，无人可用时段 {len(result.gaps)} 个，"
            f"总次数极差 {result.balanced_spread}。")

    def export_xlsx(self) -> None:
        if self.result is None:
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "导出 Excel", "值班排班表.xlsx", "Excel 文件 (*.xlsx)")
        if not path:
            return
        Path(path).write_bytes(
            export_excel(self.result.assignments, self.result.member_stats, self.result.gaps))
        self.statusBar().showMessage(f"已导出 Excel：{path}")

    def export_csv(self) -> None:
        if self.result is None:
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "导出 CSV", "值班排班表.csv", "CSV 文件 (*.csv)")
        if not path:
            return
        Path(path).write_bytes(export_csv(self.result.assignments))
        self.statusBar().showMessage(f"已导出 CSV：{path}")


def main() -> None:
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(build_style())
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
