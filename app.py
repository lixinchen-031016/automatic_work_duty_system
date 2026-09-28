"""自动值班排班系统 — PySide6 桌面应用

功能：上传成员个人课表(.xls/.xlsx) -> 解析入库(SQLite) -> 按空闲时段生成排班表
     （避免课程/长期特殊安排冲突、每人每天只值一次、均衡分配）-> 界面展示与导出(Excel/CSV)
     空闲时段总览：按周查看各成员忙闲、高亮全员空闲时段，方便安排任务

运行：python app.py
"""

from __future__ import annotations

import json
import os
import sys
from collections import Counter
from collections.abc import Callable
from datetime import date, timedelta
from itertools import pairwise
from pathlib import Path

import pandas as pd
import shiboken6
from PySide6.QtCore import (
    QDate,
    QObject,
    QRunnable,
    QSettings,
    Qt,
    QThreadPool,
    Signal,
)
from PySide6.QtGui import (
    QAction,
    QBrush,
    QColor,
    QFontMetrics,
    QKeySequence,
    QPixmap,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QSplitter,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from duty_system.calendar import (
    MAX_WEEK,
    CalendarEntry,
    TermCalendar,
    natural_date,
)
from duty_system.database import Assignment, Database, SpecialArrangement
from duty_system.exporter import (
    build_detail_df,
    build_gap_df,
    build_pivot_df,
    build_stats_df,
    export_csv,
    export_excel,
    export_leaves_excel,
    export_roster_excel,
)
from duty_system.gantt import (
    CalendarAvailabilityMatrix,
    build_calendar_availability,
    export_gantt_excel,
    slot_header,
)
from duty_system.holiday import (
    build_holiday_import_plan,
    fetch_holiday_data,
    get_builtin_holiday_data,
)
from duty_system.parser import (
    BLOCK_LABELS,
    BLOCK_SESSIONS,
    WEEKDAY_LABELS,
    parse_schedule_files,
)
from duty_system.rendering import BarChart, render_charts_png, render_table_png
from duty_system.roster import (
    TARGET_AVAILABILITY_STUDIOS,
    UNKNOWN_STUDIO,
    parse_roster_file,
)
from duty_system.scheduler import (
    ScheduleConfig,
    ScheduleResult,
    build_busy_map,
    build_special_busy_map,
    capacity_advice,
    compute_gaps,
    diagnose_gaps,
    generate_schedule,
    rebuild_member_stats,
    replacement_candidates,
    summarize_gap_causes,
)
from duty_system.theme import build_style
from duty_system.ui_utils import (
    card as _card,
)
from duty_system.ui_utils import (
    dialog_buttons,
    dialog_header,
    fill_table,
    special_sessions_label,
    special_weeks_label,
)

# 数据库默认位置：
#   源码运行 → 程序目录；打包运行（PyInstaller）→ Windows 在 exe 旁（绿色软件），
#   macOS 写入用户应用支持目录（.app 包内容不可写、签名不可破坏）；
#   应用支持目录不可写时回退到家目录 ~/.dutysystem，避免启动即崩溃
if getattr(sys, "frozen", False):
    if sys.platform == "darwin":
        _BASE = Path.home() / "Library" / "Application Support" / "DutySystem"
        try:
            _BASE.mkdir(parents=True, exist_ok=True)
        except OSError:
            _BASE = Path.home() / ".dutysystem"
            _BASE.mkdir(parents=True, exist_ok=True)
        DB_PATH = _BASE / "duty_system.db"
    else:
        preferred = Path(sys.executable).resolve().parent
        preferred_db = preferred / "duty_system.db"
        writable = (
            os.access(preferred_db, os.W_OK) if preferred_db.exists()
            else os.access(preferred, os.W_OK)
        )
        if writable:
            DB_PATH = preferred_db
        else:
            local_app_data = Path(
                os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
            _BASE = local_app_data / "DutySystem"
            try:
                _BASE.mkdir(parents=True, exist_ok=True)
            except OSError:
                _BASE = Path.home() / ".dutysystem"
                _BASE.mkdir(parents=True, exist_ok=True)
            DB_PATH = _BASE / "duty_system.db"
else:
    DB_PATH = Path(__file__).parent / "duty_system.db"

SHORTCUT_MODIFIER = "⌘" if sys.platform == "darwin" else "Ctrl+"


def resource_path(relative_path: str | Path) -> Path:
    """返回源码运行或 PyInstaller 解包目录中的资源路径。"""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative_path


# 类苹果设计语言（macOS 浅色模式）：
#   背景 #f3f6fb / 卡片白色圆角 / 主色 #2563eb / 文字 #132238·#66758a
#   分隔线 #e5e5ea / 控件边框 #d2d2d7 / 8pt 间距网格 / 分段控件式 Tab
class _TaskSignals(QObject):
    """后台任务的完成/失败信号（QRunnable 本身不能带信号，需 QObject 载体）。"""

    done = Signal(object)
    failed = Signal(str)


class _Task(QRunnable):
    """把耗时操作放到线程池执行，完成或失败后回主线程回调。"""

    def __init__(self, fn: Callable[[], object]):
        super().__init__()
        self._fn = fn
        self.signals = _TaskSignals()

    def run(self) -> None:  # pragma: no cover - Qt 线程池调用
        try:
            result = self._fn()
        except Exception as exc:  # noqa: BLE001 - 任何异常都要回传主线程展示
            self.signals.failed.emit(f"{type(exc).__name__}: {exc}")
        else:
            self.signals.done.emit(result)


class MainWindow(QMainWindow):
    def __init__(self, db_path: str | Path | None = None):
        super().__init__()
        if db_path is None:
            stored = QSettings().value("database/path", "", type=str)
            db_path = Path(stored) if stored and Path(stored).exists() else DB_PATH
        self.db = Database(db_path)
        self.result: ScheduleResult | None = None
        self.gantt_matrix: CalendarAvailabilityMatrix | None = None
        self._stale = False
        self._last_config: ScheduleConfig | None = None
        self._pivot_rows: list[tuple[int, int]] = []
        self._pivot_blocks: list[int] = []
        self._pivot_date_offset: int = 1
        # 数据缓存：成员 / 课程 / 请假在「读多写少」的界面上被反复全量查询，
        # 统一从这里取，写操作后调用 invalidate_cache() 失效即可。
        self._members_cache = None
        self._courses_cache = None
        self._leaves_cache = None
        self._specials_cache = None
        self._roster_cache = None
        self._calendar_cache = None
        self._term_calendar_cache: TermCalendar | None = None
        self._calendar_error: str | None = None
        # 忙时表由「成员 + 课程」唯一决定，展开成本高（课程数 x 周次 x 节次），
        # 空闲时段总览与微调对话框都会用到，因此与课表缓存同生命周期
        self._busy_cache = None
        self._special_busy_cache = None
        # 缺口诊断结果（按排班结果对象缓存，避免每次刷新重复诊断）
        self._gap_cache: tuple | None = None
        self._undo_stack: list[tuple[int, int, int, tuple[int, ...]]] = []
        self._redo_stack: list[tuple[int, int, int, tuple[int, ...]]] = []
        self._undo_limit = 50
        # 空闲时段总览单元格上次写入的状态：(行, 列) -> 状态元组，用于跳过无变化单元格
        self._gantt_cell_state: dict = {}
        self._dirty_schedule_sections: set[str] = set()
        self._pool = QThreadPool.globalInstance()
        self._busy = False
        # 必须持有正在执行的任务引用：QRunnable 被 Python 回收后线程池就无法再执行它
        self._active_task = None
        self._closing = False
        self.setWindowTitle("自动值班排班系统")
        self.setMinimumSize(1120, 700)
        self.resize(1440, 900)
        self._build_ui()
        self._load_settings()
        self.refresh_members()
        self._restore_result()
        self.statusBar().showMessage(
            "已恢复上次的排班结果。" if self.result is not None else "就绪。请上传成员课表。")

    # ---------- 数据缓存 ----------

    def invalidate_cache(self) -> None:
        """成员 / 课程 / 请假发生写入后调用，下次读取自动重查数据库"""
        self._members_cache = None
        self._courses_cache = None
        self._leaves_cache = None
        self._specials_cache = None
        self._roster_cache = None
        self._calendar_cache = None
        self._term_calendar_cache = None
        self._calendar_error = None
        self._busy_cache = None
        self._special_busy_cache = None
        self._gap_cache = None

    def members(self) -> list:
        if self._members_cache is None:
            self._members_cache = self.db.list_members()
        return self._members_cache

    def scheduling_members(self) -> list:
        """当前参与排班的成员；未勾选成员仍保留课表和历史数据。"""
        return [m for m in self.members() if m.participates_in_scheduling]

    def gantt_members(self) -> list:
        """空闲总览目标成员；导入花名册后仅显示短视频和图片工作室。"""
        members = self.scheduling_members()
        if not self.roster_entries():
            # 老数据未导入花名册时保持原有全量展示，避免升级后页面突然为空。
            return members
        targets = set(TARGET_AVAILABILITY_STUDIOS)
        return [m for m in members if (m.studio or UNKNOWN_STUDIO) in targets]

    def roster_entries(self) -> list:
        if self._roster_cache is None:
            self._roster_cache = self.db.list_roster_entries()
        return self._roster_cache

    def courses(self) -> list:
        if self._courses_cache is None:
            self._courses_cache = self.db.get_courses()
        return self._courses_cache

    def leaves(self) -> list:
        if self._leaves_cache is None:
            self._leaves_cache = self.db.list_leaves()
        return self._leaves_cache

    def specials(self) -> list:
        if self._specials_cache is None:
            self._specials_cache = self.db.list_special_arrangements()
        return self._specials_cache

    def calendar_entries(self) -> list[CalendarEntry]:
        if self._calendar_cache is None:
            self._calendar_cache = self.db.list_calendar()
        return self._calendar_cache

    def term_calendar(self) -> TermCalendar:
        """当前学期起始日 + 数据库日历覆盖的解析器。"""
        term_start = self._term_start()
        cached = self._term_calendar_cache
        if cached is None or cached.term_start != term_start:
            try:
                cached = TermCalendar(term_start, self.calendar_entries())
            except ValueError as exc:
                # 学期起始日调整后，历史日历项可能暂时越界。运行态先停用全部
                # 覆盖，避免空闲时段总览和排班刷新崩溃；保存/校验时仍会明确报错。
                self._calendar_error = str(exc)
                cached = TermCalendar(term_start, [])
            else:
                self._calendar_error = None
            self._term_calendar_cache = cached
        return cached

    def is_off_day(self, week: int, weekday: int) -> bool:
        """逻辑教学日是否为无真实日期的纯假日。"""
        return self.term_calendar().logical_to_date(week, weekday) is None

    def is_class_day(self, week: int, weekday: int) -> bool:
        """逻辑教学日是否对应一个周末补课日。"""
        return self.term_calendar().is_class_day(week, weekday)

    def busy_map(self) -> dict[int, set]:
        """课程忙时表（按 成员 -> {(周, 星期, 节次)}）"""
        if self._busy_cache is None:
            self._busy_cache = build_busy_map(self.members(), self.courses())
        return self._busy_cache

    def special_busy_map(self) -> dict[int, set]:
        """长期特殊安排忙时表（与课程忙时表分开，便于界面显示占用原因）。"""
        if self._special_busy_cache is None:
            self._special_busy_cache = build_special_busy_map(
                self.members(), self.specials())
        return self._special_busy_cache

    def special_set(self) -> set[tuple[int, int, int, int]]:
        """(成员, 周, 星期, 节次) 形式的长期特殊安排集合。"""
        return {
            (member_id, week, weekday, session)
            for member_id, sessions in self.special_busy_map().items()
            for week, weekday, session in sessions
        }

    # ---------- 后台任务 ----------

    @property
    def busy(self) -> bool:
        return self._busy

    def run_async(self, label: str, fn: Callable[[], object],
                  on_done: Callable[[object], None]) -> None:
        """后台执行 fn()，完成或失败后在主线程回调，期间给出忙碌提示。

        排班计算与 Excel/PNG 生成在成员较多时会阻塞主线程导致界面无响应，
        统一走线程池；同一时刻只允许一个后台任务，避免结果互相覆盖。
        """
        if self._closing or self._busy:
            if not self._closing:
                self.statusBar().showMessage("上一个任务尚未完成，请稍候…")
            return
        self._busy = True
        self.btn_generate.setEnabled(False)
        self.action_generate.setEnabled(False)
        self.centralWidget().setEnabled(False)
        self.menuBar().setEnabled(False)
        self._set_app_status(label, "info")
        self.statusBar().showMessage(f"{label}…")
        QApplication.setOverrideCursor(Qt.BusyCursor)
        task = _Task(fn)
        task.signals.done.connect(lambda result: self._finish_async(on_done, result))
        task.signals.failed.connect(lambda msg: self._finish_async(None, None, msg))
        self._active_task = task
        self._pool.start(task)

    def _finish_async(self, on_done, result, error: str | None = None) -> None:
        self._active_task = None
        if getattr(self, "_closing", False):
            return
        self._busy = False
        self.btn_generate.setEnabled(True)
        self.action_generate.setEnabled(True)
        self.centralWidget().setEnabled(True)
        self.menuBar().setEnabled(True)
        QApplication.restoreOverrideCursor()
        if error is not None:
            self._set_app_status("操作失败", "danger")
            self.statusBar().showMessage(f"操作失败：{error}")
            QMessageBox.warning(self, "操作失败", error)
            return
        self._set_app_status("处理完成", "success")
        if on_done is not None:
            on_done(result)

    # ---------- 界面构建 ----------

    def _build_ui(self) -> None:
        file_menu = self.menuBar().addMenu("文件")
        self.action_upload = QAction("上传成员课表…", self)
        self.action_upload.setShortcut(QKeySequence.StandardKey.Open)
        self.action_upload.triggered.connect(self.upload_files)
        file_menu.addAction(self.action_upload)

        self.action_import_roster = QAction("导入花名册…", self)
        self.action_import_roster.setToolTip(
            "导入全媒体中心花名册，自动补充上传课表成员的工作室归属")
        self.action_import_roster.triggered.connect(self.import_roster)
        file_menu.addAction(self.action_import_roster)

        self.action_export_roster = QAction("导出完整花名册…", self)
        self.action_export_roster.setToolTip(
            "按花名册、已上传课表和工作室分配导出完整 Excel")
        self.action_export_roster.triggered.connect(self.export_roster)
        self.action_export_roster.setEnabled(False)
        file_menu.addAction(self.action_export_roster)
        file_menu.addSeparator()

        export_menu = file_menu.addMenu("导出当前排班")
        self.action_export_excel = QAction("导出 Excel…", self)
        self.action_export_excel.setShortcut(QKeySequence("Ctrl+Shift+E"))
        self.action_export_excel.triggered.connect(self.export_xlsx)
        self.action_export_csv = QAction("导出 CSV…", self)
        self.action_export_csv.triggered.connect(self.export_csv)
        self.action_export_png = QAction("导出 PNG…", self)
        self.action_export_png.triggered.connect(self.export_png)
        for action in (
            self.action_export_excel,
            self.action_export_csv,
            self.action_export_png,
        ):
            action.setEnabled(False)
            export_menu.addAction(action)

        self.action_quit = QAction("退出", self)
        self.action_quit.setShortcut(QKeySequence.StandardKey.Quit)
        self.action_quit.triggered.connect(self.close)
        file_menu.addSeparator()
        file_menu.addAction(self.action_quit)

        schedule_menu = self.menuBar().addMenu("排班")
        self.action_participation = QAction("设置参与排班…", self)
        self.action_participation.setToolTip("选择哪些成员参与自动排班和手动微调")
        self.action_participation.triggered.connect(self.open_participation_dialog)
        schedule_menu.addAction(self.action_participation)
        schedule_menu.addSeparator()
        self.action_generate = QAction("生成排班表", self)
        self.action_generate.setShortcut(QKeySequence("Ctrl+G"))
        self.action_generate.triggered.connect(self.generate)
        schedule_menu.addAction(self.action_generate)
        self.action_clear_schedule = QAction("清空排班…", self)
        self.action_clear_schedule.setEnabled(False)
        self.action_clear_schedule.triggered.connect(self.clear_schedule)
        schedule_menu.addAction(self.action_clear_schedule)

        edit_menu = self.menuBar().addMenu("编辑")
        self.undo_action = QAction("撤销", self)
        self.undo_action.setShortcut(QKeySequence.StandardKey.Undo)
        self.undo_action.setShortcutContext(Qt.ApplicationShortcut)
        self.undo_action.setToolTip(f"撤销 {SHORTCUT_MODIFIER}Z")
        self.undo_action.setEnabled(False)
        self.undo_action.triggered.connect(self.undo_tweak)
        edit_menu.addAction(self.undo_action)
        self.redo_action = QAction("重做", self)
        if sys.platform == "darwin":
            self.redo_action.setShortcut(QKeySequence.StandardKey.Redo)
        else:
            self.redo_action.setShortcuts([
                QKeySequence.StandardKey.Redo,
                QKeySequence("Ctrl+Shift+Z"),
            ])
        self.redo_action.setShortcutContext(Qt.ApplicationShortcut)
        self.redo_action.setToolTip(
            "重做 ⌘⇧Z" if SHORTCUT_MODIFIER == "⌘"
            else "重做 Ctrl+Shift+Z")
        self.redo_action.setEnabled(False)
        self.redo_action.triggered.connect(self.redo_tweak)
        edit_menu.addAction(self.redo_action)

        self.action_focus_search = QAction("搜索成员", self)
        self.action_focus_search.setShortcut(QKeySequence.StandardKey.Find)
        self.action_focus_search.triggered.connect(self._focus_member_search)
        self.addAction(self.action_focus_search)

        root = QWidget()
        root.setObjectName("appRoot")
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        top_bar = QFrame()
        top_bar.setObjectName("topBar")
        top_layout = QHBoxLayout(top_bar)
        top_layout.setContentsMargins(18, 10, 18, 10)
        top_layout.setSpacing(11)

        brand_mark = QLabel()
        brand_mark.setObjectName("brandLogo")
        brand_mark.setAlignment(Qt.AlignCenter)
        brand_mark.setToolTip("成都工业学院")
        logo = QPixmap(str(resource_path(
            Path("assets") / "chengdu_technological_university_logo_mark.png")))
        if logo.isNull():
            brand_mark.setObjectName("brandMark")
            brand_mark.setText("值")
            brand_mark.setFixedSize(38, 38)
        else:
            scaled_logo = logo.scaledToHeight(42, Qt.SmoothTransformation)
            brand_mark.setPixmap(scaled_logo)
            brand_mark.setFixedSize(scaled_logo.size())
        self.brand_logo = brand_mark
        top_layout.addWidget(brand_mark)

        brand_text = QVBoxLayout()
        brand_text.setSpacing(1)
        title = QLabel("自动值班排班系统")
        title.setObjectName("appTitle")
        subtitle = QLabel("课表识别 · 冲突避让 · 均衡排班")
        subtitle.setObjectName("appSubtitle")
        brand_text.addWidget(title)
        brand_text.addWidget(subtitle)
        top_layout.addLayout(brand_text)

        top_layout.addSpacing(12)
        self.app_status_label = QLabel("等待生成排班")
        self.app_status_label.setObjectName("statusPill")
        top_layout.addWidget(self.app_status_label)
        top_layout.addStretch(1)

        splitter = QSplitter(Qt.Horizontal)
        left_scroll = QScrollArea()
        left_scroll.setObjectName("leftScroll")
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QFrame.NoFrame)
        left_scroll.setWidget(self._build_left_panel())
        splitter.addWidget(left_scroll)
        splitter.addWidget(self._build_right_panel())
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([340, 1020])
        root_layout.addWidget(top_bar)
        root_layout.addWidget(splitter, stretch=1)
        self.setCentralWidget(root)
        self._configure_accessibility()

    def _focus_member_search(self) -> None:
        self.member_search.setFocus(Qt.ShortcutFocusReason)
        self.member_search.selectAll()

    def _configure_accessibility(self) -> None:
        """设置可访问名称、说明和明确的键盘 Tab 顺序。"""
        self.btn_upload.setAccessibleName("上传成员课表")
        self.btn_import_roster.setAccessibleName("导入花名册")
        self.member_search.setAccessibleName("搜索成员")
        self.member_search.setAccessibleDescription(
            "按姓名、学号、班级或工作室筛选成员")
        self.member_class_filter.setAccessibleName("班级筛选")
        self.member_studio_filter.setAccessibleName("工作室筛选")
        self.member_list.setAccessibleName("成员列表")
        self.btn_edit_studio.setAccessibleName("修改选中成员工作室")
        self.btn_batch_studio.setAccessibleName("批量修改成员工作室")
        self.btn_edit_member_profile.setAccessibleName("修改成员个人信息")
        self.btn_remove.setAccessibleName("删除选中成员")
        self.week_from.setAccessibleName("起始周")
        self.week_to.setAccessibleName("结束周")
        self.per_slot.setAccessibleName("每时段值班人数")
        self.max_week.setAccessibleName("每人每周上限")
        self.max_day.setAccessibleName("每人每天上限")
        self.seed.setAccessibleName("随机种子")
        self.term_start.setAccessibleName("学期起始日")
        self.btn_calendar.setAccessibleName("管理学期日历")
        self.btn_generate.setAccessibleName("生成排班表")
        self.btn_clear_schedule.setAccessibleName("清空排班")
        self.tabs.setAccessibleName("排班功能页签")
        self.app_status_label.setAccessibleName("当前状态")
        self.brand_logo.setAccessibleName("成都工业学院校徽")
        self.pivot_table.setAccessibleName("值班排班表")

        order = [
            self.btn_upload,
            self.btn_import_roster,
            self.member_search,
            self.member_class_filter,
            self.member_studio_filter,
            self.member_list,
            self.btn_edit_studio,
            self.btn_batch_studio,
            self.btn_remove,
            self.week_from,
            self.week_to,
            *self.weekday_checks.values(),
            *self.block_checks.values(),
            self.per_slot,
            self.max_week,
            self.max_day,
            self.seed,
            self.term_start,
            self.btn_calendar,
            self.btn_generate,
            self.btn_clear_schedule,
            self.member_combo,
            self.btn_edit_member_profile,
            self.tabs,
        ]
        for first, second in pairwise(order):
            QWidget.setTabOrder(first, second)

    def _build_left_panel(self) -> QWidget:
        panel = QWidget()
        panel.setObjectName("leftPanel")
        panel.setMinimumWidth(320)
        panel.setMaximumWidth(400)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(12, 14, 8, 14)
        layout.setSpacing(12)

        self.btn_upload = QPushButton("上传成员课表")
        self.btn_upload.setObjectName("primary")
        self.btn_upload.setMinimumHeight(40)
        self.btn_upload.setToolTip("一次选择多份教务系统导出的 .xls / .xlsx 课表")
        self.btn_upload.clicked.connect(self.upload_files)
        layout.addWidget(self.btn_upload)

        self.btn_import_roster = QPushButton("导入花名册")
        self.btn_import_roster.setToolTip(
            "支持 .xlsx/.xls；按学号优先、姓名兜底自动归类工作室")
        self.btn_import_roster.clicked.connect(self.import_roster)
        layout.addWidget(self.btn_import_roster)

        grp_members = QGroupBox("成员课表")
        members_layout = QVBoxLayout(grp_members)
        members_layout.setContentsMargins(8, 4, 8, 8)
        members_layout.setSpacing(8)
        self.member_search = QLineEdit()
        self.member_search.setPlaceholderText("搜索姓名 / 学号 / 班级 / 工作室")
        self.member_search.setClearButtonEnabled(True)
        members_layout.addWidget(self.member_search)
        self.member_class_filter = QComboBox()
        self.member_class_filter.setToolTip("按班级筛选成员")
        members_layout.addWidget(self.member_class_filter)
        self.member_studio_filter = QComboBox()
        self.member_studio_filter.setToolTip("按工作室筛选成员")
        members_layout.addWidget(self.member_studio_filter)
        self.member_list = QListWidget()
        self.member_list.setSelectionMode(QAbstractItemView.SingleSelection)
        self.member_list.setMinimumHeight(120)
        self.member_list.setToolTip(
            "双击成员可在「成员课表」页查看其课表详情。\n"
            "是否参与排班请在菜单栏「排班 → 设置参与排班…」中设置。")
        self.member_list.itemDoubleClicked.connect(self.open_member_courses)
        self.member_search.textChanged.connect(self._apply_member_filter)
        self.member_class_filter.currentIndexChanged.connect(self._apply_member_filter)
        self.member_studio_filter.currentIndexChanged.connect(self._apply_member_filter)
        members_layout.addWidget(self.member_list)
        member_actions = QHBoxLayout()
        self.btn_edit_studio = QPushButton("修改工作室…")
        self.btn_edit_studio.setToolTip(
            "用于成员转工作室或纠正自动归类；手工修改后重导花名册不会覆盖")
        self.btn_edit_studio.clicked.connect(self.edit_selected_member_studio)
        member_actions.addWidget(self.btn_edit_studio)
        self.btn_remove = QPushButton("删除选中成员")
        self.btn_remove.setObjectName("danger")
        self.btn_remove.clicked.connect(self.remove_selected_member)
        member_actions.addWidget(self.btn_remove)
        members_layout.addLayout(member_actions)
        self.btn_batch_studio = QPushButton("批量修改工作室…")
        self.btn_batch_studio.setToolTip(
            "按当前搜索/筛选结果勾选多人，一次修改并锁定工作室归属")
        self.btn_batch_studio.clicked.connect(self.batch_edit_members_studio)
        members_layout.addWidget(self.btn_batch_studio)
        layout.addWidget(grp_members, stretch=1)

        grp_cfg = QGroupBox("排班参数")
        form = QFormLayout(grp_cfg)
        form.setContentsMargins(8, 4, 8, 8)
        form.setSpacing(9)
        form.setLabelAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        week_row = QHBoxLayout()
        self.week_from = QSpinBox()
        self.week_from.setRange(1, 25)
        self.week_from.setValue(1)
        self.week_from.valueChanged.connect(self._mark_stale)
        self.week_from.valueChanged.connect(self._sync_gantt_week)
        self.week_to = QSpinBox()
        self.week_to.setRange(1, 25)
        self.week_to.setValue(18)
        self.week_to.valueChanged.connect(self._mark_stale)
        week_row.addWidget(self.week_from)
        week_row.addWidget(QLabel("至"))
        week_row.addWidget(self.week_to)
        form.addRow("值班周范围", week_row)

        self.weekday_checks: dict[int, QCheckBox] = {}
        wd_grid = QGridLayout()
        wd_grid.setHorizontalSpacing(8)
        wd_grid.setVerticalSpacing(6)
        for index, weekday in enumerate((1, 2, 3, 4, 5, 6, 7)):
            cb = QCheckBox(WEEKDAY_LABELS[weekday].replace("周", ""))
            cb.setToolTip(WEEKDAY_LABELS[weekday])
            cb.setChecked(weekday <= 5)
            cb.stateChanged.connect(self._on_gantt_filter_changed)
            cb.stateChanged.connect(self._mark_stale)
            self.weekday_checks[weekday] = cb
            wd_grid.addWidget(cb, index // 4, index % 4)
        form.addRow("值班星期", wd_grid)

        self.block_checks: dict[int, QCheckBox] = {}
        blk_grid = QGridLayout()
        blk_grid.setHorizontalSpacing(8)
        blk_grid.setVerticalSpacing(6)
        for index, block in enumerate((1, 2, 3, 4, 5)):
            cb = QCheckBox(BLOCK_LABELS[block].split(" ")[0].replace("节", ""))
            cb.setChecked(True)
            cb.setToolTip(BLOCK_LABELS[block])
            cb.stateChanged.connect(self._on_gantt_filter_changed)
            cb.stateChanged.connect(self._mark_stale)
            self.block_checks[block] = cb
            blk_grid.addWidget(cb, index // 3, index % 3)
        form.addRow("值班时段", blk_grid)

        self.per_slot = QSpinBox()
        self.per_slot.setRange(1, 5)
        self.per_slot.valueChanged.connect(self._mark_stale)
        form.addRow("每时段人数", self.per_slot)

        self.max_week = QSpinBox()
        self.max_week.setRange(1, 10)
        self.max_week.setValue(3)
        self.max_week.setToolTip("每人每周最多值班次数")
        self.max_week.valueChanged.connect(self._mark_stale)
        form.addRow("每周上限", self.max_week)

        self.max_day = QSpinBox()
        self.max_day.setRange(1, 5)
        self.max_day.setValue(1)
        self.max_day.setToolTip("每人每天最多值班次数")
        self.max_day.valueChanged.connect(self._mark_stale)
        form.addRow("每天上限", self.max_day)

        self.seed = QSpinBox()
        self.seed.setRange(0, 9999)
        self.seed.setValue(42)
        self.seed.setToolTip("平手时决定分给谁，固定种子可复现")
        self.seed.valueChanged.connect(self._mark_stale)
        form.addRow("随机种子", self.seed)

        self.term_start = QDateEdit()
        self.term_start.setCalendarPopup(True)
        self.term_start.setDisplayFormat("yyyy-MM-dd")
        today = QDate.currentDate()
        self.term_start.setDate(today.addDays(-(today.dayOfWeek() - 1)))
        self.term_start.setToolTip("第一周周一的日期；值班表按此把周次换算为具体日期")
        self.term_start.dateChanged.connect(self._on_term_start_changed)
        form.addRow("学期起始日", self.term_start)
        calendar_row = QHBoxLayout()
        self.calendar_label = QLabel("未设置")
        self.calendar_label.setObjectName("secondary")
        self.btn_calendar = QPushButton("管理日历…")
        self.btn_calendar.setToolTip("批量设置法定假日与周末补课映射")
        self.btn_calendar.clicked.connect(self.open_calendar_dialog)
        calendar_row.addWidget(self.calendar_label, stretch=1)
        calendar_row.addWidget(self.btn_calendar)
        form.addRow("学期日历", calendar_row)
        layout.addWidget(grp_cfg)

        self.btn_generate = QPushButton("生成排班表")
        self.btn_generate.setObjectName("primary")
        self.btn_generate.setMinimumHeight(42)
        self.btn_generate.clicked.connect(self.generate)
        layout.addWidget(self.btn_generate)
        self.btn_clear_schedule = QPushButton("清空排班…")
        self.btn_clear_schedule.setObjectName("danger")
        self.btn_clear_schedule.setEnabled(False)
        self.btn_clear_schedule.clicked.connect(self.clear_schedule)
        layout.addWidget(self.btn_clear_schedule)

        grp_data = QGroupBox("数据存储")
        dl = QHBoxLayout(grp_data)
        dl.setContentsMargins(8, 4, 8, 8)
        dl.setSpacing(8)
        self.db_path_label = QLabel()
        self.db_path_label.setToolTip("")
        btn_change_db = QPushButton("更改位置…")
        btn_change_db.setToolTip(
            "更改数据库存储位置：现有数据自动复制到新位置（原文件保留）；\n"
            "若所选目录已有同名数据库文件，则切换为使用该文件。")
        btn_change_db.clicked.connect(self.change_db_location)
        dl.addWidget(self.db_path_label, stretch=1)
        dl.addWidget(btn_change_db)
        layout.addWidget(grp_data)
        self._update_db_path_label()
        self._refresh_calendar_summary()
        return panel

    @staticmethod
    def _metric_card(label_text: str) -> tuple[QFrame, QLabel, QLabel]:
        card = QFrame()
        card.setObjectName("metricCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(13, 11, 13, 11)
        layout.setSpacing(4)
        label = QLabel(label_text)
        label.setObjectName("metricLabel")
        value = QLabel("—")
        value.setObjectName("metricValue")
        hint = QLabel("等待排班结果")
        hint.setObjectName("metricHint")
        layout.addWidget(label)
        layout.addWidget(value)
        layout.addWidget(hint)
        return card, value, hint

    def _build_right_panel(self) -> QWidget:
        self.tabs = QTabWidget()

        # Tab1 排班总表
        tab1 = QWidget()
        v1 = QVBoxLayout(tab1)
        v1.setContentsMargins(16, 14, 16, 16)
        v1.setSpacing(12)
        summary_card = QFrame()
        summary_card.setObjectName("summaryCard")
        summary_layout = QVBoxLayout(summary_card)
        summary_layout.setContentsMargins(15, 13, 15, 13)
        summary_layout.setSpacing(10)
        self.summary_label = QLabel("尚未生成排班表。设置左侧参数后点击「生成排班表」。")
        self.summary_label.setObjectName("summary")
        self.summary_label.setWordWrap(True)
        summary_layout.addWidget(self.summary_label)

        metric_row = QHBoxLayout()
        metric_row.setSpacing(9)
        metric_specs = (
            ("duties", "本周安排", "尚未生成"),
            ("members", "参与成员", "尚未生成"),
            ("spread", "次数极差", "尚未生成"),
            ("gaps", "待处理缺口", "尚未生成"),
        )
        for key, label_text, hint_text in metric_specs:
            card, value, hint = self._metric_card(label_text)
            hint.setText(hint_text)
            setattr(self, f"metric_{key}", value)
            setattr(self, f"metric_{key}_hint", hint)
            metric_row.addWidget(card, stretch=1)
        summary_layout.addLayout(metric_row)

        btn_row = QHBoxLayout()
        export_title = QLabel("结果导出")
        export_title.setObjectName("secondary")
        btn_row.addWidget(export_title)
        self.btn_export_xlsx = QPushButton("Excel")
        self.btn_export_xlsx.setToolTip("按日期与值班人内容自动适配列宽、行高，打开即可完整查看")
        self.btn_export_xlsx.clicked.connect(self.export_xlsx)
        self.btn_export_csv = QPushButton("CSV")
        self.btn_export_csv.setToolTip("完整保留日期与姓名文本；CSV 不支持列宽，排版查看请用 Excel")
        self.btn_export_csv.clicked.connect(self.export_csv)
        self.btn_export_png = QPushButton("PNG")
        self.btn_export_png.setToolTip("按内容自动换行、调整行高，避免日期与值班人被裁切")
        self.btn_export_png.clicked.connect(self.export_png)
        self.btn_export_xlsx.setEnabled(False)
        self.btn_export_csv.setEnabled(False)
        self.btn_export_png.setEnabled(False)
        btn_row.addWidget(self.btn_export_xlsx)
        btn_row.addWidget(self.btn_export_csv)
        btn_row.addWidget(self.btn_export_png)
        btn_row.addStretch()
        summary_layout.addLayout(btn_row)
        v1.addWidget(summary_card)
        self.pivot_table = QTableWidget()
        self.pivot_table.setToolTip("双击值班单元格可手动调整该时段值班人")
        self.pivot_table.cellDoubleClicked.connect(self._on_pivot_cell_double_clicked)
        self.pivot_stack = QStackedWidget()
        self.pivot_stack.addWidget(_card(self.pivot_table))

        pivot_empty = QFrame()
        pivot_empty.setObjectName("emptyState")
        empty_layout = QVBoxLayout(pivot_empty)
        empty_layout.setContentsMargins(24, 24, 24, 24)
        empty_layout.addStretch(1)
        self.empty_title = QLabel("三步完成第一次排班")
        self.empty_title.setObjectName("pageTitle")
        self.empty_title.setAlignment(Qt.AlignCenter)
        self.empty_text = QLabel("先导入课表，再确认排班参数，最后生成结果。")
        self.empty_text.setObjectName("secondary")
        self.empty_text.setAlignment(Qt.AlignCenter)
        self.empty_text.setWordWrap(True)
        empty_layout.addWidget(self.empty_title)
        empty_layout.addWidget(self.empty_text)
        empty_layout.addSpacing(12)

        steps_row = QHBoxLayout()
        steps_row.setSpacing(10)
        self.empty_step_cards: list[QFrame] = []
        step_specs = (
            ("1", "导入课表", "上传成员的 .xls / .xlsx 文件"),
            ("2", "设置规则", "选择周次、星期、时段与上限"),
            ("3", "生成导出", "生成排班并导出 Excel / PNG"),
        )
        for number, title_text, detail_text in step_specs:
            step = QFrame()
            step.setObjectName("guideStep")
            step_layout = QVBoxLayout(step)
            step_layout.setContentsMargins(14, 12, 14, 12)
            step_layout.setSpacing(5)
            number_label = QLabel(number)
            number_label.setObjectName("guideNumber")
            number_label.setAlignment(Qt.AlignCenter)
            number_label.setFixedSize(24, 24)
            title_label = QLabel(title_text)
            title_label.setObjectName("guideTitle")
            detail_label = QLabel(detail_text)
            detail_label.setObjectName("secondary")
            detail_label.setWordWrap(True)
            step_layout.addWidget(number_label)
            step_layout.addWidget(title_label)
            step_layout.addWidget(detail_label)
            self.empty_step_cards.append(step)
            steps_row.addWidget(step, stretch=1)
        empty_layout.addLayout(steps_row)
        empty_layout.addSpacing(14)
        self.empty_action = QPushButton("上传成员课表")
        self.empty_action.setObjectName("primary")
        self.empty_action.clicked.connect(self._on_empty_action)
        empty_layout.addWidget(self.empty_action, alignment=Qt.AlignCenter)
        empty_layout.addStretch(1)
        self.pivot_stack.addWidget(pivot_empty)
        v1.addWidget(self.pivot_stack, stretch=1)
        self.tabs.addTab(tab1, "值班排班表")

        # Tab2 值班明细
        tab2 = QWidget()
        v2 = QVBoxLayout(tab2)
        v2.setContentsMargins(16, 14, 16, 16)
        v2.setSpacing(10)
        detail_title = QLabel("值班明细")
        detail_title.setObjectName("pageTitle")
        detail_hint = QLabel("逐条查看周次、星期、日期、时段与值班人")
        detail_hint.setObjectName("secondary")
        v2.addWidget(detail_title)
        v2.addWidget(detail_hint)
        self.detail_table = QTableWidget()
        v2.addWidget(_card(self.detail_table))
        self.tabs.addTab(tab2, "值班明细")

        # Tab3 成员课表
        tab3 = QWidget()
        v3 = QVBoxLayout(tab3)
        v3.setContentsMargins(16, 14, 16, 16)
        v3.setSpacing(12)
        member_title = QLabel("成员课表与不可用安排")
        member_title.setObjectName("pageTitle")
        v3.addWidget(member_title)
        sel_row = QHBoxLayout()
        sel_row.addWidget(QLabel("选择成员"))
        self.member_combo = QComboBox()
        self.member_combo.setMinimumWidth(220)
        self.member_combo.currentIndexChanged.connect(self._on_member_combo_changed)
        sel_row.addWidget(self.member_combo)
        self.btn_edit_member_profile = QPushButton("修改个人信息…")
        self.btn_edit_member_profile.setToolTip(
            "修改姓名、学号、班级、学期、专业和院系；课程与历史排班仍保留")
        self.btn_edit_member_profile.clicked.connect(self.edit_member_profile)
        sel_row.addWidget(self.btn_edit_member_profile)
        sel_row.addStretch()
        v3.addLayout(sel_row)
        self.member_info = QLabel("")
        self.member_info.setObjectName("secondary")
        v3.addWidget(self.member_info)
        self.course_table = QTableWidget()
        v3.addWidget(_card(self.course_table), stretch=1)

        grp_special = QGroupBox("长期特殊安排（修改课表）")
        grp_special.setToolTip(
            "把连续多周、每周固定星期/时段的安排补充到课表占用中；\n"
            "自动排班、空闲时段总览和手动微调都会避让。")
        sp = QHBoxLayout(grp_special)
        sp.setContentsMargins(8, 4, 8, 8)
        sp.setSpacing(8)
        self.special_table = QTableWidget()
        self.special_table.setColumnCount(4)
        self.special_table.setHorizontalHeaderLabels(["周次", "星期", "时段", "原因"])
        self.special_table.setMaximumHeight(130)
        self.special_table.verticalHeader().setVisible(False)
        self.special_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.special_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.special_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.special_table.cellDoubleClicked.connect(
            lambda _row, _col: self.edit_special_arrangement())
        sp.addWidget(self.special_table, stretch=1)
        special_btns = QVBoxLayout()
        btn_special_add = QPushButton("添加安排…")
        btn_special_add.setToolTip("添加长期固定的课表外占用")
        btn_special_add.clicked.connect(self.add_special_arrangement)
        btn_special_edit = QPushButton("编辑选中…")
        btn_special_edit.setToolTip("修改选中安排的时间范围、星期、时段或原因")
        btn_special_edit.clicked.connect(self.edit_special_arrangement)
        btn_special_del = QPushButton("删除选中")
        btn_special_del.clicked.connect(self.remove_special_arrangement)
        special_btns.addWidget(btn_special_add)
        special_btns.addWidget(btn_special_edit)
        special_btns.addWidget(btn_special_del)
        special_btns.addStretch()
        sp.addLayout(special_btns)
        v3.addWidget(grp_special)

        grp_leave = QGroupBox("请假登记（临时不可值班日）")
        grp_leave.setToolTip("登记后重新生成排班将避开该天；空闲时段总览中该天标红")
        lv = QHBoxLayout(grp_leave)
        lv.setContentsMargins(8, 4, 8, 8)
        lv.setSpacing(8)
        self.leave_table = QTableWidget()
        self.leave_table.setColumnCount(3)
        self.leave_table.setHorizontalHeaderLabels(["周次", "星期", "原因"])
        self.leave_table.setMaximumHeight(120)
        self.leave_table.verticalHeader().setVisible(False)
        self.leave_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.leave_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.leave_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        lv.addWidget(self.leave_table, stretch=1)
        leave_btns = QVBoxLayout()
        btn_leave_add = QPushButton("添加请假…")
        btn_leave_add.clicked.connect(self.add_leave)
        btn_leave_del = QPushButton("删除选中")
        btn_leave_del.clicked.connect(self.remove_leave)
        self.btn_leave_export = QPushButton("导出请假记录…")
        self.btn_leave_export.setToolTip(
            "导出所有成员的请假记录为 Excel（含周次/星期/日期/原因），方便报备与存档")
        self.btn_leave_export.clicked.connect(self.export_leaves)
        leave_btns.addWidget(btn_leave_add)
        leave_btns.addWidget(btn_leave_del)
        leave_btns.addWidget(self.btn_leave_export)
        leave_btns.addStretch()
        lv.addLayout(leave_btns)
        v3.addWidget(grp_leave)
        self.tabs.addTab(tab3, "成员课表")

        # Tab4 值班统计
        tab4 = QWidget()
        v4 = QVBoxLayout(tab4)
        v4.setContentsMargins(16, 14, 16, 16)
        v4.setSpacing(12)
        stats_title = QLabel("值班统计")
        stats_title.setObjectName("pageTitle")
        v4.addWidget(stats_title)
        self.stats_table = QTableWidget()
        v4.addWidget(_card(self.stats_table), stretch=2)
        self.gap_title = QLabel("")
        self.gap_title.setObjectName("secondary")
        v4.addWidget(self.gap_title)
        self.gap_table = QTableWidget()
        v4.addWidget(_card(self.gap_table), stretch=1)
        self.tabs.addTab(tab4, "值班统计")

        # Tab5 空闲时段总览：成员空闲时段一览，方便安排任务
        tab5 = QWidget()
        v5 = QVBoxLayout(tab5)
        v5.setContentsMargins(16, 14, 16, 16)
        v5.setSpacing(12)
        gantt_title = QLabel("空闲时段总览")
        gantt_title.setObjectName("pageTitle")
        v5.addWidget(gantt_title)
        ctl = QHBoxLayout()
        ctl.addWidget(QLabel("查看周次"))
        self.gantt_week = QSpinBox()
        self.gantt_week.setRange(1, 25)
        self.gantt_week.setValue(1)
        self.gantt_week.setToolTip("空闲时段总览按周查看（课程随周次变化）")
        self.gantt_week.valueChanged.connect(self.refresh_gantt)
        ctl.addWidget(self.gantt_week)
        legend = QFrame()
        legend.setObjectName("legendBar")
        legend_layout = QGridLayout(legend)
        legend_layout.setContentsMargins(8, 4, 8, 4)
        legend_layout.setHorizontalSpacing(10)
        legend_layout.setVerticalSpacing(3)
        legend_items = (
            ("#c9f2cf", "空闲"), ("#f2f2f7", "有课"), ("#b8d9ff", "值班"),
            ("#ffd9a8", "全员空闲"), ("#ffd6d2", "请假"), ("#e5d8ff", "其他安排"),
        )
        for index, (color, text_value) in enumerate(legend_items):
            dot = QLabel()
            dot.setFixedSize(9, 9)
            dot.setStyleSheet(f"background:{color};border-radius:3px;")
            label = QLabel(text_value)
            label.setObjectName("legendText")
            legend_layout.addWidget(dot, index // 3, (index % 3) * 2)
            legend_layout.addWidget(label, index // 3, (index % 3) * 2 + 1)
        ctl.addWidget(legend)
        ctl.addStretch()
        self.btn_export_gantt = QPushButton("导出空闲时段总览 Excel")
        self.btn_export_gantt.clicked.connect(self.export_gantt)
        self.btn_export_gantt.setEnabled(False)
        ctl.addWidget(self.btn_export_gantt)
        v5.addLayout(ctl)
        self.gantt_hint = QLabel("")
        self.gantt_hint.setObjectName("summary")
        self.gantt_hint.setWordWrap(True)
        v5.addWidget(self.gantt_hint)
        self.gantt_table = QTableWidget()
        # 表头列宽/行高只设置一次：以前每次刷新空闲时段总览都重设
        # ResizeToContents，会触发整表列宽重算，成员多时明显拖慢刷新
        gantt_header = self.gantt_table.horizontalHeader()
        gantt_header.setSectionResizeMode(QHeaderView.ResizeToContents)
        gantt_header.setStretchLastSection(False)
        gantt_header.setMinimumHeight(44)
        gantt_vertical_header = self.gantt_table.verticalHeader()
        gantt_vertical_header.setDefaultSectionSize(44)
        gantt_vertical_header.setMinimumWidth(112)
        v5.addWidget(_card(self.gantt_table), stretch=1)
        self.tabs.addTab(tab5, "空闲时段总览")

        # Tab6 统计图表：值班 / 请假情况可视化
        tab6 = QWidget()
        v6 = QVBoxLayout(tab6)
        v6.setContentsMargins(16, 14, 16, 16)
        v6.setSpacing(12)
        chart_title = QLabel("统计图表")
        chart_title.setObjectName("pageTitle")
        v6.addWidget(chart_title)
        chart_ctl = QHBoxLayout()
        self.charts_hint = QLabel("值班与请假情况一览，可导出为 PNG 汇报")
        self.charts_hint.setObjectName("secondary")
        chart_ctl.addWidget(self.charts_hint)
        chart_ctl.addStretch()
        self.btn_export_charts = QPushButton("导出统计图 PNG")
        self.btn_export_charts.setToolTip("把三张统计图渲染为一张高清图片，方便发群汇报 / 存档")
        self.btn_export_charts.clicked.connect(self.export_charts)
        self.btn_export_charts.setEnabled(False)
        chart_ctl.addWidget(self.btn_export_charts)
        v6.addLayout(chart_ctl)
        charts_row = QHBoxLayout()
        charts_row.setSpacing(12)
        self.chart_duty = BarChart("各成员值班总次数", "#2563eb")
        self.chart_weekly = BarChart("每周值班人次", "#059669")
        self.chart_leave = BarChart("各成员请假天数", "#dc2626")
        for c in (self.chart_duty, self.chart_weekly, self.chart_leave):
            charts_row.addWidget(c, stretch=1)
        v6.addLayout(charts_row, stretch=1)
        self.tabs.addTab(tab6, "统计图表")

        self.tabs.currentChanged.connect(self._on_tab_changed)
        return self.tabs

    # ---------- 数据刷新 ----------

    def refresh_members(self) -> None:
        members = self.members()
        selected_item = self.member_list.currentItem()
        selected_id = selected_item.data(Qt.UserRole) if selected_item else None
        selected_class = self.member_class_filter.currentData()
        selected_studio = self.member_studio_filter.currentData()
        self.member_list.clear()
        self.member_class_filter.blockSignals(True)
        self.member_class_filter.clear()
        self.member_class_filter.addItem("全部", "")
        for class_name in sorted({
                m.class_name.strip() for m in members if m.class_name.strip()}):
            self.member_class_filter.addItem(class_name, class_name)
        class_index = self.member_class_filter.findData(selected_class)
        self.member_class_filter.setCurrentIndex(max(class_index, 0))
        self.member_class_filter.blockSignals(False)
        self.member_studio_filter.blockSignals(True)
        self.member_studio_filter.clear()
        self.member_studio_filter.addItem("全部", "")
        for studio in sorted({
                (m.studio or UNKNOWN_STUDIO).strip()
                for m in members if (m.studio or UNKNOWN_STUDIO).strip()}):
            self.member_studio_filter.addItem(studio, studio)
        studio_index = self.member_studio_filter.findData(selected_studio)
        self.member_studio_filter.setCurrentIndex(max(studio_index, 0))
        self.member_studio_filter.blockSignals(False)
        self.member_combo.blockSignals(True)
        self.member_combo.clear()
        for m in members:
            suffix = "" if m.participates_in_scheduling else " · 不参与排班"
            studio = m.studio or UNKNOWN_STUDIO
            item = QListWidgetItem(
                f"{m.name}（{m.course_count} 门课 · {studio}）{suffix}")
            item.setData(Qt.UserRole, m.id)
            item.setData(Qt.UserRole + 1, (
                m.name, m.student_id, m.class_name, studio))
            item.setFlags(item.flags() & ~Qt.ItemIsUserCheckable)
            if not m.participates_in_scheduling:
                item.setForeground(QBrush(QColor("#8e8e93")))
                item.setToolTip("不参与排班；可通过菜单「排班 → 设置参与排班…」修改。")
            else:
                item.setToolTip("参与排班；可通过菜单「排班 → 设置参与排班…」修改。")
            item.setToolTip(
                item.toolTip() + f"\n工作室：{studio}"
                + ("（已手工锁定）" if m.studio_locked else ""))
            self.member_list.addItem(item)
            combo_status = "" if m.participates_in_scheduling else " · 不参与排班"
            self.member_combo.addItem(
                f"{m.name} · {studio}（{m.class_name or '未知班级'}{combo_status}）",
                m.id)
        self.member_combo.blockSignals(False)
        self.action_participation.setEnabled(bool(members))
        self.btn_edit_member_profile.setEnabled(bool(members))
        self.btn_batch_studio.setEnabled(bool(members))
        self.action_export_roster.setEnabled(bool(members or self.roster_entries()))
        if selected_id is not None:
            for row in range(self.member_list.count()):
                if self.member_list.item(row).data(Qt.UserRole) == selected_id:
                    self.member_list.setCurrentRow(row)
                    break
        self._apply_member_filter()
        self._clear_undo_redo()
        if self.member_combo.count():
            self.show_member_courses(0)
        else:
            self.member_info.setText("")
            fill_table(self.course_table, pd.DataFrame(columns=["星期", "课程", "教师", "周次", "节次", "地点"]))
            self.special_table.setRowCount(0)
        self.refresh_gantt()
        self._update_empty_state()
        self._mark_stale()

    def _apply_member_filter(self, *_args) -> None:
        """按姓名/学号/班级即时过滤成员列表，不修改底层数据。"""
        query = self.member_search.text().strip().casefold()
        selected_class = self.member_class_filter.currentData() or ""
        selected_studio = self.member_studio_filter.currentData() or ""
        for row in range(self.member_list.count()):
            item = self.member_list.item(row)
            name, student_id, class_name, studio = item.data(Qt.UserRole + 1)
            matches_query = (
                not query
                or query in str(name).casefold()
                or query in str(student_id).casefold()
                or query in str(class_name).casefold()
                or query in str(studio).casefold()
            )
            matches_class = not selected_class or class_name == selected_class
            matches_studio = not selected_studio or studio == selected_studio
            item.setHidden(not (matches_query and matches_class and matches_studio))

    def _on_empty_action(self) -> None:
        if self.members():
            self.generate()
        else:
            self.upload_files()

    def _update_empty_state(self) -> None:
        """无排班结果时，给出下一步引导文案"""
        if self.result is not None:
            return
        member_count = self.member_list.count()
        scheduling_count = len(self.scheduling_members())
        if member_count == 0:
            self.empty_title.setText("三步完成第一次排班")
            self.empty_text.setText("先导入课表，再确认排班参数，最后生成结果。")
            self.empty_action.setText("上传成员课表")
            self.summary_label.setText(
                "开始使用三步：① 上传成员课表（支持多选 .xls / .xlsx）→ ② 调整排班参数 → "
                "③ 点击「生成排班表」。可在「空闲时段总览」页查看全员共同空闲时段，方便安排任务。")
            self.empty_step_cards[0].setProperty("state", "active")
            self.empty_step_cards[1].setProperty("state", "")
            self.empty_step_cards[2].setProperty("state", "")
        elif scheduling_count == 0:
            self.empty_title.setText("没有可参与排班的成员")
            self.empty_text.setText(
                f"已导入 {member_count} 名成员，但当前全部设为“不参与排班”。\n"
                "请通过菜单「排班 → 设置参与排班…」重新勾选需要参与排班的成员。")
            self.empty_action.setText("生成排班表")
            self.summary_label.setText(
                "当前没有参与排班的成员，请先通过菜单设置参与范围，再点击「生成排班表」。")
            self.empty_step_cards[0].setProperty("state", "done")
            self.empty_step_cards[1].setProperty("state", "active")
            self.empty_step_cards[2].setProperty("state", "")
        else:
            self.empty_title.setText("已准备就绪，可以生成排班")
            self.empty_text.setText(
                f"已导入 {member_count} 名成员，其中 {scheduling_count} 名参与排班。"
                "确认左侧规则后，点击下方按钮生成排班表。")
            self.empty_action.setText("生成排班表")
            if scheduling_count == member_count:
                summary = (
                    f"已就绪 {member_count} 名成员（均参与排班），"
                    "点击「生成排班表」开始排班。")
            else:
                summary = (
                    f"已就绪 {member_count} 名成员，其中 {scheduling_count} 名参与排班，"
                    "点击「生成排班表」开始排班。")
            self.summary_label.setText(summary)
            self.empty_step_cards[0].setProperty("state", "done")
            self.empty_step_cards[1].setProperty("state", "active")
            self.empty_step_cards[2].setProperty("state", "")
        for step in self.empty_step_cards:
            step.style().unpolish(step)
            step.style().polish(step)
        self._reset_metric_cards()
        self.pivot_stack.setCurrentIndex(1)

    def _mark_stale(self) -> None:
        """参数或成员变化后标记结果过期，提示重新生成"""
        if self.result is None or self._stale:
            return
        self._stale = True
        self.summary_label.setText(
            "⚠ 排班参数或成员已变化，下方结果可能过期——请点击「生成排班表」重新生成。")
        self._set_app_status("结果可能过期", "warning")

    def _restore_result(self) -> None:
        """启动时从数据库恢复上次的排班结果（关闭程序不会丢失）"""
        assignments = self.db.load_assignments()
        members = self.scheduling_members()
        if not assignments or not members:
            return
        # 优先用「生成时保存的参数」判定缺口；没有记录（老版本数据）才退回当前界面参数
        config = self._last_config or self._load_last_config() or self._current_config()
        assignments = [
            a for a in assignments if not self.is_off_day(a.week, a.weekday)
        ]
        if not assignments:
            return
        result = ScheduleResult()
        result.assignments = assignments
        result.member_stats = rebuild_member_stats(members, assignments)
        result.gaps = compute_gaps(
            assignments, config, is_off=self.is_off_day,
            is_class=self.is_class_day)
        self.result = result
        self._last_config = config
        self._stale = False
        self.refresh_schedule_tabs()
        self.refresh_gantt()

    # ---------- 数据存储位置 ----------

    def _update_db_path_label(self) -> None:
        path = Path(self.db.path)
        fm = QFontMetrics(self.db_path_label.font())
        text = fm.elidedText(f"{path.name} · {path.parent}", Qt.ElideMiddle, 230)
        self.db_path_label.setText(text)
        self.db_path_label.setToolTip(str(path))

    def change_db_location(self) -> None:
        """更改数据库存储位置：迁移现有数据或切换到已有数据库文件。"""
        if self.busy:
            self.statusBar().showMessage("上一个任务尚未完成，请稍候…")
            return
        current = Path(self.db.path)
        folder = QFileDialog.getExistingDirectory(
            self, "选择数据库存储位置", str(current.parent))
        if not folder:
            return
        new_path = Path(folder) / current.name
        if new_path == current:
            QMessageBox.information(self, "提示", "数据库已位于该位置。")
            return
        if new_path.exists():
            ret = QMessageBox.question(
                self, "切换到已有数据库",
                f"所选目录已存在 {current.name}：\n{new_path}\n\n"
                "是否切换为使用该数据库文件？其中的数据将替代当前界面内容；\n"
                "当前数据库文件仍保留在原位置，不会被修改。",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if ret != QMessageBox.Yes:
                return
            migrated = False
        else:
            migrated = True

        def prepare_database() -> Database:
            # 在线备份必须使用 SQLite API；大库备份放到后台，避免界面卡顿。
            if migrated:
                self.db.backup_to(new_path)
            return Database(new_path)

        self.run_async(
            "正在准备数据库",
            prepare_database,
            lambda new_db: self._activate_database(
                current, new_path, new_db, migrated),
        )

    def _activate_database(
        self,
        current: Path,
        new_path: Path,
        new_db: Database,
        migrated: bool,
    ) -> None:
        """后台备份/校验成功后，在主线程切换数据库并刷新界面。"""
        QSettings().setValue("database/path", str(new_path))
        self.db = new_db
        self.invalidate_cache()
        self._update_db_path_label()
        self._refresh_calendar_summary()
        self.result = None
        self._last_config = None
        self._stale = False
        self.refresh_members()
        self._restore_result()
        if self.result is None:
            self._update_empty_state()
            for table in (
                self.pivot_table,
                self.detail_table,
                self.stats_table,
                self.gap_table,
            ):
                fill_table(table, pd.DataFrame())
            self.btn_export_xlsx.setEnabled(False)
            self.btn_export_csv.setEnabled(False)
            self.btn_export_png.setEnabled(False)
            self.action_export_excel.setEnabled(False)
            self.action_export_csv.setEnabled(False)
            self.action_export_png.setEnabled(False)
            self.action_clear_schedule.setEnabled(False)
            self.gap_title.setText("无人可值时段")
        self.refresh_gantt()
        self.refresh_charts()
        if migrated:
            self.statusBar().showMessage(
                f"数据库已迁移至：{new_path}（原文件保留在 {current}，可手动删除）")
        else:
            self.statusBar().showMessage(f"已切换数据库：{new_path}")

    # ---------- 参数持久化 ----------

    def _param_specs(self) -> list[tuple[str, object, str, object]]:
        """参数控件 -> QSettings 键 的登记表：(类型, 控件, 键, 默认值)。

        新增一个可持久化参数只需在这里加一行：读写逻辑共用同一份声明，
        避免以前「加了控件忘了加 _load_settings / _save_settings」的漏改问题。
        """
        specs: list[tuple[str, object, str, object]] = [
            ("spin", self.week_from, "week_from", 1),
            ("spin", self.week_to, "week_to", 18),
            ("spin", self.per_slot, "per_slot", 1),
            ("spin", self.max_week, "max_week", 3),
            ("spin", self.max_day, "max_day", 1),
            ("spin", self.seed, "seed", 42),
            ("spin", self.gantt_week, "gantt_week", 1),
            ("date", self.term_start, "term_start", None),
        ]
        specs += [("check", cb, f"weekday_{d}", d <= 5)
                  for d, cb in self.weekday_checks.items()]
        specs += [("check", cb, f"block_{b}", True)
                  for b, cb in self.block_checks.items()]
        return specs

    def _load_settings(self) -> None:
        s = QSettings()
        for kind, widget, key, default in self._param_specs():
            if kind == "spin":
                widget.setValue(int(s.value(key, default)))
            elif kind == "check":
                if s.contains(key):
                    widget.setChecked(int(s.value(key, 1 if default else 0)) == 1)
                else:
                    widget.setChecked(bool(default))
            elif kind == "date":
                parsed = QDate.fromString(str(s.value(key, "") or ""), "yyyy-MM-dd")
                if parsed.isValid():
                    widget.setDate(parsed)
        self._refresh_calendar_summary()

    def _save_settings(self) -> None:
        s = QSettings()
        for kind, widget, key, _default in self._param_specs():
            if kind == "spin":
                s.setValue(key, widget.value())
            elif kind == "check":
                s.setValue(key, 1 if widget.isChecked() else 0)
            elif kind == "date":
                s.setValue(key, widget.date().toString("yyyy-MM-dd"))

    # ---------- 上次排班参数持久化 ----------
    # 缺口判定依赖 per_slot/星期/时段等参数。只存排班结果的话，用户改过参数
    # 再重启，缺口清单就会与实际排班对不上——因此把生成时的参数一并存下来。

    def _config_key(self) -> str:
        """上次排班参数的存储键：按数据库文件区分。

        切换到另一个数据库（或恢复备份）时，各自的排班参数互不干扰，
        否则会用 A 库的 per_slot 去判断 B 库的缺口。
        """
        return f"schedule/last_config/{self.db.path}"

    def _save_last_config(self, config: ScheduleConfig) -> None:
        QSettings().setValue(self._config_key(), json.dumps({
            "weeks": [config.weeks.start, config.weeks.stop - 1],
            "weekdays": list(config.weekdays),
            "blocks": list(config.blocks),
            "per_slot": config.per_slot,
            "max_per_week": config.max_per_week,
            "max_per_day": config.max_per_day,
            "seed": config.seed,
        }))

    def _load_last_config(self) -> ScheduleConfig | None:
        raw = QSettings().value(self._config_key(), "", type=str)
        if not raw:
            return None
        try:
            data = json.loads(raw)
            lo, hi = int(data["weeks"][0]), int(data["weeks"][1])
            return ScheduleConfig(
                weeks=range(lo, hi + 1),
                weekdays=[int(x) for x in data["weekdays"]],
                blocks=[int(x) for x in data["blocks"]],
                per_slot=int(data.get("per_slot", 1)),
                max_per_week=int(data.get("max_per_week", 3)),
                max_per_day=int(data.get("max_per_day", 1)),
                seed=int(data.get("seed", 42)))
        except (KeyError, IndexError, TypeError, ValueError):
            return None

    def _on_term_start_changed(self) -> None:
        """学期起始日只影响日期显示，改完立即刷新结果表"""
        self._term_calendar_cache = None
        self._calendar_error = None
        if self.result is not None:
            self.refresh_schedule_tabs()
        self.refresh_gantt()
        self._refresh_calendar_summary()

    def _add_calendar_row(
        self,
        table: QTableWidget,
        entry: CalendarEntry | None = None,
    ) -> int:
        """向学期日历编辑表追加一行可视化控件。"""
        row = table.rowCount()
        table.insertRow(row)

        date_edit = QDateEdit()
        date_edit.setCalendarPopup(True)
        date_edit.setDisplayFormat("yyyy-MM-dd")
        term_start = self._term_start()
        term_end = term_start + timedelta(days=MAX_WEEK * 7 - 1)
        date_edit.setMinimumDate(QDate(term_start.year, term_start.month, term_start.day))
        date_edit.setMaximumDate(QDate(term_end.year, term_end.month, term_end.day))
        default_date = entry.date if entry else self._term_start()
        date_edit.setDate(QDate(default_date.year, default_date.month, default_date.day))
        table.setCellWidget(row, 0, date_edit)

        kind_combo = QComboBox()
        kind_combo.addItem("放假（off）", "off")
        kind_combo.addItem("补课（class）", "class")
        if entry is not None:
            kind_combo.setCurrentIndex(kind_combo.findData(entry.override_type))
        table.setCellWidget(row, 1, kind_combo)

        week_spin = QSpinBox()
        week_spin.setRange(1, 25)
        week_spin.setValue(entry.maps_to_week if entry and entry.maps_to_week
                           else self.week_from.value())
        table.setCellWidget(row, 2, week_spin)

        weekday_combo = QComboBox()
        for weekday in range(1, 8):
            weekday_combo.addItem(WEEKDAY_LABELS[weekday], weekday)
        if entry is not None and entry.maps_to_weekday is not None:
            weekday_combo.setCurrentIndex(
                weekday_combo.findData(entry.maps_to_weekday))
        table.setCellWidget(row, 3, weekday_combo)

        note_edit = QLineEdit(entry.note if entry else "")
        note_edit.setPlaceholderText("原因或备注（可空）")
        table.setCellWidget(row, 4, note_edit)

        delete_button = QPushButton("删除")
        delete_button.setObjectName("danger")
        delete_button.setFixedWidth(64)

        def sync_class_fields() -> None:
            is_class = kind_combo.currentData() == "class"
            week_spin.setEnabled(is_class)
            weekday_combo.setEnabled(is_class)

        def remove_row() -> None:
            for current in range(table.rowCount()):
                if table.cellWidget(current, 5) is delete_button:
                    table.removeRow(current)
                    return

        kind_combo.currentIndexChanged.connect(sync_class_fields)
        delete_button.clicked.connect(remove_row)
        table.setCellWidget(row, 5, delete_button)
        sync_class_fields()
        table.setRowHeight(row, 38)
        return row

    @staticmethod
    def _calendar_entries_from_table(table: QTableWidget) -> list[CalendarEntry]:
        """从可视化编辑表读取并构造日历覆盖项。"""
        entries: list[CalendarEntry] = []
        seen_dates: set[date] = set()
        for row in range(table.rowCount()):
            date_edit = table.cellWidget(row, 0)
            kind_combo = table.cellWidget(row, 1)
            week_spin = table.cellWidget(row, 2)
            weekday_combo = table.cellWidget(row, 3)
            note_edit = table.cellWidget(row, 4)
            assert isinstance(date_edit, QDateEdit)
            assert isinstance(kind_combo, QComboBox)
            assert isinstance(week_spin, QSpinBox)
            assert isinstance(weekday_combo, QComboBox)
            assert isinstance(note_edit, QLineEdit)

            entry_date = date_edit.date().toPython()
            if entry_date in seen_dates:
                raise ValueError(f"第 {row + 1} 行日期重复：{entry_date.isoformat()}")
            seen_dates.add(entry_date)
            kind = str(kind_combo.currentData())
            if kind == "off":
                entries.append(CalendarEntry(
                    entry_date, "off", note=note_edit.text().strip()))
            else:
                entries.append(CalendarEntry(
                    entry_date, "class", week_spin.value(),
                    int(weekday_combo.currentData()), note_edit.text().strip()))
        return sorted(entries, key=lambda item: item.date)

    @staticmethod
    def _makeup_entries(
        term_start: date,
        week: int,
        weekday: int,
        makeup_date: date,
    ) -> list[CalendarEntry]:
        """构造一组互补的放假/补课覆盖项。"""
        return [
            CalendarEntry(natural_date(term_start, week, weekday), "off", note="调休放假"),
            CalendarEntry(makeup_date, "class", week, weekday, "周末补课"),
        ]

    def _open_add_makeup_dialog(self, table: QTableWidget) -> None:
        """一键生成互补的 off/class，并在加入编辑表前校验 maps_to。"""
        dlg = QDialog(self)
        dlg.setWindowTitle("添加调休")
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(10)
        layout.addWidget(dialog_header(
            "添加调休",
            "选择代表周次和星期，系统会自动生成放假与周末补课映射。"))
        form = QFormLayout()
        form.setSpacing(10)
        layout.addLayout(form, stretch=1)

        week_spin = QSpinBox()
        week_spin.setRange(1, 25)
        week_spin.setValue(self.week_from.value())
        form.addRow("代表周次", week_spin)

        weekday_combo = QComboBox()
        for weekday in range(1, 8):
            weekday_combo.addItem(WEEKDAY_LABELS[weekday], weekday)
        weekday_combo.setCurrentIndex(weekday_combo.findData(4))
        form.addRow("代表星期", weekday_combo)

        off_date_edit = QDateEdit()
        off_date_edit.setCalendarPopup(True)
        off_date_edit.setDisplayFormat("yyyy-MM-dd")
        off_date_edit.setEnabled(False)
        form.addRow("放假日期", off_date_edit)

        class_date_edit = QDateEdit()
        class_date_edit.setCalendarPopup(True)
        class_date_edit.setDisplayFormat("yyyy-MM-dd")
        term_start = self._term_start()
        term_end = term_start + timedelta(days=MAX_WEEK * 7 - 1)
        class_date_edit.setMinimumDate(
            QDate(term_start.year, term_start.month, term_start.day))
        class_date_edit.setMaximumDate(
            QDate(term_end.year, term_end.month, term_end.day))
        form.addRow("补课日期", class_date_edit)

        def sync_off_date(*_args) -> None:
            off_date = natural_date(
                self._term_start(), week_spin.value(),
                int(weekday_combo.currentData()))
            off_date_edit.setDate(QDate(off_date.year, off_date.month, off_date.day))
            days_to_saturday = (5 - off_date.weekday()) % 7 or 7
            makeup = off_date + timedelta(days=days_to_saturday)
            class_date_edit.setDate(QDate(makeup.year, makeup.month, makeup.day))

        week_spin.valueChanged.connect(sync_off_date)
        weekday_combo.currentIndexChanged.connect(sync_off_date)
        sync_off_date()

        buttons = dialog_buttons("添加")
        layout.addWidget(buttons)

        def confirm() -> None:
            week = week_spin.value()
            weekday = int(weekday_combo.currentData())
            makeup_date = class_date_edit.date().toPython()
            entries = self._makeup_entries(
                self._term_start(), week, weekday, makeup_date)
            try:
                self.db.validate_calendar(
                    self._term_start(), entries=entries)
                existing = self._calendar_entries_from_table(table)
                existing_dates = {entry.date for entry in existing}
                duplicated = [entry for entry in entries if entry.date in existing_dates]
                if duplicated:
                    dates = "、".join(entry.date.isoformat() for entry in duplicated)
                    raise ValueError(f"日期已存在：{dates}")
            except ValueError as exc:
                QMessageBox.warning(dlg, "无法添加调休", str(exc))
                return
            for entry in entries:
                self._add_calendar_row(table, entry)
            dlg.accept()

        buttons.accepted.connect(confirm)
        buttons.rejected.connect(dlg.reject)
        dlg.exec()

    def _open_national_holiday_import_dialog(self, table: QTableWidget) -> None:
        """预览国家调休数据，确认后合并到当前学期日历编辑表。"""
        dlg = QDialog(self)
        dlg.setWindowTitle("导入国家调休日历")
        dlg.resize(960, 620)
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(10)
        layout.addWidget(dialog_header(
            "导入国家调休日历",
            "加载法定节假日和调休上班日；补课日代表的逻辑周和星期会自动推断，请逐行确认。"))

        controls = QHBoxLayout()
        controls.addWidget(QLabel("数据年份"))
        year_spin = QSpinBox()
        year_spin.setRange(2004, 2100)
        year_spin.setValue(self._term_start().year)
        controls.addWidget(year_spin)
        load_button = QPushButton("加载内置数据")
        online_button = QPushButton("在线更新")
        controls.addWidget(load_button)
        controls.addWidget(online_button)
        controls.addStretch(1)
        source_label = QLabel("尚未加载")
        source_label.setObjectName("secondary")
        controls.addWidget(source_label)
        layout.addLayout(controls)

        preview = QTableWidget(0, 6)
        preview.setHorizontalHeaderLabels(
            ["日期", "类型", "代表周", "代表星期", "备注", "操作"])
        preview.verticalHeader().setVisible(False)
        preview.setSelectionMode(QAbstractItemView.NoSelection)
        preview.setEditTriggers(QAbstractItemView.NoEditTriggers)
        preview_header = preview.horizontalHeader()
        preview_header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        preview_header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        preview_header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        preview_header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        preview_header.setSectionResizeMode(4, QHeaderView.Stretch)
        preview_header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        layout.addWidget(preview, stretch=1)

        warning_label = QLabel("")
        warning_label.setObjectName("secondary")
        warning_label.setWordWrap(True)
        layout.addWidget(warning_label)

        buttons = dialog_buttons("导入到日历")
        layout.addWidget(buttons)

        def populate(data) -> None:
            if not shiboken6.isValid(dlg):
                return
            preview.setRowCount(0)
            plan = build_holiday_import_plan(self._term_start(), data)
            for entry in plan.entries:
                self._add_calendar_row(preview, entry)
            source_label.setText(f"来源：{data.source}")
            warning_label.setText(
                "\n".join(plan.warnings) if plan.warnings else
                f"共生成 {len(plan.entries)} 项，可在预览中修改或删除。")

        def load_builtin(*_args) -> None:
            data = get_builtin_holiday_data(year_spin.value())
            if data is None:
                source_label.setText(f"{year_spin.value()} 年暂无内置数据")
                warning_label.setText("请点击「在线更新」获取该年度数据。")
                return
            populate(data)

        def update_online() -> None:
            year = year_spin.value()
            self.run_async(
                f"正在获取 {year} 年国家调休数据",
                lambda: fetch_holiday_data(year),
                populate,
            )

        def import_preview() -> None:
            try:
                incoming = self._calendar_entries_from_table(preview)
                if not incoming:
                    raise ValueError("预览中没有可导入的日历项")
                existing = self._calendar_entries_from_table(table)
                existing_by_date = {entry.date: entry for entry in existing}
                conflicts = [
                    entry for entry in incoming
                    if entry.date in existing_by_date
                    and existing_by_date[entry.date] != entry
                ]
                if conflicts:
                    dates = "、".join(entry.date.isoformat() for entry in conflicts)
                    raise ValueError(f"以下日期已存在，请先在预览中删除：{dates}")
                incoming = [
                    entry for entry in incoming if entry.date not in existing_by_date
                ]
                self.db.validate_calendar(
                    self._term_start(), entries=[*existing, *incoming],
                    allow_unpaired_off=True)
            except ValueError as exc:
                QMessageBox.warning(dlg, "无法导入", str(exc))
                return
            for entry in incoming:
                self._add_calendar_row(table, entry)
            dlg.accept()

        load_button.clicked.connect(load_builtin)
        online_button.clicked.connect(update_online)
        year_spin.valueChanged.connect(load_builtin)
        buttons.accepted.connect(import_preview)
        buttons.rejected.connect(dlg.reject)
        load_builtin()
        dlg.exec()

    def _refresh_calendar_summary(self) -> None:
        self.term_calendar()
        if self._calendar_error:
            self.calendar_label.setText("配置异常，已暂停应用")
            self.calendar_label.setToolTip(self._calendar_error)
            return
        entries = self.calendar_entries()
        self.calendar_label.setToolTip("")
        if not entries:
            self.calendar_label.setText("未设置")
            return
        class_count = sum(e.override_type == "class" for e in entries)
        off_count = len(entries) - class_count
        self.calendar_label.setText(f"{len(entries)} 项（放假 {off_count} / 补课 {class_count}）")

    def open_calendar_dialog(self) -> None:
        """通过日期选择框编辑学期日历覆盖，保存前执行成对一致性校验。"""
        dlg = QDialog(self)
        dlg.setWindowTitle("学期日历覆盖")
        dlg.resize(900, 520)
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(10)
        layout.addWidget(dialog_header(
            "学期日历覆盖",
            "统一维护法定假日与周末补课映射，保存前会自动检查成对关系。"))
        tip = QLabel(
            "「放假 off」表示逻辑教学日放假；「补课 class」选择周末日期，"
            "并指定它代表的逻辑教学周与星期。")
        if self._calendar_error:
            tip.setText(f"当前日历配置无效，已暂停应用：{self._calendar_error}")
            tip.setStyleSheet("color: #b3261e;")
        tip.setObjectName("secondary")
        tip.setWordWrap(True)
        layout.addWidget(tip)

        table = QTableWidget(0, 6)
        table.setHorizontalHeaderLabels(
            ["日期", "类型", "代表周", "代表星期", "备注", "操作"])
        table.verticalHeader().setVisible(False)
        table.setSelectionMode(QAbstractItemView.NoSelection)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.Stretch)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        for entry in self.calendar_entries():
            self._add_calendar_row(table, entry)
        layout.addWidget(table, stretch=1)

        add_button = QPushButton("添加覆盖项")
        add_button.clicked.connect(lambda: self._add_calendar_row(table))
        makeup_button = QPushButton("一键添加调休")
        makeup_button.clicked.connect(
            lambda: self._open_add_makeup_dialog(table))
        holiday_button = QPushButton("从国家调休导入…")
        holiday_button.clicked.connect(
            lambda: self._open_national_holiday_import_dialog(table))
        validate_button = QPushButton("校验")
        allow_unpaired = QCheckBox("允许纯法定假日（off 无需补课配对）")
        allow_unpaired.setChecked(True)
        allow_unpaired.setToolTip(
            "默认勾选：off 可表示无补课日的纯法定假日。"
            "取消勾选后，每个 off 都必须有唯一 class 项与其成对。")

        def validate_current() -> None:
            try:
                entries = self._calendar_entries_from_table(table)
                self.db.validate_calendar(
                    self._term_start(), entries=entries,
                    allow_unpaired_off=allow_unpaired.isChecked())
            except ValueError as exc:
                QMessageBox.warning(dlg, "日历校验失败", str(exc))
                return
            QMessageBox.information(dlg, "日历校验", "日历覆盖项校验通过。")

        validate_button.clicked.connect(validate_current)
        options = QHBoxLayout()
        options.addWidget(add_button)
        options.addWidget(makeup_button)
        options.addWidget(holiday_button)
        options.addWidget(validate_button)
        options.addStretch(1)
        options.addWidget(allow_unpaired)
        layout.addLayout(options)
        buttons = dialog_buttons("保存")
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        layout.addWidget(buttons)
        if dlg.exec() != QDialog.Accepted:
            return
        try:
            entries = self._calendar_entries_from_table(table)
            self.db.validate_calendar(
                self._term_start(), entries=entries,
                allow_unpaired_off=allow_unpaired.isChecked())
        except ValueError as exc:
            QMessageBox.warning(self, "日历校验失败", str(exc))
            return
        self.db.replace_calendar(entries)
        self.invalidate_cache()
        self._refresh_calendar_summary()
        self.refresh_gantt()
        if self.result is not None:
            self.refresh_schedule_tabs()
            self._mark_stale()
        self.statusBar().showMessage(f"学期日历已更新（{len(entries)} 项）。")

    def closeEvent(self, event) -> None:
        # 关闭时可能有后台任务在跑：置位后回调直接返回，避免回写到已关闭的界面
        self._closing = True
        self._save_settings()
        super().closeEvent(event)

    def _set_app_status(self, text: str, tone: str = "neutral") -> None:
        tones = {
            "neutral": ("#475569", "#eef2f7", "#dbe3ee"),
            "info": ("#1d4ed8", "#dbeafe", "#bfdbfe"),
            "success": ("#047857", "#d1fae5", "#a7f3d0"),
            "warning": ("#b45309", "#fef3c7", "#fde68a"),
            "danger": ("#b91c1c", "#fee2e2", "#fecaca"),
        }
        foreground, background, border = tones.get(tone, tones["neutral"])
        self.app_status_label.setText(text)
        self.app_status_label.setStyleSheet(
            f"color:{foreground};background:{background};border:1px solid {border};")

    def _reset_metric_cards(self) -> None:
        for key in ("duties", "members", "spread", "gaps"):
            getattr(self, f"metric_{key}").setText("—")
            hint = getattr(self, f"metric_{key}_hint")
            hint.setText("等待排班结果")
            hint.setStyleSheet("")
        self._set_app_status("等待生成排班")

    def _refresh_schedule_summary(self) -> None:
        result = self.result
        if result is None:
            self._reset_metric_cards()
            return
        n = len(result.assignments)
        member_count = len(result.member_stats)
        slot_count = len({(a.week, a.weekday, a.block) for a in result.assignments})
        self.metric_duties.setText(f"{n} 人次")
        self.metric_duties_hint.setText(f"覆盖 {slot_count} 个时段")
        self.metric_members.setText(f"{member_count} 人")
        self.metric_members_hint.setText(
            f"人均 {n / max(member_count, 1):.1f} 次")
        self.metric_spread.setText(f"{result.balanced_spread} 次")
        self.metric_spread_hint.setText("越小越均衡")
        gap_count = len(result.gaps)
        self.metric_gaps.setText(f"{gap_count} 个")
        self.metric_gaps_hint.setText(
            "已全部覆盖" if gap_count == 0 else "需要处理")
        self.metric_gaps_hint.setStyleSheet(
            "" if gap_count == 0 else "color:#d97706;")

        summary = (
            f"当前结果共 {n} 人次安排，参与成员 {member_count} 人，"
            f"总次数极差 {result.balanced_spread}。")
        if result.gaps:
            summary += capacity_advice(
                summarize_gap_causes(self._gap_diagnoses(result)),
                self._config_per_slot(result),
            )
        self.summary_label.setText(summary)
        self._set_app_status(
            "排班已生成" if gap_count == 0 else "存在待处理缺口",
            "success" if gap_count == 0 else "warning",
        )
        has_data = bool(result.assignments)
        self.btn_export_xlsx.setEnabled(has_data)
        self.btn_export_csv.setEnabled(has_data)
        self.btn_export_png.setEnabled(has_data)
        self.btn_clear_schedule.setEnabled(has_data)
        self.action_export_excel.setEnabled(has_data)
        self.action_export_csv.setEnabled(has_data)
        self.action_export_png.setEnabled(has_data)
        self.action_clear_schedule.setEnabled(has_data)

    def _refresh_pivot_section(self) -> None:
        if self.result is None:
            self.pivot_stack.setCurrentIndex(1)
            return
        self.pivot_stack.setCurrentIndex(0)
        self._refresh_schedule_summary()
        fill_table(self.pivot_table, self._pivot_display())

    def _refresh_detail_section(self) -> None:
        if self.result is None:
            return
        fill_table(self.detail_table, build_detail_df(
            self.result.assignments,
            start_date=self._term_start(),
            calendar=self.term_calendar(),
        ))

    def _refresh_stats_section(self) -> None:
        result = self.result
        if result is None:
            return
        fill_table(self.stats_table, build_stats_df(result.member_stats))
        if result.gaps:
            self.gap_title.setText(
                f"无人可值时段（{len(result.gaps)} 个，按成因分类，可对照处理）")
            fill_table(self.gap_table, build_gap_df(self._gap_diagnoses(result)))
        else:
            self.gap_title.setText("所有值班时段均已安排到位，无缺口。")
            fill_table(self.gap_table, pd.DataFrame(
                columns=["周次", "星期", "时段", "主要原因", "无课人数"]))

    def _refresh_schedule_section(self, section: str) -> None:
        if section == "pivot":
            self._refresh_pivot_section()
        elif section == "detail":
            self._refresh_detail_section()
        elif section == "stats":
            self._refresh_stats_section()
        elif section == "gantt":
            self.refresh_gantt()
        elif section == "charts":
            self.refresh_charts()
        self._dirty_schedule_sections.discard(section)

    def refresh_schedule_tabs(
        self,
        sections: set[str] | None = None,
    ) -> None:
        """刷新排班相关区域；指定 sections 时只重算所需部分。"""
        result = self.result
        if result is None:
            for table in (self.pivot_table, self.detail_table,
                          self.stats_table, self.gap_table):
                fill_table(table, pd.DataFrame())
            self._update_empty_state()
            self.gap_title.setText("无人可值时段")
            self.btn_export_xlsx.setEnabled(False)
            self.btn_export_csv.setEnabled(False)
            self.btn_export_png.setEnabled(False)
            has_stored = bool(self.db.load_assignments())
            self.btn_clear_schedule.setEnabled(has_stored)
            self.action_export_excel.setEnabled(False)
            self.action_export_csv.setEnabled(False)
            self.action_export_png.setEnabled(False)
            self.action_clear_schedule.setEnabled(has_stored)
            self.refresh_charts()
            self._dirty_schedule_sections.clear()
            return

        selected = sections or {"pivot", "detail", "stats", "charts"}
        if "pivot" in selected:
            self._refresh_pivot_section()
        if "detail" in selected:
            self._refresh_detail_section()
        if "stats" in selected:
            self._refresh_stats_section()
        if "charts" in selected:
            self.refresh_charts()
        self._dirty_schedule_sections.difference_update(selected)

    def _refresh_after_schedule_change(self) -> None:
        """微调后只刷新当前可见区域，其余页签按需延迟刷新。"""
        section_by_tab = {
            0: "pivot",
            1: "detail",
            3: "stats",
            4: "gantt",
            5: "charts",
        }
        affected = {"pivot", "detail", "stats", "gantt", "charts"}
        current = section_by_tab.get(self.tabs.currentIndex())
        if current in affected:
            self._refresh_schedule_section(current)
        self._dirty_schedule_sections.update(affected - {current})

    # ---------- 缺口诊断 ----------

    def _config_per_slot(self, result: ScheduleResult) -> int:
        cfg = self._last_config or self._load_last_config()
        return cfg.per_slot if cfg is not None else 1

    def _gap_diagnoses(self, result: ScheduleResult) -> list:
        """缺口成因诊断（带缓存：参数不变时同一结果只算一次）"""
        cached = self._gap_cache
        if cached is not None and cached[0] is result:
            return cached[1]
        cfg = self._last_config or self._load_last_config() or self._current_config()
        diagnoses = diagnose_gaps(
            self.scheduling_members(), self.busy_map(), self.leaves(), result.assignments,
            cfg, special_arrangements=self.specials(), is_off=self.is_off_day,
            is_class=self.is_class_day)
        self._gap_cache = (result, diagnoses)
        return diagnoses

    def _term_start(self) -> date:
        return self.term_start.date().toPython()

    def _pivot_frame(self, assignments: list[Assignment]) -> pd.DataFrame:
        """扁平化透视表（周次/星期 + 日期 + 各时段），供表格展示与 PNG 导出"""
        pivot = build_pivot_df(
            assignments, start_date=self._term_start(), calendar=self.term_calendar())
        return pd.concat(
            [pd.DataFrame(pivot.index.tolist(), columns=["周次", "星期"]),
             pivot.reset_index(drop=True)], axis=1)

    def _pivot_display(self) -> pd.DataFrame:
        """值班排班表的展示数据：周次/星期 + 日期 + 各时段，并记录微调定位映射"""
        result = self.result
        # build_pivot_df 返回前会把索引字符串化（"第1周"/"周一"），
        # 这里按相同顺序用原始整数重建行映射，供微调对话框定位
        weeks = sorted({a.week for a in result.assignments})
        weekdays = sorted({a.weekday for a in result.assignments})
        calendar = self.term_calendar()
        self._pivot_rows = [
            (w, d) for w in weeks for d in weekdays
            if calendar.logical_to_date(w, d) is not None
        ]
        self._pivot_blocks = sorted({a.block for a in result.assignments})
        self._pivot_date_offset = 1  # 展示列顺序：周次、星期、日期、各时段
        return self._pivot_frame(result.assignments)

    def show_member_courses(self, index: int) -> None:
        if index < 0:
            return
        member_id = self.member_combo.itemData(index)
        m = self.db.get_member(member_id)
        if m is None:
            return
        participation = "参与排班" if m.participates_in_scheduling else "不参与排班"
        self.member_info.setText(
            f"学号 {m.student_id or '—'} | 工作室 {m.studio or UNKNOWN_STUDIO} | "
            f"{m.term or '—'} | {m.major or '—'} | "
            f"{m.department or '—'} | {participation} | 来源：{m.file_name or '—'}")
        # 整周集中安排（军训/思政实践）没有星期与节次，排序键与展示都要单独处理
        courses = sorted(self.db.get_courses(member_id),
                         key=lambda c: (c.weekday or 99, min(c.session_list or [0]),
                                        c.course_name))
        df = pd.DataFrame([{
            "星期": WEEKDAY_LABELS.get(c.weekday, "整周"), "课程": c.course_name,
            "教师": c.teacher or "—", "周次": c.weeks_text,
            "节次": f"{c.sessions_text}节" if c.sessions_text else "整周",
            "地点": c.location or "—",
        } for c in courses])
        fill_table(self.course_table, df if not df.empty else
                   pd.DataFrame(columns=["星期", "课程", "教师", "周次", "节次", "地点"]))
        self._refresh_special_arrangements(member_id)
        self._refresh_leaves(member_id)

    def _on_member_combo_changed(self, index: int) -> None:
        """切换查看成员时清空微调历史，避免跨成员撤销。"""
        self._clear_undo_redo()
        self.show_member_courses(index)

    # ---------- 空闲时段总览 ----------

    def _on_tab_changed(self, index: int) -> None:
        tab = self.tabs.tabText(index)
        section_by_tab = {
            0: "pivot",
            1: "detail",
            3: "stats",
            4: "gantt",
            5: "charts",
        }
        section = section_by_tab.get(index)
        if section in self._dirty_schedule_sections:
            self._refresh_schedule_section(section)
        elif tab == "空闲时段总览":
            self.refresh_gantt()
        elif tab == "统计图表":
            self.refresh_charts()

    def _on_gantt_filter_changed(self) -> None:
        if self.tabs.tabText(self.tabs.currentIndex()) == "空闲时段总览":
            self.refresh_gantt()

    def _sync_gantt_week(self, week: int) -> None:
        """值班起始周变化时空闲时段总览跟随，切到总览页即在排班起始周"""
        self.gantt_week.setValue(week)

    def refresh_gantt(self) -> None:
        members = self.gantt_members()
        all_member_count = len(self.members())
        checked_weekdays = [
            d for d, cb in self.weekday_checks.items() if cb.isChecked()]
        blocks = [b for b, cb in self.block_checks.items() if cb.isChecked()]
        calendar = self.term_calendar()
        week = self.gantt_week.value()
        if not members or not blocks:
            self.gantt_matrix = None
            self.gantt_table.clearContents()
            self.gantt_table.setRowCount(0)
            self.gantt_table.setColumnCount(0)
            if all_member_count and not members and self.roster_entries():
                self.gantt_hint.setText(
                    "空闲时段总览仅显示短视频工作室和图片工作室成员。"
                    "当前没有这两个工作室的参与成员；如归属有误，请在成员列表使用"
                    "「修改工作室…」调整。")
            elif all_member_count and not members:
                self.gantt_hint.setText(
                    "当前没有参与排班的成员；请通过菜单「排班 → 设置参与排班…」"
                    "选择成员后查看空闲时段总览。")
            else:
                self.gantt_hint.setText(
                    "请先上传成员课表，并至少勾选一个值班星期和值班时段"
                    "（空闲时段总览跟随左侧筛选）。")
            self.btn_export_gantt.setEnabled(False)
            return

        matrix = build_calendar_availability(
            members, self.courses(), calendar,
            week, checked_weekdays, blocks,
            leaves=self.leaves(),
            assignments=self.result.assignments if self.result else None,
            busy=self.busy_map(),
            special_arrangements=self.specials(),
        )
        self.gantt_matrix = matrix
        self._fill_gantt_table(matrix)

        all_free = matrix.all_free_columns
        if all_free:
            text = "、".join(
                f"{column.date.month}月{column.date.day}日"
                f" {WEEKDAY_LABELS[column.date.isoweekday()]}"
                f" {BLOCK_LABELS[column.block].split(' ')[0]}"
                for column in all_free)
            self.gantt_hint.setText(
                f"第{matrix.week}周全员空闲时段共 {len(all_free)} 个（橙色高亮行）：{text}。"
                "适合安排需要全员参加的任务。")
        else:
            self.gantt_hint.setText(
                f"第{matrix.week}周没有全员空闲的时段；汇总列为各时段空闲人数，"
                "选择空闲人数最多的时段最容易凑齐人。")
        self.btn_export_gantt.setEnabled(True)

    @staticmethod
    def _apply_cell(item: QTableWidgetItem, state: tuple) -> None:
        """把期望状态写入单元格（仅在状态变化时调用）"""
        text, bg, fg, tooltip, centered = state
        item.setText(text)
        item.setBackground(QBrush(bg))
        item.setForeground(QBrush(fg) if fg is not None else QBrush())
        item.setToolTip(tooltip)
        if centered:
            item.setTextAlignment(Qt.AlignCenter)
        else:
            item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)

    @staticmethod
    def _cell_state(item: QTableWidgetItem) -> tuple:
        """读取单元格当前状态，与 _apply_cell 的元组一一对应"""
        brush = item.foreground()
        fg = None if brush.style() == Qt.NoBrush else brush.color()
        centered = bool(item.textAlignment() & Qt.AlignCenter)
        return (item.text(), item.background().color(), fg, item.toolTip(), centered)

    def _fill_cell(self, t: QTableWidget, r: int, c: int, state: tuple,
                   cache: dict) -> None:
        """带 Python 侧缓存地设置单元格

        先在 Python 里比较期望状态与上次写入的状态：完全一致就直接跳过，
        一次 Qt 属性访问都不做。读取 Qt 属性（text/background/foreground/
        toolTip/alignment）看似便宜，但成员多时每帧要读数万次，实测是
        空闲时段总览刷新里最大的一块开销。
        """
        key = (r, c)
        if cache.get(key) == state:
            return
        item = t.item(r, c)
        if item is None:
            item = QTableWidgetItem()
            item.setFlags(Qt.ItemIsEnabled)
            t.setItem(r, c, item)
        elif item.flags() & Qt.ItemIsEditable:
            item.setFlags(Qt.ItemIsEnabled)
        self._apply_cell(item, state)
        cache[key] = state

    def _fill_gantt_table(self, m: CalendarAvailabilityMatrix) -> None:
        """行=自然日期x时段，列=成员（末列为汇总）；
        空闲绿色、有课灰色、请假红色、特殊安排紫色、值班蓝色、全员空闲橙色。

        复用已有 QTableWidgetItem：成员多时表格有数千个单元格，
        每次重建 item 会造成明显卡顿，因此只在结构变化时重建行列表头。
        """
        t = self.gantt_table
        FREE, BUSY, ALL_FREE, LEAVE, SPECIAL, DUTY = (
            QColor("#c9f2cf"), QColor("#f2f2f7"), QColor("#ffd9a8"),
            QColor("#ffd6d2"), QColor("#e5d8ff"), QColor("#b8d9ff"))
        NO_COLOR = QColor("#1d1d1f")
        BUSY_COLOR = QColor("#aeaeb2")
        LEAVE_COLOR = QColor("#b3261e")
        SPECIAL_COLOR = QColor("#6e3dc2")
        DUTY_COLOR = QColor("#0a5aa8")
        ALL_FREE_COLOR = QColor("#b25e00")

        columns = m.columns
        rows, cols = len(columns), m.member_count + 1
        if t.rowCount() != rows or t.columnCount() != cols:
            t.clearContents()
            t.setRowCount(rows)
            t.setColumnCount(cols)
            cache = {}
            self._gantt_cell_state = cache
        else:
            cache = self._gantt_cell_state
        t.setHorizontalHeaderLabels(m.member_names + ["空闲人数"])
        t.setVerticalHeaderLabels([
            slot_header(c.date.isoweekday(), c.block, c.date, is_off=c.is_off)
            for c in columns])

        week = m.week
        for r in range(m.member_count):
            free_row = m.free[r]
            for i, column in enumerate(columns):
                actual = column.date
                weekday = actual.isoweekday()
                b = column.block
                leave_reason = m.leave_info.get((r, actual))
                if column.is_off:
                    date_text = f"（{actual.month}月{actual.day}日）"
                    state = ("假", LEAVE, LEAVE_COLOR,
                             f"第{week}周 {WEEKDAY_LABELS[weekday]}{date_text}：法定假日放假", True)
                elif leave_reason is not None:
                    reason = f"（{leave_reason}）" if leave_reason else ""
                    state = ("假", LEAVE, LEAVE_COLOR,
                             f"第{week}周 {WEEKDAY_LABELS[weekday]}：请假{reason}", True)
                elif (r, actual, b) in m.duty_cells:
                    state = ("值", DUTY, DUTY_COLOR,
                             f"第{week}周 {WEEKDAY_LABELS[weekday]} {BLOCK_LABELS[b]}：已排值班", True)
                elif (r, actual, b) in m.special_info:
                    reasons = m.special_info[(r, actual, b)]
                    state = ("特", SPECIAL, SPECIAL_COLOR,
                             f"第{week}周 {WEEKDAY_LABELS[weekday]} {BLOCK_LABELS[b]}：其他安排\n"
                             + "\n".join(reasons), True)
                else:
                    if free_row[i]:
                        state = ("", FREE, None,
                                 f"第{week}周 {WEEKDAY_LABELS[weekday]} {BLOCK_LABELS[b]}：空闲",
                                 False)
                    else:
                        names = m.busy_courses.get((r, actual, b), [])
                        state = ("课", BUSY, BUSY_COLOR,
                                 f"第{week}周 {WEEKDAY_LABELS[weekday]} {BLOCK_LABELS[b]}：\n"
                                 + "\n".join(names), True)
                self._fill_cell(t, i, r, state, cache)

        summary_col = m.member_count
        for i, column in enumerate(columns):
            if column.is_off:
                state = ("放假", LEAVE, LEAVE_COLOR,
                         f"第{week}周 {WEEKDAY_LABELS[column.date.isoweekday()]}：法定假日放假",
                         True)
            else:
                count = m.free_counts[i]
                all_free = count == m.member_count
                state = (f"{count}/{m.member_count}",
                         ALL_FREE if all_free else QColor(0, 0, 0, 0),
                         ALL_FREE_COLOR if all_free else NO_COLOR, "", True)
            self._fill_cell(t, i, summary_col, state, cache)
            head = t.verticalHeaderItem(i)
            if head is not None:
                highlighted = column.is_off or (
                    not column.is_off and m.free_counts[i] == m.member_count)
                head.setForeground(QBrush(QColor(
                    "#ff9500" if highlighted else "#86868b")))

    def export_gantt(self) -> None:
        m = self.gantt_matrix
        if m is None:
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "导出空闲时段总览 Excel", f"空闲时段总览_第{m.week}周.xlsx", "Excel 文件 (*.xlsx)")
        if not path:
            return
        target = Path(path)
        self.run_async(
            "正在导出空闲时段总览",
            lambda: target.write_bytes(export_gantt_excel(m)),
            lambda _r: self.statusBar().showMessage(f"已导出空闲时段总览 Excel：{target}"))

    # ---------- 统计图表 ----------

    def refresh_charts(self) -> None:
        """刷新三张统计图：成员值班次数 / 每周值班人次 / 成员请假天数"""
        members = self.scheduling_members()
        member_ids = {m.id for m in members}
        if self.result is not None:
            assignments = [
                a for a in self.result.assignments if a.member_id in member_ids]
            stats = rebuild_member_stats(members, assignments)
            self.chart_duty.set_data([(s["name"], s["total"]) for s in stats.values()])
            week_cnt = Counter(a.week for a in assignments)
            self.chart_weekly.set_data(
                [(str(w), week_cnt[w])
                 for w in sorted({a.week for a in assignments})])
        else:
            self.chart_duty.set_data([])
            self.chart_weekly.set_data([])
        leave_cnt = Counter(
            l.member_id for l in self.leaves() if l.member_id in member_ids)
        name_of = {m.id: m.name for m in members}
        self.chart_leave.set_data(sorted(
            ((name_of.get(mid, "已删除成员"), n) for mid, n in leave_cnt.items()),
            key=lambda t: (-t[1], t[0])))
        self.btn_export_charts.setEnabled(
            bool(members) and any(c.data() for c in
                                  (self.chart_duty, self.chart_weekly, self.chart_leave)))

    def export_charts(self) -> None:
        charts = [
            ("各成员值班总次数", self.chart_duty.data(), "#2563eb"),
            ("每周值班人次", self.chart_weekly.data(), "#059669"),
            ("各成员请假天数", self.chart_leave.data(), "#dc2626"),
        ]
        path, _ = QFileDialog.getSaveFileName(
            self, "导出统计图", "值班请假统计图.png", "PNG 图片 (*.png)")
        if not path:
            return
        target = Path(path)
        self.run_async(
            "正在导出统计图",
            lambda: render_charts_png(charts, target),
            lambda paths: self.statusBar().showMessage(
                f"已导出统计图（{len(paths)} 张）：{target.parent}"))

    # ---------- 长期特殊安排（修改课表） ----------

    def _refresh_special_arrangements(self, member_id: int) -> None:
        t = self.special_table
        t.setRowCount(0)
        for a in (x for x in self.specials() if x.member_id == member_id):
            r = t.rowCount()
            t.insertRow(r)
            values = (
                special_weeks_label(a.week_start, a.week_end),
                WEEKDAY_LABELS[a.weekday],
                special_sessions_label(a.session_list),
                a.reason or "其他安排",
            )
            for c, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                t.setItem(r, c, item)
            t.item(r, 0).setData(Qt.UserRole, a.id)

    def _selected_special_arrangement(self) -> SpecialArrangement | None:
        row = self.special_table.currentRow()
        if row < 0:
            return None
        arrangement_id = self.special_table.item(row, 0).data(Qt.UserRole)
        return next((a for a in self.specials() if a.id == arrangement_id), None)

    def _open_special_arrangement_dialog(
        self,
        arrangement: SpecialArrangement | None = None,
    ) -> None:
        member_id = self.member_combo.currentData()
        if member_id is None:
            QMessageBox.information(self, "提示", "请先选择成员。")
            return
        member = self.db.get_member(member_id)
        if member is None:
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("修改特殊安排" if arrangement else "添加特殊安排")
        dlg.setMinimumWidth(500)
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(10)
        layout.addWidget(dialog_header(
            "修改特殊安排" if arrangement else "添加特殊安排",
            "补充连续多周的课表外占用，排班和空闲总览会统一避让。"))
        form = QFormLayout()
        form.setSpacing(10)
        layout.addLayout(form, stretch=1)

        tips = QLabel(
            f"成员：{member.name}\n"
            "在连续周范围内，每周所选星期与时段都会作为「其他安排」占用。")
        tips.setObjectName("secondary")
        tips.setWordWrap(True)
        form.addRow(tips)

        week_row = QHBoxLayout()
        week_start = QSpinBox()
        week_start.setRange(1, 25)
        week_start.setValue(arrangement.week_start if arrangement else self.week_from.value())
        week_end = QSpinBox()
        week_end.setRange(1, 25)
        week_end.setValue(arrangement.week_end if arrangement else self.week_to.value())
        week_row.addWidget(week_start)
        week_row.addWidget(QLabel("至"))
        week_row.addWidget(week_end)
        form.addRow("连续周次", week_row)

        weekday = QComboBox()
        for day in range(1, 8):
            weekday.addItem(WEEKDAY_LABELS[day], day)
        if arrangement:
            weekday.setCurrentIndex(weekday.findData(arrangement.weekday))
        form.addRow("星期", weekday)

        block_row = QHBoxLayout()
        block_checks: dict[int, QCheckBox] = {}
        for block in sorted(BLOCK_SESSIONS):
            checkbox = QCheckBox(BLOCK_LABELS[block].split(" ")[0])
            checkbox.setToolTip(BLOCK_LABELS[block])
            if arrangement and set(BLOCK_SESSIONS[block]) <= set(arrangement.session_list):
                checkbox.setChecked(True)
            block_checks[block] = checkbox
            block_row.addWidget(checkbox)
        form.addRow("占用时段", block_row)

        reason = QLineEdit(arrangement.reason if arrangement else "")
        reason.setPlaceholderText("如：固定实习、球队训练、长期治疗（可空）")
        form.addRow("原因", reason)

        buttons = dialog_buttons("保存")
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        layout.addWidget(buttons)

        if dlg.exec() != QDialog.Accepted:
            return
        if week_start.value() > week_end.value():
            QMessageBox.warning(self, "无法保存", "起始周不能大于结束周。")
            return
        selected_blocks = [b for b, checkbox in block_checks.items() if checkbox.isChecked()]
        if not selected_blocks:
            QMessageBox.warning(self, "无法保存", "请至少选择一个占用时段。")
            return
        sessions = sorted({
            session
            for block in selected_blocks
            for session in BLOCK_SESSIONS[block]
        })
        reason_text = reason.text().strip()
        if arrangement is None:
            self.db.add_special_arrangement(
                member_id, week_start.value(), week_end.value(),
                weekday.currentData(), sessions, reason_text)
        else:
            self.db.update_special_arrangement(
                arrangement.id, member_id, week_start.value(), week_end.value(),
                weekday.currentData(), sessions, reason_text)
        self.invalidate_cache()
        self._refresh_special_arrangements(member_id)
        self.refresh_gantt()
        self._mark_stale()
        verb = "已修改" if arrangement else "已添加"
        self.statusBar().showMessage(
            f"{verb} {member.name} 的课表特殊安排；重新生成排班后将自动避让。")

    def add_special_arrangement(self) -> None:
        self._open_special_arrangement_dialog()

    def edit_special_arrangement(self) -> None:
        arrangement = self._selected_special_arrangement()
        if arrangement is None:
            QMessageBox.information(self, "提示", "请先在特殊安排表中选择一条记录。")
            return
        self._open_special_arrangement_dialog(arrangement)

    def remove_special_arrangement(self) -> None:
        arrangement = self._selected_special_arrangement()
        if arrangement is None:
            QMessageBox.information(self, "提示", "请先在特殊安排表中选择一条记录。")
            return
        member_id = self.member_combo.currentData()
        self.db.remove_special_arrangement(arrangement.id)
        self.invalidate_cache()
        self._refresh_special_arrangements(member_id)
        self.refresh_gantt()
        self._mark_stale()
        self.statusBar().showMessage("已删除该课表特殊安排；重新生成排班后将恢复该时段可排。")

    # ---------- 请假登记 ----------

    def _refresh_leaves(self, member_id: int) -> None:
        t = self.leave_table
        t.setRowCount(0)
        for l in self.db.list_leaves(member_id):
            r = t.rowCount()
            t.insertRow(r)
            for c, v in enumerate((f"第{l.week}周", WEEKDAY_LABELS[l.weekday], l.reason or "—")):
                item = QTableWidgetItem(v)
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                t.setItem(r, c, item)
            t.item(r, 0).setData(Qt.UserRole, l.id)

    def add_leave(self) -> None:
        member_id = self.member_combo.currentData()
        if member_id is None:
            QMessageBox.information(self, "提示", "请先选择成员。")
            return
        dlg = QDialog(self)
        dlg.setWindowTitle("添加请假")
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(10)
        layout.addWidget(dialog_header(
            "添加请假",
            "登记后该成员当天不再参与排班，空闲时段总览会将整行标红。"))
        form = QFormLayout()
        form.setSpacing(10)
        layout.addLayout(form, stretch=1)
        week = QSpinBox()
        week.setRange(1, 25)
        week.setValue(self.gantt_week.value())
        wd = QComboBox()
        for d in range(1, 8):
            wd.addItem(WEEKDAY_LABELS[d], d)
        reason = QLineEdit()
        reason.setPlaceholderText("如：生病、比赛、社团活动（可空）")
        form.addRow("周次", week)
        form.addRow("星期", wd)
        form.addRow("原因", reason)
        buttons = dialog_buttons("添加")
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        layout.addWidget(buttons)
        if dlg.exec() != QDialog.Accepted:
            return
        m = self.db.get_member(member_id)
        self.db.add_leave(member_id, week.value(), wd.currentData(), reason.text().strip())
        self.invalidate_cache()
        self._refresh_leaves(member_id)
        self.refresh_gantt()
        self.refresh_charts()
        self._mark_stale()
        self.statusBar().showMessage(
            f"已登记 {m.name} 第{week.value()}周{WEEKDAY_LABELS[wd.currentData()]}请假，"
            "重新生成排班将避开该天。")

    def remove_leave(self) -> None:
        row = self.leave_table.currentRow()
        if row < 0:
            QMessageBox.information(self, "提示", "请先在请假记录中选择一条。")
            return
        leave_id = self.leave_table.item(row, 0).data(Qt.UserRole)
        member_id = self.member_combo.currentData()
        self.db.remove_leave(leave_id)
        self.invalidate_cache()
        self._refresh_leaves(member_id)
        self.refresh_gantt()
        self.refresh_charts()
        self._mark_stale()
        self.statusBar().showMessage("已删除该请假记录。")

    def export_leaves(self) -> None:
        """导出所有成员的请假记录为 Excel（按钮位于请假登记卡片）"""
        leaves = self.leaves()
        if not leaves:
            QMessageBox.information(self, "提示", "暂无请假记录可导出。")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "导出请假记录", "请假记录.xlsx", "Excel 文件 (*.xlsx)")
        if not path:
            return
        members = self.members()
        start = self._term_start()
        calendar = self.term_calendar()
        target = Path(path)
        self.run_async(
            "正在导出请假记录",
            lambda: target.write_bytes(
                export_leaves_excel(
                    leaves, members, start_date=start, calendar=calendar)),
            lambda _r: self.statusBar().showMessage(
                f"已导出请假记录（{len(leaves)} 条）：{target}"))

    # ---------- 手动微调 ----------

    def _on_pivot_cell_double_clicked(self, row: int, col: int) -> None:
        if self.result is None or not self._pivot_rows or col < 2 + self._pivot_date_offset:
            return
        block_idx = col - 2 - self._pivot_date_offset
        if not 0 <= block_idx < len(self._pivot_blocks):
            return
        week, weekday = self._pivot_rows[row]
        self._open_tweak_dialog(week, weekday, self._pivot_blocks[block_idx])

    def _open_tweak_dialog(self, week: int, weekday: int, block: int) -> None:
        cfg = self._last_config
        if cfg is None:
            return
        slot_assignments = [a for a in self.result.assignments
                            if a.week == week and a.weekday == weekday and a.block == block]
        current = {a.member_id: a for a in slot_assignments}
        members = self.scheduling_members()
        busy = self.busy_map()
        leave_set = {(l.member_id, l.week, l.weekday) for l in self.leaves()}
        cands = replacement_candidates(
            members, busy, leave_set, self.result.assignments,
            week, weekday, block, cfg.max_per_week, cfg.max_per_day,
            special_set=self.special_set())

        dlg = QDialog(self)
        dlg.setWindowTitle("手动微调")
        dlg.setMinimumWidth(440)
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(10)
        layout.addWidget(dialog_header(
            "手动微调",
            f"第{week}周 {WEEKDAY_LABELS[weekday]} · {BLOCK_LABELS[block]}　"
            f"当前值班：{'、'.join(a.member_name for a in slot_assignments) or '（空缺）'}"))
        form = QFormLayout()
        form.setSpacing(10)
        layout.addLayout(form, stretch=1)

        target = QComboBox()
        for a in slot_assignments:
            target.addItem(a.member_name, a.member_id)
        if len(current) < cfg.per_slot:
            target.addItem("＋ 新增一人", -1)
        if target.count() == 0:
            QMessageBox.information(self, "提示", "该时段暂无可调整的值班安排。")
            return
        form.addRow("调整对象", target)

        cand_list = QListWidget()
        cand_list.setMinimumHeight(240)
        self._fill_tweak_candidates(cand_list, cands, target.currentData())
        target.currentIndexChanged.connect(
            lambda _: self._fill_tweak_candidates(cand_list, cands, target.currentData()))
        form.addRow("新值班人", cand_list)

        buttons = dialog_buttons("应用")
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        layout.addWidget(buttons)
        if dlg.exec() != QDialog.Accepted:
            return
        sel = cand_list.currentItem()
        if sel is None:
            return
        data = sel.data(Qt.UserRole)
        target_mid = target.currentData()
        if data == "remove":
            self._apply_tweak(week, weekday, block, target_mid, None)
        elif isinstance(data, int):
            out_id = target_mid if target_mid != -1 else None
            self._apply_tweak(week, weekday, block, out_id, data)

    @staticmethod
    def _fill_tweak_candidates(
        lst: QListWidget,
        cands: list[tuple],
        target_mid,
    ) -> None:
        lst.clear()
        if target_mid is not None and target_mid != -1:
            remove_item = QListWidgetItem("（移除该值班人）")
            remove_item.setData(Qt.UserRole, "remove")
            lst.addItem(remove_item)
        if not cands:
            empty = QListWidgetItem("（暂无其他成员可换）")
            empty.setFlags(Qt.ItemIsEnabled)
            lst.addItem(empty)
            return
        for m, reason in cands:
            soft = reason == "本周已达上限"
            item = QListWidgetItem(f"{m.name} — {reason}" if reason else m.name)
            item.setData(Qt.UserRole, m.id)
            if reason and not soft:
                item.setFlags(Qt.ItemIsEnabled)  # 仅展示不可选
                item.setToolTip(reason)
            elif soft:
                item.setToolTip("课程/特殊安排/请假/当天条件均满足，手动换入将超过每周上限，请知悉")
            lst.addItem(item)

    def _slot_member_ids(self, week: int, weekday: int, block: int) -> tuple[int, ...]:
        if self.result is None:
            return ()
        return tuple(sorted(
            a.member_id for a in self.result.assignments
            if a.week == week and a.weekday == weekday and a.block == block
        ))

    def _push_undo(self, snapshot: tuple[int, int, int, tuple[int, ...]]) -> None:
        self._undo_stack.append(snapshot)
        if len(self._undo_stack) > self._undo_limit:
            del self._undo_stack[:len(self._undo_stack) - self._undo_limit]
        self._update_undo_redo_actions()

    def _push_redo(self, snapshot: tuple[int, int, int, tuple[int, ...]]) -> None:
        self._redo_stack.append(snapshot)
        if len(self._redo_stack) > self._undo_limit:
            del self._redo_stack[:len(self._redo_stack) - self._undo_limit]
        self._update_undo_redo_actions()

    def _clear_undo_redo(self) -> None:
        self._undo_stack.clear()
        self._redo_stack.clear()
        self._update_undo_redo_actions()

    def _update_undo_redo_actions(self) -> None:
        if hasattr(self, "undo_action"):
            self.undo_action.setEnabled(bool(self._undo_stack) and self.result is not None)
        if hasattr(self, "redo_action"):
            self.redo_action.setEnabled(bool(self._redo_stack) and self.result is not None)

    def _restore_slot_snapshot(
        self,
        week: int,
        weekday: int,
        block: int,
        member_ids: tuple[int, ...],
    ) -> None:
        """把一个微调快照恢复到结果对象和数据库。"""
        if self.result is None or self._last_config is None:
            return
        members = {m.id: m for m in self.members()}
        kept = [
            a for a in self.result.assignments
            if not (a.week == week and a.weekday == weekday and a.block == block)
        ]
        for member_id in member_ids:
            member = members[member_id]
            kept.append(Assignment(
                week=week, weekday=weekday, block=block,
                member_id=member_id, member_name=member.name))
        self.db.replace_slot(week, weekday, block, list(member_ids))
        self.result.assignments = kept
        self.result.member_stats = rebuild_member_stats(list(members.values()), kept)
        self.result.gaps = compute_gaps(
            kept, self._last_config, is_off=self.is_off_day,
            is_class=self.is_class_day)
        self._gap_cache = None
        self._stale = False
        self._refresh_after_schedule_change()

    def undo_tweak(self, *_args) -> None:
        """撤销一次手动微调。"""
        if not self._undo_stack or self.result is None:
            return
        week, weekday, block, old_ids = self._undo_stack.pop()
        current_ids = self._slot_member_ids(week, weekday, block)
        self._push_redo((week, weekday, block, current_ids))
        self._restore_slot_snapshot(week, weekday, block, old_ids)
        self._update_undo_redo_actions()
        self.statusBar().showMessage(
            f"已撤销第{week}周{WEEKDAY_LABELS[weekday]}"
            f"{BLOCK_LABELS[block].split(' ')[0]}的微调。")

    def redo_tweak(self, *_args) -> None:
        """重做一次手动微调。"""
        if not self._redo_stack or self.result is None:
            return
        week, weekday, block, new_ids = self._redo_stack.pop()
        current_ids = self._slot_member_ids(week, weekday, block)
        self._push_undo((week, weekday, block, current_ids))
        self._restore_slot_snapshot(week, weekday, block, new_ids)
        self._update_undo_redo_actions()
        self.statusBar().showMessage(
            f"已重做第{week}周{WEEKDAY_LABELS[weekday]}"
            f"{BLOCK_LABELS[block].split(' ')[0]}的微调。")

    def _apply_tweak(self, week: int, weekday: int, block: int, out_id: int | None, in_id: int | None) -> None:
        """执行微调：out_id 换出（None=纯新增），in_id 换入（None=纯移除）"""
        if self.result is None or self._last_config is None:
            return
        all_members = {m.id: m for m in self.members()}
        members = self.scheduling_members()
        cfg = self._last_config
        if in_id is not None:
            busy = self.busy_map()
            leave_set = {(l.member_id, l.week, l.weekday) for l in self.leaves()}
            # 硬约束（课程/特殊安排/请假/每天一次）必须满足；周上限允许手动越限
            eligible = {m.id for m, reason in replacement_candidates(
                members, busy, leave_set, self.result.assignments,
                week, weekday, block, cfg.max_per_week, cfg.max_per_day,
                special_set=self.special_set())
                if reason in ("", "本周已达上限")}
            if in_id not in eligible:
                QMessageBox.warning(self, "无法调整", "该成员在此时段不满足值班条件，请重新选择。")
                return
        old_ids = self._slot_member_ids(week, weekday, block)
        kept = []
        for a in self.result.assignments:
            if (a.week == week and a.weekday == weekday and a.block == block
                    and a.member_id in (out_id, in_id)):
                continue
            kept.append(a)
        if in_id is not None:
            kept.append(Assignment(
                week=week, weekday=weekday, block=block,
                member_id=in_id, member_name=all_members[in_id].name))
        new_ids = tuple(sorted(
            a.member_id for a in kept
            if a.week == week and a.weekday == weekday and a.block == block
        ))
        if new_ids != old_ids:
            self._push_undo((week, weekday, block, old_ids))
            self._redo_stack.clear()
            self._update_undo_redo_actions()
        self.result.assignments = kept
        self.result.member_stats = rebuild_member_stats(members, kept)
        self.result.gaps = compute_gaps(
            kept, cfg, is_off=self.is_off_day, is_class=self.is_class_day)
        self._gap_cache = None
        # 只重写被调整的那一个时段，避免全表清空 + 全量重写
        self.db.replace_slot(week, weekday, block,
                             [a.member_id for a in kept
                              if (a.week, a.weekday, a.block) == (week, weekday, block)])
        self._stale = False
        self._refresh_after_schedule_change()
        out_name = all_members[out_id].name if out_id is not None else None
        in_name = all_members[in_id].name if in_id is not None else None
        slot = f"第{week}周{WEEKDAY_LABELS[weekday]}{BLOCK_LABELS[block].split(' ')[0]}"
        if out_name and in_name:
            change = f"{out_name} → {in_name}"
        elif out_name:
            change = f"移除 {out_name}"
        else:
            change = f"新增 {in_name}"
        self.statusBar().showMessage(f"已手动调整 {slot}：{change}。")

    # ---------- 动作 ----------

    def open_participation_dialog(self) -> None:
        """从菜单批量设置哪些课表成员参与排班。"""
        if self.busy:
            self.statusBar().showMessage("上一个任务尚未完成，请稍候…")
            return
        members = self.members()
        if not members:
            QMessageBox.information(self, "提示", "请先上传成员课表。")
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("设置参与排班")
        dlg.setMinimumWidth(460)
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(10)
        layout.addWidget(dialog_header(
            "设置参与排班",
            "勾选的成员会参与自动排班和手动微调；取消勾选只保留课表，不参与排班。"))

        select_row = QHBoxLayout()
        btn_select_all = QPushButton("全选")
        btn_select_none = QPushButton("全不选")
        select_row.addWidget(btn_select_all)
        select_row.addWidget(btn_select_none)
        select_row.addStretch(1)
        layout.addLayout(select_row)

        member_list = QListWidget()
        member_list.setMinimumHeight(280)
        for m in members:
            item = QListWidgetItem(
                f"{m.name}（学号 {m.student_id or '—'} · "
                f"{m.class_name or '未知班级'}）")
            item.setData(Qt.UserRole, m.id)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(
                Qt.Checked if m.participates_in_scheduling else Qt.Unchecked)
            member_list.addItem(item)
        layout.addWidget(member_list, stretch=1)

        def set_all(checked: bool) -> None:
            state = Qt.Checked if checked else Qt.Unchecked
            for row in range(member_list.count()):
                member_list.item(row).setCheckState(state)

        btn_select_all.clicked.connect(lambda: set_all(True))
        btn_select_none.clicked.connect(lambda: set_all(False))
        buttons = dialog_buttons("保存")
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        layout.addWidget(buttons)

        if dlg.exec() != QDialog.Accepted:
            return
        old_state = {m.id: m.participates_in_scheduling for m in members}
        updates: dict[int, bool] = {}
        for row in range(member_list.count()):
            item = member_list.item(row)
            member_id = item.data(Qt.UserRole)
            participates = item.checkState() == Qt.Checked
            if old_state[member_id] != participates:
                updates[member_id] = participates
        if not updates:
            self.statusBar().showMessage("参与排班设置未变化。")
            return
        self.db.set_members_participation(updates)
        self.invalidate_cache()
        self.refresh_members()
        participating = len(self.scheduling_members())
        self.statusBar().showMessage(
            f"已更新参与排班设置（{len(updates)} 名成员变化）；"
            f"当前 {participating} 名成员参与，重新生成排班后生效。")

    def upload_files(self) -> None:
        if self.busy:
            self.statusBar().showMessage("上一个任务尚未完成，请稍候…")
            return
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择成员课表文件", str(Path.home()),
            "课表文件 (*.xls *.xlsx);;所有文件 (*)")
        if not files:
            return
        known = {(m.student_id, m.name) for m in self.members()}

        def apply_import(payload: tuple[list, list[str], list[str]]) -> None:
            schedules, errors, parse_warnings = payload
            added, updated = 0, 0
            for schedule in schedules:
                is_update = (schedule.student_id, schedule.name) in known
                self.db.upsert_member(schedule)
                known.add((schedule.student_id, schedule.name))
                added += 0 if is_update else 1
                updated += 1 if is_update else 0
            self.invalidate_cache()
            self.refresh_members()
            if errors:
                QMessageBox.warning(self, "部分文件导入失败", "\n".join(errors))
            if parse_warnings:
                QMessageBox.warning(
                    self, "课表存在未定位的集中安排",
                    "以下备注行课程没有具体星期/节次，已按整周避让：\n"
                    + "\n".join(parse_warnings),
                )
            parts = []
            if added:
                parts.append(f"新增 {added} 名成员")
            if updated:
                parts.append(f"更新 {updated} 份课表（同学号同名，课程整体替换）")
            if parts:
                self.statusBar().showMessage(
                    "、".join(parts) + f"；当前共 {len(self.members())} 名成员。")
            else:
                self.statusBar().showMessage("未导入任何文件。")

        self.run_async(
            "正在解析课表",
            lambda: parse_schedule_files(files),
            apply_import,
        )

    def import_roster(self) -> None:
        if self.busy:
            self.statusBar().showMessage("上一个任务尚未完成，请稍候…")
            return
        default_dir = Path.home() / "Downloads"
        preferred = default_dir / "2025届全媒体中心花名册.xlsx"
        start = preferred if preferred.exists() else default_dir
        path, _ = QFileDialog.getOpenFileName(
            self, "选择全媒体中心花名册", str(start),
            "花名册文件 (*.xlsx *.xls);;所有文件 (*)")
        if not path:
            return

        def apply_import(result) -> None:
            count, changed = self.db.replace_roster(
                result.entries, result.source_file)
            self.invalidate_cache()
            self.refresh_members()
            matched = sum(
                1 for member in self.members()
                if (member.studio or UNKNOWN_STUDIO) != UNKNOWN_STUDIO)
            message = (
                f"已导入花名册 {count} 条；自动更新 {changed} 名已上传成员，"
                f"当前 {matched} 人已匹配工作室。")
            self.statusBar().showMessage(message)
            if result.warnings:
                QMessageBox.information(
                    self, "花名册导入完成",
                    message + "\n" + "\n".join(result.warnings))

        self.run_async(
            "正在解析花名册",
            lambda: parse_roster_file(path),
            apply_import,
        )

    def _studio_options(self, *current_values: str) -> list[str]:
        options = set(self.db.roster_studios())
        options.update(TARGET_AVAILABILITY_STUDIOS)
        options.add(UNKNOWN_STUDIO)
        options.update(
            (member.studio or UNKNOWN_STUDIO).strip() for member in self.members())
        options.update(
            (studio or "").strip() for studio in current_values if studio)
        return sorted(option for option in options if option)

    def edit_selected_member_studio(self) -> None:
        item = self.member_list.currentItem()
        if item is None:
            QMessageBox.information(self, "提示", "请先在成员列表中选择成员。")
            return
        member_id = item.data(Qt.UserRole)
        member = self.db.get_member(member_id)
        if member is None:
            return
        current = member.studio or UNKNOWN_STUDIO
        options = self._studio_options(current)
        studio, accepted = QInputDialog.getItem(
            self, "修改工作室",
            f"设置「{member.name}」的工作室归属：\n"
            "可直接输入新工作室名称，保存后不会被后续花名册重导覆盖。",
            options, options.index(current) if current in options else 0, True)
        if not accepted:
            return
        studio = studio.strip()
        if not studio:
            QMessageBox.information(self, "提示", "工作室名称不能为空。")
            return
        self.db.set_member_studio(member_id, studio)
        self.invalidate_cache()
        self.refresh_members()
        self.statusBar().showMessage(
            f"已将 {member.name} 的工作室修改为“{studio}”，并锁定该归属。")

    def batch_edit_members_studio(self) -> None:
        visible_members = [
            self.member_list.item(row).data(Qt.UserRole)
            for row in range(self.member_list.count())
            if not self.member_list.item(row).isHidden()
        ]
        member_by_id = {member.id: member for member in self.members()}
        visible_members = [
            member_by_id[member_id] for member_id in visible_members
            if member_id in member_by_id
        ]
        if not visible_members:
            QMessageBox.information(self, "提示", "当前筛选结果中没有成员。")
            return

        selected_item = self.member_list.currentItem()
        selected_id = selected_item.data(Qt.UserRole) if selected_item else None
        current_member = member_by_id.get(selected_id) if selected_id else None
        current_studio = (
            current_member.studio if current_member is not None
            else visible_members[0].studio
        ) or UNKNOWN_STUDIO

        dlg = QDialog(self)
        dlg.setWindowTitle("批量修改工作室")
        dlg.setMinimumWidth(480)
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(10)
        layout.addWidget(dialog_header(
            "批量修改工作室",
            f"列表显示当前搜索/筛选结果（{len(visible_members)} 人）。"
            "勾选成员后统一设置工作室，并锁定归属以防重导花名册覆盖。"))

        select_row = QHBoxLayout()
        btn_select_all = QPushButton("全选")
        btn_select_none = QPushButton("清空")
        select_row.addWidget(btn_select_all)
        select_row.addWidget(btn_select_none)
        select_row.addStretch()
        layout.addLayout(select_row)

        member_list = QListWidget()
        member_list.setMinimumHeight(260)
        for member in visible_members:
            studio = member.studio or UNKNOWN_STUDIO
            item = QListWidgetItem(
                f"{member.name}（学号 {member.student_id or '—'} · "
                f"当前 {studio}）")
            item.setData(Qt.UserRole, member.id)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            checked = member.id == selected_id if selected_id is not None else False
            item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
            member_list.addItem(item)
        layout.addWidget(member_list, stretch=1)

        form = QFormLayout()
        studio_combo = QComboBox()
        studio_combo.setEditable(True)
        studio_combo.addItems(self._studio_options(current_studio))
        studio_combo.setCurrentText(current_studio)
        form.addRow("目标工作室", studio_combo)
        layout.addLayout(form)

        def set_all(checked: bool) -> None:
            state = Qt.Checked if checked else Qt.Unchecked
            for row in range(member_list.count()):
                member_list.item(row).setCheckState(state)

        btn_select_all.clicked.connect(lambda: set_all(True))
        btn_select_none.clicked.connect(lambda: set_all(False))

        buttons = dialog_buttons("保存")

        def accept_batch() -> None:
            selected_ids = [
                member_list.item(row).data(Qt.UserRole)
                for row in range(member_list.count())
                if member_list.item(row).checkState() == Qt.Checked
            ]
            studio = studio_combo.currentText().strip()
            if not selected_ids:
                QMessageBox.information(dlg, "提示", "请至少勾选一名成员。")
                return
            if not studio:
                QMessageBox.information(dlg, "提示", "工作室名称不能为空。")
                return
            dlg.accept()

        buttons.accepted.connect(accept_batch)
        buttons.rejected.connect(dlg.reject)
        layout.addWidget(buttons)
        if dlg.exec() != QDialog.Accepted:
            return

        selected_ids = [
            member_list.item(row).data(Qt.UserRole)
            for row in range(member_list.count())
            if member_list.item(row).checkState() == Qt.Checked
        ]
        studio = studio_combo.currentText().strip()
        changed = self.db.set_members_studio(selected_ids, studio)
        self.invalidate_cache()
        self.refresh_members()
        self.statusBar().showMessage(
            f"已将 {changed} 名成员的工作室批量修改为“{studio}”，并锁定归属。")

    def edit_member_profile(self) -> None:
        if self.member_combo.currentIndex() < 0:
            QMessageBox.information(self, "提示", "请先上传成员课表。")
            return
        member_id = self.member_combo.currentData()
        member = self.db.get_member(member_id)
        if member is None:
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("修改个人信息")
        dlg.setMinimumWidth(460)
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(18, 16, 18, 14)
        layout.setSpacing(10)
        layout.addWidget(dialog_header(
            "修改个人信息",
            f"正在修改：{member.name}。课程、值班、请假与特殊安排记录不会丢失。"))
        form = QFormLayout()
        form.setSpacing(10)
        fields = {
            "name": ("姓名", member.name),
            "student_id": ("学号", member.student_id),
            "class_name": ("班级", member.class_name),
            "term": ("学期", member.term),
            "major": ("专业", member.major),
            "department": ("院系", member.department),
        }
        edits: dict[str, QLineEdit] = {}
        for key, (label_text, value) in fields.items():
            edit = QLineEdit(value)
            if key == "name":
                edit.setPlaceholderText("必填")
            edits[key] = edit
            form.addRow(label_text, edit)
        studio_combo = QComboBox()
        studio_combo.setEditable(True)
        studio_combo.addItems(self._studio_options(
            member.studio or UNKNOWN_STUDIO))
        studio_combo.setCurrentText(member.studio or UNKNOWN_STUDIO)
        form.addRow("工作室", studio_combo)
        layout.addLayout(form)

        buttons = dialog_buttons("保存")

        def save_profile() -> None:
            try:
                self.db.update_member_profile(
                    member_id,
                    name=edits["name"].text(),
                    student_id=edits["student_id"].text(),
                    class_name=edits["class_name"].text(),
                    term=edits["term"].text(),
                    major=edits["major"].text(),
                    department=edits["department"].text(),
                )
            except ValueError as exc:
                QMessageBox.warning(dlg, "无法保存", str(exc))
                return
            studio = studio_combo.currentText().strip() or UNKNOWN_STUDIO
            if studio != (member.studio or UNKNOWN_STUDIO):
                self.db.set_member_studio(member_id, studio)
            dlg.accept()

        buttons.accepted.connect(save_profile)
        buttons.rejected.connect(dlg.reject)
        layout.addWidget(buttons)
        if dlg.exec() != QDialog.Accepted:
            return

        self.invalidate_cache()
        self.refresh_members()
        index = self.member_combo.findData(member_id)
        if index >= 0:
            self.member_combo.setCurrentIndex(index)
        self.statusBar().showMessage(f"已更新 {edits['name'].text().strip()} 的个人信息。")

    def export_roster(self) -> None:
        members = self.members()
        entries = self.roster_entries()
        if not members and not entries:
            QMessageBox.information(self, "提示", "请先导入花名册或上传成员课表。")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "导出完整花名册", "全媒体中心完整花名册.xlsx",
            "Excel 文件 (*.xlsx)")
        if not path:
            return
        courses = self.courses()
        target = Path(path)
        self.run_async(
            "正在导出完整花名册",
            lambda: target.write_bytes(
                export_roster_excel(entries, members, courses)),
            lambda _result: self.statusBar().showMessage(
                f"已导出完整花名册：{target}"))

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
                f"确定删除成员「{m.name}」？其课程与排班记录将一并移除。",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No) == QMessageBox.Yes:
            self.db.delete_member(member_id)
            self.invalidate_cache()
            self.refresh_members()
            self.statusBar().showMessage(f"已删除成员 {m.name}。")

    def open_member_courses(self, item: QListWidgetItem) -> None:
        """双击成员列表 → 跳转「成员课表」页并定位到该成员"""
        mid = item.data(Qt.UserRole)
        idx = self.member_combo.findData(mid)
        if idx < 0:
            return
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == "成员课表":
                self.tabs.setCurrentIndex(i)
                break
        self.member_combo.setCurrentIndex(idx)

    def _current_config(self) -> ScheduleConfig:
        weekdays = [d for d, cb in self.weekday_checks.items() if cb.isChecked()]
        blocks = [b for b, cb in self.block_checks.items() if cb.isChecked()]
        return ScheduleConfig(
            weeks=range(self.week_from.value(), self.week_to.value() + 1),
            weekdays=sorted(weekdays), blocks=sorted(blocks),
            per_slot=self.per_slot.value(),
            max_per_week=self.max_week.value(),
            max_per_day=self.max_day.value(),
            seed=self.seed.value(),
        )

    def generate(self) -> None:
        all_members = self.members()
        if not all_members:
            QMessageBox.information(self, "提示", "请先上传成员课表。")
            return
        members = self.scheduling_members()
        if not members:
            QMessageBox.information(
                self, "提示",
                "当前没有参与排班的成员，请先通过菜单「排班 → 设置参与排班…」设置。")
            return
        weekdays = [d for d, cb in self.weekday_checks.items() if cb.isChecked()]
        blocks = [b for b, cb in self.block_checks.items() if cb.isChecked()]
        if not blocks:
            QMessageBox.warning(self, "提示", "请至少选择一个值班时段。")
            return
        if self.week_from.value() > self.week_to.value():
            QMessageBox.warning(self, "提示", "起始周不能大于结束周。")
            return
        calendar = self.term_calendar()
        has_class_day = any(
            calendar.is_class_day(week, weekday)
            for week in range(self.week_from.value(), self.week_to.value() + 1)
            for weekday in range(1, 8)
        )
        if not weekdays and not has_class_day:
            QMessageBox.warning(
                self, "提示", "请至少选择一个值班星期；若该范围有补课日，也可自动纳入。")
            return
        config = self._current_config()
        # 按周增量：只重排所选周范围，范围外的历史排班保留并作为均衡基数
        existing = self.db.load_assignments()
        base = [
            a for a in existing
            if a.week not in config.weeks and not self.is_off_day(a.week, a.weekday)
        ]
        courses, leaves, specials = self.courses(), self.leaves(), self.specials()

        def is_off(week: int, weekday: int) -> bool:
            return calendar.logical_to_date(week, weekday) is None

        def is_class(week: int, weekday: int) -> bool:
            return calendar.is_class_day(week, weekday)

        def work() -> dict:
            # 后台线程内自行构建忙时表：避免跨线程读取主线程缓存
            result = generate_schedule(members, courses, config,
                                       leaves=leaves, base_assignments=base,
                                       special_arrangements=specials,
                                       is_off=is_off, is_class=is_class)
            return {
                "result": result,
                "fresh": [a for a in result.assignments if a.week in config.weeks],
                "base_n": len(base),
                "leaves_n": len(leaves),
                "specials_n": len(specials),
            }

        self.run_async("正在排班",
                       work,
                       lambda payload: self._apply_generation(payload, config))

    def _apply_generation(self, payload: dict, config: ScheduleConfig) -> None:
        """排班计算完成后回到主线程：落库 + 刷新界面"""
        result = payload["result"]
        fresh = payload["fresh"]
        # 只同步发生变化的行：未变动的历史安排保持原行，避免整周删除重建
        added, removed = self.db.sync_assignments_for_weeks(list(config.weeks), fresh)
        self.result = result
        self._last_config = config
        self._stale = False
        self._clear_undo_redo()
        self._save_last_config(config)
        self.refresh_schedule_tabs()
        self.refresh_gantt()
        self.tabs.setCurrentIndex(0)
        scope = (f"第{config.weeks.start}–{config.weeks.stop - 1}周" if len(config.weeks) > 1
                 else f"第{config.weeks.start}周")
        kept = f"，范围外历史排班 {payload['base_n']} 人次保留" if payload["base_n"] else ""
        repaired = f"，缺口修复补上 {result.repaired} 个时段" if getattr(result, "repaired", 0) else ""
        advice = capacity_advice(
            summarize_gap_causes(self._gap_diagnoses(result)), config.per_slot)
        self.statusBar().showMessage(
            f"已重新排班 {scope}：本次安排 {len(fresh)} 人次（新增 {added} / 移除 {removed}）"
            f"{kept}，无人可用时段 {len(result.gaps)} 个{repaired}，"
            f"总次数极差 {result.balanced_spread}"
            + (f"，已避让 {payload['leaves_n']} 条请假记录。" if payload["leaves_n"] else "。")
            + (f"已应用 {payload['specials_n']} 条课表特殊安排。"
               if payload["specials_n"] else "")
            + (advice if advice else ""))

    def clear_schedule(self) -> None:
        """清空全部排班或当前所选周范围的排班。"""
        assignments = self.db.load_assignments()
        if not assignments:
            QMessageBox.information(self, "提示", "当前没有可清空的排班记录。")
            return
        weeks = list(self._current_config().weeks)
        if len(weeks) > 1:
            week_label = f"第{weeks[0]}–{weeks[-1]}周"
        else:
            week_label = f"第{weeks[0]}周"
        scope_items = [f"清空所选周范围（{week_label}）", "清空全部排班"]
        mode, ok = QInputDialog.getItem(
            self, "清空排班", "请选择清空范围：", scope_items, 0, False)
        if not ok:
            return
        if mode == scope_items[0]:
            week_set = set(weeks)
            count = sum(a.week in week_set for a in assignments)
            prompt = f"确定清空{week_label}内的 {count} 条值班安排？"
        else:
            count = len(assignments)
            prompt = f"确定清空全部 {count} 条值班安排？此操作不可恢复。"
        if count == 0:
            QMessageBox.information(self, "提示", "所选范围内没有排班记录。")
            return
        if QMessageBox.question(
                self, "确认清空", prompt,
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No) != QMessageBox.Yes:
            return
        if mode == scope_items[0]:
            self.db.delete_assignments_for_weeks(weeks)
        else:
            self.db.clear_assignments()
        self.invalidate_cache()
        self.result = None
        self._last_config = None
        self._stale = False
        self._clear_undo_redo()
        self.refresh_schedule_tabs()
        self.refresh_gantt()
        self.statusBar().showMessage(f"已清空 {count} 条值班安排。")

    def _export_scope(self) -> tuple | None:
        """导出前选择周数：返回 (assignments, stats, gaps, 周次文本) 或 None（取消导出）"""
        weeks = sorted({a.week for a in self.result.assignments})
        if not weeks:
            return None
        label_all = (f"全部周（第{weeks[0]}–{weeks[-1]}周）" if len(weeks) > 1
                     else f"第{weeks[0]}周")
        if len(weeks) > 1:
            items = [label_all] + [f"第{w}周" for w in weeks]
            sel, ok = QInputDialog.getItem(
                self, "选择导出周数", "要导出哪些周的值班表？", items, 0, False)
            if not ok:
                return None
        else:
            sel = label_all
        if sel == label_all:
            weeks_txt = (f"第{weeks[0]}–{weeks[-1]}周" if len(weeks) > 1 else f"第{weeks[0]}周")
            return self.result.assignments, self.result.member_stats, self.result.gaps, weeks_txt
        week = int(sel[1:-1])
        assignments = [a for a in self.result.assignments if a.week == week]
        members = self.members()
        stats = rebuild_member_stats(members, assignments)
        gaps = [g for g in self.result.gaps if g[0] == week]
        return assignments, stats, gaps, f"第{week}周"

    def _export_bytes(self, label: str, filename: str, suffix: str, filt: str,
                      weeks_txt: str, build: Callable[[], bytes]) -> None:
        """导出通用流程：选路径（主线程）-> 生成字节并写盘（后台线程）。

        openpyxl 生成大表、PNG 渲染在成员多时耗时明显，放到线程池避免界面卡顿。
        周次选择已由调用方完成，这里不再弹窗（否则会重复询问）。
        """
        path, _ = QFileDialog.getSaveFileName(self, label, filename, filt)
        if not path:
            return
        target = Path(path)

        def work() -> bytes:
            data = build()
            target.write_bytes(data)
            return data

        self.run_async(f"正在导出{suffix}",
                       work,
                       lambda _data: self.statusBar().showMessage(
                           f"已导出{suffix}（{weeks_txt}）：{target}"))

    def export_xlsx(self) -> None:
        scope = self._export_scope()
        if scope is None:
            return
        assignments, stats, gaps, weeks_txt = scope
        # 界面控件只能在主线程读：先生成函数参数，再交给后台线程
        start = self._term_start()
        calendar = self.term_calendar()
        self._export_bytes(
            "导出 Excel", f"值班排班表_{weeks_txt}.xlsx", "Excel", "Excel 文件 (*.xlsx)",
            weeks_txt,
            lambda: export_excel(
                assignments, stats, gaps, start_date=start, calendar=calendar))

    def export_csv(self) -> None:
        scope = self._export_scope()
        if scope is None:
            return
        assignments, _, _, weeks_txt = scope
        start = self._term_start()
        calendar = self.term_calendar()
        self._export_bytes(
            "导出 CSV", f"值班排班表_{weeks_txt}.csv", "CSV", "CSV 文件 (*.csv)", weeks_txt,
            lambda: export_csv(
                assignments, start_date=start, calendar=calendar))

    def export_png(self) -> None:
        if self.result is None:
            return
        scope = self._export_scope()
        if scope is None:
            return
        assignments, _, _, weeks_txt = scope
        path, _ = QFileDialog.getSaveFileName(
            self, "导出图片", f"值班排班表_{weeks_txt}.png", "PNG 图片 (*.png)")
        if not path:
            return
        start = self._term_start()
        calendar = self.term_calendar()
        frame = self._pivot_frame(assignments)
        if weeks_txt.startswith("第") and "–" not in weeks_txt:
            week = int(weeks_txt[1:-1])
            weekdays = sorted({a.weekday for a in assignments}) or [1]
            dates = [calendar.logical_to_date(week, d) for d in weekdays]
            dates = [d for d in dates if d is not None]
            if dates:
                fd, ld = min(dates), max(dates)
                subtitle = f"{weeks_txt}（{fd.month}月{fd.day}日–{ld.month}月{ld.day}日）"
            else:
                subtitle = f"{weeks_txt}（法定假日，无排班日期）"
        else:
            subtitle = (f"{weeks_txt} · {start.year}年{start.month}月{start.day}日起")
        target = Path(path)
        # QImage/QPainter 在非 GUI 线程绘到图片是允许的，这里直接写入目标文件
        self.run_async(
            "正在生成图片",
            lambda: render_table_png(frame, "值班排班表", subtitle, target),
            lambda paths: self.statusBar().showMessage(
                f"已导出图片（{weeks_txt}，{len(paths)} 张）：{target.parent}"))


def main() -> None:
    app = QApplication(sys.argv)
    app.setOrganizationName("DutySystem")
    app.setApplicationName("DutyScheduler")
    app.setStyle("Fusion")
    app.setStyleSheet(build_style())
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
