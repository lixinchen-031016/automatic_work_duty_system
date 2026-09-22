"""自动值班排班系统 — PySide6 桌面应用

功能：上传成员个人课表(.xls/.xlsx) -> 解析入库(SQLite) -> 按空闲时段生成排班表
     （避免课程/长期特殊安排冲突、每人每天只值一次、均衡分配）-> 界面展示与导出(Excel/CSV)
     空闲甘特图：按周查看各成员忙闲、高亮全员空闲时段，方便安排任务

运行：python app.py
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
from collections import Counter
from collections.abc import Callable
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from PySide6.QtCore import (
    QDate,
    QObject,
    QRect,
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
    QFont,
    QFontMetrics,
    QImage,
    QKeySequence,
    QPainter,
    QPen,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QFrame,
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
    QSpinBox,
    QSplitter,
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
    parse_schedule_file,
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
    return f"第{week_start}周" if week_start == week_end else f"第{week_start}–{week_end}周"


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
    """把 DataFrame 填入只读 QTableWidget（苹果风格：无网格线、交替行、隐藏行号）

    尺寸未变时复用既有单元格、只更新变化的文本：明细/统计等表格可能有数千个
    单元格，反复重建 item 会让刷新明显变慢。
    """
    rows, cols = len(df), len(df.columns)
    if table.rowCount() != rows or table.columnCount() != cols:
        table.clearContents()
        table.setRowCount(rows)
        table.setColumnCount(cols)
    table.setHorizontalHeaderLabels([str(c) for c in df.columns])
    table.setShowGrid(False)
    table.setAlternatingRowColors(True)
    table.verticalHeader().setVisible(False)
    table.verticalHeader().setDefaultSectionSize(30)
    # 一次性把 DataFrame 取成字符串矩阵：pandas 的逐格标量访问（df.iat）
    # 在单元格较多时开销极大，这里改成一次遍历后再写表。
    values = [tuple(str(v) for v in row)
              for row in df.itertuples(index=False, name=None)]
    for r, row_values in enumerate(values):
        for c, text in enumerate(row_values):
            item = table.item(r, c)
            if item is None:
                item = QTableWidgetItem()
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                table.setItem(r, c, item)
            if item.text() != text:
                item.setText(text)
    # 列宽：用「填完后一次性自适应」代替常驻 ResizeToContents。
    # 后者在每次 setText/setItem 时都会重算列宽，表格越大越吃亏；
    # 一次性 resizeColumnsToContents() 只在本次刷新结束时扫描一遍。
    header = table.horizontalHeader()
    if cols and header.sectionResizeMode(0) != QHeaderView.Interactive:
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setStretchLastSection(True)
    if cols:
        table.resizeColumnsToContents()


def render_table_png(df: pd.DataFrame, title: str, subtitle: str, path: Path) -> None:
    """把表格渲染为高清 PNG（2x 缩放、类苹果表格样式），方便发群通知 / 打印张贴"""
    margin, pad, row_h, head_h = 28, 12, 34, 38
    font = QFont()
    font.setPixelSize(13)
    bold = QFont()
    bold.setPixelSize(13)
    bold.setBold(True)
    title_font = QFont()
    title_font.setPixelSize(21)
    title_font.setBold(True)
    sub_font = QFont()
    sub_font.setPixelSize(12)
    fm, hfm, tfm, sfm = (QFontMetrics(f) for f in (font, bold, title_font, sub_font))

    cols = [str(c) for c in df.columns]
    widths = []
    for c in range(len(cols)):
        w = hfm.horizontalAdvance(cols[c])
        if len(df):
            w = max(w, max(fm.horizontalAdvance(str(df.iat[r, c])) for r in range(len(df))))
        widths.append(max(w + pad * 2, 56))

    img_w = margin * 2 + sum(widths)
    img_h = margin + tfm.height() + 8 + sfm.height() + 16 + head_h + row_h * len(df) + margin
    img = QImage(img_w * 2, img_h * 2, QImage.Format_ARGB32)
    img.fill(Qt.white)
    p = QPainter(img)
    p.setRenderHint(QPainter.Antialiasing)
    p.setRenderHint(QPainter.TextAntialiasing)
    p.scale(2, 2)

    p.setPen(QPen(QColor("#1d1d1f")))
    p.setFont(title_font)
    p.drawText(QRect(margin, margin, img_w - margin * 2, tfm.height()),
               Qt.AlignLeft | Qt.AlignVCenter, title)
    p.setPen(QPen(QColor("#86868b")))
    p.setFont(sub_font)
    p.drawText(QRect(margin, margin + tfm.height() + 8, img_w - margin * 2, sfm.height()),
               Qt.AlignLeft | Qt.AlignVCenter, subtitle)

    y0 = margin + tfm.height() + 8 + sfm.height() + 16
    table_h = head_h + row_h * len(df)
    p.fillRect(QRect(margin, y0, img_w - margin * 2, head_h), QColor("#f2f2f7"))
    p.setFont(bold)
    x = margin
    for w, c in zip(widths, cols):
        p.setPen(QPen(QColor("#1d1d1f")))
        p.drawText(QRect(x, y0, w, head_h), Qt.AlignCenter, c)
        x += w

    p.setFont(font)
    for r in range(len(df)):
        y = y0 + head_h + r * row_h
        if r % 2:
            p.fillRect(QRect(margin, y, img_w - margin * 2, row_h), QColor("#f7f7f9"))
        x = margin
        for c, w in enumerate(widths):
            text = str(df.iat[r, c])
            p.setPen(QPen(QColor("#aeaeb2") if text == "—" else "#1d1d1f"))
            p.drawText(QRect(x, y, w, row_h), Qt.AlignCenter, text)
            x += w

    p.setPen(QPen(QColor("#e5e5ea"), 1))
    x = margin
    for w in widths[:-1]:
        x += w
        p.drawLine(x, y0, x, y0 + table_h)
    y = y0 + head_h
    for _ in range(len(df) - 1):
        y += row_h
        p.drawLine(margin, y, img_w - margin, y)
    p.setPen(QPen(QColor("#d2d2d7"), 1))
    p.drawRect(QRect(margin, y0, img_w - margin * 2, table_h))
    p.end()
    img.save(str(path))


def draw_bar_chart(
    painter: QPainter, rect: QRect, title: str,
    items: list[tuple[str, int]], color: str = "#007aff",
) -> None:
    """在 rect 内绘制标题 + 柱状图（类苹果风格），供界面与 PNG 导出共用"""
    title_font = QFont()
    title_font.setPixelSize(13)
    title_font.setBold(True)
    tfm = QFontMetrics(title_font)
    label_font = QFont()
    label_font.setPixelSize(12)
    lfm = QFontMetrics(label_font)

    painter.setFont(title_font)
    painter.setPen(QPen(QColor("#1d1d1f")))
    painter.drawText(QRect(rect.left(), rect.top(), rect.width(), tfm.height()),
                     Qt.AlignLeft | Qt.AlignVCenter, title)
    if not items:
        painter.setFont(label_font)
        painter.setPen(QPen(QColor("#aeaeb2")))
        painter.drawText(rect.adjusted(0, tfm.height() + 8, 0, 0),
                         Qt.AlignCenter, "暂无数据")
        return

    top = rect.top() + tfm.height() + 14
    bottom = rect.bottom() - lfm.height() - 8
    left = rect.left() + 30
    right = rect.right() - 8
    plot_h = bottom - top
    if plot_h <= 20:
        return
    max_v = max(v for _, v in items)
    top_tick = max_v if max_v % 4 == 0 else (max_v // 4 + 1) * 4
    top_tick = max(top_tick, 4)

    # 横向网格线（0/25/50/75/100% 五档）与左侧刻度值
    painter.setFont(label_font)
    grid_pen = QPen(QColor("#e5e5ea"), 1)
    for i in range(5):
        y = bottom - plot_h * i / 4
        painter.setPen(grid_pen)
        painter.drawLine(left, y, right, y)
        painter.setPen(QPen(QColor("#aeaeb2")))
        painter.drawText(QRect(left - 30, y - lfm.height() / 2, 26, lfm.height()),
                         Qt.AlignRight | Qt.AlignVCenter, str(top_tick * i // 4))

    # 柱体 + 柱顶数值 + 底部标签
    slot_w = (right - left) / len(items)
    bar_w = min(slot_w * 0.52, 64)
    bar_color = QColor(color)
    for i, (label, v) in enumerate(items):
        cx = left + slot_w * (i + 0.5)
        h = plot_h * v / top_tick if top_tick else 0
        bar_rect = QRect(cx - bar_w / 2, bottom - h, bar_w, h)
        painter.setPen(Qt.NoPen)
        painter.setBrush(bar_color)
        painter.drawRoundedRect(bar_rect, 3, 3)
        painter.setPen(QPen(QColor("#1d1d1f")))
        painter.drawText(QRect(cx - slot_w / 2, bottom - h - lfm.height() - 2,
                               slot_w, lfm.height()),
                         Qt.AlignCenter, str(v))
        painter.setPen(QPen(QColor("#86868b")))
        painter.drawText(QRect(cx - slot_w / 2, bottom + 2, slot_w, lfm.height()),
                         Qt.AlignCenter, label)


def render_charts_png(charts: list[tuple[str, list[tuple[str, int]], str]], path: Path) -> None:
    """多张柱状图渲染为一张高清 PNG（2x 缩放，纵向排列）"""
    margin, chart_h, img_w = 28, 280, 980
    img_h = margin * 2 + chart_h * len(charts) + 16 * (len(charts) - 1)
    img = QImage(img_w * 2, img_h * 2, QImage.Format_ARGB32)
    img.fill(Qt.white)
    p = QPainter(img)
    p.setRenderHint(QPainter.Antialiasing)
    p.setRenderHint(QPainter.TextAntialiasing)
    p.scale(2, 2)
    for i, (title, items, color) in enumerate(charts):
        draw_bar_chart(p, QRect(margin, margin + i * (chart_h + 16),
                                img_w - margin * 2, chart_h), title, items, color)
    p.end()
    img.save(str(path))


class BarChart(QFrame):
    """类苹果风格柱状图（值班 / 请假情况可视化）"""

    def __init__(self, title: str, color: str = "#007aff", parent: QWidget | None = None):
        super().__init__(parent)
        self._title = title
        self._color = color
        self._items: list[tuple[str, int]] = []
        self.setMinimumHeight(220)
        self.setStyleSheet(
            "BarChart { background: #ffffff; border: 1px solid #e8e8ed; border-radius: 10px; }")

    def set_data(self, items: list[tuple[str, int]]) -> None:
        self._items = list(items)
        self.update()

    def data(self) -> list[tuple[str, int]]:
        return self._items

    def paintEvent(self, event) -> None:
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setRenderHint(QPainter.TextAntialiasing)
        draw_bar_chart(p, self.rect().adjusted(16, 12, -16, -10),
                       self._title, self._items, self._color)
        p.end()


class _TaskSignals(QObject):
    """后台任务的完成/失败信号（QRunnable 本身不能带信号，需 QObject 载体）"""

    done = Signal(object)
    failed = Signal(str)


class _Task(QRunnable):
    """把耗时操作（排班计算、Excel/PNG 生成）放到线程池执行，避免界面假死。

    只负责执行纯计算与取数据；任何界面刷新与数据库写入都在主线程回调里做。
    """

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
        self._calendar_cache = None
        self._term_calendar_cache: TermCalendar | None = None
        self._calendar_error: str | None = None
        # 忙时表由「成员 + 课程」唯一决定，展开成本高（课程数 x 周次 x 节次），
        # 甘特图与微调对话框都会用到，因此与课表缓存同生命周期
        self._busy_cache = None
        self._special_busy_cache = None
        # 缺口诊断结果（按排班结果对象缓存，避免每次刷新重复诊断）
        self._gap_cache: tuple | None = None
        self._undo_stack: list[tuple[int, int, int, tuple[int, ...]]] = []
        self._redo_stack: list[tuple[int, int, int, tuple[int, ...]]] = []
        self._undo_limit = 50
        # 甘特图单元格上次写入的状态：(行, 列) -> 状态元组，用于跳过无变化单元格
        self._gantt_cell_state: dict = {}
        self._pool = QThreadPool.globalInstance()
        self._busy = False
        # 必须持有正在执行的任务引用：QRunnable 被 Python 回收后线程池就无法再执行它
        self._active_task = None
        self._closing = False
        self.setWindowTitle("自动值班排班系统")
        self.resize(1280, 800)
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
                # 覆盖，避免甘特图和排班刷新崩溃；保存/校验时仍会明确报错。
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
        QApplication.restoreOverrideCursor()
        if error is not None:
            self.statusBar().showMessage(f"操作失败：{error}")
            QMessageBox.warning(self, "操作失败", error)
            return
        if on_done is not None:
            on_done(result)

    # ---------- 界面构建 ----------

    def _build_ui(self) -> None:
        edit_menu = self.menuBar().addMenu("编辑")
        self.undo_action = QAction("撤销", self)
        self.undo_action.setShortcut(QKeySequence.StandardKey.Undo)
        self.undo_action.setShortcutContext(Qt.ApplicationShortcut)
        self.undo_action.setToolTip(
            f"撤销 {SHORTCUT_MODIFIER}Z" if SHORTCUT_MODIFIER == "⌘"
            else f"撤销 {SHORTCUT_MODIFIER}Z")
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
        self.member_search = QLineEdit()
        self.member_search.setPlaceholderText("搜索姓名 / 学号 / 班级")
        self.member_search.setClearButtonEnabled(True)
        members_layout.addWidget(self.member_search)
        self.member_class_filter = QComboBox()
        self.member_class_filter.setToolTip("按班级筛选成员")
        members_layout.addWidget(self.member_class_filter)
        self.member_list = QListWidget()
        self.member_list.setSelectionMode(QAbstractItemView.SingleSelection)
        self.member_list.setMinimumHeight(120)
        self.member_list.setToolTip("双击成员可在「成员课表」页查看其课表详情")
        self.member_list.itemDoubleClicked.connect(self.open_member_courses)
        self.member_search.textChanged.connect(self._apply_member_filter)
        self.member_class_filter.currentIndexChanged.connect(self._apply_member_filter)
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
        wd_row = QHBoxLayout()
        for d in (1, 2, 3, 4, 5, 6, 7):
            cb = QCheckBox(WEEKDAY_LABELS[d])
            cb.setChecked(d <= 5)
            cb.stateChanged.connect(self._on_gantt_filter_changed)
            cb.stateChanged.connect(self._mark_stale)
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
            cb.stateChanged.connect(self._mark_stale)
            self.block_checks[b] = cb
            blk_row.addWidget(cb)
        form.addRow("值班时段", blk_row)

        num_row1 = QHBoxLayout()
        self.per_slot = QSpinBox()
        self.per_slot.setRange(1, 5)
        self.per_slot.valueChanged.connect(self._mark_stale)
        self.max_week = QSpinBox()
        self.max_week.setRange(1, 10)
        self.max_week.setValue(3)
        self.max_week.setToolTip("每人每周最多值班次数")
        self.max_week.valueChanged.connect(self._mark_stale)
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
        self.max_day.valueChanged.connect(self._mark_stale)
        self.seed = QSpinBox()
        self.seed.setRange(0, 9999)
        self.seed.setValue(42)
        self.seed.setToolTip("平手时决定分给谁，固定种子可复现")
        self.seed.valueChanged.connect(self._mark_stale)
        num_row2.addWidget(QLabel("每天上限"))
        num_row2.addWidget(self.max_day)
        num_row2.addWidget(QLabel("随机种子"))
        num_row2.addWidget(self.seed)
        form.addRow("", num_row2)

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
        btn_calendar = QPushButton("管理日历…")
        btn_calendar.setToolTip("批量设置法定假日与周末补课映射")
        btn_calendar.clicked.connect(self.open_calendar_dialog)
        calendar_row.addWidget(self.calendar_label, stretch=1)
        calendar_row.addWidget(btn_calendar)
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
        self.btn_export_png = QPushButton("导出图片")
        self.btn_export_png.setToolTip("把值班表导出为 PNG 图片，方便发群通知 / 打印张贴")
        self.btn_export_png.clicked.connect(self.export_png)
        self.btn_export_xlsx.setEnabled(False)
        self.btn_export_csv.setEnabled(False)
        self.btn_export_png.setEnabled(False)
        btn_row.addWidget(self.btn_export_xlsx)
        btn_row.addWidget(self.btn_export_csv)
        btn_row.addWidget(self.btn_export_png)
        btn_row.addStretch()
        v1.addLayout(btn_row)
        self.pivot_table = QTableWidget()
        self.pivot_table.setToolTip("双击值班单元格可手动调整该时段值班人")
        self.pivot_table.cellDoubleClicked.connect(self._on_pivot_cell_double_clicked)
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
        self.member_combo.currentIndexChanged.connect(self._on_member_combo_changed)
        sel_row.addWidget(self.member_combo)
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
            "自动排班、甘特图和手动微调都会避让。")
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
        grp_leave.setToolTip("登记后重新生成排班将避开该天；甘特图中该天标红")
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
            '<span style="background:#b8d9ff;">&nbsp;&nbsp;&nbsp;&nbsp;</span> 已排值班&nbsp;&nbsp;'
            '<span style="background:#ffd9a8;">&nbsp;&nbsp;&nbsp;&nbsp;</span> 全员空闲&nbsp;&nbsp;'
            '<span style="background:#ffd6d2;">&nbsp;&nbsp;&nbsp;&nbsp;</span> 请假&nbsp;&nbsp;'
            '<span style="background:#e5d8ff;">&nbsp;&nbsp;&nbsp;&nbsp;</span> 其他安排')
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
        # 表头列宽/行高只设置一次：以前每次刷新甘特图都重设
        # ResizeToContents，会触发整表列宽重算，成员多时明显拖慢刷新
        gantt_header = self.gantt_table.horizontalHeader()
        gantt_header.setSectionResizeMode(QHeaderView.ResizeToContents)
        gantt_header.setStretchLastSection(False)
        gantt_header.setMinimumHeight(44)
        gantt_vertical_header = self.gantt_table.verticalHeader()
        gantt_vertical_header.setDefaultSectionSize(44)
        gantt_vertical_header.setMinimumWidth(112)
        v5.addWidget(_card(self.gantt_table), stretch=1)
        self.tabs.addTab(tab5, "空闲甘特图")

        # Tab6 统计图表：值班 / 请假情况可视化
        tab6 = QWidget()
        v6 = QVBoxLayout(tab6)
        v6.setContentsMargins(16, 14, 16, 16)
        v6.setSpacing(12)
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
        self.chart_duty = BarChart("各成员值班总次数", "#007aff")
        self.chart_weekly = BarChart("每周值班人次", "#34c759")
        self.chart_leave = BarChart("各成员请假天数", "#ff453a")
        for c in (self.chart_duty, self.chart_weekly, self.chart_leave):
            charts_row.addWidget(c, stretch=1)
        v6.addLayout(charts_row, stretch=1)
        self.tabs.addTab(tab6, "统计图表")

        self.tabs.currentChanged.connect(self._on_tab_changed)
        return self.tabs

    # ---------- 数据刷新 ----------

    def refresh_members(self) -> None:
        members = self.members()
        selected_class = self.member_class_filter.currentData()
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
        self.member_combo.blockSignals(True)
        self.member_combo.clear()
        for m in members:
            self.member_list.addItem(f"{m.name}（{m.course_count} 门课）")
            item = self.member_list.item(self.member_list.count() - 1)
            item.setData(Qt.UserRole, m.id)
            item.setData(Qt.UserRole + 1, (
                m.name, m.student_id, m.class_name))
            self.member_combo.addItem(f"{m.name}（{m.class_name or '未知班级'}）", m.id)
        self.member_combo.blockSignals(False)
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
        for row in range(self.member_list.count()):
            item = self.member_list.item(row)
            name, student_id, class_name = item.data(Qt.UserRole + 1)
            matches_query = (
                not query
                or query in str(name).casefold()
                or query in str(student_id).casefold()
                or query in str(class_name).casefold()
            )
            matches_class = not selected_class or class_name == selected_class
            item.setHidden(not (matches_query and matches_class))

    def _update_empty_state(self) -> None:
        """无排班结果时，给出下一步引导文案"""
        if self.result is not None:
            return
        if self.member_list.count() == 0:
            self.summary_label.setText(
                "开始使用三步：① 上传成员课表（支持多选 .xls / .xlsx）→ ② 调整排班参数 → "
                "③ 点击「生成排班表」。可在「空闲甘特图」页查看全员共同空闲时段，方便安排任务。")
        else:
            self.summary_label.setText(
                f"已就绪 {self.member_list.count()} 名成员，点击左下角「生成排班表」开始排班。")

    def _mark_stale(self) -> None:
        """参数或成员变化后标记结果过期，提示重新生成"""
        if self.result is None or self._stale:
            return
        self._stale = True
        self.summary_label.setText(
            "⚠ 排班参数或成员已变化，下方结果可能过期——请点击「生成排班表」重新生成。")

    def _restore_result(self) -> None:
        """启动时从数据库恢复上次的排班结果（关闭程序不会丢失）"""
        assignments = self.db.load_assignments()
        members = self.members()
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
        """更改数据库存储位置：迁移现有数据或切换到已有数据库文件"""
        current = Path(self.db.path)
        folder = QFileDialog.getExistingDirectory(self, "选择数据库存储位置", str(current.parent))
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
            # 用 SQLite 在线备份而非文件复制：WAL 模式下最新数据与 schema
            # 可能还在 -wal 文件里，直接复制主库会得到损坏/缺失内容的副本
            self.db.backup_to(new_path)
            migrated = True
        QSettings().setValue("database/path", str(new_path))
        self.db = Database(new_path)
        self.invalidate_cache()
        self._update_db_path_label()
        self._refresh_calendar_summary()
        self.result = None
        self._last_config = None
        self._stale = False
        self.refresh_members()
        self._restore_result()
        if self.result is None:
            self.summary_label.setText("尚未生成排班表。设置左侧参数后点击「生成排班表」。")
            for tb in (self.pivot_table, self.detail_table, self.stats_table, self.gap_table):
                fill_table(tb, pd.DataFrame())
            self.btn_export_xlsx.setEnabled(False)
            self.btn_export_csv.setEnabled(False)
            self.btn_export_png.setEnabled(False)
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
        form = QFormLayout(dlg)
        form.setContentsMargins(16, 16, 16, 12)
        form.setSpacing(10)

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

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("添加")
        buttons.button(QDialogButtonBox.Cancel).setText("取消")
        form.addRow(buttons)

        def confirm() -> None:
            week = week_spin.value()
            weekday = int(weekday_combo.currentData())
            off_date = natural_date(self._term_start(), week, weekday)
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

        tip = QLabel(
            "选择年份后加载国家法定节假日和调休上班日。补课日对应的逻辑周/星期"
            "为自动推断结果，请逐行确认后再导入。")
        tip.setObjectName("secondary")
        tip.setWordWrap(True)
        layout.addWidget(tip)

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

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("导入到日历")
        buttons.button(QDialogButtonBox.Cancel).setText("取消")
        layout.addWidget(buttons)

        def populate(data) -> None:
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
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("保存")
        buttons.button(QDialogButtonBox.Cancel).setText("取消")
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
        self.db.clear_calendar()
        self.db.upsert_calendar(entries)
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

    def refresh_schedule_tabs(self) -> None:
        result = self.result
        if result is None:
            for table in (self.pivot_table, self.detail_table,
                          self.stats_table, self.gap_table):
                fill_table(table, pd.DataFrame())
            self.summary_label.setText("尚未生成排班表。设置左侧参数后点击「生成排班表」。")
            self.gap_title.setText("无人可值时段")
            self.btn_export_xlsx.setEnabled(False)
            self.btn_export_csv.setEnabled(False)
            self.btn_export_png.setEnabled(False)
            self.btn_clear_schedule.setEnabled(bool(self.db.load_assignments()))
            self.refresh_charts()
            return
        display = self._pivot_display()
        fill_table(self.pivot_table, display)
        fill_table(self.detail_table, build_detail_df(
            result.assignments, start_date=self._term_start(),
            calendar=self.term_calendar()))
        fill_table(self.stats_table, build_stats_df(result.member_stats))

        n = len(result.assignments)
        summary = (
            f"排班完成：共 {n} 人次安排 | 参与成员 {len(result.member_stats)} 人 | "
            f"人均 {n / max(len(result.member_stats), 1):.1f} 次 | "
            f"总次数极差 {result.balanced_spread}（越小越均衡）| "
            f"无人可用时段 {len(result.gaps)} 个")
        if result.gaps:
            # 区分「排不了」（课程/特殊安排/请假冲突）与「排不下」（上限不足）：
            # 前者只能增加成员，后者提高上限即可，给出的建议必须能落地
            summary += "。" + capacity_advice(
                summarize_gap_causes(self._gap_diagnoses(result)), self._config_per_slot(result))
        self.summary_label.setText(summary)
        has_data = bool(result.assignments)
        self.btn_export_xlsx.setEnabled(has_data)
        self.btn_export_csv.setEnabled(has_data)
        self.btn_export_png.setEnabled(has_data)
        self.btn_clear_schedule.setEnabled(has_data)

        if result.gaps:
            self.gap_title.setText(
                f"无人可值时段（{len(result.gaps)} 个，按成因分类，可对照处理）")
            fill_table(self.gap_table, build_gap_df(self._gap_diagnoses(result)))
        else:
            self.gap_title.setText("所有值班时段均已安排到位，无缺口。")
            fill_table(self.gap_table, pd.DataFrame(
                columns=["周次", "星期", "时段", "主要原因", "无课人数"]))
        self.refresh_charts()

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
            self.members(), self.busy_map(), self.leaves(), result.assignments,
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
        self.member_info.setText(
            f"学号 {m.student_id or '—'} | {m.term or '—'} | {m.major or '—'} | "
            f"{m.department or '—'} | 来源：{m.file_name or '—'}")
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

    # ---------- 空闲甘特图 ----------

    def _on_tab_changed(self, index: int) -> None:
        tab = self.tabs.tabText(index)
        if tab == "空闲甘特图":
            self.refresh_gantt()
        elif tab == "统计图表":
            self.refresh_charts()

    def _on_gantt_filter_changed(self) -> None:
        if self.tabs.tabText(self.tabs.currentIndex()) == "空闲甘特图":
            self.refresh_gantt()

    def _sync_gantt_week(self, week: int) -> None:
        """值班起始周变化时甘特图跟随，切到甘特页即在排班起始周"""
        self.gantt_week.setValue(week)

    def refresh_gantt(self) -> None:
        members = self.members()
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
            self.gantt_hint.setText("请先上传成员课表，并至少勾选一个值班星期和值班时段（甘特图跟随左侧筛选）。")
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
        甘特图刷新里最大的一块开销。
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
            self, "导出甘特图 Excel", f"空闲甘特图_第{m.week}周.xlsx", "Excel 文件 (*.xlsx)")
        if not path:
            return
        target = Path(path)
        self.run_async(
            "正在导出甘特图",
            lambda: target.write_bytes(export_gantt_excel(m)),
            lambda _r: self.statusBar().showMessage(f"已导出甘特图 Excel：{target}"))

    # ---------- 统计图表 ----------

    def refresh_charts(self) -> None:
        """刷新三张统计图：成员值班次数 / 每周值班人次 / 成员请假天数"""
        members = self.members()
        if self.result is not None:
            stats = rebuild_member_stats(members, self.result.assignments)
            self.chart_duty.set_data([(s["name"], s["total"]) for s in stats.values()])
            week_cnt = Counter(a.week for a in self.result.assignments)
            self.chart_weekly.set_data(
                [(str(w), week_cnt[w])
                 for w in sorted({a.week for a in self.result.assignments})])
        else:
            self.chart_duty.set_data([])
            self.chart_weekly.set_data([])
        leave_cnt = Counter(l.member_id for l in self.leaves())
        name_of = {m.id: m.name for m in members}
        self.chart_leave.set_data(sorted(
            ((name_of.get(mid, "已删除成员"), n) for mid, n in leave_cnt.items()),
            key=lambda t: (-t[1], t[0])))
        self.btn_export_charts.setEnabled(
            bool(members) and any(c.data() for c in
                                  (self.chart_duty, self.chart_weekly, self.chart_leave)))

    def export_charts(self) -> None:
        charts = [
            ("各成员值班总次数", self.chart_duty.data(), "#007aff"),
            ("每周值班人次", self.chart_weekly.data(), "#34c759"),
            ("各成员请假天数", self.chart_leave.data(), "#ff453a"),
        ]
        path, _ = QFileDialog.getSaveFileName(
            self, "导出统计图", "值班请假统计图.png", "PNG 图片 (*.png)")
        if not path:
            return
        target = Path(path)
        self.run_async(
            "正在导出统计图",
            lambda: render_charts_png(charts, target),
            lambda _r: self.statusBar().showMessage(f"已导出统计图：{target}"))

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
        form = QFormLayout(dlg)
        form.setContentsMargins(16, 16, 16, 12)

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

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("保存")
        buttons.button(QDialogButtonBox.Cancel).setText("取消")
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        form.addRow(buttons)

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
        form = QFormLayout(dlg)
        form.setContentsMargins(16, 16, 16, 12)
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
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("添加")
        buttons.button(QDialogButtonBox.Cancel).setText("取消")
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        form.addRow(buttons)
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
        members = self.members()
        busy = self.busy_map()
        leave_set = {(l.member_id, l.week, l.weekday) for l in self.leaves()}
        cands = replacement_candidates(
            members, busy, leave_set, self.result.assignments,
            week, weekday, block, cfg.max_per_week, cfg.max_per_day,
            special_set=self.special_set())

        dlg = QDialog(self)
        dlg.setWindowTitle("手动微调")
        dlg.setMinimumWidth(440)
        form = QFormLayout(dlg)
        form.setContentsMargins(16, 16, 16, 12)
        info = QLabel(
            f"第{week}周 {WEEKDAY_LABELS[weekday]} · {BLOCK_LABELS[block]}　"
            f"当前值班：{'、'.join(a.member_name for a in slot_assignments) or '（空缺）'}")
        info.setObjectName("secondary")
        info.setWordWrap(True)
        form.addRow(info)

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

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("应用")
        buttons.button(QDialogButtonBox.Cancel).setText("取消")
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        form.addRow(buttons)
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
        self.refresh_schedule_tabs()
        self.refresh_gantt()

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
        members = {m.id: m for m in self.members()}
        cfg = self._last_config
        if in_id is not None:
            busy = self.busy_map()
            leave_set = {(l.member_id, l.week, l.weekday) for l in self.leaves()}
            # 硬约束（课程/特殊安排/请假/每天一次）必须满足；周上限允许手动越限
            eligible = {m.id for m, reason in replacement_candidates(
                list(members.values()), busy, leave_set, self.result.assignments,
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
                member_id=in_id, member_name=members[in_id].name))
        new_ids = tuple(sorted(
            a.member_id for a in kept
            if a.week == week and a.weekday == weekday and a.block == block
        ))
        if new_ids != old_ids:
            self._push_undo((week, weekday, block, old_ids))
            self._redo_stack.clear()
            self._update_undo_redo_actions()
        self.result.assignments = kept
        self.result.member_stats = rebuild_member_stats(list(members.values()), kept)
        self.result.gaps = compute_gaps(
            kept, cfg, is_off=self.is_off_day, is_class=self.is_class_day)
        self._gap_cache = None
        # 只重写被调整的那一个时段，避免全表清空 + 全量重写
        self.db.replace_slot(week, weekday, block,
                             [a.member_id for a in kept
                              if (a.week, a.weekday, a.block) == (week, weekday, block)])
        self._stale = False
        self.refresh_schedule_tabs()
        self.refresh_gantt()
        out_name = members[out_id].name if out_id is not None else None
        in_name = members[in_id].name if in_id is not None else None
        slot = f"第{week}周{WEEKDAY_LABELS[weekday]}{BLOCK_LABELS[block].split(' ')[0]}"
        if out_name and in_name:
            change = f"{out_name} → {in_name}"
        elif out_name:
            change = f"移除 {out_name}"
        else:
            change = f"新增 {in_name}"
        self.statusBar().showMessage(f"已手动调整 {slot}：{change}。")

    # ---------- 动作 ----------

    def upload_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择成员课表文件", str(Path.home()),
            "课表文件 (*.xls *.xlsx);;所有文件 (*)")
        if not files:
            return
        known = {(m.student_id, m.name) for m in self.members()}
        added, updated, errors, parse_warnings = 0, 0, [], []
        for f in files:
            try:
                schedule = parse_schedule_file(Path(f).read_bytes(), Path(f).name)
            except Exception as e:
                errors.append(f"{Path(f).name}：{e}")
                continue
            if not schedule.courses and not schedule.name:
                errors.append(f"{Path(f).name}：未解析到课程信息，请确认是教务系统导出的个人课表")
                continue
            parse_warnings.extend(
                f"{Path(f).name}：{warning}" for warning in schedule.warnings)
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
            self.statusBar().showMessage("、".join(parts) + f"；当前共 {len(self.members())} 名成员。")
        else:
            self.statusBar().showMessage("未导入任何文件。")

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
        members = self.members()
        if not members:
            QMessageBox.information(self, "提示", "请先上传成员课表。")
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
            lambda _r: self.statusBar().showMessage(f"已导出图片（{weeks_txt}）：{target}"))


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
