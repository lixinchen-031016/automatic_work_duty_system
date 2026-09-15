"""自动值班排班系统 — PySide6 桌面应用

功能：上传成员个人课表(.xls/.xlsx) -> 解析入库(SQLite) -> 按空闲时段生成排班表
     （避免课程冲突、每人每天只值一次、均衡分配）-> 界面展示与导出(Excel/CSV)
     空闲甘特图：按周查看各成员忙闲、高亮全员空闲时段，方便安排任务

运行：python app.py
"""

from __future__ import annotations

import shutil
import sys
import tempfile
from collections import Counter
from datetime import date
from pathlib import Path

import pandas as pd
from PySide6.QtCore import QDate, QRect, QSettings, Qt
from PySide6.QtGui import QBrush, QColor, QFont, QFontMetrics, QImage, QPainter, QPen
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QComboBox, QDateEdit, QDialog,
    QDialogButtonBox, QFileDialog, QFormLayout, QFrame, QGroupBox, QHBoxLayout,
    QHeaderView, QInputDialog, QLabel, QLineEdit, QListWidget, QListWidgetItem,
    QMainWindow, QMessageBox, QPushButton, QSpinBox, QSplitter, QTabWidget,
    QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

from duty_system.database import Assignment, Database
from duty_system.exporter import (
    build_detail_df, build_pivot_df, build_stats_df, export_csv, export_excel,
    export_leaves_excel, week_date,
)
from duty_system.gantt import AvailabilityMatrix, build_availability, export_gantt_excel, slot_header
from duty_system.parser import BLOCK_LABELS, WEEKDAY_LABELS, parse_schedule_file
from duty_system.scheduler import (
    ScheduleConfig, ScheduleResult, build_busy_map, generate_schedule,
    rebuild_member_stats, replacement_candidates,
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
        DB_PATH = Path(sys.executable).resolve().parent / "duty_system.db"
else:
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


class MainWindow(QMainWindow):
    def __init__(self, db_path: str | Path | None = None):
        super().__init__()
        if db_path is None:
            stored = QSettings().value("database/path", "", type=str)
            db_path = Path(stored) if stored and Path(stored).exists() else DB_PATH
        self.db = Database(db_path)
        self.result: ScheduleResult | None = None
        self.gantt_matrix: AvailabilityMatrix | None = None
        self._stale = False
        self._last_config: ScheduleConfig | None = None
        self._pivot_rows: list[tuple[int, int]] = []
        self._pivot_blocks: list[int] = []
        self._pivot_date_offset: int = 1
        self.setWindowTitle("自动值班排班系统")
        self.resize(1280, 800)
        self._build_ui()
        self._load_settings()
        self.refresh_members()
        self._restore_result()
        self.statusBar().showMessage(
            "已恢复上次的排班结果。" if self.result is not None else "就绪。请上传成员课表。")

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
        self.member_list.setToolTip("双击成员可在「成员课表」页查看其课表详情")
        self.member_list.itemDoubleClicked.connect(self.open_member_courses)
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
        layout.addWidget(grp_cfg)

        self.btn_generate = QPushButton("生成排班表")
        self.btn_generate.setObjectName("primary")
        self.btn_generate.setMinimumHeight(42)
        self.btn_generate.clicked.connect(self.generate)
        layout.addWidget(self.btn_generate)

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
        self.member_combo.currentIndexChanged.connect(self.show_member_courses)
        sel_row.addWidget(self.member_combo)
        sel_row.addStretch()
        v3.addLayout(sel_row)
        self.member_info = QLabel("")
        self.member_info.setObjectName("secondary")
        v3.addWidget(self.member_info)
        self.course_table = QTableWidget()
        v3.addWidget(_card(self.course_table), stretch=1)

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
            '<span style="background:#ffd6d2;">&nbsp;&nbsp;&nbsp;&nbsp;</span> 请假')
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
        self._update_empty_state()
        self._mark_stale()

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
        members = self.db.list_members()
        if not assignments or not members:
            return
        config = self._current_config()
        result = ScheduleResult()
        result.assignments = assignments
        result.member_stats = rebuild_member_stats(members, assignments)
        grid = {(w, d, b) for w in config.weeks for d in config.weekdays for b in config.blocks}
        covered = {(a.week, a.weekday, a.block) for a in assignments}
        result.gaps = sorted(g for g in grid if g not in covered)
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
            shutil.copy2(current, new_path)
            migrated = True
        QSettings().setValue("database/path", str(new_path))
        self.db = Database(new_path)
        self._update_db_path_label()
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

    def _load_settings(self) -> None:
        s = QSettings()
        if not s.contains("week_from"):
            return
        self.week_from.setValue(int(s.value("week_from", 1)))
        self.week_to.setValue(int(s.value("week_to", 18)))
        self.per_slot.setValue(int(s.value("per_slot", 1)))
        self.max_week.setValue(int(s.value("max_week", 3)))
        self.max_day.setValue(int(s.value("max_day", 1)))
        self.seed.setValue(int(s.value("seed", 42)))
        self.gantt_week.setValue(int(s.value("gantt_week", 1)))
        ts = QDate.fromString(str(s.value("term_start", "")), "yyyy-MM-dd")
        if ts.isValid():
            self.term_start.setDate(ts)
        for d, cb in self.weekday_checks.items():
            cb.setChecked(int(s.value(f"weekday_{d}", 1 if d <= 5 else 0)) == 1)
        for b, cb in self.block_checks.items():
            cb.setChecked(int(s.value(f"block_{b}", 1)) == 1)

    def _save_settings(self) -> None:
        s = QSettings()
        s.setValue("week_from", self.week_from.value())
        s.setValue("week_to", self.week_to.value())
        s.setValue("per_slot", self.per_slot.value())
        s.setValue("max_week", self.max_week.value())
        s.setValue("max_day", self.max_day.value())
        s.setValue("seed", self.seed.value())
        s.setValue("gantt_week", self.gantt_week.value())
        s.setValue("term_start", self.term_start.date().toString("yyyy-MM-dd"))
        for d, cb in self.weekday_checks.items():
            s.setValue(f"weekday_{d}", 1 if cb.isChecked() else 0)
        for b, cb in self.block_checks.items():
            s.setValue(f"block_{b}", 1 if cb.isChecked() else 0)

    def _on_term_start_changed(self) -> None:
        """学期起始日只影响日期显示，改完立即刷新结果表"""
        if self.result is not None:
            self.refresh_schedule_tabs()

    def closeEvent(self, event) -> None:
        self._save_settings()
        super().closeEvent(event)

    def refresh_schedule_tabs(self) -> None:
        result = self.result
        if result is None:
            return
        display = self._pivot_display()
        fill_table(self.pivot_table, display)
        fill_table(self.detail_table, build_detail_df(result.assignments, start_date=self._term_start()))
        fill_table(self.stats_table, build_stats_df(result.member_stats))

        n = len(result.assignments)
        self.summary_label.setText(
            f"排班完成：共 {n} 人次安排 | 参与成员 {len(result.member_stats)} 人 | "
            f"人均 {n / max(len(result.member_stats), 1):.1f} 次 | "
            f"总次数极差 {result.balanced_spread}（越小越均衡）| "
            f"无人可用时段 {len(result.gaps)} 个（成员有课、请假、已达每周/每天上限或当天已值过）")
        has_data = bool(result.assignments)
        self.btn_export_xlsx.setEnabled(has_data)
        self.btn_export_csv.setEnabled(has_data)
        self.btn_export_png.setEnabled(has_data)

        if result.gaps:
            self.gap_title.setText(f"无人可值时段（{len(result.gaps)} 个）")
            gap_df = pd.DataFrame([{
                "周次": f"第{w}周", "星期": WEEKDAY_LABELS[d], "时段": BLOCK_LABELS[b],
            } for w, d, b in result.gaps])
            fill_table(self.gap_table, gap_df)
        else:
            self.gap_title.setText("所有值班时段均已安排到位，无缺口。")
            fill_table(self.gap_table, pd.DataFrame(columns=["周次", "星期", "时段"]))
        self.refresh_charts()

    def _term_start(self) -> date:
        return self.term_start.date().toPython()

    def _pivot_frame(self, assignments: list[Assignment]) -> pd.DataFrame:
        """扁平化透视表（周次/星期 + 日期 + 各时段），供表格展示与 PNG 导出"""
        pivot = build_pivot_df(assignments, start_date=self._term_start())
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
        self._pivot_rows = [(w, d) for w in weeks for d in weekdays]
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
        courses = sorted(self.db.get_courses(member_id),
                         key=lambda c: (c.weekday, min(c.session_list), c.course_name))
        df = pd.DataFrame([{
            "星期": WEEKDAY_LABELS[c.weekday], "课程": c.course_name,
            "教师": c.teacher or "—", "周次": c.weeks_text,
            "节次": f"{c.sessions_text}节", "地点": c.location or "—",
        } for c in courses])
        fill_table(self.course_table, df if not df.empty else
                   pd.DataFrame(columns=["星期", "课程", "教师", "周次", "节次", "地点"]))
        self._refresh_leaves(member_id)

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
            self.gantt_week.value(), weekdays, blocks,
            leaves=self.db.list_leaves(),
            assignments=self.result.assignments if self.result else None)
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
        """行=成员（末行为汇总），列=星期x时段；
        空闲绿色、有课灰色、请假红色、值班蓝色、全员空闲橙色"""
        t = self.gantt_table
        FREE, BUSY, ALL_FREE, LEAVE, DUTY = (
            QColor("#c9f2cf"), QColor("#f2f2f7"), QColor("#ffd9a8"),
            QColor("#ffd6d2"), QColor("#b8d9ff"))
        t.clearContents()
        t.setRowCount(m.member_count + 1)
        t.setColumnCount(len(m.slots))
        t.setVerticalHeaderLabels(m.member_names + ["空闲人数"])
        t.setHorizontalHeaderLabels([slot_header(d, b) for d, b in m.slots])

        for r in range(m.member_count):
            for i, (d, b) in enumerate(m.slots):
                item = QTableWidgetItem("")
                item.setFlags(Qt.ItemIsEnabled)
                leave_reason = m.leave_info.get((r, d))
                if leave_reason is not None:
                    item.setBackground(LEAVE)
                    item.setText("假")
                    item.setForeground(QBrush(QColor("#b3261e")))
                    item.setTextAlignment(Qt.AlignCenter)
                    reason = f"（{leave_reason}）" if leave_reason else ""
                    item.setToolTip(f"第{m.week}周 {WEEKDAY_LABELS[d]}：请假{reason}")
                elif (r, d, b) in m.duty_cells:
                    item.setBackground(DUTY)
                    item.setText("值")
                    item.setForeground(QBrush(QColor("#0a5aa8")))
                    item.setTextAlignment(Qt.AlignCenter)
                    item.setToolTip(f"第{m.week}周 {WEEKDAY_LABELS[d]} {BLOCK_LABELS[b]}：已排值班")
                elif m.free[r][i]:
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

    # ---------- 统计图表 ----------

    def refresh_charts(self) -> None:
        """刷新三张统计图：成员值班次数 / 每周值班人次 / 成员请假天数"""
        members = self.db.list_members()
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
        leave_cnt = Counter(l.member_id for l in self.db.list_leaves())
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
        render_charts_png(charts, Path(path))
        self.statusBar().showMessage(f"已导出统计图：{path}")

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
        self._refresh_leaves(member_id)
        self.refresh_gantt()
        self.refresh_charts()
        self._mark_stale()
        self.statusBar().showMessage("已删除该请假记录。")

    def export_leaves(self) -> None:
        """导出所有成员的请假记录为 Excel（按钮位于请假登记卡片）"""
        leaves = self.db.list_leaves()
        if not leaves:
            QMessageBox.information(self, "提示", "暂无请假记录可导出。")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "导出请假记录", "请假记录.xlsx", "Excel 文件 (*.xlsx)")
        if not path:
            return
        members = self.db.list_members()
        Path(path).write_bytes(
            export_leaves_excel(leaves, members, start_date=self._term_start()))
        self.statusBar().showMessage(f"已导出请假记录（{len(leaves)} 条）：{path}")

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
        members = self.db.list_members()
        busy = build_busy_map(members, self.db.get_courses())
        leave_set = {(l.member_id, l.week, l.weekday) for l in self.db.list_leaves()}
        cands = replacement_candidates(
            members, busy, leave_set, self.result.assignments,
            week, weekday, block, cfg.max_per_week, cfg.max_per_day)

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
                item.setToolTip("课程/请假/当天条件均满足，手动换入将超过每周上限，请知悉")
            lst.addItem(item)

    def _apply_tweak(self, week: int, weekday: int, block: int, out_id: int | None, in_id: int | None) -> None:
        """执行微调：out_id 换出（None=纯新增），in_id 换入（None=纯移除）"""
        if self.result is None or self._last_config is None:
            return
        members = {m.id: m for m in self.db.list_members()}
        cfg = self._last_config
        if in_id is not None:
            busy = build_busy_map(list(members.values()), self.db.get_courses())
            leave_set = {(l.member_id, l.week, l.weekday) for l in self.db.list_leaves()}
            # 硬约束（课程/请假/每天一次）必须满足；周上限允许手动越限
            eligible = {m.id for m, reason in replacement_candidates(
                list(members.values()), busy, leave_set, self.result.assignments,
                week, weekday, block, cfg.max_per_week, cfg.max_per_day)
                if reason in ("", "本周已达上限")}
            if in_id not in eligible:
                QMessageBox.warning(self, "无法调整", "该成员在此时段不满足值班条件，请重新选择。")
                return
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
        self.result.assignments = kept
        self.result.member_stats = rebuild_member_stats(list(members.values()), kept)
        grid = {(w, d, b) for w in cfg.weeks for d in cfg.weekdays for b in cfg.blocks}
        covered = {(a.week, a.weekday, a.block) for a in kept}
        self.result.gaps = sorted(g for g in grid if g not in covered)
        self.db.clear_assignments()
        self.db.save_assignments(kept)
        self._stale = False
        self.refresh_schedule_tabs()
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
        known = {(m.student_id, m.name) for m in self.db.list_members()}
        added, updated, errors = 0, 0, []
        for f in files:
            try:
                schedule = parse_schedule_file(Path(f).read_bytes(), Path(f).name)
            except Exception as e:
                errors.append(f"{Path(f).name}：{e}")
                continue
            if not schedule.courses and not schedule.name:
                errors.append(f"{Path(f).name}：未解析到课程信息，请确认是教务系统导出的个人课表")
                continue
            is_update = (schedule.student_id, schedule.name) in known
            self.db.upsert_member(schedule)
            known.add((schedule.student_id, schedule.name))
            added += 0 if is_update else 1
            updated += 1 if is_update else 0
        self.refresh_members()
        if errors:
            QMessageBox.warning(self, "部分文件导入失败", "\n".join(errors))
        parts = []
        if added:
            parts.append(f"新增 {added} 名成员")
        if updated:
            parts.append(f"更新 {updated} 份课表（同学号同名，课程整体替换）")
        if parts:
            self.statusBar().showMessage("、".join(parts) + f"；当前共 {len(self.db.list_members())} 名成员。")
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
        config = self._current_config()
        # 按周增量：只重排所选周范围，范围外的历史排班保留并作为均衡基数
        existing = self.db.load_assignments()
        base = [a for a in existing if a.week not in config.weeks]
        result = generate_schedule(
            members, self.db.get_courses(), config,
            leaves=self.db.list_leaves(), base_assignments=base)
        fresh = [a for a in result.assignments if a.week in config.weeks]
        self.db.delete_assignments_for_weeks(list(config.weeks))
        self.db.save_assignments(fresh)
        self.result = result
        self._last_config = config
        self._stale = False
        self.refresh_schedule_tabs()
        self.refresh_gantt()
        self.tabs.setCurrentIndex(0)
        leave_n = len(self.db.list_leaves())
        scope = (f"第{config.weeks.start}–{config.weeks.stop - 1}周" if len(config.weeks) > 1
                 else f"第{config.weeks.start}周")
        kept = f"，范围外历史排班 {len(base)} 人次保留" if base else ""
        self.statusBar().showMessage(
            f"已重新排班 {scope}：本次安排 {len(fresh)} 人次{kept}，"
            f"无人可用时段 {len(result.gaps)} 个，总次数极差 {result.balanced_spread}"
            + (f"，已避让 {leave_n} 条请假记录。" if leave_n else "。"))

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
        members = self.db.list_members()
        stats = rebuild_member_stats(members, assignments)
        gaps = [g for g in self.result.gaps if g[0] == week]
        return assignments, stats, gaps, f"第{week}周"

    def export_xlsx(self) -> None:
        if self.result is None:
            return
        scope = self._export_scope()
        if scope is None:
            return
        assignments, stats, gaps, weeks_txt = scope
        path, _ = QFileDialog.getSaveFileName(
            self, "导出 Excel", f"值班排班表_{weeks_txt}.xlsx", "Excel 文件 (*.xlsx)")
        if not path:
            return
        Path(path).write_bytes(export_excel(assignments, stats, gaps, start_date=self._term_start()))
        self.statusBar().showMessage(f"已导出 Excel（{weeks_txt}）：{path}")

    def export_csv(self) -> None:
        if self.result is None:
            return
        scope = self._export_scope()
        if scope is None:
            return
        assignments, _, _, weeks_txt = scope
        path, _ = QFileDialog.getSaveFileName(
            self, "导出 CSV", f"值班排班表_{weeks_txt}.csv", "CSV 文件 (*.csv)")
        if not path:
            return
        Path(path).write_bytes(export_csv(assignments, start_date=self._term_start()))
        self.statusBar().showMessage(f"已导出 CSV（{weeks_txt}）：{path}")

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
        if weeks_txt.startswith("第") and "–" not in weeks_txt:
            week = int(weeks_txt[1:-1])
            weekdays = sorted({a.weekday for a in assignments}) or [1]
            fd = week_date(start, week, weekdays[0])
            ld = week_date(start, week, weekdays[-1])
            subtitle = f"{weeks_txt}（{fd.month}月{fd.day}日–{ld.month}月{ld.day}日）"
        else:
            subtitle = (f"{weeks_txt} · {start.year}年{start.month}月{start.day}日起")
        render_table_png(self._pivot_frame(assignments), "值班排班表", subtitle, Path(path))
        self.statusBar().showMessage(f"已导出图片（{weeks_txt}）：{path}")


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
