"""表格与统计图 PNG 渲染。

大表格和大成员量统计图采用分页输出，避免构造超出 QImage 限制的超宽/超高图片。
"""

from __future__ import annotations

from math import ceil
from pathlib import Path

import pandas as pd
from PySide6.QtCore import QRect, Qt
from PySide6.QtGui import QColor, QFont, QFontMetrics, QImage, QPainter, QPen
from PySide6.QtWidgets import QFrame, QWidget

_IMAGE_SCALE = 2
_MAX_PAGE_WIDTH = 2400
_MAX_TABLE_CONTENT_HEIGHT = 4200
_MAX_CHART_CONTENT_WIDTH = 2200
_TABLE_CELL_MAX_WIDTH = 280


def _wrap_text(text: str, metrics: QFontMetrics, max_width: int) -> list[str]:
    """按像素宽度把文本拆成多行，避免 PNG 内容被裁切。"""
    result: list[str] = []
    for raw_line in str(text).split("\n"):
        if not raw_line:
            result.append("")
            continue
        current = ""
        for char in raw_line:
            candidate = current + char
            if current and metrics.horizontalAdvance(candidate) > max_width:
                result.append(current)
                current = char
            else:
                current = candidate
        result.append(current)
    return result or [""]


def _page_paths(path: Path, total: int) -> list[Path]:
    if total <= 1:
        return [path]
    return [
        path.with_name(f"{path.stem}_{index + 1}{path.suffix or '.png'}")
        for index in range(total)
    ]


def _column_pages(widths: list[int], fixed_count: int) -> list[list[int]]:
    available = _MAX_PAGE_WIDTH - 56
    pages: list[list[int]] = []
    current = list(range(fixed_count))
    current_width = sum(widths[:fixed_count])
    for column_index in range(fixed_count, len(widths)):
        width = widths[column_index]
        if len(current) > fixed_count and current_width + width > available:
            pages.append(current)
            current = list(range(fixed_count))
            current_width = sum(widths[:fixed_count])
        current.append(column_index)
        current_width += width
    if current or not pages:
        pages.append(current)
    return pages


def _row_pages(row_heights: list[int]) -> list[list[int]]:
    pages: list[list[int]] = []
    current: list[int] = []
    current_height = 0
    for row_index, height in enumerate(row_heights):
        if current and current_height + height > _MAX_TABLE_CONTENT_HEIGHT:
            pages.append(current)
            current = []
            current_height = 0
        current.append(row_index)
        current_height += height
    if current or not pages:
        pages.append(current)
    return pages


def _fixed_column_count(columns: list[str]) -> int:
    if len(columns) >= 3 and columns[:2] == ["周次", "星期"] and columns[2] == "日期":
        return 3
    if len(columns) >= 2 and columns[:2] == ["周次", "星期"]:
        return 2
    return 0


def _render_table_page(
    df: pd.DataFrame,
    title: str,
    subtitle: str,
    page_path: Path,
    column_indices: list[int],
    row_indices: list[int],
    widths: list[int],
) -> None:
    margin, pad, min_row_h, min_head_h = 28, 12, 34, 38
    line_h = 18
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

    page_widths = [widths[index] for index in column_indices]
    column_names = [str(df.columns[index]) for index in column_indices]
    header_lines = [
        _wrap_text(name, hfm, page_widths[index] - pad * 2)
        for index, name in enumerate(column_names)
    ]
    head_h = max(
        min_head_h,
        max((len(lines) for lines in header_lines), default=1) * line_h + pad * 2,
    )

    body_lines: list[list[list[str]]] = []
    row_heights: list[int] = []
    for row_index in row_indices:
        current_row: list[list[str]] = []
        max_lines = 1
        for column_index in column_indices:
            lines = _wrap_text(
                str(df.iat[row_index, column_index]),
                fm,
                widths[column_index] - pad * 2,
            )
            current_row.append(lines)
            max_lines = max(max_lines, len(lines))
        body_lines.append(current_row)
        row_heights.append(max(min_row_h, max_lines * line_h + pad * 2))

    table_w = sum(page_widths)
    title_w = max(tfm.horizontalAdvance(title), sfm.horizontalAdvance(subtitle))
    content_w = max(table_w, title_w)
    img_w = margin * 2 + content_w
    table_y = margin + tfm.height() + 8 + sfm.height() + 16
    table_h = head_h + sum(row_heights)
    img_h = table_y + table_h + margin

    img = QImage(img_w * _IMAGE_SCALE, img_h * _IMAGE_SCALE, QImage.Format_ARGB32)
    img.fill(Qt.white)
    painter = QPainter(img)
    painter.setRenderHint(QPainter.Antialiasing)
    painter.setRenderHint(QPainter.TextAntialiasing)
    painter.scale(_IMAGE_SCALE, _IMAGE_SCALE)

    painter.setPen(QPen(QColor("#1d1d1f")))
    painter.setFont(title_font)
    painter.drawText(QRect(margin, margin, content_w, tfm.height()),
                     Qt.AlignLeft | Qt.AlignVCenter, title)
    painter.setPen(QPen(QColor("#86868b")))
    painter.setFont(sub_font)
    painter.drawText(
        QRect(margin, margin + tfm.height() + 8, content_w, sfm.height()),
        Qt.AlignLeft | Qt.AlignVCenter,
        subtitle,
    )

    painter.fillRect(QRect(margin, table_y, table_w, head_h), QColor("#f2f2f7"))
    painter.setFont(bold)
    x = margin
    for width, lines in zip(page_widths, header_lines):
        text_y = table_y + (head_h - len(lines) * line_h) / 2
        for line in lines:
            painter.setPen(QPen(QColor("#1d1d1f")))
            painter.drawText(
                QRect(x + pad, text_y, width - pad * 2, line_h),
                Qt.AlignCenter,
                line,
            )
            text_y += line_h
        x += width

    painter.setFont(font)
    y = table_y + head_h
    for display_row, height in enumerate(row_heights):
        if display_row % 2:
            painter.fillRect(QRect(margin, y, table_w, height), QColor("#f7f7f9"))
        x = margin
        for display_column, width in enumerate(page_widths):
            lines = body_lines[display_row][display_column]
            text_y = y + (height - len(lines) * line_h) / 2
            for line in lines:
                painter.setPen(QPen(
                    QColor("#aeaeb2") if line == "—" else "#1d1d1f"))
                painter.drawText(
                    QRect(x + pad, text_y, width - pad * 2, line_h),
                    Qt.AlignCenter,
                    line,
                )
                text_y += line_h
            x += width
        y += height

    painter.setPen(QPen(QColor("#e5e5ea"), 1))
    x = margin
    for width in page_widths[:-1]:
        x += width
        painter.drawLine(x, table_y, x, table_y + table_h)
    y = table_y + head_h
    for height in row_heights[:-1]:
        y += height
        painter.drawLine(margin, y, margin + table_w, y)
    painter.setPen(QPen(QColor("#d2d2d7"), 1))
    painter.drawRect(QRect(margin, table_y, table_w, table_h))
    painter.end()
    img.save(str(page_path))


def render_table_png(
    df: pd.DataFrame,
    title: str,
    subtitle: str,
    path: Path,
) -> list[Path]:
    """渲染表格 PNG；内容超限时按行列分页，返回实际生成的文件列表。"""
    cols = [str(column) for column in df.columns]
    if not cols:
        return []

    font = QFont()
    font.setPixelSize(13)
    bold = QFont()
    bold.setPixelSize(13)
    bold.setBold(True)
    fm = QFontMetrics(font)
    hfm = QFontMetrics(bold)

    widths: list[int] = []
    for column_index, column_name in enumerate(cols):
        natural_width = hfm.horizontalAdvance(column_name)
        for row_index in range(len(df)):
            text = str(df.iat[row_index, column_index])
            natural_width = max(
                natural_width,
                max(
                    (fm.horizontalAdvance(line) for line in text.split("\n")),
                    default=0,
                ),
            )
        widths.append(
            min(max(natural_width + 24, 56), _TABLE_CELL_MAX_WIDTH))

    row_heights: list[int] = []
    for row_index in range(len(df)):
        max_lines = 1
        for column_index, width in enumerate(widths):
            lines = _wrap_text(
                str(df.iat[row_index, column_index]),
                fm,
                width - 24,
            )
            max_lines = max(max_lines, len(lines))
        row_heights.append(max(34, max_lines * 18 + 24))

    column_page_indices = _column_pages(widths, _fixed_column_count(cols))
    row_page_indices = _row_pages(row_heights)
    total_pages = len(column_page_indices) * len(row_page_indices)
    output_paths = _page_paths(path, total_pages)

    page_number = 0
    for row_indices in row_page_indices:
        for column_indices in column_page_indices:
            page_number += 1
            page_title = title
            page_subtitle = subtitle
            if total_pages > 1:
                page_title = f"{title}（第 {page_number}/{total_pages} 页）"
                page_subtitle = f"{subtitle} · 表头与日期列为跨页重复列"
            _render_table_page(
                df,
                page_title,
                page_subtitle,
                output_paths[page_number - 1],
                column_indices,
                row_indices,
                widths,
            )
    return output_paths


def _draw_bar_chart(
    painter: QPainter,
    rect: QRect,
    title: str,
    items: list[tuple[str, int]],
    color: str = "#2563eb",
) -> None:
    """在 rect 内绘制标题 + 柱状图（类苹果风格）。"""
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

    label_area_h = lfm.height() * 2
    top = rect.top() + tfm.height() + 14
    bottom = rect.bottom() - label_area_h - 8
    left = rect.left() + 30
    right = rect.right() - 8
    plot_h = bottom - top
    if plot_h <= 20:
        return
    max_v = max(value for _, value in items)
    top_tick = max_v if max_v % 4 == 0 else (max_v // 4 + 1) * 4
    top_tick = max(top_tick, 4)

    painter.setFont(label_font)
    grid_pen = QPen(QColor("#e5e5ea"), 1)
    for index in range(5):
        y = bottom - plot_h * index / 4
        painter.setPen(grid_pen)
        painter.drawLine(left, y, right, y)
        painter.setPen(QPen(QColor("#aeaeb2")))
        painter.drawText(
            QRect(left - 30, y - lfm.height() / 2, 26, lfm.height()),
            Qt.AlignRight | Qt.AlignVCenter,
            str(top_tick * index // 4),
        )

    slot_w = (right - left) / len(items)
    bar_w = min(slot_w * 0.52, 64)
    bar_color = QColor(color)
    for index, (label, value) in enumerate(items):
        center_x = left + slot_w * (index + 0.5)
        height = plot_h * value / top_tick if top_tick else 0
        bar_rect = QRect(center_x - bar_w / 2, bottom - height, bar_w, height)
        painter.setPen(Qt.NoPen)
        painter.setBrush(bar_color)
        painter.drawRoundedRect(bar_rect, 3, 3)
        painter.setPen(QPen(QColor("#1d1d1f")))
        painter.drawText(
            QRect(center_x - slot_w / 2, bottom - height - lfm.height() - 2,
                  slot_w, lfm.height()),
            Qt.AlignCenter,
            str(value),
        )
        painter.setPen(QPen(QColor("#86868b")))
        painter.drawText(
            QRect(center_x - slot_w / 2, bottom + 2, slot_w, label_area_h),
            Qt.AlignHCenter | Qt.AlignTop | Qt.TextWordWrap,
            label,
        )


def _render_charts_page(
    charts: list[tuple[str, list[tuple[str, int]], str]],
    path: Path,
) -> None:
    margin, chart_h = 28, 300
    label_font = QFont()
    label_font.setPixelSize(12)
    title_font = QFont()
    title_font.setPixelSize(13)
    title_font.setBold(True)
    label_fm = QFontMetrics(label_font)
    title_fm = QFontMetrics(title_font)

    max_items = max((len(items) for _, items, _ in charts), default=1)
    max_label_w = max(
        (
            label_fm.horizontalAdvance(label)
            for _, items, _ in charts
            for label, _ in items
        ),
        default=64,
    )
    slot_w = max(64, min(max_label_w + 18, 140))
    content_w = 38 + slot_w * max_items
    content_w = max(
        content_w,
        *(title_fm.horizontalAdvance(title) for title, _, _ in charts),
    )
    img_w = max(980, content_w + margin * 2)
    img_h = margin * 2 + chart_h * len(charts) + 16 * (len(charts) - 1)
    img = QImage(img_w * _IMAGE_SCALE, img_h * _IMAGE_SCALE, QImage.Format_ARGB32)
    img.fill(Qt.white)
    painter = QPainter(img)
    painter.setRenderHint(QPainter.Antialiasing)
    painter.setRenderHint(QPainter.TextAntialiasing)
    painter.scale(_IMAGE_SCALE, _IMAGE_SCALE)
    for index, (title, items, color) in enumerate(charts):
        _draw_bar_chart(
            painter,
            QRect(margin, margin + index * (chart_h + 16),
                  img_w - margin * 2, chart_h),
            title,
            items,
            color,
        )
    painter.end()
    img.save(str(path))


def render_charts_png(
    charts: list[tuple[str, list[tuple[str, int]], str]],
    path: Path,
) -> list[Path]:
    """渲染统计图 PNG；成员过多时按每页标签数分页，返回文件列表。"""
    if not charts:
        return []
    label_font = QFont()
    label_font.setPixelSize(12)
    label_fm = QFontMetrics(label_font)

    max_label_w = max(
        (
            label_fm.horizontalAdvance(label)
            for _, items, _ in charts
            for label, _ in items
        ),
        default=64,
    )
    slot_w = max(64, min(max_label_w + 18, 140))
    available = _MAX_CHART_CONTENT_WIDTH - 38 - 56
    items_per_page = max(1, int(available // slot_w))
    max_items = max((len(items) for _, items, _ in charts), default=1)
    total_pages = max(1, ceil(max_items / items_per_page))
    output_paths = _page_paths(path, total_pages)

    for page_index in range(total_pages):
        start = page_index * items_per_page
        end = start + items_per_page
        page_charts = [
            (
                f"{title}（第 {page_index + 1}/{total_pages} 页）"
                if total_pages > 1 else title,
                items[start:end],
                color,
            )
            for title, items, color in charts
        ]
        _render_charts_page(page_charts, output_paths[page_index])
    return output_paths


class BarChart(QFrame):
    """类苹果风格柱状图（值班 / 请假情况可视化）"""

    def __init__(
        self,
        title: str,
        color: str = "#2563eb",
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self._title = title
        self._color = color
        self._items: list[tuple[str, int]] = []
        self.setMinimumHeight(220)
        self.setStyleSheet(
            "BarChart { background: #ffffff; border: 1px solid #dde5f0; "
            "border-radius: 14px; }")

    def set_data(self, items: list[tuple[str, int]]) -> None:
        self._items = list(items)
        self.update()

    def data(self) -> list[tuple[str, int]]:
        return self._items

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.TextAntialiasing)
        _draw_bar_chart(
            painter,
            self.rect().adjusted(16, 12, -16, -10),
            self._title,
            self._items,
            self._color,
        )
        painter.end()
