"""成员空闲时段甘特图

数据层：计算第 W 周内各成员在每个 (星期, 时段块) 的忙闲状态，
空闲判定与排班算法完全一致（时段块内每一节均无课才算空闲）。
展示：app.py 用 QTableWidget 染色绘制甘特图；本模块提供导出
带填充色的 Excel 甘特图（行=成员，列=星期x时段，绿=空闲）。
"""

from __future__ import annotations

import io
from dataclasses import dataclass, field

from .database import Assignment, CourseRecord, Leave, Member
from .parser import BLOCK_LABELS, BLOCK_SESSIONS, WEEKDAY_LABELS
from .scheduler import build_busy_map


@dataclass
class AvailabilityMatrix:
    """第 week 周各成员忙闲矩阵（甘特图数据源）"""
    week: int
    member_names: list[str]                                 # 行标签
    slots: list[tuple[int, int]]                            # 列 (星期, 时段块)
    free: list[list[bool]]                                  # [成员行][时段列] 是否空闲
    busy_courses: dict[tuple[int, int, int], list[str]]     # (行号, 星期, 时段) -> 忙时课程名
    free_counts: list[int]                                 # 各时段空闲人数
    leave_info: dict[tuple[int, int], str] = field(default_factory=dict)  # (行号, 星期) -> 请假原因
    duty_cells: set[tuple[int, int, int]] = field(default_factory=set)  # (行号, 星期, 时段) -> 已排值班

    @property
    def member_count(self) -> int:
        return len(self.member_names)

    def slot_column(self, weekday: int, block: int) -> int:
        return self.slots.index((weekday, block))

    @property
    def all_free_slots(self) -> list[tuple[int, int]]:
        """全员共同空闲的 (星期, 时段)，方便安排集体任务"""
        return [slot for i, slot in enumerate(self.slots)
                if self.member_count > 0 and self.free_counts[i] == self.member_count]


def slot_header(weekday: int, block: int) -> str:
    """甘特图列标签，两行显示：星期 / 节次"""
    return f"{WEEKDAY_LABELS[weekday]}\n{BLOCK_LABELS[block].split(' ')[0]}"


def build_availability(
    members: list[Member],
    courses: list[CourseRecord],
    week: int,
    weekdays: list[int],
    blocks: list[int],
    leaves: list[Leave] | None = None,
    assignments: list[Assignment] | None = None,
) -> AvailabilityMatrix:
    """计算第 week 周各成员忙闲矩阵（请假成员当天整行不可用，已排值班单元格标蓝）"""
    busy = build_busy_map(members, courses)
    leave_map = {(l.member_id, l.weekday): l.reason
                 for l in (leaves or []) if l.week == week}
    duty_set = {(a.member_id, a.weekday, a.block)
                for a in (assignments or []) if a.week == week}
    # 课程名索引：(成员id, 星期, 节次) -> 课程名集合（供 tooltip / 单元格详情）
    course_index: dict[tuple[int, int, int], set[str]] = {}
    member_ids = {m.id for m in members}
    for c in courses:
        if c.member_id not in member_ids or week not in c.week_list:
            continue
        for s in c.session_list:
            course_index.setdefault((c.member_id, c.weekday, s), set()).add(c.course_name)

    slots = [(d, b) for d in sorted(weekdays) for b in sorted(blocks)]
    free: list[list[bool]] = []
    busy_courses: dict[tuple[int, int, int], list[str]] = {}
    leave_info: dict[tuple[int, int], str] = {}
    duty_cells: set[tuple[int, int, int]] = set()
    for m in members:
        row = len(free)  # 当前行号
        row_free: list[bool] = []
        for d, b in slots:
            if (m.id, d) in leave_map:
                row_free.append(False)
                leave_info[(row, d)] = leave_map[(m.id, d)]
                continue
            if (m.id, d, b) in duty_set:
                row_free.append(False)
                duty_cells.add((row, d, b))
                continue
            names: set[str] = set()
            is_free = True
            for s in BLOCK_SESSIONS[b]:
                if (week, d, s) in busy.get(m.id, ()):
                    is_free = False
                    names |= course_index.get((m.id, d, s), set())
            row_free.append(is_free)
            if names:
                busy_courses[(row, d, b)] = sorted(names)
        free.append(row_free)

    free_counts = [sum(1 for row in free if row[i]) for i in range(len(slots))]
    return AvailabilityMatrix(
        week=week,
        member_names=[m.name for m in members],
        slots=slots,
        free=free,
        busy_courses=busy_courses,
        free_counts=free_counts,
        leave_info=leave_info,
        duty_cells=duty_cells,
    )


def export_gantt_excel(matrix: AvailabilityMatrix) -> bytes:
    """导出带填充色的 Excel 甘特图，返回文件字节"""
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    FREE_FILL = PatternFill("solid", fgColor="C9F2CF")
    BUSY_FILL = PatternFill("solid", fgColor="F2F2F7")
    ALL_FREE_FILL = PatternFill("solid", fgColor="FFD9A8")
    LEAVE_FILL = PatternFill("solid", fgColor="FFD6D2")
    DUTY_FILL = PatternFill("solid", fgColor="B8D9FF")
    HEADER_FILL = PatternFill("solid", fgColor="007AFF")
    THIN = Side(style="thin", color="FFFFFF")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"第{matrix.week}周空闲甘特图"

    # 表头两行：第1行星期（横向合并），第2行节次
    for r in (1, 2):
        corner = ws.cell(row=r, column=1, value="成员" if r == 1 else f"第{matrix.week}周")
        corner.font = Font(bold=True, color="FFFFFF")
        corner.fill = HEADER_FILL
        corner.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)

    col = 2
    for d in sorted({d for d, _ in matrix.slots}):
        n = sum(1 for dd, _ in matrix.slots if dd == d)
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + n - 1)
        cell = ws.cell(row=1, column=col, value=WEEKDAY_LABELS[d])
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        col += n
    for i, (d, b) in enumerate(matrix.slots):
        cell = ws.cell(row=2, column=2 + i, value=BLOCK_LABELS[b].split(" ")[0])
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # 成员行：空闲绿色、有课灰色、请假红色、值班蓝色
    border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    for r, name in enumerate(matrix.member_names):
        row = 3 + r
        head = ws.cell(row=row, column=1, value=name)
        head.font = Font(bold=True)
        head.alignment = Alignment(horizontal="center", vertical="center")
        for i, (d, b) in enumerate(matrix.slots):
            leave_reason = matrix.leave_info.get((r, d))
            is_free = matrix.free[r][i]
            names = matrix.busy_courses.get((r, d, b), [])
            if leave_reason is not None:
                cell = ws.cell(row=row, column=2 + i,
                               value=f"请假：{leave_reason}" if leave_reason else "请假")
                cell.fill = LEAVE_FILL
                cell.font = Font(color="B3261E")
            elif (r, d, b) in matrix.duty_cells:
                cell = ws.cell(row=row, column=2 + i, value="值班")
                cell.fill = DUTY_FILL
                cell.font = Font(bold=True, color="0A5AA8")
            else:
                cell = ws.cell(row=row, column=2 + i, value="" if is_free else "、".join(names))
                cell.fill = FREE_FILL if is_free else BUSY_FILL
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)

    # 汇总行：各时段空闲人数，全员空闲橙色高亮
    row = 3 + len(matrix.member_names)
    head = ws.cell(row=row, column=1, value="空闲人数")
    head.font = Font(bold=True)
    head.alignment = Alignment(horizontal="center", vertical="center")
    for i, (d, b) in enumerate(matrix.slots):
        all_free = matrix.member_count > 0 and matrix.free_counts[i] == matrix.member_count
        cell = ws.cell(row=row, column=2 + i, value=f"{matrix.free_counts[i]}/{matrix.member_count}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
        if all_free:
            cell.fill = ALL_FREE_FILL
            cell.font = Font(bold=True, color="B25E00")

    ws.freeze_panes = "B3"
    ws.column_dimensions["A"].width = 12
    for i in range(len(matrix.slots)):
        ws.column_dimensions[get_column_letter(2 + i)].width = 10

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
