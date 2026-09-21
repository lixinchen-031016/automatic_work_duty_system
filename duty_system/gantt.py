"""成员空闲时段甘特图

数据层：计算第 W 周内各成员在每个 (星期, 时段块) 的忙闲状态，
空闲判定与排班算法完全一致（时段块内每一节均无课、无长期特殊安排才算空闲）。
展示：app.py 用 QTableWidget 染色绘制甘特图；本模块提供导出
带填充色的 Excel 甘特图（行=成员，列=星期x时段，绿=空闲）。

性能：build_availability 接受调用方传入的 busy 忙时表（缓存后按周切换
无需重建）；课程名索引复用忙时表判断「本周是否上课」，避免在 week_list
上做线性查找。
"""

from __future__ import annotations

import io
from dataclasses import dataclass, field

from .database import Assignment, CourseRecord, Leave, Member, SpecialArrangement
from .parser import (
    BLOCK_LABELS,
    BLOCK_SESSIONS,
    WEEKDAY_LABELS,
    WHOLE_WEEK_SESSIONS,
    WHOLE_WEEK_WEEKDAY,
)
from .scheduler import build_busy_map, build_special_busy_map


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
    special_info: dict[tuple[int, int, int], list[str]] = field(default_factory=dict)  # (行号, 星期, 时段) -> 特殊安排原因

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
    busy: dict[int, set] | None = None,
    special_arrangements: list[SpecialArrangement] | None = None,
) -> AvailabilityMatrix:
    """计算第 week 周各成员忙闲矩阵（请假成员当天整行不可用，已排值班单元格标蓝）

    busy 可传入调用方缓存的忙时表（见 scheduler.build_busy_map）：
    该表展开成本高，界面按周切换甘特图时无需重复构建。
    """
    if busy is None:
        busy = build_busy_map(members, courses)
    special_busy = build_special_busy_map(members, special_arrangements)
    leave_map = {(l.member_id, l.weekday): l.reason
                 for l in (leaves or []) if l.week == week}
    duty_set = {(a.member_id, a.weekday, a.block)
                for a in (assignments or []) if a.week == week}
    # 课程名索引：(成员id, 星期, 节次) -> 课程名集合（供 tooltip / 单元格详情）。
    # 先按 week_list 过滤课程，再用忙时表定位本门课的具体节次；不能只看忙时表，
    # 因为不同课程可能占用同一星期节次，但只在其他周上课。
    course_index: dict[tuple[int, int, int], set[str]] = {}
    special_index: dict[tuple[int, int, int], set[str]] = {}
    for a in special_arrangements or []:
        if week not in a.week_list:
            continue
        label = a.reason.strip() or "其他安排"
        for session in a.session_list:
            special_index.setdefault(
                (a.member_id, a.weekday, session), set()).add(label)
    member_ids = {m.id for m in members}
    for c in courses:
        if c.member_id not in member_ids:
            continue
        if week not in c.week_list:
            continue
        member_busy = busy.get(c.member_id)
        if not member_busy:
            continue
        if c.weekday == WHOLE_WEEK_WEEKDAY:
            # 整周集中安排（军训/思政实践）没有星期与节次：该周每一天每一节都算被
            # 占用，课程名也要挂上去——否则甘特图上是一整片「忙但没有原因」的格子，
            # 用户看不出为什么这个人整周都排不了班。
            for day in range(1, 8):
                for sec in WHOLE_WEEK_SESSIONS:
                    course_index.setdefault((c.member_id, day, sec), set()).add(c.course_name)
            continue
        weekday = c.weekday
        active = [sec for sec in c.session_list if (week, weekday, sec) in member_busy]
        if not active:
            continue
        for session in active:
            course_index.setdefault((c.member_id, weekday, session), set()).add(c.course_name)

    slots = [(d, b) for d in sorted(weekdays) for b in sorted(blocks)]
    free: list[list[bool]] = []
    busy_courses: dict[tuple[int, int, int], list[str]] = {}
    leave_info: dict[tuple[int, int], str] = {}
    duty_cells: set[tuple[int, int, int]] = set()
    special_info: dict[tuple[int, int, int], list[str]] = {}
    for m in members:
        row = len(free)  # 当前行号
        member_busy = busy.get(m.id, ())
        member_special = special_busy.get(m.id, ())
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
            special_names: set[str] = set()
            is_free = True
            for s in BLOCK_SESSIONS[b]:
                if (week, d, s) in member_busy:
                    is_free = False
                    names |= course_index.get((m.id, d, s), set())
                if (week, d, s) in member_special:
                    is_free = False
                    special_names |= special_index.get((m.id, d, s), set())
            row_free.append(is_free)
            if names:
                busy_courses[(row, d, b)] = sorted(names)
            if special_names:
                special_info[(row, d, b)] = sorted(special_names)
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
        special_info=special_info,
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
    SPECIAL_FILL = PatternFill("solid", fgColor="E5D8FF")
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

    # 成员行：空闲绿色、有课灰色、特殊安排紫色、请假红色、值班蓝色
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
            special_names = matrix.special_info.get((r, d, b), [])
            if leave_reason is not None:
                cell = ws.cell(row=row, column=2 + i,
                               value=f"请假：{leave_reason}" if leave_reason else "请假")
                cell.fill = LEAVE_FILL
                cell.font = Font(color="B3261E")
            elif (r, d, b) in matrix.duty_cells:
                cell = ws.cell(row=row, column=2 + i, value="值班")
                cell.fill = DUTY_FILL
                cell.font = Font(bold=True, color="0A5AA8")
            elif special_names:
                cell = ws.cell(row=row, column=2 + i,
                               value="其他安排：" + "、".join(special_names))
                cell.fill = SPECIAL_FILL
                cell.font = Font(color="6E3DC2")
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
