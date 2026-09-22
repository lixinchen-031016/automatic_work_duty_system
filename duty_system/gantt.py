"""成员空闲时段甘特图

数据层：计算第 W 周内各成员在每个 (星期, 时段块) 的忙闲状态，
空闲判定与排班算法完全一致（时段块内每一节均无课、无长期特殊安排才算空闲）。
展示：app.py 用 QTableWidget 染色绘制甘特图；本模块提供导出
带填充色的 Excel 甘特图（行=星期x时段，列=成员，绿=空闲）。

性能：build_availability 接受调用方传入的 busy 忙时表（缓存后按周切换
无需重建）；课程名索引复用忙时表判断「本周是否上课」，避免在 week_list
上做线性查找。
"""

from __future__ import annotations

import io
from dataclasses import dataclass, field
from datetime import date, timedelta

from .calendar import TermCalendar
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
    calendar_dates: dict[int, date | None] = field(default_factory=dict)  # 逻辑星期 -> 真实日期
    off_dates: dict[int, date] = field(default_factory=dict)  # 逻辑星期 -> 被放假的自然日期
    off_weekdays: set[int] = field(default_factory=set)  # 纯法定假日：当天所有成员均不可用

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


@dataclass(frozen=True)
class CalendarGanttColumn:
    """自然日期驱动的甘特图时段项。"""

    date: date
    block: int
    logical_week: int | None = None
    logical_weekday: int | None = None
    is_off: bool = False


@dataclass
class CalendarAvailabilityMatrix:
    """自然周甘特图：补课日归属其真实日期所在周。"""

    week: int
    member_names: list[str]
    columns: list[CalendarGanttColumn]
    free: list[list[bool]]
    busy_courses: dict[tuple[int, date, int], list[str]]
    free_counts: list[int]
    leave_info: dict[tuple[int, date], str] = field(default_factory=dict)
    duty_cells: set[tuple[int, date, int]] = field(default_factory=set)
    special_info: dict[tuple[int, date, int], list[str]] = field(default_factory=dict)

    @property
    def member_count(self) -> int:
        return len(self.member_names)

    @property
    def all_free_columns(self) -> list[CalendarGanttColumn]:
        return [
            column for i, column in enumerate(self.columns)
            if self.member_count > 0 and self.free_counts[i] == self.member_count
        ]


def slot_header(
    weekday: int,
    block: int,
    actual_date: date | None = None,
    *,
    is_off: bool = False,
) -> str:
    """甘特图时段标签，两行显示：真实日期 / 节次。"""
    if actual_date is None:
        return f"{WEEKDAY_LABELS[weekday]}\n{BLOCK_LABELS[block].split(' ')[0]}"
    day = WEEKDAY_LABELS[actual_date.isoweekday()]
    suffix = "（放假）" if is_off else ""
    return (f"{actual_date.month:02d}-{actual_date.day:02d}\n"
            f"{day}{suffix} · {BLOCK_LABELS[block].split(' ')[0]}")


@dataclass(frozen=True)
class GanttColumn:
    """甘特图展示时段项；补课时可额外保留原始放假日项。"""

    weekday: int
    block: int
    date: date | None
    is_off: bool = False
    slot_index: int | None = None


def display_columns(matrix: AvailabilityMatrix) -> list[GanttColumn]:
    """逻辑排班时段项 + 调休时额外显示的原始放假日项。"""
    columns: list[GanttColumn] = []
    for slot_index, (weekday, block) in enumerate(matrix.slots):
        is_off = weekday in matrix.off_weekdays
        actual_date = (matrix.off_dates.get(weekday) if is_off else
                       matrix.calendar_dates.get(weekday))
        columns.append(GanttColumn(
            weekday, block, actual_date, is_off, slot_index))

    blocks = sorted({block for _, block in matrix.slots})
    for weekday in sorted(matrix.off_dates):
        if weekday in matrix.off_weekdays:
            continue
        for block in blocks:
            columns.append(GanttColumn(
                weekday, block, matrix.off_dates[weekday], True))
    if any(column.date is not None for column in columns):
        columns.sort(key=lambda column: (
            column.date or date.max, column.block, column.is_off))
    return columns


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
    calendar: TermCalendar | None = None,
) -> AvailabilityMatrix:
    """计算第 week 周各成员忙闲矩阵（请假成员当天整行不可用，已排值班单元格标蓝）

    busy 可传入调用方缓存的忙时表（见 scheduler.build_busy_map）：
    该表展开成本高，界面按周切换甘特图时无需重复构建。
    """
    if busy is None:
        busy = build_busy_map(members, courses)
    calendar_dates: dict[int, date | None] = {}
    off_dates: dict[int, date] = {}
    off_weekdays: set[int] = set()
    if calendar is not None:
        for weekday in sorted(weekdays):
            actual_date = calendar.logical_to_date(week, weekday)
            calendar_dates[weekday] = actual_date
            off_date = calendar.off_natural_date(week, weekday)
            if off_date is not None:
                off_dates[weekday] = off_date
            if actual_date is None:
                off_weekdays.add(weekday)
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
            if d in off_weekdays:
                row_free.append(False)
                continue
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
        calendar_dates=calendar_dates,
        off_dates=off_dates,
        off_weekdays=off_weekdays,
    )


def build_calendar_availability(
    members: list[Member],
    courses: list[CourseRecord],
    calendar: TermCalendar,
    week: int,
    weekdays: list[int],
    blocks: list[int],
    leaves: list[Leave] | None = None,
    assignments: list[Assignment] | None = None,
    busy: dict[int, set] | None = None,
    special_arrangements: list[SpecialArrangement] | None = None,
) -> CalendarAvailabilityMatrix:
    """按自然日期周构建甘特图；补课日自动纳入实际周末。"""
    if busy is None:
        busy = build_busy_map(members, courses)
    special_busy = build_special_busy_map(members, special_arrangements)
    leave_map = {
        (leave.member_id, leave.week, leave.weekday): leave.reason
        for leave in (leaves or [])
    }
    duty_set = {
        (a.member_id, a.week, a.weekday, a.block)
        for a in (assignments or [])
    }
    week_start = calendar.term_start + timedelta(days=(week - 1) * 7)
    class_dates = {
        entry.date for entry in calendar.entries if entry.override_type == "class"
    }
    columns: list[CalendarGanttColumn] = []
    logical_by_column: list[tuple[int, int] | None] = []
    for day_offset in range(7):
        actual_date = week_start + timedelta(days=day_offset)
        logical = calendar.date_to_logical(actual_date)
        include = actual_date.isoweekday() in weekdays or actual_date in class_dates
        if not include:
            continue
        for block in sorted(blocks):
            columns.append(CalendarGanttColumn(
                date=actual_date,
                block=block,
                logical_week=logical[0] if logical else None,
                logical_weekday=logical[1] if logical else None,
                is_off=logical is None,
            ))
            logical_by_column.append(logical)

    # 只索引当前自然周实际出现的逻辑日，避免 300 人场景把全部课程展开到
    # 所有周次/节次（会放大成数十万次无用的集合写入）。
    needed_weeks_by_weekday: dict[int, set[int]] = {}
    for logical in logical_by_column:
        if logical is None:
            continue
        logical_week, logical_weekday = logical
        needed_weeks_by_weekday.setdefault(logical_weekday, set()).add(logical_week)
    member_ids = {member.id for member in members}
    course_index: dict[tuple[int, int, int, int], set[str]] = {}
    for course in courses:
        if course.member_id not in member_ids:
            continue
        needed_weeks = needed_weeks_by_weekday.get(course.weekday, set())
        if not needed_weeks:
            continue
        active_weeks = needed_weeks.intersection(course.week_list)
        if not active_weeks:
            continue
        for week_number in active_weeks:
            for session in course.session_list:
                course_index.setdefault(
                    (course.member_id, week_number, course.weekday, session),
                    set()).add(course.course_name)
    special_index: dict[tuple[int, int, int, int], set[str]] = {}
    for arrangement in special_arrangements or []:
        needed_weeks = needed_weeks_by_weekday.get(arrangement.weekday, set())
        if (arrangement.member_id not in member_ids or not needed_weeks):
            continue
        active_weeks = needed_weeks.intersection(arrangement.week_list)
        if not active_weeks:
            continue
        label = arrangement.reason.strip() or "其他安排"
        for week_number in active_weeks:
            for session in arrangement.session_list:
                special_index.setdefault(
                    (arrangement.member_id, week_number,
                     arrangement.weekday, session), set()).add(label)

    free: list[list[bool]] = []
    busy_courses: dict[tuple[int, date, int], list[str]] = {}
    leave_info: dict[tuple[int, date], str] = {}
    duty_cells: set[tuple[int, date, int]] = set()
    special_info: dict[tuple[int, date, int], list[str]] = {}
    for row, member in enumerate(members):
        member_busy = busy.get(member.id, ())
        member_special = special_busy.get(member.id, ())
        row_free: list[bool] = []
        for column, logical in zip(columns, logical_by_column):
            if logical is None:
                row_free.append(False)
                continue
            logical_week, logical_weekday = logical
            if (member.id, logical_week, logical_weekday) in leave_map:
                row_free.append(False)
                leave_info[(row, column.date)] = leave_map[
                    (member.id, logical_week, logical_weekday)]
                continue
            if (member.id, logical_week, logical_weekday, column.block) in duty_set:
                row_free.append(False)
                duty_cells.add((row, column.date, column.block))
                continue
            names: set[str] = set()
            special_names: set[str] = set()
            is_free = True
            for session in BLOCK_SESSIONS[column.block]:
                if (logical_week, logical_weekday, session) in member_busy:
                    is_free = False
                    names |= course_index.get(
                        (member.id, logical_week, logical_weekday, session),
                        set())
                if (logical_week, logical_weekday, session) in member_special:
                    is_free = False
                    special_names |= special_index.get(
                        (member.id, logical_week, logical_weekday, session),
                        set())
            row_free.append(is_free)
            if names:
                busy_courses[(row, column.date, column.block)] = sorted(names)
            if special_names:
                special_info[(row, column.date, column.block)] = sorted(special_names)
        free.append(row_free)

    free_counts = [
        sum(1 for row in free if row[index])
        for index in range(len(columns))
    ]
    return CalendarAvailabilityMatrix(
        week=week,
        member_names=[member.name for member in members],
        columns=columns,
        free=free,
        busy_courses=busy_courses,
        free_counts=free_counts,
        leave_info=leave_info,
        duty_cells=duty_cells,
        special_info=special_info,
    )


def export_gantt_excel(matrix: AvailabilityMatrix) -> bytes:
    """导出带填充色的 Excel 甘特图，返回文件字节。"""
    if isinstance(matrix, CalendarAvailabilityMatrix):
        return _export_calendar_gantt_excel(matrix)
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    FREE_FILL = PatternFill("solid", fgColor="C9F2CF")
    BUSY_FILL = PatternFill("solid", fgColor="F2F2F7")
    ALL_FREE_FILL = PatternFill("solid", fgColor="FFD9A8")
    LEAVE_FILL = PatternFill("solid", fgColor="FFD6D2")
    OFF_FILL = PatternFill("solid", fgColor="FFD6D2")
    SPECIAL_FILL = PatternFill("solid", fgColor="E5D8FF")
    DUTY_FILL = PatternFill("solid", fgColor="B8D9FF")
    HEADER_FILL = PatternFill("solid", fgColor="007AFF")
    THIN = Side(style="thin", color="FFFFFF")
    border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"第{matrix.week}周空闲甘特图"

    columns = display_columns(matrix)
    summary_col = 2 + matrix.member_count

    # 横轴=成员，纵轴=日期x时段；汇总人数放最后一列。
    corner = ws.cell(row=1, column=1, value="日期 / 时段")
    corner.font = Font(bold=True, color="FFFFFF")
    corner.fill = HEADER_FILL
    corner.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    for index, name in enumerate(matrix.member_names):
        cell = ws.cell(row=1, column=2 + index, value=name)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(start_row=1, start_column=2 + index,
                       end_row=2, end_column=2 + index)
    summary_head = ws.cell(row=1, column=summary_col, value="空闲人数")
    summary_head.font = Font(bold=True, color="FFFFFF")
    summary_head.fill = HEADER_FILL
    summary_head.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=1, start_column=summary_col,
                   end_row=2, end_column=summary_col)

    for row_index, column in enumerate(columns):
        row = 3 + row_index
        if column.date is None:
            day_label = WEEKDAY_LABELS[column.weekday]
        else:
            day_label = (f"{column.date.month:02d}-{column.date.day:02d} "
                         f"{WEEKDAY_LABELS[column.date.isoweekday()]}")
        if column.is_off:
            day_label += "（放假）"
        block_label = BLOCK_LABELS[column.block].split(" ")[0]
        head = ws.cell(row=row, column=1, value=f"{day_label}\n{block_label}")
        head.font = Font(bold=True)
        head.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        head.border = border

        d, b = column.weekday, column.block
        for member_index in range(matrix.member_count):
            cell = ws.cell(row=row, column=2 + member_index)
            if column.is_off:
                cell.value = "放假"
                cell.fill = OFF_FILL
                cell.font = Font(bold=True, color="B3261E")
            elif matrix.leave_info.get((member_index, d)) is not None:
                leave_reason = matrix.leave_info[(member_index, d)]
                cell.value = f"请假：{leave_reason}" if leave_reason else "请假"
                cell.fill = LEAVE_FILL
                cell.font = Font(color="B3261E")
            elif (member_index, d, b) in matrix.duty_cells:
                cell.value = "值班"
                cell.fill = DUTY_FILL
                cell.font = Font(bold=True, color="0A5AA8")
            elif matrix.special_info.get((member_index, d, b)):
                special_names = matrix.special_info[(member_index, d, b)]
                cell.value = "其他安排：" + "、".join(special_names)
                cell.fill = SPECIAL_FILL
                cell.font = Font(color="6E3DC2")
            else:
                assert column.slot_index is not None
                is_free = matrix.free[member_index][column.slot_index]
                names = matrix.busy_courses.get((member_index, d, b), [])
                cell.value = "" if is_free else "、".join(names)
                cell.fill = FREE_FILL if is_free else BUSY_FILL
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)

        summary = ws.cell(row=row, column=summary_col)
        if column.is_off:
            summary.value = "放假"
            summary.fill = OFF_FILL
            summary.font = Font(bold=True, color="B3261E")
        else:
            assert column.slot_index is not None
            all_free = (matrix.member_count > 0
                        and matrix.free_counts[column.slot_index] == matrix.member_count)
            summary.value = (f"{matrix.free_counts[column.slot_index]}/"
                             f"{matrix.member_count}")
            if all_free:
                summary.fill = ALL_FREE_FILL
                summary.font = Font(bold=True, color="B25E00")
        summary.alignment = Alignment(horizontal="center", vertical="center")
        summary.border = border

    ws.freeze_panes = "B3"
    ws.column_dimensions["A"].width = 24
    for index in range(matrix.member_count):
        ws.column_dimensions[get_column_letter(2 + index)].width = 12
    ws.column_dimensions[get_column_letter(summary_col)].width = 12
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _export_calendar_gantt_excel(matrix: CalendarAvailabilityMatrix) -> bytes:
    """导出自然日期周版本的空闲甘特图（横轴=成员，纵轴=日期x时段）。"""
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    FREE_FILL = PatternFill("solid", fgColor="C9F2CF")
    BUSY_FILL = PatternFill("solid", fgColor="F2F2F7")
    ALL_FREE_FILL = PatternFill("solid", fgColor="FFD9A8")
    LEAVE_FILL = PatternFill("solid", fgColor="FFD6D2")
    OFF_FILL = PatternFill("solid", fgColor="FFD6D2")
    SPECIAL_FILL = PatternFill("solid", fgColor="E5D8FF")
    DUTY_FILL = PatternFill("solid", fgColor="B8D9FF")
    HEADER_FILL = PatternFill("solid", fgColor="007AFF")
    THIN = Side(style="thin", color="FFFFFF")
    border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"第{matrix.week}周空闲甘特图"
    summary_col = 2 + matrix.member_count

    # 横轴=成员，纵轴=自然日期x时段；汇总人数放最后一列。
    corner = ws.cell(row=1, column=1, value="日期 / 时段")
    corner.font = Font(bold=True, color="FFFFFF")
    corner.fill = HEADER_FILL
    corner.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    for index, name in enumerate(matrix.member_names):
        cell = ws.cell(row=1, column=2 + index, value=name)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(start_row=1, start_column=2 + index,
                       end_row=2, end_column=2 + index)
    summary_head = ws.cell(row=1, column=summary_col, value="空闲人数")
    summary_head.font = Font(bold=True, color="FFFFFF")
    summary_head.fill = HEADER_FILL
    summary_head.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=1, start_column=summary_col,
                   end_row=2, end_column=summary_col)

    for index, column in enumerate(matrix.columns):
        excel_row = 3 + index
        weekday = WEEKDAY_LABELS[column.date.isoweekday()]
        suffix = "（放假）" if column.is_off else ""
        block = BLOCK_LABELS[column.block].split(" ")[0]
        label = (f"{column.date.month:02d}-{column.date.day:02d}\n"
                 f"{weekday}{suffix} · {block}")
        head = ws.cell(row=excel_row, column=1, value=label)
        head.font = Font(bold=True)
        head.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        head.border = border

        for member_index in range(matrix.member_count):
            cell = ws.cell(row=excel_row, column=2 + member_index)
            if column.is_off:
                cell.value = "放假"
                cell.fill = OFF_FILL
                cell.font = Font(bold=True, color="B3261E")
            elif (member_index, column.date) in matrix.leave_info:
                reason = matrix.leave_info[(member_index, column.date)]
                cell.value = f"请假：{reason}" if reason else "请假"
                cell.fill = LEAVE_FILL
                cell.font = Font(color="B3261E")
            elif (member_index, column.date, column.block) in matrix.duty_cells:
                cell.value = "值班"
                cell.fill = DUTY_FILL
                cell.font = Font(bold=True, color="0A5AA8")
            elif matrix.special_info.get((member_index, column.date, column.block)):
                names = matrix.special_info[(member_index, column.date, column.block)]
                cell.value = "其他安排：" + "、".join(names)
                cell.fill = SPECIAL_FILL
                cell.font = Font(color="6E3DC2")
            else:
                is_free = matrix.free[member_index][index]
                names = matrix.busy_courses.get(
                    (member_index, column.date, column.block), [])
                cell.value = "" if is_free else "、".join(names)
                cell.fill = FREE_FILL if is_free else BUSY_FILL
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")

        summary = ws.cell(row=excel_row, column=summary_col)
        if column.is_off:
            summary.value = "放假"
            summary.fill = OFF_FILL
            summary.font = Font(bold=True, color="B3261E")
        else:
            count = matrix.free_counts[index]
            summary.value = f"{count}/{matrix.member_count}"
            if matrix.member_count > 0 and count == matrix.member_count:
                summary.fill = ALL_FREE_FILL
                summary.font = Font(bold=True, color="B25E00")
        summary.border = border
        summary.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "B3"
    ws.column_dimensions["A"].width = 24
    for index in range(matrix.member_count):
        ws.column_dimensions[get_column_letter(2 + index)].width = 12
    ws.column_dimensions[get_column_letter(summary_col)].width = 12
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
