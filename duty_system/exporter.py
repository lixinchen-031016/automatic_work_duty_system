"""排班结果的结构化表格构建与导出（Excel / CSV）"""

from __future__ import annotations

import csv
import io
from collections import defaultdict
from datetime import date, timedelta

import pandas as pd

from .calendar import TermCalendar
from .database import Assignment, CourseRecord, Leave, Member
from .excel_layout import fit_excel_layout
from .parser import BLOCK_LABELS, WEEKDAY_LABELS
from .roster import (
    UNKNOWN_STUDIO,
    RosterEntry,
    normalize_name,
    normalize_student_id,
)


def week_date(
    start_date: date,
    week: int,
    weekday: int,
    calendar: TermCalendar | None = None,
) -> date | None:
    """学期第 week 周 weekday（1=周一..7=周日）对应的真实日期。"""
    if calendar is not None:
        return calendar.logical_to_date(week, weekday)
    return start_date + timedelta(days=(week - 1) * 7 + (weekday - 1))


def _date_label(
    start_date: date,
    week: int,
    weekday: int,
    calendar: TermCalendar | None = None,
) -> str:
    d = week_date(start_date, week, weekday, calendar)
    if d is None:
        return "—"
    return f"{d.month}月{d.day}日"


def _weekday_label(weekday: int, actual_date: date | None) -> str:
    if actual_date is None:
        return WEEKDAY_LABELS[weekday]
    return WEEKDAY_LABELS[actual_date.isoweekday()]


def build_detail_df(
    assignments: list[Assignment],
    per_slot: int = 1,
    start_date: date | None = None,
    calendar: TermCalendar | None = None,
) -> pd.DataFrame:
    """明细表：每行一个值班任务（周次 x 星期 x 时段），多人在岗合并显示"""
    grouped: dict[tuple[int, int, int], list[str]] = defaultdict(list)
    for a in assignments:
        grouped[(a.week, a.weekday, a.block)].append(a.member_name)

    rows = []
    for (week, weekday, block), names in sorted(grouped.items()):
        actual_date = calendar.logical_to_date(week, weekday) if calendar else None
        if calendar is not None and actual_date is None:
            continue
        rows.append({
            "周次": f"第{week}周",
            "星期": _weekday_label(weekday, actual_date),
            "日期": (_date_label(start_date, week, weekday, calendar)
                     if start_date else ""),
            "值班时段": BLOCK_LABELS[block],
            "值班人": "、".join(sorted(names)),
            "人数": len(names),
        })
    columns = ["周次", "星期"] + (["日期"] if start_date else []) + ["值班时段", "值班人", "人数"]
    return pd.DataFrame(rows, columns=columns)


def build_pivot_df(
    assignments: list[Assignment],
    start_date: date | None = None,
    calendar: TermCalendar | None = None,
) -> pd.DataFrame:
    """透视表：行=周次x星期，列=值班时段（多人用顿号连接），适合界面展示与打印"""
    grouped: dict[tuple[int, int, int], list[str]] = defaultdict(list)
    for a in assignments:
        grouped[(a.week, a.weekday, a.block)].append(a.member_name)

    blocks = sorted({b for (_, _, b) in grouped})
    weeks = sorted({w for (w, _, _) in grouped})
    weekdays = sorted({d for (_, d, _) in grouped})

    index: list[tuple[int, int]] = []
    actual_dates: dict[tuple[int, int], date] = {}
    for w in weeks:
        for d in weekdays:
            actual = calendar.logical_to_date(w, d) if calendar else None
            if calendar is not None and actual is None:
                continue
            index.append((w, d))
            if actual is not None:
                actual_dates[(w, d)] = actual
    data = {}
    for b in blocks:
        col = []
        for (w, d) in index:
            names = grouped.get((w, d, b))
            col.append("、".join(sorted(names)) if names else "—")
        data[BLOCK_LABELS[b].split(" ")[0]] = col  # 列名如 "1-2节"

    df = pd.DataFrame(data, index=pd.MultiIndex.from_tuples(
        index, names=["周次", "星期"]))
    if start_date:
        df.insert(0, "日期", [_date_label(start_date, w, d, calendar) for (w, d) in index])
    df.index = df.index.map(
        lambda t: (f"第{t[0]}周", _weekday_label(t[1], actual_dates.get(t))))
    return df


def build_gap_df(diagnoses: list) -> pd.DataFrame:
    """无人可值时段表：周次/星期/时段 + 成因与无课人数（诊断结果的展示层）

    diagnoses 为 scheduler.GapDiagnosis 列表；只用属性访问，
    避免 exporter 依赖 scheduler 造成循环导入。
    """
    rows = [{
        "周次": f"第{d.week}周",
        "星期": WEEKDAY_LABELS[d.weekday],
        "时段": BLOCK_LABELS[d.block],
        "主要原因": d.cause,
        "无课人数": d.free_members,
    } for d in diagnoses]
    return pd.DataFrame(rows, columns=["周次", "星期", "时段", "主要原因", "无课人数"])



def _member_studio_text(member: Member) -> str:
    studios = getattr(member, "studios", None) or []
    positions = getattr(member, "studio_positions", {}) or {}
    labels = [
        f"{studio}（{positions[studio]}）" if positions.get(studio) else studio
        for studio in studios
    ]
    return "、".join(labels) or member.studio or UNKNOWN_STUDIO


def _member_position_text(member: Member) -> str:
    studios = getattr(member, "studios", None) or []
    positions = getattr(member, "studio_positions", {}) or {}
    values = [
        (studio, positions.get(studio, "")) for studio in studios
    ]
    if len(values) == 1:
        return values[0][1]
    return "；".join(
        f"{studio}：{position or '未指定'}" for studio, position in values
    )


def _member_match_index(
    members: list[Member],
) -> tuple[dict[tuple[str, str], Member], dict[str, list[Member]], dict[str, list[Member]]]:
    by_pair: dict[tuple[str, str], Member] = {}
    by_student: dict[str, list[Member]] = defaultdict(list)
    by_name: dict[str, list[Member]] = defaultdict(list)
    for member in members:
        student_id = (member.student_id or "").strip()
        name = member.name.strip()
        by_pair[(student_id, name)] = member
        if student_id:
            by_student[student_id].append(member)
        if name:
            by_name[name].append(member)
    return by_pair, by_student, by_name


def _match_member(
    entry: RosterEntry,
    indexes: tuple[dict[tuple[str, str], Member], dict[str, list[Member]], dict[str, list[Member]]],
) -> Member | None:
    by_pair, by_student, by_name = indexes
    student_id = (entry.student_id or "").strip()
    name = entry.name.strip()
    member = by_pair.get((student_id, name))
    if member is not None:
        return member
    if student_id:
        options = by_student.get(student_id, [])
        if len(options) == 1:
            return options[0]
    options = by_name.get(name, [])
    return options[0] if len(options) == 1 else None


def build_roster_df(
    entries: list[RosterEntry],
    members: list[Member],
) -> pd.DataFrame:
    """合并花名册原始字段与课表上传/工作室分配状态。

    已上传成员优先使用数据库中的工作室；这样手工修改会体现在导出中。
    花名册中一行都未匹配到的上传成员会追加到末尾，避免遗漏“未指定工作室”。
    """
    indexes = _member_match_index(members)
    rows: list[dict] = []
    matched_member_ids: set[int] = set()
    seen_identities: set[tuple[str, str]] = set()
    for entry in entries:
        identity = (
            "id", normalize_student_id(entry.student_id)
        ) if normalize_student_id(entry.student_id) else (
            "name", normalize_name(entry.name)
        )
        if identity in seen_identities:
            continue
        seen_identities.add(identity)
        member = _match_member(entry, indexes)
        if member is not None:
            matched_member_ids.add(member.id)
        studio = (
            _member_studio_text(member) if member is not None
            else entry.studio or UNKNOWN_STUDIO
        )
        rows.append({
            "工作室": studio,
            "职位": (
                _member_position_text(member) if member is not None
                else entry.position
            ),
            "姓名": member.name if member is not None else entry.name,
            "学号": member.student_id if member is not None else entry.student_id,
            "电话": entry.phone,
            "学院+专业": (
                entry.college_major
                or (f"{member.department}{member.major}" if member is not None else "")
            ),
            "是否已上传课表": "是" if member is not None and member.course_count else "否",
            "课程数": member.course_count if member is not None else 0,
            "是否参与排班": (
                "是" if member is not None and member.participates_in_scheduling
                else "否" if member is not None else ""
            ),
            "课表文件": member.file_name if member is not None else "",
        })

    for member in members:
        if member.id in matched_member_ids:
            continue
        rows.append({
            "工作室": _member_studio_text(member),
            "职位": "",
            "姓名": member.name,
            "学号": member.student_id,
            "电话": "",
            "学院+专业": f"{member.department}{member.major}".strip(),
            "是否已上传课表": "是" if member.course_count else "否",
            "课程数": member.course_count,
            "是否参与排班": "是" if member.participates_in_scheduling else "否",
            "课表文件": member.file_name,
        })
    columns = [
        "工作室", "职位", "姓名", "学号", "电话", "学院+专业",
        "是否已上传课表", "课程数", "是否参与排班", "课表文件",
    ]
    return pd.DataFrame(rows, columns=columns)


def build_roster_course_df(
    members: list[Member],
    courses: list[CourseRecord],
) -> pd.DataFrame:
    """完整课表明细表：按成员展开课程名、周次、节次与地点。"""
    info = {member.id: member for member in members}
    rows: list[dict] = []
    for course in courses:
        member = info.get(course.member_id)
        if member is None:
            continue
        rows.append({
            "工作室": _member_studio_text(member),
            "学号": member.student_id,
            "姓名": member.name,
            "课程名称": course.course_name,
            "教师": course.teacher,
            "星期": WEEKDAY_LABELS.get(course.weekday, "整周"),
            "周次": course.weeks_text,
            "节次": course.sessions_text,
            "地点": course.location,
        })
    columns = [
        "工作室", "学号", "姓名", "课程名称", "教师", "星期",
        "周次", "节次", "地点",
    ]
    return pd.DataFrame(rows, columns=columns)


def _roster_format_groups(
    entries: list[RosterEntry],
    members: list[Member],
) -> list[tuple[str, list[dict]]]:
    """按上传花名册顺序与工作室分组构建六列导出行。"""
    indexes = _member_match_index(members)
    groups: dict[str, list[dict]] = {}
    studio_order: list[str] = []
    seen: set[tuple] = set()
    metadata: dict[tuple[int, str], RosterEntry] = {}

    def ensure_studio(studio: str) -> None:
        if studio not in groups:
            groups[studio] = []
            studio_order.append(studio)

    for entry in entries:
        studio = entry.studio or UNKNOWN_STUDIO
        ensure_studio(studio)
        member = _match_member(entry, indexes)
        if member is None:
            identity = (
                "id", normalize_student_id(entry.student_id)
            ) if normalize_student_id(entry.student_id) else (
                "name", normalize_name(entry.name)
            )
            key = (studio.casefold(), identity)
            if key in seen:
                continue
            seen.add(key)
            groups[studio].append({
                "studio": studio,
                "position": entry.position,
                "name": entry.name,
                "student_id": entry.student_id,
                "phone": entry.phone,
                "college_major": entry.college_major,
            })
            continue

        member_studios = getattr(member, "studios", None) or []
        current_studio = next(
            (value for value in member_studios
             if value.casefold() == studio.casefold()),
            None,
        )
        if current_studio is None:
            continue
        key = (member.id, studio.casefold())
        if key in seen:
            continue
        seen.add(key)
        metadata[(member.id, studio.casefold())] = entry
        positions = getattr(member, "studio_positions", {}) or {}
        groups[studio].append({
            "studio": current_studio,
            "position": positions.get(current_studio, ""),
            "name": member.name,
            "student_id": member.student_id,
            "phone": entry.phone,
            "college_major": (
                entry.college_major
                or f"{member.department}{member.major}".strip()
            ),
        })

    # 名单外上传成员或后续手工新增的工作室归属追加到对应分组末尾。
    for member in members:
        positions = getattr(member, "studio_positions", {}) or {}
        for studio in getattr(member, "studios", None) or []:
            key = (member.id, studio.casefold())
            if key in seen:
                continue
            seen.add(key)
            ensure_studio(studio)
            roster_entry = metadata.get(key)
            groups[studio].append({
                "studio": studio,
                "position": positions.get(studio, ""),
                "name": member.name,
                "student_id": member.student_id,
                "phone": roster_entry.phone if roster_entry else "",
                "college_major": (
                    roster_entry.college_major if roster_entry
                    else f"{member.department}{member.major}".strip()
                ),
            })

    return [
        (studio, groups[studio])
        for studio in studio_order
        if groups.get(studio)
    ]

def _write_uploaded_roster_sheet(
    ws,
    entries: list[RosterEntry],
    members: list[Member],
) -> None:
    """把主表写成与上传花名册一致的结构与视觉版式。"""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    title = "全媒体中心花名册"
    headers = ["工作室", "职位", "姓名", "学号", "电话", "学院+专业"]
    ws.merge_cells("A1:F1")
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=16)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    header_fill = PatternFill("solid", fgColor="4472C4")
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for column, header in enumerate(headers, 1):
        cell = ws.cell(2, column, header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    row_index = 3
    for studio, rows in _roster_format_groups(entries, members):
        studio_start = row_index
        position_start = row_index
        previous_position = None
        for item in rows:
            values = [
                item["studio"], item["position"], item["name"],
                item["student_id"], item["phone"], item["college_major"],
            ]
            for column, value in enumerate(values, 1):
                cell = ws.cell(row_index, column, value)
                cell.border = border
                cell.alignment = Alignment(
                    horizontal="center" if column in (1, 2, 4) else "left",
                    vertical="center",
                    wrap_text=column == 6,
                )
            position = item["position"]
            if previous_position is not None and position != previous_position:
                if row_index - 1 > position_start:
                    ws.merge_cells(
                        start_row=position_start, start_column=2,
                        end_row=row_index - 1, end_column=2)
                position_start = row_index
            previous_position = position
            row_index += 1

        if row_index - 1 > studio_start:
            ws.merge_cells(
                start_row=studio_start, start_column=1,
                end_row=row_index - 1, end_column=1)
        if row_index - 1 > position_start:
            ws.merge_cells(
                start_row=position_start, start_column=2,
                end_row=row_index - 1, end_column=2)

    for column, width in enumerate([18, 12, 12, 16, 16, 34], 1):
        ws.column_dimensions[get_column_letter(column)].width = width
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 24
    ws.freeze_panes = "A3"


def _write_dataframe_sheet(ws, frame: pd.DataFrame, title: str) -> None:
    """把 DataFrame 写入工作表并应用统一表头与自适应布局。"""
    from openpyxl.styles import Font, PatternFill

    ws.title = title
    ws.append([str(column) for column in frame.columns])
    for row in frame.itertuples(index=False, name=None):
        ws.append(list(row))
    fit_excel_layout(ws, header_rows=1, min_row_height=24.0)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
    ws.freeze_panes = "A2"


def export_roster_excel(
    entries: list[RosterEntry],
    members: list[Member],
    courses: list[CourseRecord],
) -> bytes:
    """导出上传花名册同版式主表、课表状态和课程明细。"""
    from openpyxl import Workbook

    workbook = Workbook()
    roster_sheet = workbook.active
    roster_sheet.title = "完整花名册"
    _write_uploaded_roster_sheet(roster_sheet, entries, members)

    _write_dataframe_sheet(
        workbook.create_sheet(),
        build_roster_df(entries, members),
        "课表状态",
    )
    _write_dataframe_sheet(
        workbook.create_sheet(),
        build_roster_course_df(members, courses),
        "课表明细",
    )

    buf = io.BytesIO()
    workbook.save(buf)
    return buf.getvalue()

def build_stats_df(member_stats: dict[int, dict]) -> pd.DataFrame:
    """值班统计表：每人总次数与值班周分布"""
    rows = [{
        "成员": s["name"],
        "总值班次数": s["total"],
        "值班周次": "、".join(f"第{w}周" for w in s["weeks"]) or "—",
    } for s in member_stats.values()]
    return pd.DataFrame(rows, columns=["成员", "总值班次数", "值班周次"]).sort_values(
        ["总值班次数", "成员"], ascending=[False, True]).reset_index(drop=True)


def export_excel(
    assignments: list[Assignment],
    member_stats: dict[int, dict],
    gaps: list[tuple[int, int, int]],
    start_date: date | None = None,
    calendar: TermCalendar | None = None,
) -> bytes:
    """导出 Excel：排班总表(透视) + 值班明细 + 值班统计，返回文件字节"""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pivot = build_pivot_df(assignments, start_date=start_date, calendar=calendar)
        pivot.to_excel(writer, sheet_name="排班总表", merge_cells=False)

        detail = build_detail_df(assignments, start_date=start_date, calendar=calendar)
        detail.to_excel(writer, sheet_name="值班明细", index=False)

        stats = build_stats_df(member_stats)
        stats.to_excel(writer, sheet_name="值班统计", index=False)

        if gaps:
            gap_df = pd.DataFrame([{
                "周次": f"第{w}周", "星期": WEEKDAY_LABELS[d], "时段": BLOCK_LABELS[b],
            } for w, d, b in gaps])
            gap_df.to_excel(writer, sheet_name="无人可用时段", index=False)

        _beautify(writer)
    return buf.getvalue()


def _beautify(writer: pd.ExcelWriter) -> None:
    """美化表格：内容自适应列宽/行高，表头加粗并冻结。"""
    from openpyxl.styles import Font, PatternFill

    for ws in writer.book.worksheets:
        fit_excel_layout(ws, header_rows=1, min_row_height=24.0)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="4472C4")
        ws.freeze_panes = "A2"


def export_csv(
    assignments: list[Assignment],
    start_date: date | None = None,
    calendar: TermCalendar | None = None,
) -> bytes:
    """导出 CSV（UTF-8 BOM、CRLF、全字段引号，完整保留日期与姓名文本）。

    CSV 格式本身不保存列宽/行高；需要自动适配显示时请使用 Excel 导出。
    """
    return build_detail_df(
        assignments, start_date=start_date, calendar=calendar
    ).to_csv(
        index=False,
        quoting=csv.QUOTE_ALL,
        lineterminator="\r\n",
    ).encode("utf-8-sig")


def build_leaves_df(
    leaves: list[Leave],
    members: list[Member],
    start_date: date | None = None,
    calendar: TermCalendar | None = None,
) -> pd.DataFrame:
    """请假记录表：每行一条，含成员、周次/星期/日期与原因"""
    info_of = {m.id: m for m in members}
    rows = []
    for l in sorted(leaves, key=lambda x: (x.week, x.weekday)):
        actual_date = calendar.logical_to_date(l.week, l.weekday) if calendar else None
        if calendar is not None and actual_date is None:
            continue
        m = info_of.get(l.member_id)
        rows.append({
            "成员": m.name if m else "已删除成员",
            "学号": m.student_id if m else "",
            "周次": f"第{l.week}周",
            "星期": _weekday_label(l.weekday, actual_date),
            "日期": (_date_label(start_date, l.week, l.weekday, calendar)
                     if start_date else ""),
            "原因": l.reason,
        })
    columns = ["成员", "学号", "周次", "星期"] + (["日期"] if start_date else []) + ["原因"]
    return pd.DataFrame(rows, columns=columns)


def export_leaves_excel(
    leaves: list[Leave],
    members: list[Member],
    start_date: date | None = None,
    calendar: TermCalendar | None = None,
) -> bytes:
    """导出请假记录 Excel（单表，供请假登记处一键导出）"""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        build_leaves_df(
            leaves, members, start_date=start_date, calendar=calendar
        ).to_excel(
            writer, sheet_name="请假记录", index=False)
        _beautify(writer)
    return buf.getvalue()
