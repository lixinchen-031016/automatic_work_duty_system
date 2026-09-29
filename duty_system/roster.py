"""全媒体中心花名册解析与工作室分类规则。"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

UNKNOWN_STUDIO = "未指定工作室"
POSITION_LEADER = "部长"
POSITION_DEPUTY = "副部长"
POSITION_MEMBER = "成员"
POSITION_SENIOR_ADVISOR = "高级顾问"
POSITION_OPTIONS = (
    "", POSITION_LEADER, POSITION_DEPUTY, POSITION_MEMBER,
    POSITION_SENIOR_ADVISOR,
)


def position_options_for_studio(studio: str) -> tuple[str, ...]:
    """运营工作室支持高级顾问；其他工作室使用标准职位集合。"""
    if (studio or "").strip() == "运营工作室":
        return POSITION_OPTIONS
    return tuple(
        position for position in POSITION_OPTIONS
        if position != POSITION_SENIOR_ADVISOR
    )
# 空闲时段总览的业务分组顺序与界面下拉框保持一致。
AVAILABILITY_BUSINESS_GROUPS = (
    ("微信", ("微信工作室",)),
    ("博q", ("博Q工作室",)),
    ("设计", ("设计工作室",)),
    ("图片/短视频", ("图片工作室", "短视频工作室")),
)
TARGET_AVAILABILITY_STUDIOS = AVAILABILITY_BUSINESS_GROUPS[-1][1]

_STUDIO_ALIASES = {"工作室", "部门", "所属工作室", "所属部门"}
_POSITION_ALIASES = {"职位", "职务", "岗位"}
_NAME_ALIASES = {"姓名", "名字"}
_STUDENT_ID_ALIASES = {"学号", "学生学号", "学生编号"}
_PHONE_ALIASES = {"电话", "手机", "手机号", "联系电话"}
_COLLEGE_MAJOR_ALIASES = {
    "学院+专业", "学院专业", "学院及专业", "学院/专业", "院系专业",
}


def normalize_text(value: Any) -> str:
    """把表格单元格转成干净文本，同时处理 Excel 产生的 .0 数值。"""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value)
    text = text.replace("\u200b", "").replace("\ufeff", "").replace("\xa0", " ")
    text = text.replace("０", "0").replace("１", "1").replace("２", "2")
    text = text.replace("３", "3").replace("４", "4").replace("５", "5")
    text = text.replace("６", "6").replace("７", "7").replace("８", "8")
    text = text.replace("９", "9")
    return re.sub(r"[ \t\r\n]+", " ", text).strip()


def normalize_header(value: Any) -> str:
    return re.sub(r"\s+", "", normalize_text(value)).replace("＋", "+")


def normalize_student_id(value: Any) -> str:
    text = normalize_text(value)
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    return text


def normalize_name(value: Any) -> str:
    return re.sub(r"\s+", "", normalize_text(value))


@dataclass(frozen=True)
class RosterEntry:
    """花名册中的一行；允许同一人同时出现在多个工作室。"""

    studio: str
    name: str
    student_id: str = ""
    position: str = ""
    phone: str = ""
    college_major: str = ""
    source_file: str = ""
    row_number: int = 0


@dataclass
class RosterImportResult:
    entries: list[RosterEntry] = field(default_factory=list)
    source_file: str = ""
    warnings: list[str] = field(default_factory=list)


def _read_xlsx_rows(path: Path) -> list[list[Any]]:
    from openpyxl import load_workbook

    workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        sheet = workbook.worksheets[0]
        return [list(row) for row in sheet.iter_rows(values_only=True)]
    finally:
        workbook.close()


def _read_xls_rows(path: Path) -> list[list[Any]]:
    import xlrd

    workbook = xlrd.open_workbook(str(path), on_demand=True)
    try:
        sheet = workbook.sheet_by_index(0)
        return [
            [sheet.cell_value(row, col) for col in range(sheet.ncols)]
            for row in range(sheet.nrows)
        ]
    finally:
        workbook.release_resources()


def _read_rows(path: Path) -> list[list[Any]]:
    # 教务与办公软件常把 xlsx 改名为 xls；按文件头嗅探比扩展名可靠。
    if zipfile.is_zipfile(path):
        return _read_xlsx_rows(path)
    return _read_xls_rows(path)


def _column_index(headers: list[str], aliases: set[str]) -> int | None:
    for index, header in enumerate(headers):
        if header in aliases:
            return index
    for index, header in enumerate(headers):
        if header and any(alias in header for alias in aliases):
            return index
    return None


def _find_header(rows: list[list[Any]]) -> tuple[int, dict[str, int]]:
    for row_index, row in enumerate(rows[:50]):
        headers = [normalize_header(value) for value in row]
        if (
            _column_index(headers, _NAME_ALIASES) is None
            or _column_index(headers, _STUDENT_ID_ALIASES) is None
        ):
            continue
        mapping: dict[str, int] = {}
        for key, aliases in (
            ("studio", _STUDIO_ALIASES),
            ("position", _POSITION_ALIASES),
            ("name", _NAME_ALIASES),
            ("student_id", _STUDENT_ID_ALIASES),
            ("phone", _PHONE_ALIASES),
            ("college_major", _COLLEGE_MAJOR_ALIASES),
        ):
            index = _column_index(headers, aliases)
            if index is not None:
                mapping[key] = index
        if {"name", "student_id"}.issubset(mapping):
            return row_index, mapping
    raise ValueError("未找到包含“姓名”和“学号”的表头，请检查花名册格式")


def _cell(row: list[Any], index: int | None) -> str:
    if index is None or index >= len(row):
        return ""
    return normalize_text(row[index])


def parse_roster_file(path: str | Path) -> RosterImportResult:
    """解析花名册首表，向下继承合并单元格中的工作室/职位。"""
    source = Path(path)
    rows = _read_rows(source)
    header_index, columns = _find_header(rows)
    entries: list[RosterEntry] = []
    current_studio = ""
    current_position = ""
    missing_student_ids = 0

    for row_index in range(header_index + 1, len(rows)):
        row = rows[row_index]
        name = _cell(row, columns.get("name"))
        student_id = normalize_student_id(
            _cell(row, columns.get("student_id")))
        studio = _cell(row, columns.get("studio"))
        position = _cell(row, columns.get("position"))
        if studio:
            current_studio = studio
        if position:
            current_position = position
        if not name or name in _NAME_ALIASES:
            continue
        if not student_id:
            missing_student_ids += 1
        entries.append(RosterEntry(
            studio=current_studio or UNKNOWN_STUDIO,
            position=current_position,
            name=name,
            student_id=student_id,
            phone=_cell(row, columns.get("phone")),
            college_major=_cell(row, columns.get("college_major")),
            source_file=source.name,
            row_number=row_index + 1,
        ))

    warnings: list[str] = []
    if missing_student_ids:
        warnings.append(f"{missing_student_ids} 行没有学号，将使用姓名进行匹配")
    if not entries:
        raise ValueError("花名册中没有可导入的成员记录")
    return RosterImportResult(
        entries=entries, source_file=source.name, warnings=warnings)


@dataclass
class RosterImportRecord:
    """一次导入中某位成员在“当前/导入”版本之间的比较结果。"""

    key: tuple[str, ...]
    member_id: int | None
    display_name: str
    kind: str
    changed: bool
    default_incoming: bool
    current_text: str
    incoming_text: str
    current_entries: list[RosterEntry] = field(default_factory=list)
    incoming_entries: list[RosterEntry] = field(default_factory=list)
    current_fields: dict[str, str] = field(default_factory=dict)
    incoming_fields: dict[str, str] = field(default_factory=dict)
    current_studios: dict[str, str] = field(default_factory=dict)
    incoming_studios: dict[str, str] = field(default_factory=dict)


@dataclass
class RosterMemberUpdate:
    """冲突确认后需要显式写回成员表的内容。"""

    member_id: int
    name: str
    student_id: str
    phone: str
    studios: list[str]
    positions: dict[str, str]


def _entry_identity_key(entry: RosterEntry) -> tuple[str, ...]:
    student_id = normalize_student_id(entry.student_id)
    if student_id:
        return ("id", student_id)
    return ("name", normalize_name(entry.name))


def _member_identity_key(member: object) -> tuple[str, ...]:
    return ("member", str(member.id))


def _member_match_indexes(members: list[object]) -> tuple[dict, dict]:
    by_pair: dict[tuple[str, str], object] = {}
    by_student: dict[str, list[object]] = {}
    by_name: dict[str, list[object]] = {}
    for member in members:
        name = normalize_name(getattr(member, "name", ""))
        student_id = normalize_student_id(getattr(member, "student_id", ""))
        by_pair[(student_id, name)] = member
        if student_id:
            by_student.setdefault(student_id, []).append(member)
        if name:
            by_name.setdefault(name, []).append(member)
    return by_pair, by_student, by_name


def _entry_member(entry: RosterEntry, indexes: tuple[dict, dict]) -> object | None:
    by_pair, by_student, by_name = indexes
    student_id = normalize_student_id(entry.student_id)
    name = normalize_name(entry.name)
    member = by_pair.get((student_id, name))
    if member is not None:
        return member
    if student_id:
        options = by_student.get(student_id, [])
        if len(options) == 1:
            return options[0]
    options = by_name.get(name, [])
    return options[0] if len(options) == 1 else None


def _entry_studios(entries: list[RosterEntry]) -> dict[str, str]:
    result: dict[str, str] = {}
    for entry in entries:
        studio = (entry.studio or UNKNOWN_STUDIO).strip() or UNKNOWN_STUDIO
        result.setdefault(studio, (entry.position or "").strip())
    return _strip_unknown_studio(result)


def _member_studios(member: object) -> dict[str, str]:
    studios = list(getattr(member, "studios", None) or [])
    positions = dict(getattr(member, "studio_positions", None) or {})
    result = {studio: positions.get(studio, "") for studio in studios}
    return _strip_unknown_studio(result)


def _strip_unknown_studio(studios: dict[str, str]) -> dict[str, str]:
    if len(studios) <= 1:
        return studios
    return {
        studio: position for studio, position in studios.items()
        if studio != UNKNOWN_STUDIO
    }


def _profile_fields(
    entries: list[RosterEntry],
    member: object | None,
) -> dict[str, str]:
    first = entries[0] if entries else None
    if member is not None:
        college = (
            f"{getattr(member, 'department', '')}{getattr(member, 'major', '')}".strip()
            or (first.college_major if first else "")
        )
        return {
            "name": getattr(member, "name", "") or "",
            "student_id": getattr(member, "student_id", "") or "",
            "phone": getattr(member, "phone", "") or "",
            "college_major": college,
        }
    return {
        "name": first.name if first else "",
        "student_id": first.student_id if first else "",
        "phone": first.phone if first else "",
        "college_major": first.college_major if first else "",
    }


def _profile_text(fields: dict[str, str], studios: dict[str, str]) -> str:
    lines = [
        f"姓名：{fields.get('name') or '（空）'}",
        f"学号：{fields.get('student_id') or '（空）'}",
        f"手机号：{fields.get('phone') or '（空）'}",
        f"学院+专业：{fields.get('college_major') or '（空）'}",
        "工作室与职位：",
    ]
    if studios:
        lines.extend(
            f"  {studio}：{position or '未指定'}"
            for studio, position in studios.items()
        )
    else:
        lines.append("  （无）")
    return "\n".join(lines)


def build_roster_import_preview(
    current_entries: list[RosterEntry],
    incoming_entries: list[RosterEntry],
    members: list[object],
) -> list[RosterImportRecord]:
    """构建花名册导入预览；不修改数据库。"""
    indexes = _member_match_indexes(members)
    groups: dict[tuple[str, ...], dict] = {}

    def group_for(
        key: tuple[str, ...],
        *,
        display_name: str,
        member: object | None,
    ) -> dict:
        group = groups.setdefault(key, {
            "display_name": display_name,
            "member": member,
            "current": [],
            "incoming": [],
        })
        if member is not None:
            group["member"] = member
            group["display_name"] = getattr(member, "name", "") or display_name
        return group

    for entry in current_entries:
        member = _entry_member(entry, indexes)
        key = _member_identity_key(member) if member is not None else _entry_identity_key(entry)
        group_for(key, display_name=entry.name, member=member)["current"].append(entry)
    for entry in incoming_entries:
        member = _entry_member(entry, indexes)
        key = _member_identity_key(member) if member is not None else _entry_identity_key(entry)
        group_for(key, display_name=entry.name, member=member)["incoming"].append(entry)

    records: list[RosterImportRecord] = []
    for key, group in groups.items():
        member = group["member"]
        current_entries = group["current"]
        incoming_entries = group["incoming"]
        current_fields = _profile_fields(current_entries, member)
        incoming_fields = _profile_fields(incoming_entries, member if not incoming_entries else None)
        current_studios = (
            _member_studios(member) if member is not None
            else _entry_studios(current_entries)
        )
        incoming_studios = _entry_studios(incoming_entries)
        has_current = bool(current_entries) or (
            member is not None
            and any(studio != UNKNOWN_STUDIO for studio in current_studios)
        )
        has_incoming = bool(incoming_entries)

        changed = (
            not has_current
            or not has_incoming
            or current_fields != incoming_fields
            or current_studios != incoming_studios
        )
        if not has_current:
            kind = "新增"
            default_incoming = True
        elif not has_incoming:
            kind = "仅当前"
            default_incoming = False
        else:
            kind = "冲突"
            default_incoming = False
        records.append(RosterImportRecord(
            key=key,
            member_id=getattr(member, "id", None) if member is not None else None,
            display_name=group["display_name"],
            kind=kind,
            changed=changed,
            default_incoming=default_incoming,
            current_text=(
                _profile_text(current_fields, current_studios)
                if has_current else "（当前花名册中不存在）"
            ),
            incoming_text=(
                _profile_text(incoming_fields, incoming_studios)
                if has_incoming else "（本次导入文件中不存在）"
            ),
            current_entries=current_entries,
            incoming_entries=incoming_entries,
            current_fields=current_fields,
            incoming_fields=incoming_fields,
            current_studios=current_studios,
            incoming_studios=incoming_studios,
        ))
    return records


def merge_roster_import(
    records: list[RosterImportRecord],
    decisions: dict[tuple[str, ...], bool],
) -> tuple[list[RosterEntry], list[RosterMemberUpdate]]:
    """按冲突选择生成新的花名册记录和成员更新。"""
    merged: list[RosterEntry] = []
    updates: list[RosterMemberUpdate] = []

    for record in records:
        use_incoming = decisions.get(record.key, record.default_incoming)
        if use_incoming:
            if not record.incoming_entries:
                continue
            source_entries = record.incoming_entries
            fields = record.incoming_fields
            studios = record.incoming_studios
        elif record.current_entries or record.current_studios:
            source_entries = record.current_entries
            fields = record.current_fields
            studios = record.current_studios
        else:
            continue

        source_by_studio = {
            (entry.studio or UNKNOWN_STUDIO).casefold(): entry
            for entry in source_entries
        }
        for index, (studio, position) in enumerate(studios.items()):
            source = source_by_studio.get(studio.casefold())
            if source is None and source_entries:
                source = source_entries[min(index, len(source_entries) - 1)]
            merged.append(RosterEntry(
                studio=studio,
                position=position,
                name=fields.get("name", ""),
                student_id=fields.get("student_id", ""),
                phone=fields.get("phone", ""),
                college_major=(
                    source.college_major if source is not None
                    else fields.get("college_major", "")
                ),
                source_file=source.source_file if source is not None else "",
                row_number=source.row_number if source is not None else 0,
            ))

        if record.member_id is not None and record.changed:
            updates.append(RosterMemberUpdate(
                member_id=record.member_id,
                name=fields.get("name", ""),
                student_id=fields.get("student_id", ""),
                phone=fields.get("phone", ""),
                studios=list(studios),
                positions=dict(studios),
            ))
    return merged, updates
