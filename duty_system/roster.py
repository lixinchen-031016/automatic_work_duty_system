"""全媒体中心花名册解析与工作室分类规则。"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

UNKNOWN_STUDIO = "未指定工作室"
TARGET_AVAILABILITY_STUDIOS = ("短视频工作室", "图片工作室")

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
