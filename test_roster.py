"""花名册解析、工作室归类与完整名册导出。"""

from __future__ import annotations

import io
from pathlib import Path

import pytest

from duty_system.database import Database
from duty_system.exporter import build_roster_df, export_roster_excel
from duty_system.parser import Course, ParsedSchedule
from duty_system.roster import UNKNOWN_STUDIO, RosterEntry, parse_roster_file


def test_parse_roster_inherits_merged_studio_and_keeps_blank_student_id(
    tmp_path: Path,
) -> None:
    openpyxl = pytest.importorskip("openpyxl")
    path = tmp_path / "roster.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "全媒体中心花名册"
    ws.append(["工作室", "职位", "姓名", "学号", "电话", "学院+专业"])
    ws.append(["短视频工作室", "部长", "甲同学", 20250001, "13800000000", "计算机学院软件工程"])
    ws.merge_cells("A3:A4")
    ws.append([None, "成员", "乙同学", None, None, "设计学院视觉传达"])
    wb.save(path)

    result = parse_roster_file(path)

    assert [entry.name for entry in result.entries] == ["甲同学", "乙同学"]
    assert result.entries[0].studio == "短视频工作室"
    assert result.entries[1].studio == "短视频工作室"
    assert result.entries[0].student_id == "20250001"
    assert result.entries[1].student_id == ""
    assert result.warnings and "没有学号" in result.warnings[0]


def test_roster_classification_and_manual_transfer_are_persistent(
    tmp_path: Path,
) -> None:
    db = Database(tmp_path / "test.db")
    member_id = db.upsert_member(ParsedSchedule(
        name="甲同学", student_id="20250001", courses=[]))
    assert db.get_member(member_id).studio == UNKNOWN_STUDIO

    count, changed = db.replace_roster([
        RosterEntry(studio="短视频工作室", name="甲同学", student_id="20250001"),
        RosterEntry(studio="未在课表名单中", name="乙同学", student_id="20250002"),
    ])
    assert (count, changed) == (2, 1)
    assert db.get_member(member_id).studio == "短视频工作室"

    db.set_member_studio(member_id, "图片工作室")
    db.replace_roster([
        RosterEntry(studio="短视频工作室", name="甲同学", student_id="20250001"),
    ])
    member = db.get_member(member_id)
    assert member is not None
    assert member.studio == "图片工作室"
    assert member.studio_locked is True

    unknown_id = db.upsert_member(ParsedSchedule(
        name="未录入同学", student_id="20250003", courses=[]))
    assert db.get_member(unknown_id).studio == UNKNOWN_STUDIO


def test_complete_roster_export_merges_roster_and_uploaded_members(
    tmp_path: Path,
) -> None:
    db = Database(tmp_path / "test.db")
    db.replace_roster([
        RosterEntry(
            studio="短视频工作室", position="成员", name="已在名单",
            student_id="20250001", phone="13800000000",
            college_major="计算机学院软件工程"),
        # 同一学号在花名册中第二次出现可表示转组；完整导出只保留一条当前归属。
        RosterEntry(
            studio="图片工作室", position="成员", name="已在名单",
            student_id="20250001"),
    ])
    member_id = db.upsert_member(ParsedSchedule(
        name="已在名单", student_id="20250001", term="2026-2027-1",
        class_name="软件1班", major="软件工程", department="计算机学院",
        file_name="已在名单.xls",
        courses=[Course(
            course_name="数据库", weekday=1, weeks_text="1-16",
            week_list=list(range(1, 17)), sessions_text="01-02",
            session_list=[1, 2])],
    ))
    db.upsert_member(ParsedSchedule(
        name="未录入同学", student_id="20250002", term="2026-2027-1",
        class_name="设计1班", major="视觉传达", department="设计学院",
        file_name="未录入同学.xls", courses=[]))

    entries = db.list_roster_entries()
    members = db.list_members()
    frame = build_roster_df(entries, members)
    assert list(frame.columns) == [
        "工作室", "职位", "姓名", "学号", "电话", "学院+专业",
        "是否已上传课表", "课程数", "是否参与排班", "课表文件",
    ]
    assert set(frame["姓名"]) == {"已在名单", "未录入同学"}
    assert list(frame["姓名"]).count("已在名单") == 1
    uploaded = frame[frame["姓名"] == "已在名单"].iloc[0]
    assert uploaded["是否已上传课表"] == "是"
    assert uploaded["课程数"] == 1
    assert uploaded["工作室"] == "短视频工作室"
    unknown = frame[frame["姓名"] == "未录入同学"].iloc[0]
    assert unknown["工作室"] == UNKNOWN_STUDIO

    output = export_roster_excel(entries, members, db.get_courses())
    wb = pytest.importorskip("openpyxl").load_workbook(io.BytesIO(output))
    assert wb.sheetnames == ["完整花名册", "课表明细"]
    assert wb["课表明细"].max_row == 2
    assert wb["课表明细"].cell(2, 4).value == "数据库"
    assert db.get_member(member_id) is not None
