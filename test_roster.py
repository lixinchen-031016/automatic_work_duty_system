"""花名册解析、工作室归类与完整名册导出。"""

from __future__ import annotations

import io
import sqlite3
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
        RosterEntry(studio="短视频工作室", position="部长",
                    name="甲同学", student_id="20250001",
                    phone="13800000000"),
        RosterEntry(studio="图片工作室", position="副部长",
                    name="甲同学", student_id="20250001"),
        RosterEntry(studio="未在课表名单中", name="乙同学", student_id="20250002"),
    ])
    assert (count, changed) == (3, 1)
    assert db.get_member(member_id).studios == [
        "短视频工作室", "图片工作室"
    ]
    assert db.get_member(member_id).studio_positions == {
        "短视频工作室": "部长", "图片工作室": "副部长"}
    assert db.get_member(member_id).phone == "13800000000"

    db.set_member_studios(
        member_id, ["设计工作室", "微信工作室"],
        {"设计工作室": "副部长", "微信工作室": "成员"})
    db.replace_roster([
        RosterEntry(studio="短视频工作室", name="甲同学", student_id="20250001"),
    ])
    member = db.get_member(member_id)
    assert member is not None
    assert member.studio == "设计工作室"
    assert member.studios == ["设计工作室", "微信工作室"]
    assert member.studio_positions == {
        "设计工作室": "副部长", "微信工作室": "成员"}
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
            studio="短视频工作室", position="部长", name="已在名单",
            student_id="20250001", phone="13800000000",
            college_major="计算机学院软件工程"),
        # 同一学号在花名册中第二次出现表示兼任；完整导出合并全部工作室。
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
    db.update_member_profile(
        member_id, name="已在名单", student_id="20250001",
        term="2026-2027-1", class_name="软件1班", major="软件工程",
        department="计算机学院", phone="13900139000")
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
    assert uploaded["电话"] == "13900139000"
    assert uploaded["工作室"] == "短视频工作室（部长）、图片工作室（成员）"
    unknown = frame[frame["姓名"] == "未录入同学"].iloc[0]
    assert uploaded["职位"] == "短视频工作室：部长；图片工作室：成员"
    assert unknown["工作室"] == UNKNOWN_STUDIO

    output = export_roster_excel(entries, members, db.get_courses())
    wb = pytest.importorskip("openpyxl").load_workbook(io.BytesIO(output))
    assert wb.sheetnames == ["完整花名册", "课表状态", "课表明细"]
    ws = wb["完整花名册"]
    assert ws["A1"].value == "全媒体中心花名册"
    assert "A1:F1" in {str(cell_range) for cell_range in ws.merged_cells.ranges}
    assert [ws.cell(2, column).value for column in range(1, 7)] == [
        "工作室", "职位", "姓名", "学号", "电话", "学院+专业"]
    assert ws.max_column == 6
    assert ws["A3"].value == "短视频工作室"
    assert ws["B3"].value == "部长"
    assert ws["C3"].value == "已在名单"
    assert ws["E3"].value == "13900139000"
    assert ws["A4"].value == "图片工作室"
    assert ws["B4"].value == "成员"
    assert ws.freeze_panes == "A3"
    assert wb["课表明细"].max_row == 2
    assert wb["课表明细"].cell(2, 4).value == "数据库"
    assert db.get_member(member_id) is not None


def test_existing_single_studio_database_migrates_to_multi_studio(tmp_path: Path) -> None:
    """老库升级后，应根据花名册恢复同一成员的全部工作室归属。"""
    path = tmp_path / "legacy.db"
    db = Database(path)
    member_id = db.upsert_member(ParsedSchedule(
        name="兼岗同学", student_id="20250001", courses=[]))
    db.replace_roster([
        RosterEntry(studio="短视频工作室", position="部长",
                    name="兼岗同学", student_id="20250001"),
        RosterEntry(studio="设计工作室", position="副部长",
                    name="兼岗同学", student_id="20250001"),
    ])

    conn = sqlite3.connect(path)
    conn.execute("DELETE FROM member_studios")
    conn.execute("PRAGMA user_version = 5")
    conn.commit()
    conn.close()

    upgraded = Database(path).get_member(member_id)
    assert upgraded is not None
    assert upgraded.studios == ["短视频工作室", "设计工作室"]
    assert upgraded.studio_positions == {
        "短视频工作室": "部长", "设计工作室": "副部长"}
