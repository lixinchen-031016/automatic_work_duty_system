"""存储层测试：幂等写入、增量同步、索引与迁移、老库兼容"""

from __future__ import annotations

import sqlite3
from datetime import date
from pathlib import Path

from duty_system.calendar import CalendarEntry
from duty_system.database import MIGRATIONS, SCHEMA, Assignment, Database
from duty_system.parser import ParsedSchedule


def new_db(tmp_path: Path) -> Database:
    return Database(tmp_path / "test.db")


def add_members(db: Database, names: list[str]) -> list[int]:
    for i, name in enumerate(names):
        db.upsert_member(ParsedSchedule(name=name, student_id=str(i + 1), courses=[]))
    return [m.id for m in db.list_members()]


def test_idempotent_assignment_write(tmp_path: Path) -> None:
    """重复写入相同安排不应产生新行，也不应让自增 id 膨胀"""
    db = new_db(tmp_path)
    a, b = add_members(db, ["甲", "乙"])
    assignments = [Assignment(1, 1, 1, a, "甲"), Assignment(1, 1, 2, b, "乙")]

    db.save_assignments(assignments)
    conn = sqlite3.connect(db.path)
    ids_first = [r[0] for r in conn.execute("SELECT id FROM duty_assignments ORDER BY id")]
    conn.close()

    db.save_assignments(assignments)   # 再写一次
    db.save_assignments(assignments)   # 第三次
    conn = sqlite3.connect(db.path)
    ids_after = [r[0] for r in conn.execute("SELECT id FROM duty_assignments ORDER BY id")]
    conn.close()

    assert ids_first == ids_after, "重复写入不应重建行（INSERT OR REPLACE 会使 id 膨胀）"
    assert len(db.load_assignments()) == 2


def test_sync_assignments_keeps_unchanged_rows(tmp_path: Path) -> None:
    """按周同步：未变化的安排保持原行，只增删差异部分"""
    db = new_db(tmp_path)
    a, b, c = add_members(db, ["甲", "乙", "丙"])
    db.save_assignments([Assignment(1, 1, 1, a), Assignment(1, 1, 2, b),
                         Assignment(2, 1, 1, c)])
    conn = sqlite3.connect(db.path)
    before = {r[0]: r[1] for r in conn.execute("SELECT id, member_id FROM duty_assignments")}
    conn.close()

    added, removed = db.sync_assignments_for_weeks(
        [1, 2],
        [Assignment(1, 1, 1, a), Assignment(1, 1, 2, b), Assignment(2, 1, 1, c)])
    assert (added, removed) == (0, 0), "完全相同的排班不应产生写入"

    conn = sqlite3.connect(db.path)
    after = {r[0]: r[1] for r in conn.execute("SELECT id, member_id FROM duty_assignments")}
    conn.close()
    assert before == after, "未变化的行应保持原 id"

    added, removed = db.sync_assignments_for_weeks([1], [Assignment(1, 1, 1, c)])
    assert (added, removed) == (1, 2), f"期望新增 1 行 / 删除 2 行，实际 {added}/{removed}"
    got = {(x.week, x.weekday, x.block, x.member_id) for x in db.load_assignments()}
    assert got == {(1, 1, 1, c), (2, 1, 1, c)}, "只应改动第 1 周"


def test_replace_slot_only_touches_target(tmp_path: Path) -> None:
    """槽位替换只动目标时段，其他时段的行保持不变"""
    db = new_db(tmp_path)
    a, b, c = add_members(db, ["甲", "乙", "丙"])
    db.save_assignments([Assignment(1, 1, 1, a), Assignment(1, 1, 2, b)])
    conn = sqlite3.connect(db.path)
    other_id = conn.execute(
        "SELECT id FROM duty_assignments WHERE block = 2").fetchone()[0]
    conn.close()

    db.replace_slot(1, 1, 1, [c])
    rows = {(x.weekday, x.block, x.member_id) for x in db.load_assignments()}
    assert rows == {(1, 1, c), (1, 2, b)}, "只应替换目标槽位"

    conn = sqlite3.connect(db.path)
    still_there = conn.execute(
        "SELECT id FROM duty_assignments WHERE block = 2").fetchone()[0]
    conn.close()
    assert still_there == other_id, "未涉及的时段不应被重写"


def test_indexes_wal_and_migration(tmp_path: Path) -> None:
    """新建库具备索引与 WAL；老库（无索引）能被迁移补齐"""
    db = new_db(tmp_path)
    conn = sqlite3.connect(db.path)
    indexes = {r[0] for r in conn.execute(
        "SELECT name FROM sqlite_master WHERE type='index' AND name LIKE 'idx_%'")}
    version = conn.execute("PRAGMA user_version").fetchone()[0]
    journal = conn.execute("PRAGMA journal_mode").fetchone()[0]
    conn.close()

    expected = {"idx_courses_member", "idx_assignments_week",
                "idx_assignments_member", "idx_leaves_week",
                "idx_special_arrangements_member",
                "idx_special_arrangements_weeks"}
    assert expected <= indexes, f"缺少索引: {expected - indexes}"
    assert version == MIGRATIONS[-1][0], "user_version 未推进到最新"
    assert journal.lower() == "wal", "应启用 WAL 以提升并发写入稳定性"

    # 模拟旧版本数据库：结构在、但没有索引、user_version=0
    legacy = tmp_path / "legacy.db"
    conn = sqlite3.connect(legacy)
    conn.executescript(SCHEMA.replace("CREATE INDEX IF NOT EXISTS", "-- CREATE INDEX IF NOT EXISTS"))
    conn.execute("DROP TABLE IF EXISTS term_calendar")
    conn.execute("PRAGMA user_version = 0")
    conn.commit()
    conn.close()

    Database(legacy)   # 打开即触发迁移
    conn = sqlite3.connect(legacy)
    migrated = {r[0] for r in conn.execute(
        "SELECT name FROM sqlite_master WHERE type='index' AND name LIKE 'idx_%'")}
    version = conn.execute("PRAGMA user_version").fetchone()[0]
    calendar_table = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='term_calendar'"
    ).fetchone()
    conn.close()
    assert expected <= migrated, "老库迁移后应补齐索引"
    assert version == MIGRATIONS[-1][0], "老库迁移后应记录版本号"
    assert calendar_table is not None, "老库迁移后应创建学期日历表"


def test_term_calendar_crud_and_validation(tmp_path: Path) -> None:
    """日历覆盖可增改查删，严格校验缺对/错对，纯假日可显式放行。"""
    db = new_db(tmp_path)
    term_start = date(2026, 9, 14)
    off = CalendarEntry(date(2026, 10, 15), "off", note="调休")
    makeup = CalendarEntry(date(2026, 10, 17), "class", 5, 4, "补第5周周四")

    db.upsert_calendar([off, makeup])
    assert db.list_calendar() == [off, makeup]
    db.upsert_calendar([CalendarEntry(date(2026, 10, 15), "off", note="更新")])
    assert db.list_calendar()[0].note == "更新"
    db.validate_calendar()
    db.validate_calendar(term_start)

    db.clear_calendar()
    db.upsert_calendar([off])
    try:
        db.validate_calendar(term_start)
    except ValueError as exc:
        assert "缺少对应的 class" in str(exc)
    else:
        raise AssertionError("严格模式应拦截缺失补课对的 off 项")
    db.validate_calendar(term_start, allow_unpaired_off=True)

    db.clear_calendar()
    db.upsert_calendar([
        CalendarEntry(date(2026, 10, 16), "off"),
        CalendarEntry(date(2026, 10, 17), "class", 5, 4),
    ])
    try:
        db.validate_calendar(term_start)
    except ValueError as exc:
        assert "未设为 off" in str(exc)
    else:
        raise AssertionError("补课项指向的自然工作日不是 off 时应报错")


def test_upsert_replaces_courses_and_cascades(tmp_path: Path) -> None:
    """重传课表整体替换课程；删除成员级联清理课程/排班/请假"""
    from duty_system.parser import Course

    db = new_db(tmp_path)
    schedule = ParsedSchedule(
        name="甲", student_id="1", file_name="a.xls",
        courses=[Course(course_name=f"课{i}", weekday=1, week_list=[1],
                        session_list=[1, 2], weeks_text="1", sessions_text="01-02")
                 for i in range(3)])
    mid = db.upsert_member(schedule)
    assert len(db.get_courses(mid)) == 3

    schedule.courses = schedule.courses[:1]
    assert db.upsert_member(schedule) == mid, "同学号同名应更新而非新增"
    assert len(db.get_courses(mid)) == 1, "重传应整体替换课程"
    assert len(db.list_members()) == 1

    db.save_assignments([Assignment(1, 1, 1, mid, "甲")])
    db.add_leave(mid, 1, 1, "事假")
    db.add_special_arrangement(mid, 1, 8, 2, [3, 4], "固定实习")
    db.delete_member(mid)
    assert db.list_members() == []
    assert db.load_assignments() == [], "删除成员应级联清理排班"
    assert db.list_leaves() == [], "删除成员应级联清理请假"
    assert db.list_special_arrangements() == [], "删除成员应级联清理特殊安排"


def test_special_arrangement_crud_is_normalized_and_idempotent(tmp_path: Path) -> None:
    """长期特殊安排支持新增、覆盖、编辑、查询和删除，节次会规范化排序。"""
    db = new_db(tmp_path)
    mid, other = add_members(db, ["甲", "乙"])

    arrangement_id = db.add_special_arrangement(
        mid, 2, 10, 3, [4, 3, 3], "训练")
    again = db.add_special_arrangement(mid, 2, 10, 3, [3, 4], "训练（调整）")
    assert again == arrangement_id, "相同范围/星期/节次重复登记应覆盖而非新增"
    rows = db.list_special_arrangements(mid)
    assert len(rows) == 1 and rows[0].session_list == [3, 4]
    assert rows[0].reason == "训练（调整）" and list(rows[0].week_list) == list(range(2, 11))

    db.update_special_arrangement(arrangement_id, other, 5, 5, 6, [7, 8], "比赛")
    updated = db.list_special_arrangements(other)[0]
    assert updated.id == arrangement_id and updated.member_id == other
    assert (updated.week_start, updated.week_end, updated.weekday) == (5, 5, 6)
    assert updated.session_list == [7, 8] and updated.reason == "比赛"

    db.remove_special_arrangement(arrangement_id)
    assert db.list_special_arrangements() == []

def test_backup_to_is_wal_safe(tmp_path: Path) -> None:
    """WAL 模式下备份必须用 SQLite 在线备份 API

    回归用例：启用 WAL 后 schema/最新数据可能只在 -wal 文件里，
    用 shutil.copy2 复制主库会得到「no such table」的损坏副本。
    """
    db = new_db(tmp_path)
    ids = add_members(db, ["甲", "乙"])
    db.save_assignments([Assignment(1, 1, 1, ids[0]), Assignment(1, 1, 2, ids[1])])

    target = tmp_path / "nested" / "backup.db"
    db.backup_to(target)   # 目标目录不存在时应自动创建

    assert target.exists()
    assert not (tmp_path / "nested" / "backup.db-wal").exists(), "备份不应遗留 -wal"
    restored = Database(target)
    assert len(restored.list_members()) == 2
    assert len(restored.load_assignments()) == 2
    # 副本是一个可用数据库，不只是能读
    assert restored.sync_assignments_for_weeks([1], [Assignment(1, 1, 1, ids[1])]) == (1, 2)
