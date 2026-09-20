"""SQLite 存储层：成员、课程、特殊安排、请假与值班安排的持久化"""

from __future__ import annotations

import json
import sqlite3
from dataclasses import dataclass
from pathlib import Path

from .parser import ParsedSchedule

SCHEMA = """
CREATE TABLE IF NOT EXISTS members (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id TEXT NOT NULL DEFAULT '',
    name TEXT NOT NULL,
    term TEXT NOT NULL DEFAULT '',
    class_name TEXT NOT NULL DEFAULT '',
    major TEXT NOT NULL DEFAULT '',
    department TEXT NOT NULL DEFAULT '',
    file_name TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    UNIQUE(student_id, name)
);

CREATE TABLE IF NOT EXISTS courses (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    member_id INTEGER NOT NULL REFERENCES members(id) ON DELETE CASCADE,
    course_name TEXT NOT NULL,
    teacher TEXT NOT NULL DEFAULT '',
    weekday INTEGER NOT NULL,
    weeks_text TEXT NOT NULL DEFAULT '',
    week_list TEXT NOT NULL DEFAULT '[]',
    sessions_text TEXT NOT NULL DEFAULT '',
    session_list TEXT NOT NULL DEFAULT '[]',
    location TEXT NOT NULL DEFAULT '',
    UNIQUE(member_id, course_name, teacher, weekday, weeks_text, sessions_text, location)
);

CREATE TABLE IF NOT EXISTS duty_assignments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    week INTEGER NOT NULL,
    weekday INTEGER NOT NULL,
    block INTEGER NOT NULL,
    member_id INTEGER NOT NULL REFERENCES members(id) ON DELETE CASCADE,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    UNIQUE(week, weekday, block, member_id)
);

CREATE TABLE IF NOT EXISTS leaves (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    member_id INTEGER NOT NULL REFERENCES members(id) ON DELETE CASCADE,
    week INTEGER NOT NULL,
    weekday INTEGER NOT NULL,
    reason TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    UNIQUE(member_id, week, weekday)
);

CREATE TABLE IF NOT EXISTS special_arrangements (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    member_id INTEGER NOT NULL REFERENCES members(id) ON DELETE CASCADE,
    week_start INTEGER NOT NULL,
    week_end INTEGER NOT NULL,
    weekday INTEGER NOT NULL,
    session_list TEXT NOT NULL DEFAULT '[]',
    reason TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    UNIQUE(member_id, week_start, week_end, weekday, session_list)
);

CREATE INDEX IF NOT EXISTS idx_courses_member ON courses(member_id);
CREATE INDEX IF NOT EXISTS idx_assignments_week ON duty_assignments(week);
CREATE INDEX IF NOT EXISTS idx_assignments_member ON duty_assignments(member_id);
CREATE INDEX IF NOT EXISTS idx_leaves_week ON leaves(week);
CREATE INDEX IF NOT EXISTS idx_special_arrangements_member ON special_arrangements(member_id);
CREATE INDEX IF NOT EXISTS idx_special_arrangements_weeks ON special_arrangements(week_start, week_end);
"""

# 结构迁移：只对已有数据库执行增量 DDL。
# 每条迁移用 (版本号, SQL 列表)；执行后写入 PRAGMA user_version，
# 避免「CREATE TABLE IF NOT EXISTS 对老库静默跳过」导致新字段不生效。
MIGRATIONS: list[tuple[int, list[str]]] = [
    (1, [
        "CREATE INDEX IF NOT EXISTS idx_courses_member ON courses(member_id)",
        "CREATE INDEX IF NOT EXISTS idx_assignments_week ON duty_assignments(week)",
        "CREATE INDEX IF NOT EXISTS idx_assignments_member ON duty_assignments(member_id)",
        "CREATE INDEX IF NOT EXISTS idx_leaves_week ON leaves(week)",
    ]),
    (2, [
        """CREATE TABLE IF NOT EXISTS special_arrangements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            member_id INTEGER NOT NULL REFERENCES members(id) ON DELETE CASCADE,
            week_start INTEGER NOT NULL,
            week_end INTEGER NOT NULL,
            weekday INTEGER NOT NULL,
            session_list TEXT NOT NULL DEFAULT '[]',
            reason TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
            updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
            UNIQUE(member_id, week_start, week_end, weekday, session_list)
        )""",
        """CREATE INDEX IF NOT EXISTS idx_special_arrangements_member
            ON special_arrangements(member_id)""",
        """CREATE INDEX IF NOT EXISTS idx_special_arrangements_weeks
            ON special_arrangements(week_start, week_end)""",
    ]),
]


@dataclass
class Member:
    id: int
    name: str
    student_id: str
    term: str
    class_name: str
    major: str
    department: str
    file_name: str
    course_count: int = 0


@dataclass
class CourseRecord:
    """从数据库读出的课程记录（周次/节次为已展开的列表）"""
    id: int
    member_id: int
    course_name: str
    teacher: str
    weekday: int
    week_list: list[int]
    session_list: list[int]
    location: str
    weeks_text: str
    sessions_text: str


@dataclass
class Assignment:
    week: int
    weekday: int
    block: int
    member_id: int
    member_name: str = ""


@dataclass
class Leave:
    """请假/临时占用：某成员第 week 周 weekday 全天不可值班"""
    id: int
    member_id: int
    week: int
    weekday: int
    reason: str = ""


@dataclass
class SpecialArrangement:
    """课表外的长期固定占用：连续周范围内每周同一天、同一时段不可值班。"""
    id: int
    member_id: int
    week_start: int
    week_end: int
    weekday: int
    session_list: list[int]
    reason: str = ""

    @property
    def week_list(self) -> range:
        return range(self.week_start, self.week_end + 1)


class Database:
    def __init__(self, path: str | Path = "duty_system.db"):
        self.path = str(path)
        with self._connect() as conn:
            conn.executescript(SCHEMA)
            self._migrate(conn)

    @staticmethod
    def _migrate(conn: sqlite3.Connection) -> None:
        """按 PRAGMA user_version 增量升级老数据库结构"""
        version = conn.execute("PRAGMA user_version").fetchone()[0]
        for target, statements in MIGRATIONS:
            if version >= target:
                continue
            for sql in statements:
                conn.execute(sql)
            conn.execute(f"PRAGMA user_version = {target}")

    def _connect(self) -> sqlite3.Connection:
        # WAL：桌面单用户场景下读写并发更稳、写入更快；
        # busy_timeout 避免与后台线程同时写时直接抛「database is locked」
        conn = sqlite3.connect(self.path, timeout=10)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA journal_mode = WAL")
        conn.execute("PRAGMA synchronous = NORMAL")
        return conn

    # ---------- 备份 / 迁移 ----------

    def backup_to(self, target: str | Path) -> None:
        """把当前数据库完整复制到 target（供「更改数据库位置」使用）。

        必须用 SQLite 的在线备份 API，不能用 shutil.copy2：
        启用 WAL 后最新数据与 schema 可能仍在 -wal 文件里，
        只复制主库文件会得到一个打不开或丢数据的副本。
        """
        target = Path(target)
        target.parent.mkdir(parents=True, exist_ok=True)
        src = self._connect()
        dst = sqlite3.connect(target)
        try:
            src.backup(dst)
        finally:
            dst.close()
            src.close()
        # 备份结果自带完整数据，清掉目标可能存在的伴随文件避免混淆
        for suffix in ("-wal", "-shm"):
            Path(f"{target}{suffix}").unlink(missing_ok=True)

    # ---------- 成员 ----------

    def upsert_member(self, schedule: ParsedSchedule) -> int:
        """新增或更新成员（按 学号+姓名 唯一），并整体替换其课程，返回成员 id"""
        with self._connect() as conn:
            row = conn.execute(
                "SELECT id FROM members WHERE student_id = ? AND name = ?",
                (schedule.student_id, schedule.name),
            ).fetchone()
            if row:
                member_id = row["id"]
                conn.execute(
                    """UPDATE members SET term = ?, class_name = ?, major = ?,
                       department = ?, file_name = ? WHERE id = ?""",
                    (schedule.term, schedule.class_name, schedule.major,
                     schedule.department, schedule.file_name, member_id),
                )
                conn.execute("DELETE FROM courses WHERE member_id = ?", (member_id,))
            else:
                cur = conn.execute(
                    """INSERT INTO members
                       (student_id, name, term, class_name, major, department, file_name)
                       VALUES (?, ?, ?, ?, ?, ?, ?)""",
                    (schedule.student_id, schedule.name, schedule.term,
                     schedule.class_name, schedule.major, schedule.department,
                     schedule.file_name),
                )
                member_id = cur.lastrowid

            conn.executemany(
                """INSERT OR IGNORE INTO courses
                   (member_id, course_name, teacher, weekday, weeks_text, week_list,
                    sessions_text, session_list, location)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                [
                    (member_id, c.course_name, c.teacher, c.weekday, c.weeks_text,
                     json.dumps(c.week_list), c.sessions_text,
                     json.dumps(c.session_list), c.location)
                    for c in [*schedule.courses, *schedule.whole_week_courses]
                ],
            )
            return member_id

    def list_members(self) -> list[Member]:
        with self._connect() as conn:
            rows = conn.execute(
                """SELECT m.*, COUNT(c.id) AS course_count FROM members m
                   LEFT JOIN courses c ON c.member_id = m.id
                   GROUP BY m.id ORDER BY m.id"""
            ).fetchall()
            return [Member(
                id=r["id"], name=r["name"], student_id=r["student_id"],
                term=r["term"], class_name=r["class_name"], major=r["major"],
                department=r["department"], file_name=r["file_name"],
                course_count=r["course_count"],
            ) for r in rows]

    def delete_member(self, member_id: int) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM members WHERE id = ?", (member_id,))

    def get_member(self, member_id: int) -> Member | None:
        with self._connect() as conn:
            r = conn.execute("SELECT * FROM members WHERE id = ?", (member_id,)).fetchone()
            return Member(
                id=r["id"], name=r["name"], student_id=r["student_id"], term=r["term"],
                class_name=r["class_name"], major=r["major"], department=r["department"],
                file_name=r["file_name"],
            ) if r else None

    # ---------- 课程 ----------

    def get_courses(self, member_id: int | None = None) -> list[CourseRecord]:
        sql = "SELECT * FROM courses"
        params: tuple = ()
        if member_id is not None:
            sql += " WHERE member_id = ?"
            params = (member_id,)
        sql += " ORDER BY member_id, weekday, course_name"
        with self._connect() as conn:
            rows = conn.execute(sql, params).fetchall()
            return [CourseRecord(
                id=r["id"], member_id=r["member_id"], course_name=r["course_name"],
                teacher=r["teacher"], weekday=r["weekday"],
                week_list=json.loads(r["week_list"]),
                session_list=json.loads(r["session_list"]),
                location=r["location"], weeks_text=r["weeks_text"],
                sessions_text=r["sessions_text"],
            ) for r in rows]

    # ---------- 值班安排 ----------

    def clear_assignments(self) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM duty_assignments")

    def delete_assignments_for_weeks(self, weeks: list[int]) -> None:
        """删除指定周的值班安排（按周增量重排时清空所选范围）"""
        if not weeks:
            return
        ph = ",".join("?" * len(weeks))
        with self._connect() as conn:
            conn.execute(f"DELETE FROM duty_assignments WHERE week IN ({ph})", weeks)

    def save_assignments(self, assignments: list[Assignment]) -> None:
        """幂等写入：唯一键冲突时不做任何改动。
        不再使用 INSERT OR REPLACE（那会先删后插，使自增 id 每轮膨胀）。"""
        with self._connect() as conn:
            conn.executemany(
                """INSERT INTO duty_assignments (week, weekday, block, member_id)
                   VALUES (?, ?, ?, ?)
                   ON CONFLICT(week, weekday, block, member_id) DO NOTHING""",
                [(a.week, a.weekday, a.block, a.member_id) for a in assignments],
            )

    def sync_assignments_for_weeks(
        self,
        weeks: list[int],
        assignments: list[Assignment],
    ) -> tuple[int, int]:
        """把指定周的排班同步成 assignments：只删该周多余的行、只插新增的行。

        相比「先删整周再全量插入」，未变化的安排保持原行不动——
        自增 id 不膨胀，写入量也更小。返回 (新增数, 删除数)。
        """
        if not weeks:
            return 0, 0
        want = {(a.week, a.weekday, a.block, a.member_id) for a in assignments}
        ph = ",".join("?" * len(weeks))
        with self._connect() as conn:
            rows = conn.execute(
                f"""SELECT week, weekday, block, member_id FROM duty_assignments
                    WHERE week IN ({ph})""",
                weeks,
            ).fetchall()
            have = {(r["week"], r["weekday"], r["block"], r["member_id"]) for r in rows}
            stale = have - want
            if stale:
                conn.executemany(
                    """DELETE FROM duty_assignments
                       WHERE week = ? AND weekday = ? AND block = ? AND member_id = ?""",
                    [tuple(x) for x in stale],
                )
            fresh = want - have
            if fresh:
                conn.executemany(
                    """INSERT INTO duty_assignments (week, weekday, block, member_id)
                       VALUES (?, ?, ?, ?)
                       ON CONFLICT(week, weekday, block, member_id) DO NOTHING""",
                    [tuple(x) for x in fresh],
                )
        return len(fresh), len(stale)

    def replace_slot(
        self,
        week: int,
        weekday: int,
        block: int,
        member_ids: list[int],
    ) -> None:
        """只重写单个 (周, 星期, 时段) 的安排——手动微调用，
        避免全表 clear + 全量重写的开销与中途失败风险。"""
        with self._connect() as conn:
            conn.execute(
                "DELETE FROM duty_assignments WHERE week = ? AND weekday = ? AND block = ?",
                (week, weekday, block),
            )
            conn.executemany(
                """INSERT INTO duty_assignments (week, weekday, block, member_id)
                   VALUES (?, ?, ?, ?)
                   ON CONFLICT(week, weekday, block, member_id) DO NOTHING""",
                [(week, weekday, block, mid) for mid in member_ids],
            )

    def load_assignments(self) -> list[Assignment]:
        with self._connect() as conn:
            rows = conn.execute(
                """SELECT a.week, a.weekday, a.block, a.member_id, m.name AS member_name
                   FROM duty_assignments a JOIN members m ON m.id = a.member_id
                   ORDER BY a.week, a.weekday, a.block, m.name"""
            ).fetchall()
            return [Assignment(
                week=r["week"], weekday=r["weekday"], block=r["block"],
                member_id=r["member_id"], member_name=r["member_name"],
            ) for r in rows]

    # ---------- 请假 ----------

    def add_leave(self, member_id: int, week: int, weekday: int, reason: str = "") -> None:
        """登记请假（同成员同周同星期重复登记则覆盖）"""
        with self._connect() as conn:
            conn.execute(
                """INSERT OR REPLACE INTO leaves (member_id, week, weekday, reason)
                   VALUES (?, ?, ?, ?)""",
                (member_id, week, weekday, reason),
            )

    def remove_leave(self, leave_id: int) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM leaves WHERE id = ?", (leave_id,))

    def list_leaves(self, member_id: int | None = None) -> list[Leave]:
        sql = "SELECT * FROM leaves"
        params: tuple = ()
        if member_id is not None:
            sql += " WHERE member_id = ?"
            params = (member_id,)
        sql += " ORDER BY week, weekday, member_id"
        with self._connect() as conn:
            rows = conn.execute(sql, params).fetchall()
            return [Leave(
                id=r["id"], member_id=r["member_id"], week=r["week"],
                weekday=r["weekday"], reason=r["reason"],
            ) for r in rows]

    # ---------- 长期特殊安排（课表修改） ----------

    @staticmethod
    def _normalize_special(
        week_start: int,
        week_end: int,
        weekday: int,
        session_list: list[int],
    ) -> tuple[int, int, int, list[int], str]:
        if not 1 <= week_start <= week_end <= 25:
            raise ValueError("周次范围应为 1–25，且起始周不能大于结束周")
        if not 1 <= weekday <= 7:
            raise ValueError("星期应为 1–7")
        sessions = sorted({int(s) for s in session_list})
        if not sessions or any(s < 1 or s > 11 for s in sessions):
            raise ValueError("请至少选择一个有效节次")
        return week_start, week_end, weekday, sessions, json.dumps(
            sessions, ensure_ascii=False, separators=(",", ":"))

    def add_special_arrangement(
        self,
        member_id: int,
        week_start: int,
        week_end: int,
        weekday: int,
        session_list: list[int],
        reason: str = "",
    ) -> int:
        """新增长期特殊安排；相同范围/星期/节次重复登记时覆盖原因。"""
        week_start, week_end, weekday, _sessions, sessions_json = self._normalize_special(
            week_start, week_end, weekday, session_list)
        with self._connect() as conn:
            conn.execute(
                """INSERT INTO special_arrangements
                       (member_id, week_start, week_end, weekday, session_list, reason)
                   VALUES (?, ?, ?, ?, ?, ?)
                   ON CONFLICT(member_id, week_start, week_end, weekday, session_list)
                   DO UPDATE SET reason = excluded.reason,
                                 updated_at = datetime('now', 'localtime')""",
                (member_id, week_start, week_end, weekday, sessions_json, reason.strip()),
            )
            row = conn.execute(
                """SELECT id FROM special_arrangements
                   WHERE member_id = ? AND week_start = ? AND week_end = ?
                     AND weekday = ? AND session_list = ?""",
                (member_id, week_start, week_end, weekday, sessions_json),
            ).fetchone()
            return int(row["id"])

    def update_special_arrangement(
        self,
        arrangement_id: int,
        member_id: int,
        week_start: int,
        week_end: int,
        weekday: int,
        session_list: list[int],
        reason: str = "",
    ) -> None:
        """修改一条长期特殊安排。"""
        week_start, week_end, weekday, _sessions, sessions_json = self._normalize_special(
            week_start, week_end, weekday, session_list)
        with self._connect() as conn:
            try:
                conn.execute(
                    """UPDATE special_arrangements
                       SET member_id = ?, week_start = ?, week_end = ?, weekday = ?,
                           session_list = ?, reason = ?,
                           updated_at = datetime('now', 'localtime')
                       WHERE id = ?""",
                    (member_id, week_start, week_end, weekday, sessions_json,
                     reason.strip(), arrangement_id),
                )
            except sqlite3.IntegrityError:
                # 修改后与另一条安排完全重合时合并：保留当前记录，删除重复项。
                conn.execute(
                    """DELETE FROM special_arrangements
                       WHERE member_id = ? AND week_start = ? AND week_end = ?
                         AND weekday = ? AND session_list = ? AND id != ?""",
                    (member_id, week_start, week_end, weekday, sessions_json,
                     arrangement_id),
                )
                conn.execute(
                    """UPDATE special_arrangements
                       SET member_id = ?, week_start = ?, week_end = ?, weekday = ?,
                           session_list = ?, reason = ?,
                           updated_at = datetime('now', 'localtime')
                       WHERE id = ?""",
                    (member_id, week_start, week_end, weekday, sessions_json,
                     reason.strip(), arrangement_id),
                )

    def remove_special_arrangement(self, arrangement_id: int) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM special_arrangements WHERE id = ?", (arrangement_id,))

    def list_special_arrangements(
        self,
        member_id: int | None = None,
    ) -> list[SpecialArrangement]:
        sql = "SELECT * FROM special_arrangements"
        params: tuple = ()
        if member_id is not None:
            sql += " WHERE member_id = ?"
            params = (member_id,)
        sql += " ORDER BY member_id, week_start, week_end, weekday, session_list"
        with self._connect() as conn:
            rows = conn.execute(sql, params).fetchall()
            return [SpecialArrangement(
                id=r["id"], member_id=r["member_id"], week_start=r["week_start"],
                week_end=r["week_end"], weekday=r["weekday"],
                session_list=json.loads(r["session_list"]), reason=r["reason"],
            ) for r in rows]
