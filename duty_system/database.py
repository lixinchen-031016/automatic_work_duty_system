"""SQLite 存储层：成员、课程、值班安排的持久化"""

from __future__ import annotations

import json
import sqlite3
from dataclasses import dataclass
from pathlib import Path

from .parser import Course, ParsedSchedule

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
"""


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


class Database:
    def __init__(self, path: str | Path = "duty_system.db"):
        self.path = str(path)
        with self._connect() as conn:
            conn.executescript(SCHEMA)

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

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
                    for c in schedule.courses
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

    def save_assignments(self, assignments: list[Assignment]) -> None:
        with self._connect() as conn:
            conn.executemany(
                """INSERT OR REPLACE INTO duty_assignments
                   (week, weekday, block, member_id) VALUES (?, ?, ?, ?)""",
                [(a.week, a.weekday, a.block, a.member_id) for a in assignments],
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
