"""SQLite 存储层：成员、课程、特殊安排、请假与值班安排的持久化"""

from __future__ import annotations

import json
import sqlite3
from collections.abc import Iterable, Iterator
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from .calendar import CalendarEntry, validate_calendar_entries
from .parser import ParsedSchedule
from .roster import (
    POSITION_LEADER,
    UNKNOWN_STUDIO,
    RosterEntry,
    normalize_name,
    normalize_student_id,
)

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
    phone TEXT NOT NULL DEFAULT '',
    studio TEXT NOT NULL DEFAULT '未指定工作室',
    studio_locked INTEGER NOT NULL DEFAULT 0
        CHECK (studio_locked IN (0, 1)),
    participates_in_scheduling INTEGER NOT NULL DEFAULT 1
        CHECK (participates_in_scheduling IN (0, 1)),
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
    UNIQUE(student_id, name)
);

CREATE TABLE IF NOT EXISTS member_studios (
    member_id INTEGER NOT NULL REFERENCES members(id) ON DELETE CASCADE,
    studio TEXT NOT NULL,
    position TEXT NOT NULL DEFAULT '',
    PRIMARY KEY (member_id, studio)
);

CREATE TABLE IF NOT EXISTS roster_entries (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    studio TEXT NOT NULL DEFAULT '未指定工作室',
    position TEXT NOT NULL DEFAULT '',
    name TEXT NOT NULL,
    student_id TEXT NOT NULL DEFAULT '',
    phone TEXT NOT NULL DEFAULT '',
    college_major TEXT NOT NULL DEFAULT '',
    source_file TEXT NOT NULL DEFAULT '',
    row_number INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
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

CREATE TABLE IF NOT EXISTS term_calendar (
    date TEXT PRIMARY KEY,
    override_type TEXT NOT NULL CHECK(override_type IN ('off', 'class')),
    maps_to_week INTEGER,
    maps_to_weekday INTEGER,
    note TEXT NOT NULL DEFAULT '',
    CHECK(
        (override_type = 'off' AND maps_to_week IS NULL AND maps_to_weekday IS NULL)
        OR
        (override_type = 'class' AND maps_to_week IS NOT NULL AND maps_to_weekday IS NOT NULL)
    )
);

CREATE INDEX IF NOT EXISTS idx_member_studios_studio ON member_studios(studio);
CREATE INDEX IF NOT EXISTS idx_roster_entries_student ON roster_entries(student_id);
CREATE INDEX IF NOT EXISTS idx_roster_entries_name ON roster_entries(name);
CREATE INDEX IF NOT EXISTS idx_roster_entries_studio ON roster_entries(studio);
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
    (3, [
        """CREATE TABLE IF NOT EXISTS term_calendar (
            date TEXT PRIMARY KEY,
            override_type TEXT NOT NULL CHECK(override_type IN ('off', 'class')),
            maps_to_week INTEGER,
            maps_to_weekday INTEGER,
            note TEXT NOT NULL DEFAULT '',
            CHECK(
                (override_type = 'off' AND maps_to_week IS NULL AND maps_to_weekday IS NULL)
                OR
                (override_type = 'class' AND maps_to_week IS NOT NULL AND maps_to_weekday IS NOT NULL)
            )
        )""",
    ]),
    (4, [
        """ALTER TABLE members ADD COLUMN participates_in_scheduling
           INTEGER NOT NULL DEFAULT 1
           CHECK (participates_in_scheduling IN (0, 1))""",
    ]),
    (5, [
        """ALTER TABLE members ADD COLUMN studio
           TEXT NOT NULL DEFAULT '未指定工作室'""",
        """ALTER TABLE members ADD COLUMN studio_locked
           INTEGER NOT NULL DEFAULT 0 CHECK (studio_locked IN (0, 1))""",
        """CREATE TABLE IF NOT EXISTS roster_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            studio TEXT NOT NULL DEFAULT '未指定工作室',
            position TEXT NOT NULL DEFAULT '',
            name TEXT NOT NULL,
            student_id TEXT NOT NULL DEFAULT '',
            phone TEXT NOT NULL DEFAULT '',
            college_major TEXT NOT NULL DEFAULT '',
            source_file TEXT NOT NULL DEFAULT '',
            row_number INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
        )""",
        """CREATE INDEX IF NOT EXISTS idx_roster_entries_student
            ON roster_entries(student_id)""",
        """CREATE INDEX IF NOT EXISTS idx_roster_entries_name
            ON roster_entries(name)""",
        """CREATE INDEX IF NOT EXISTS idx_roster_entries_studio
            ON roster_entries(studio)""",
    ]),
    (6, [
        """CREATE TABLE IF NOT EXISTS member_studios (
            member_id INTEGER NOT NULL REFERENCES members(id) ON DELETE CASCADE,
            studio TEXT NOT NULL,
            PRIMARY KEY (member_id, studio)
        )""",
        """INSERT OR IGNORE INTO member_studios (member_id, studio)
           SELECT id, CASE WHEN TRIM(studio) = '' THEN '未指定工作室'
                           ELSE studio END
           FROM members""",
        """CREATE INDEX IF NOT EXISTS idx_member_studios_studio
            ON member_studios(studio)""",
    ]),
    (7, [
        """ALTER TABLE member_studios ADD COLUMN position
           TEXT NOT NULL DEFAULT ''""",
        """UPDATE member_studios
           SET position = COALESCE((
               SELECT r.position FROM roster_entries r
               JOIN members m ON m.id = member_studios.member_id
               WHERE r.studio = member_studios.studio
                 AND (
                     (m.student_id <> '' AND r.student_id = m.student_id)
                     OR r.name = m.name
                 )
               ORDER BY CASE WHEN r.student_id = m.student_id THEN 0 ELSE 1 END,
                        r.id
               LIMIT 1
           ), '')""",
    ]),
    (8, [
        """ALTER TABLE members ADD COLUMN phone TEXT NOT NULL DEFAULT ''""",
        """UPDATE members
           SET phone = COALESCE((
               SELECT r.phone FROM roster_entries r
               WHERE members.student_id <> ''
                 AND r.student_id = members.student_id
               ORDER BY r.id LIMIT 1
           ), '')""",
        """UPDATE members
           SET phone = COALESCE((
               SELECT r.phone FROM roster_entries r
               WHERE r.name = members.name
               ORDER BY r.id LIMIT 1
           ), phone)
           WHERE phone = ''""",
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
    participates_in_scheduling: bool = True
    studio: str = UNKNOWN_STUDIO
    studio_locked: bool = False
    studios: list[str] = field(default_factory=list)
    studio_positions: dict[str, str] = field(default_factory=dict)
    phone: str = ""


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
        with self._connection() as conn:
            conn.execute("PRAGMA journal_mode = WAL")
            conn.executescript(SCHEMA)
            self._migrate(conn)
            # v6 迁移后按花名册重建多工作室归属；已手工锁定的成员保持原样。
            self._sync_member_studios(conn)

    @staticmethod
    def _migrate(conn: sqlite3.Connection) -> None:
        """按 PRAGMA user_version 增量升级老数据库结构"""
        version = conn.execute("PRAGMA user_version").fetchone()[0]
        for target, statements in MIGRATIONS:
            if version >= target:
                continue
            for sql in statements:
                # 新建数据库的 SCHEMA 已包含新列；老库才需要执行 ALTER TABLE。
                if "ALTER TABLE " in sql and " ADD COLUMN " in sql:
                    table = sql.split("ALTER TABLE ", 1)[1].split()[0]
                    column = sql.split(" ADD COLUMN ", 1)[1].split()[0]
                    columns = {
                        row["name"]
                        for row in conn.execute(f"PRAGMA table_info({table})")
                    }
                    if column in columns:
                        continue
                conn.execute(sql)
            conn.execute(f"PRAGMA user_version = {target}")

    def _open_connection(self) -> sqlite3.Connection:
        """创建连接并设置连接级 PRAGMA。"""
        conn = sqlite3.connect(self.path, timeout=10)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA synchronous = NORMAL")
        return conn

    @contextmanager
    def _connection(self) -> Iterator[sqlite3.Connection]:
        """数据库写事务上下文：异常回滚，结束时始终关闭连接。"""
        conn = self._open_connection()
        try:
            yield conn
            conn.commit()
        except BaseException:
            conn.rollback()
            raise
        finally:
            conn.close()

    # ---------- 备份 / 迁移 ----------

    def backup_to(self, target: str | Path) -> None:
        """把当前数据库完整复制到 target（供「更改数据库位置」使用）。

        必须用 SQLite 的在线备份 API，不能用 shutil.copy2：
        启用 WAL 后最新数据与 schema 可能仍在 -wal 文件里，
        只复制主库文件会得到一个打不开或丢数据的副本。
        """
        target = Path(target)
        target.parent.mkdir(parents=True, exist_ok=True)
        src = self._open_connection()
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

    @staticmethod
    def _roster_studio_lookup(conn: sqlite3.Connection) -> tuple[dict, dict]:
        """构建学号/姓名到 (工作室, 职位) 的匹配索引。"""
        by_student: dict[str, list[tuple[str, str]]] = {}
        by_name: dict[str, list[tuple[str, str]]] = {}
        rows = conn.execute(
            "SELECT studio, position, name, student_id FROM roster_entries "
            "ORDER BY id"
        ).fetchall()
        for row in rows:
            studio = (row["studio"] or UNKNOWN_STUDIO).strip() or UNKNOWN_STUDIO
            position = (row["position"] or "").strip()
            student_id = normalize_student_id(row["student_id"])
            name = normalize_name(row["name"])
            membership = (studio, position)
            for index, source in ((student_id, by_student), (name, by_name)):
                if not index:
                    continue
                options = source.setdefault(index, [])
                existing = next(
                    (i for i, item in enumerate(options)
                     if item[0].casefold() == studio.casefold()),
                    None,
                )
                if existing is None:
                    options.append(membership)
                elif not options[existing][1] and position:
                    options[existing] = membership
        return by_student, by_name

    @staticmethod
    def _roster_phone_lookup(conn: sqlite3.Connection) -> tuple[dict, dict]:
        by_student: dict[str, str] = {}
        by_name: dict[str, str] = {}
        rows = conn.execute(
            "SELECT name, student_id, phone FROM roster_entries ORDER BY id"
        ).fetchall()
        for row in rows:
            phone = (row["phone"] or "").strip()
            if not phone:
                continue
            student_id = normalize_student_id(row["student_id"])
            name = normalize_name(row["name"])
            if student_id:
                by_student.setdefault(student_id, phone)
            if name:
                by_name.setdefault(name, phone)
        return by_student, by_name

    @staticmethod
    def _resolve_roster_phone(
        lookup: tuple[dict, dict],
        student_id: str,
        name: str,
    ) -> str:
        by_student, by_name = lookup
        normalized_id = normalize_student_id(student_id)
        if normalized_id and by_student.get(normalized_id):
            return by_student[normalized_id]
        return by_name.get(normalize_name(name), "")

    @staticmethod
    def _normalize_studios(studios: Iterable[str]) -> list[str]:
        """去重并保持顺序；空集合按未指定工作室处理。"""
        normalized: list[str] = []
        seen: set[str] = set()
        for raw in studios:
            studio = (raw or "").strip()
            key = studio.casefold()
            if not studio or key in seen:
                continue
            seen.add(key)
            normalized.append(studio)
        return normalized or [UNKNOWN_STUDIO]

    @staticmethod
    def _normalize_memberships(
        memberships: Iterable[tuple[str, str]],
    ) -> list[tuple[str, str]]:
        """规范化工作室-职位组合，重复工作室保留首个非空职位。"""
        normalized: list[tuple[str, str]] = []
        indexes: dict[str, int] = {}
        for raw_studio, raw_position in memberships:
            studio = (raw_studio or "").strip()
            position = (raw_position or "").strip()
            if not studio:
                continue
            key = studio.casefold()
            if key in indexes:
                index = indexes[key]
                if not normalized[index][1] and position:
                    normalized[index] = (normalized[index][0], position)
                continue
            indexes[key] = len(normalized)
            normalized.append((studio, position))
        return normalized or [(UNKNOWN_STUDIO, "")]

    @staticmethod
    def _member_memberships_map(
        conn: sqlite3.Connection,
    ) -> dict[int, list[tuple[str, str]]]:
        result: dict[int, list[tuple[str, str]]] = {}
        rows = conn.execute(
            "SELECT member_id, studio, position FROM member_studios "
            "ORDER BY member_id, rowid"
        ).fetchall()
        for row in rows:
            result.setdefault(row["member_id"], []).append(
                (row["studio"], row["position"] or ""))
        return result

    @classmethod
    def _member_studios_map(cls, conn: sqlite3.Connection) -> dict[int, list[str]]:
        return {
            member_id: [studio for studio, _position in memberships]
            for member_id, memberships in cls._member_memberships_map(conn).items()
        }

    @classmethod
    def _member_positions_map(
        cls,
        conn: sqlite3.Connection,
    ) -> dict[int, dict[str, str]]:
        return {
            member_id: {studio: position for studio, position in memberships}
            for member_id, memberships in cls._member_memberships_map(conn).items()
        }

    @staticmethod
    def _replace_member_memberships(
        conn: sqlite3.Connection,
        member_id: int,
        memberships: Iterable[tuple[str, str]],
    ) -> list[tuple[str, str]]:
        normalized = Database._normalize_memberships(memberships)
        conn.execute("DELETE FROM member_studios WHERE member_id = ?", (member_id,))
        conn.executemany(
            """INSERT INTO member_studios (member_id, studio, position)
               VALUES (?, ?, ?)""",
            [(member_id, studio, position) for studio, position in normalized],
        )
        return normalized

    @staticmethod
    def _replace_member_studios(
        conn: sqlite3.Connection,
        member_id: int,
        studios: Iterable[str],
        positions: dict[str, str] | None = None,
    ) -> list[str]:
        positions = {key.casefold(): value for key, value in (positions or {}).items()}
        memberships = [
            (studio, positions.get(studio.casefold(), ""))
            for studio in Database._normalize_studios(studios)
        ]
        Database._replace_member_memberships(conn, member_id, memberships)
        return [studio for studio, _position in memberships]

    @staticmethod
    def _resolve_memberships(
        lookup: tuple[dict, dict],
        student_id: str,
        name: str,
    ) -> list[tuple[str, str]]:
        """学号优先、姓名兜底，返回花名册中全部工作室及职位。"""
        by_student, by_name = lookup
        normalized_id = normalize_student_id(student_id)
        if normalized_id and by_student.get(normalized_id):
            return Database._normalize_memberships(by_student[normalized_id])
        options = by_name.get(normalize_name(name), [])
        return Database._normalize_memberships(options)

    @staticmethod
    def _resolve_studios(
        lookup: tuple[dict, dict],
        student_id: str,
        name: str,
    ) -> list[str]:
        return [
            studio for studio, _position in Database._resolve_memberships(
                lookup, student_id, name)
        ]

    @classmethod
    def _sync_member_studios(cls, conn: sqlite3.Connection) -> int:
        """把未手工锁定的成员按当前花名册重建为工作室-职位归属。"""
        lookup = cls._roster_studio_lookup(conn)
        phone_lookup = cls._roster_phone_lookup(conn)
        existing = cls._member_memberships_map(conn)
        changed = 0
        rows = conn.execute(
            "SELECT id, name, student_id, phone FROM members "
            "WHERE studio_locked = 0"
        ).fetchall()
        for row in rows:
            memberships = cls._resolve_memberships(
                lookup, row["student_id"], row["name"])
            phone = row["phone"] or cls._resolve_roster_phone(
                phone_lookup, row["student_id"], row["name"])
            membership_changed = memberships != existing.get(row["id"], [])
            if membership_changed or phone != (row["phone"] or ""):
                cls._replace_member_memberships(conn, row["id"], memberships)
                conn.execute(
                    "UPDATE members SET studio = ?, phone = ? WHERE id = ?",
                    (memberships[0][0], phone, row["id"]),
                )
                changed += 1
        return changed

    def upsert_member(self, schedule: ParsedSchedule) -> int:
        """新增或更新成员（按 学号+姓名 唯一），并整体替换其课程。"""
        with self._connection() as conn:
            row = conn.execute(
                "SELECT id, studio, studio_locked FROM members "
                "WHERE student_id = ? AND name = ?",
                (schedule.student_id, schedule.name),
            ).fetchone()
            lookup = self._roster_studio_lookup(conn)
            roster_phone = self._resolve_roster_phone(
                self._roster_phone_lookup(conn), schedule.student_id, schedule.name)
            resolved_memberships = self._resolve_memberships(
                lookup, schedule.student_id, schedule.name)
            existing_memberships = self._member_memberships_map(conn)
            if row:
                member_id = row["id"]
                if row["studio_locked"]:
                    memberships = existing_memberships.get(member_id) or [
                        (row["studio"] or UNKNOWN_STUDIO, "")]
                else:
                    memberships = resolved_memberships
                    self._replace_member_memberships(
                        conn, member_id, memberships)
                primary_studio = memberships[0][0]
                conn.execute(
                    """UPDATE members SET term = ?, class_name = ?, major = ?,
                       department = ?, file_name = ?, studio = ?,
                       phone = CASE WHEN phone = '' THEN ? ELSE phone END
                       WHERE id = ?""",
                    (schedule.term, schedule.class_name, schedule.major,
                     schedule.department, schedule.file_name, primary_studio,
                     roster_phone, member_id),
                )
                conn.execute("DELETE FROM courses WHERE member_id = ?", (member_id,))
            else:
                cur = conn.execute(
                    """INSERT INTO members
                       (student_id, name, term, class_name, major, department,
                        file_name, studio, phone)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (schedule.student_id, schedule.name, schedule.term,
                     schedule.class_name, schedule.major, schedule.department,
                     schedule.file_name, resolved_memberships[0][0], roster_phone),
                )
                member_id = cur.lastrowid
                self._replace_member_memberships(
                    conn, member_id, resolved_memberships)

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

    def replace_roster(
        self,
        entries: Iterable[RosterEntry],
        source_file: str = "",
    ) -> tuple[int, int]:
        """整体替换花名册，并自动刷新未手工锁定的成员工作室。"""
        rows = list(entries)
        with self._connection() as conn:
            conn.execute("DELETE FROM roster_entries")
            conn.executemany(
                """INSERT INTO roster_entries
                   (studio, position, name, student_id, phone, college_major,
                    source_file, row_number)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                [(
                    entry.studio or UNKNOWN_STUDIO,
                    entry.position,
                    entry.name,
                    normalize_student_id(entry.student_id),
                    entry.phone,
                    entry.college_major,
                    entry.source_file or source_file,
                    entry.row_number,
                ) for entry in rows],
            )
            changed = self._sync_member_studios(conn)
        return len(rows), changed

    def list_roster_entries(self) -> list[RosterEntry]:
        with self._connection() as conn:
            rows = conn.execute(
                "SELECT * FROM roster_entries ORDER BY id"
            ).fetchall()
            return [RosterEntry(
                studio=r["studio"],
                position=r["position"],
                name=r["name"],
                student_id=r["student_id"],
                phone=r["phone"],
                college_major=r["college_major"],
                source_file=r["source_file"],
                row_number=r["row_number"],
            ) for r in rows]

    def roster_studios(self) -> list[str]:
        with self._connection() as conn:
            rows = conn.execute(
                "SELECT DISTINCT studio FROM roster_entries "
                "WHERE TRIM(studio) != '' ORDER BY studio"
            ).fetchall()
            return [row["studio"] for row in rows]

    def set_members_studios(
        self,
        member_ids: Iterable[int],
        studios: Iterable[str],
        positions: dict[str, str] | None = None,
    ) -> int:
        """批量替换成员的工作室-职位集合并锁定，返回处理人数。"""
        ids = list(dict.fromkeys(int(member_id) for member_id in member_ids))
        if not ids:
            return 0
        normalized_studios = self._normalize_studios(studios)
        requested_positions = {
            studio.casefold(): (position or "").strip()
            for studio, position in (positions or {}).items()
        }
        with self._connection() as conn:
            existing_map = self._member_memberships_map(conn)
            assignments: dict[int, list[tuple[str, str]]] = {}
            for member_id in ids:
                existing_positions = {
                    studio.casefold(): position
                    for studio, position in existing_map.get(member_id, [])
                }
                memberships = [
                    (
                        studio,
                        requested_positions.get(
                            studio.casefold(),
                            existing_positions.get(studio.casefold(), ""),
                        ),
                    )
                    for studio in normalized_studios
                ]
                assignments[member_id] = self._normalize_memberships(memberships)

            self._validate_leader_assignments(conn, ids, assignments)
            for member_id, memberships in assignments.items():
                conn.execute(
                    "UPDATE members SET studio = ?, studio_locked = 1 WHERE id = ?",
                    (memberships[0][0], member_id),
                )
                self._replace_member_memberships(conn, member_id, memberships)
        return len(ids)

    @staticmethod
    def _validate_leader_assignments(
        conn: sqlite3.Connection,
        member_ids: list[int],
        assignments: dict[int, list[tuple[str, str]]],
    ) -> None:
        """同一工作室最多只能指定一名部长。"""
        requested: dict[str, int] = {}
        for memberships in assignments.values():
            for studio, position in memberships:
                if position != POSITION_LEADER:
                    continue
                key = studio.casefold()
                requested[key] = requested.get(key, 0) + 1
                if requested[key] > 1:
                    raise ValueError(f"{studio} 只能指定一名部长")
        if not requested:
            return
        placeholders = ",".join("?" * len(member_ids))
        for studio in requested:
            existing = conn.execute(
                f"""SELECT COUNT(*) AS n FROM member_studios
                    WHERE LOWER(studio) = ? AND position = ?
                      AND member_id NOT IN ({placeholders})""",
                [studio, POSITION_LEADER, *member_ids],
            ).fetchone()["n"]
            if existing:
                raise ValueError(f"{studio} 已有部长，不能重复指定")

    def set_members_studio(self, member_ids: Iterable[int], studio: str) -> int:
        """兼容单工作室调用：批量替换为一个工作室并锁定。"""
        return self.set_members_studios(member_ids, [studio])

    def set_member_studios(
        self,
        member_id: int,
        studios: Iterable[str],
        positions: dict[str, str] | None = None,
    ) -> None:
        """手工设置成员的全部工作室和职位并锁定，重导花名册不会覆盖。"""
        self.set_members_studios([member_id], studios, positions)

    def set_member_studio(self, member_id: int, studio: str) -> None:
        """兼容单工作室调用。"""
        self.set_member_studios(member_id, [studio])

    def update_member_profile(
        self,
        member_id: int,
        *,
        name: str,
        student_id: str,
        term: str = "",
        class_name: str = "",
        major: str = "",
        department: str = "",
        phone: str = "",
    ) -> None:
        """修改成员个人信息；课程与历史排班仍按内部 id 保留。"""
        normalized_name = (name or "").strip()
        if not normalized_name:
            raise ValueError("姓名不能为空")
        normalized_id = normalize_student_id(student_id)
        try:
            with self._connection() as conn:
                cursor = conn.execute(
                    """UPDATE members
                       SET name = ?, student_id = ?, term = ?, class_name = ?,
                           major = ?, department = ?, phone = ?
                       WHERE id = ?""",
                    (normalized_name, normalized_id, (term or "").strip(),
                     (class_name or "").strip(), (major or "").strip(),
                     (department or "").strip(), (phone or "").strip(),
                     member_id),
                )
                if cursor.rowcount == 0:
                    raise ValueError("成员不存在或已被删除")
        except sqlite3.IntegrityError as exc:
            raise ValueError("已存在相同学号和姓名的成员") from exc

    def list_members(self) -> list[Member]:
        with self._connection() as conn:
            rows = conn.execute(
                """SELECT m.*, COUNT(c.id) AS course_count FROM members m
                   LEFT JOIN courses c ON c.member_id = m.id
                   GROUP BY m.id ORDER BY m.id"""
            ).fetchall()
            membership_map = self._member_memberships_map(conn)
            return [Member(
                id=r["id"], name=r["name"], student_id=r["student_id"],
                term=r["term"], class_name=r["class_name"], major=r["major"],
                department=r["department"], file_name=r["file_name"],
                phone=r["phone"],
                course_count=r["course_count"],
                participates_in_scheduling=bool(r["participates_in_scheduling"]),
                studio=r["studio"],
                studio_locked=bool(r["studio_locked"]),
                studios=[studio for studio, _position in membership_map.get(
                    r["id"], [(r["studio"] or UNKNOWN_STUDIO, "")])],
                studio_positions={
                    studio: position
                    for studio, position in membership_map.get(
                        r["id"], [(r["studio"] or UNKNOWN_STUDIO, "")])
                },
            ) for r in rows]

    def set_members_participation(self, settings: dict[int, bool]) -> None:
        """批量设置成员是否参与排班；课表和历史排班记录均保留。"""
        if not settings:
            return
        with self._connection() as conn:
            conn.executemany(
                "UPDATE members SET participates_in_scheduling = ? WHERE id = ?",
                [(1 if participates else 0, member_id)
                 for member_id, participates in settings.items()],
            )

    def set_member_participation(self, member_id: int, participates: bool) -> None:
        """设置单个成员是否参与排班，供单点调用复用批量接口。"""
        self.set_members_participation({member_id: participates})

    def delete_member(self, member_id: int) -> None:
        with self._connection() as conn:
            conn.execute("DELETE FROM members WHERE id = ?", (member_id,))

    def get_member(self, member_id: int) -> Member | None:
        with self._connection() as conn:
            r = conn.execute("SELECT * FROM members WHERE id = ?", (member_id,)).fetchone()
            memberships = self._member_memberships_map(conn).get(
                r["id"], [(r["studio"] or UNKNOWN_STUDIO, "")]) if r else []
            return Member(
                id=r["id"], name=r["name"], student_id=r["student_id"], term=r["term"],
                class_name=r["class_name"], major=r["major"], department=r["department"],
                file_name=r["file_name"], phone=r["phone"],
                participates_in_scheduling=bool(r["participates_in_scheduling"]),
                studio=r["studio"],
                studio_locked=bool(r["studio_locked"]),
                studios=[studio for studio, _position in memberships],
                studio_positions={
                    studio: position for studio, position in memberships
                },
            ) if r else None

    # ---------- 课程 ----------

    def get_courses(self, member_id: int | None = None) -> list[CourseRecord]:
        sql = "SELECT * FROM courses"
        params: tuple = ()
        if member_id is not None:
            sql += " WHERE member_id = ?"
            params = (member_id,)
        sql += " ORDER BY member_id, weekday, course_name"
        with self._connection() as conn:
            rows = conn.execute(sql, params).fetchall()
            return [CourseRecord(
                id=r["id"], member_id=r["member_id"], course_name=r["course_name"],
                teacher=r["teacher"], weekday=r["weekday"],
                week_list=json.loads(r["week_list"]),
                session_list=json.loads(r["session_list"]),
                location=r["location"], weeks_text=r["weeks_text"],
                sessions_text=r["sessions_text"],
            ) for r in rows]

    # ---------- 学期日历覆盖（调休 / 法定假日） ----------

    @staticmethod
    def _coerce_calendar_entry(entry: CalendarEntry | dict | tuple | list) -> CalendarEntry:
        if isinstance(entry, CalendarEntry):
            return entry
        if isinstance(entry, dict):
            return CalendarEntry(
                date=entry["date"],
                override_type=entry["override_type"],
                maps_to_week=entry.get("maps_to_week"),
                maps_to_weekday=entry.get("maps_to_weekday"),
                note=entry.get("note", ""),
            )
        if isinstance(entry, (tuple, list)):
            if len(entry) < 2:
                raise ValueError("日历条目至少需要日期和类型")
            values = list(entry) + [None, None, ""]
            return CalendarEntry(
                date=values[0], override_type=values[1],
                maps_to_week=values[2], maps_to_weekday=values[3],
                note=values[4] or "",
            )
        raise TypeError("日历条目必须是 CalendarEntry、dict 或 tuple/list")

    _UPSERT_CALENDAR_SQL = """INSERT INTO term_calendar
        (date, override_type, maps_to_week, maps_to_weekday, note)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(date) DO UPDATE SET
            override_type = excluded.override_type,
            maps_to_week = excluded.maps_to_week,
            maps_to_weekday = excluded.maps_to_weekday,
            note = excluded.note"""

    @staticmethod
    def _calendar_values(
        rows: Iterable[CalendarEntry],
    ) -> list[tuple[str, str, int | None, int | None, str]]:
        return [
            (entry.date.isoformat(), entry.override_type, entry.maps_to_week,
             entry.maps_to_weekday, entry.note)
            for entry in rows
        ]

    def upsert_calendar(
        self,
        entries: Iterable[CalendarEntry | dict | tuple | list],
    ) -> None:
        """按日期新增或覆盖学期日历项。"""
        rows = [self._coerce_calendar_entry(entry) for entry in entries]
        with self._connection() as conn:
            conn.executemany(self._UPSERT_CALENDAR_SQL, self._calendar_values(rows))

    def replace_calendar(
        self,
        entries: Iterable[CalendarEntry | dict | tuple | list],
    ) -> None:
        """原子替换全部学期日历项，失败时保留原数据。"""
        rows = [self._coerce_calendar_entry(entry) for entry in entries]
        values = self._calendar_values(rows)
        with self._connection() as conn:
            conn.execute("DELETE FROM term_calendar")
            if values:
                conn.executemany(self._UPSERT_CALENDAR_SQL, values)

    def list_calendar(self) -> list[CalendarEntry]:
        with self._connection() as conn:
            rows = conn.execute(
                "SELECT * FROM term_calendar ORDER BY date, override_type"
            ).fetchall()
            return [CalendarEntry(
                date=date.fromisoformat(r["date"]),
                override_type=r["override_type"],
                maps_to_week=r["maps_to_week"],
                maps_to_weekday=r["maps_to_weekday"],
                note=r["note"],
            ) for r in rows]

    def clear_calendar(self) -> None:
        with self._connection() as conn:
            conn.execute("DELETE FROM term_calendar")

    def validate_calendar(
        self,
        term_start: date | None = None,
        *,
        entries: Iterable[CalendarEntry | dict | tuple | list] | None = None,
        allow_unpaired_off: bool = False,
    ) -> None:
        """校验日历覆盖；非法项抛 ``ValueError``，不会修改数据库。"""
        rows = (self.list_calendar() if entries is None else
                [self._coerce_calendar_entry(entry) for entry in entries])
        validate_calendar_entries(
            rows, term_start,
            allow_unpaired_off=allow_unpaired_off,
        )

    # ---------- 值班安排 ----------

    def clear_assignments(self) -> None:
        with self._connection() as conn:
            conn.execute("DELETE FROM duty_assignments")

    def delete_assignments_for_weeks(self, weeks: list[int]) -> None:
        """删除指定周的值班安排（按周增量重排时清空所选范围）"""
        if not weeks:
            return
        ph = ",".join("?" * len(weeks))
        with self._connection() as conn:
            conn.execute(f"DELETE FROM duty_assignments WHERE week IN ({ph})", weeks)

    def save_assignments(self, assignments: list[Assignment]) -> None:
        """幂等写入：唯一键冲突时不做任何改动。
        不再使用 INSERT OR REPLACE（那会先删后插，使自增 id 每轮膨胀）。"""
        with self._connection() as conn:
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
        with self._connection() as conn:
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
        with self._connection() as conn:
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
        with self._connection() as conn:
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
        with self._connection() as conn:
            conn.execute(
                """INSERT OR REPLACE INTO leaves (member_id, week, weekday, reason)
                   VALUES (?, ?, ?, ?)""",
                (member_id, week, weekday, reason),
            )

    def remove_leave(self, leave_id: int) -> None:
        with self._connection() as conn:
            conn.execute("DELETE FROM leaves WHERE id = ?", (leave_id,))

    def list_leaves(self, member_id: int | None = None) -> list[Leave]:
        sql = "SELECT * FROM leaves"
        params: tuple = ()
        if member_id is not None:
            sql += " WHERE member_id = ?"
            params = (member_id,)
        sql += " ORDER BY week, weekday, member_id"
        with self._connection() as conn:
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
        with self._connection() as conn:
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
        with self._connection() as conn:
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
        with self._connection() as conn:
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
        with self._connection() as conn:
            rows = conn.execute(sql, params).fetchall()
            return [SpecialArrangement(
                id=r["id"], member_id=r["member_id"], week_start=r["week_start"],
                week_end=r["week_end"], weekday=r["weekday"],
                session_list=json.loads(r["session_list"]), reason=r["reason"],
            ) for r in rows]
