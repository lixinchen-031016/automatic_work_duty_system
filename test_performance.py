"""性能回归测试：为关键路径设定宽松上限，捕捉数量级级别的退化

阈值刻意留出 5~10 倍余量（CI 机器性能差异大），只拦住「从毫秒变秒级」这类
真正影响体验的退化，不做精细基准。
"""

from __future__ import annotations

import random
import time
from pathlib import Path

import pytest


from duty_system.database import CourseRecord, Database, Member, Assignment
from duty_system.gantt import build_availability
from duty_system.parser import ParsedSchedule, Course
from duty_system.scheduler import ScheduleConfig, build_busy_map, generate_schedule

BUDGET_S = {
    "generate_300": 5.0,
    "generate_dense_20": 3.0,
    "availability_300": 2.0,
    "courses_query_300": 2.0,
    "stats_300": 0.5,
}


def timed(fn):
    t0 = time.perf_counter()
    result = fn()
    return result, time.perf_counter() - t0


def make_dataset(n_members: int, per_member: int = 40, weeks: int = 18, seed: int = 5):
    rng = random.Random(seed)
    members: list[Member] = []
    courses: list[CourseRecord] = []
    for i in range(1, n_members + 1):
        members.append(Member(id=i, name=f"成员{i}", student_id=str(i), term="",
                              class_name="", major="", department="", file_name=""))
        for c in range(per_member):
            courses.append(CourseRecord(
                id=len(courses) + 1, member_id=i, course_name=f"课程{c}", teacher="",
                weekday=rng.randint(1, 7),
                week_list=sorted(rng.sample(range(1, weeks + 1), k=rng.randint(4, weeks))),
                session_list=sorted(rng.sample(range(1, 12), k=2)),
                location="", weeks_text="", sessions_text=""))
    return members, courses


def test_generate_schedule_scales() -> None:
    """300 名成员 / 12000 门课的全学期排班应在预算内完成"""
    members, courses = make_dataset(300)
    config = ScheduleConfig(weeks=range(1, 19), per_slot=1)
    result, elapsed = timed(lambda: generate_schedule(members, courses, config))
    assert result.assignments, "应产出排班"
    assert elapsed < BUDGET_S["generate_300"], \
        f"300 人排班耗时 {elapsed:.2f}s，超出预算 {BUDGET_S['generate_300']}s"


def test_generate_dense_repair_scales() -> None:
    """人手紧张（缺口多 -> 触发大量修复尝试）时仍需在预算内完成"""
    members, courses = make_dataset(20, seed=3)
    config = ScheduleConfig(weeks=range(1, 19), per_slot=2)
    result, elapsed = timed(lambda: generate_schedule(members, courses, config))
    assert elapsed < BUDGET_S["generate_dense_20"], \
        f"密集修复耗时 {elapsed:.2f}s，超出预算 {BUDGET_S['generate_dense_20']}s"
    assert result.gaps == sorted(result.gaps)


def test_availability_and_stats_scale() -> None:
    """甘特矩阵与统计重算应保持线性规模"""
    members, courses = make_dataset(300)
    config = ScheduleConfig(weeks=range(1, 19), per_slot=1)
    result = generate_schedule(members, courses, config)

    m, elapsed = timed(lambda: build_availability(
        members, courses, 5, [1, 2, 3, 4, 5], [1, 2, 3, 4, 5],
        assignments=result.assignments))
    assert m.member_count == 300
    assert elapsed < BUDGET_S["availability_300"], \
        f"甘特矩阵构建耗时 {elapsed:.2f}s，超出预算 {BUDGET_S['availability_300']}s"

    from duty_system.scheduler import rebuild_member_stats
    stats, elapsed = timed(lambda: rebuild_member_stats(members, result.assignments))
    assert sum(s["total"] for s in stats.values()) == len(result.assignments)
    assert elapsed < BUDGET_S["stats_300"], \
        f"统计重算耗时 {elapsed:.3f}s，超出预算 {BUDGET_S['stats_300']}s"


def test_busy_map_and_query_cost(tmp_path: Path) -> None:
    """课表查询（含 JSON 解析）与忙时表构建不应成为瓶颈"""
    db = Database(tmp_path / "perf.db")
    rng = random.Random(2)
    for i in range(150):
        db.upsert_member(ParsedSchedule(
            name=f"成员{i}", student_id=str(i), term="2026-2027-1",
            courses=[Course(course_name=f"课程{c}", teacher="教师",
                            weekday=rng.randint(1, 7), weeks_text="1-18",
                            week_list=list(range(1, 19)), sessions_text="01-02",
                            session_list=[1, 2]) for c in range(40)]))
    members = db.list_members()

    courses, elapsed = timed(db.get_courses)
    assert len(courses) == 150 * 40
    assert elapsed < BUDGET_S["courses_query_300"], \
        f"全量课表查询耗时 {elapsed:.2f}s，超出预算 {BUDGET_S['courses_query_300']}s"

    _, elapsed = timed(lambda: build_busy_map(members, courses))
    assert elapsed < BUDGET_S["availability_300"], \
        f"忙时表构建耗时 {elapsed:.2f}s，超出预算 {BUDGET_S['availability_300']}s"


def test_assignment_write_scales(tmp_path: Path) -> None:
    """批量写入 5000 条安排应远快于 1 秒（WAL + 幂等键）"""
    db = Database(tmp_path / "write.db")
    for i in range(300):
        db.upsert_member(ParsedSchedule(name=f"成员{i}", student_id=str(i), courses=[]))
    ids = [m.id for m in db.list_members()]
    assignments = [Assignment(w, d, b, ids[(w * 25 + d * 5 + b) % 300])
                   for w in range(1, 19) for d in range(1, 6) for b in range(1, 6)]

    _, elapsed = timed(lambda: db.save_assignments(assignments))
    assert len(db.load_assignments()) >= 450
    assert elapsed < 2.0, f"写入 {len(assignments)} 条安排耗时 {elapsed:.2f}s"


# --------------------------------------------------------------------------- #
# 界面刷新预算（offscreen）：锁住已优化的路径，防止被改回慢实现
# --------------------------------------------------------------------------- #

def _build_window(tmp_path, n_members: int):
    """建一个带排班结果的窗口（界面测试需要 QApplication）"""
    import os

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    pytest.importorskip("PySide6.QtWidgets")
    from PySide6.QtCore import QSettings
    from PySide6.QtWidgets import QApplication

    import app as appmod
    from duty_system.parser import Course as ParserCourse

    app = QApplication.instance() or QApplication([])
    QSettings.setDefaultFormat(QSettings.IniFormat)
    QSettings.setPath(QSettings.IniFormat, QSettings.UserScope, str(tmp_path / "qs"))

    win = appmod.MainWindow(db_path=tmp_path / f"perf_{n_members}.db")
    rng = random.Random(11)
    for i in range(n_members):
        win.db.upsert_member(ParsedSchedule(
            name=f"成员{i:03d}", student_id=str(i), term="t", class_name="C",
            major="M", department="D", file_name="f.xls",
            courses=[ParserCourse(course_name=f"课程{c}", teacher="T",
                                  weekday=rng.randint(1, 7), weeks_text="1-18",
                                  week_list=list(range(1, 19)), sessions_text="01-02",
                                  session_list=sorted(rng.sample(range(1, 12), 2)))
                     for c in range(40)]))
    win.invalidate_cache()
    win.refresh_members()
    win.week_to.setValue(18)
    win.generate()
    deadline = time.time() + 60
    while win.busy and time.time() < deadline:
        app.processEvents()
        time.sleep(0.005)
    app.processEvents()
    assert win.result is not None and win.result.assignments
    return app, win


def test_ui_refresh_budget_with_300_members(tmp_path) -> None:
    """300 名成员时，排班页与甘特图刷新应保持在同一量级（各 < 250ms）

    覆盖两条已优化的路径：表格单元格复用（fill_table / _fill_gantt_table）
    与忙时表、单元格状态缓存。
    """
    app, win = _build_window(tmp_path, 300)

    win.refresh_schedule_tabs()
    _, tabs = timed(win.refresh_schedule_tabs)

    win.gantt_week.setValue(5)
    win.refresh_gantt()
    _, gantt = timed(win.refresh_gantt)

    win.close()
    assert tabs < 0.25, f"排班页刷新耗时 {tabs:.3f}s，超出预算"
    assert gantt < 0.25, f"甘特图刷新耗时 {gantt:.3f}s，超出预算"


def test_gantt_row_toggle_is_cheap(tmp_path) -> None:
    """切换「查看周次」不应重建整表的单元格对象（缓存 + 复用生效）"""
    app, win = _build_window(tmp_path, 120)
    win.refresh_gantt()
    before = win.gantt_table.item(0, 0)

    win.gantt_week.setValue(7)
    win.refresh_gantt()
    assert win.gantt_table.item(0, 0) is before, "换周不应重建单元格对象"

    _, elapsed = timed(lambda: [win.refresh_gantt() for _ in range(3)])
    win.close()
    assert elapsed / 3 < 0.15, f"单次甘特刷新 {elapsed / 3:.3f}s，超出预算"


def test_gap_diagnosis_is_cached(tmp_path) -> None:
    """缺口诊断（逐时段逐成员判定）应缓存，重复刷新不重复计算"""
    app, win = _build_window(tmp_path, 60)
    win.refresh_schedule_tabs()
    first = win._gap_cache
    assert first is not None, "首次刷新应完成诊断"

    win.refresh_schedule_tabs()
    assert win._gap_cache is first, "参数与结果未变时不应重复诊断"
    win.close()
