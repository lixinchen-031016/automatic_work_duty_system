"""缺口成因诊断：区分「排不了」（课程/请假冲突）与「排不下」（上限不足）"""

from __future__ import annotations

from collections import Counter

from duty_system.database import CourseRecord, Member
from duty_system.exporter import build_gap_df
from duty_system.scheduler import (
    GAP_DAY_CAP, GAP_ELIGIBLE, GAP_MIXED_CAP, GAP_NO_FREE, GAP_WEEK_CAP,
    ScheduleConfig, build_busy_map, capacity_advice, compute_gaps, diagnose_gaps,
    generate_schedule, summarize_gap_causes,
)


def member(i: int) -> Member:
    return Member(id=i, name=f"成员{i}", student_id=str(i), term="", class_name="",
                  major="", department="", file_name="")


def course(mid: int, weekday: int, sessions: list[int], weeks: list[int] | None = None):
    return CourseRecord(id=mid * 100 + weekday, member_id=mid, course_name="课程", teacher="",
                        weekday=weekday, week_list=weeks or list(range(1, 19)),
                        session_list=sessions, location="", weeks_text="", sessions_text="")


def test_all_busy_slot_is_reported_as_no_free_member() -> None:
    """全员在该时段有课 -> 归因「成员均有课或请假」，而不是上限问题"""
    members = [member(i) for i in (1, 2, 3)]
    # 3 人第 1 周周一 1-2 节全都有课，其余时段空闲
    courses = [course(m.id, 1, [1, 2]) for m in members]
    config = ScheduleConfig(weeks=range(1, 2), weekdays=[1], blocks=[1], per_slot=1)
    result = generate_schedule(members, courses, config)

    assert result.gaps == [(1, 1, 1)], "该时段应必然缺人"
    diags = diagnose_gaps(members, build_busy_map(members, courses), None,
                          result.assignments, config)
    assert len(diags) == 1
    d = diags[0]
    assert d.cause == GAP_NO_FREE
    assert d.blocked_by_course == 3 and d.free_members == 0 and d.eligible == 0
    assert "增加成员" in capacity_advice(summarize_gap_causes(diags), 1), \
        "全员冲突类缺口只能靠增加成员，建议里应说明"


def test_capacity_limited_gap_and_advice_is_actionable() -> None:
    """无课成员被每天/每周上限挡住 -> 归因上限，且提高上限后缺口应消失"""
    members = [member(i) for i in (1, 2, 3)]
    # 3 人全天空闲；每天 3 个时段 × 每时段 1 人 = 3 人次，正好等于每天上限
    # 再加一个时段就会超出每天上限 -> 该时段必然缺人
    config = ScheduleConfig(weeks=range(1, 2), weekdays=[1], blocks=[1, 2, 3, 4],
                            per_slot=1, max_per_day=1, max_per_week=3)
    result = generate_schedule(members, [], config)

    diags = diagnose_gaps(members, {}, None, result.assignments, config)
    causes = {d.cause for d in diags}
    assert causes, "应存在缺口"
    assert causes <= {GAP_DAY_CAP, GAP_MIXED_CAP, GAP_WEEK_CAP}, f"应归因为上限不足: {causes}"
    assert all(d.free_members > 0 for d in diags), "这些时段其实有空闲成员"
    assert all(d.blocked_by_course == 0 for d in diags)

    advice = capacity_advice(summarize_gap_causes(diags), 1)
    assert "提高上限" in advice, "上限类缺口应给出可操作建议"

    # 提高每天/每周上限后，同样的成员应能排满 —— 验证建议确实有效
    relaxed = ScheduleConfig(weeks=range(1, 2), weekdays=[1], blocks=[1, 2, 3, 4],
                             per_slot=1, max_per_day=4, max_per_week=4)
    filled = generate_schedule(members, [], relaxed)
    assert filled.gaps == [], f"放宽上限后仍缺人: {filled.gaps}"


def test_no_missed_slots_after_generation() -> None:
    """一次完整生成后不应存在「仍有可用成员却没排」的时段"""
    import random

    rng = random.Random(17)
    members = [member(i) for i in range(1, 13)]
    courses = [
        CourseRecord(id=len(members) * 100 + i, member_id=rng.choice(members).id,
                     course_name="课程", teacher="", weekday=rng.randint(1, 7),
                     week_list=sorted(rng.sample(range(1, 19), k=8)),
                     session_list=sorted(rng.sample(range(1, 12), k=2)),
                     location="", weeks_text="", sessions_text="")
        for i in range(120)]
    config = ScheduleConfig(weeks=range(1, 5), per_slot=1, max_per_week=3, max_per_day=1)
    result = generate_schedule(members, courses, config)

    busy = build_busy_map(members, courses)
    diags = diagnose_gaps(members, busy, None, result.assignments, config)
    missed = [d.slot for d in diags if d.cause == GAP_ELIGIBLE]
    assert not missed, f"存在仍可安排却被漏掉的时段: {missed[:5]}"

    # 诊断覆盖的时段必须与 gaps 完全一致（口径不漂移）
    assert [d.slot for d in diags] == compute_gaps(result.assignments, config)


def test_diagnosis_uses_same_context_as_scheduler() -> None:
    """诊断的分类结果必须与逐成员查看 reason() 的结论一致"""
    from duty_system.scheduler import ScheduleContext, REASON_COURSE, REASON_DAY, REASON_WEEK

    members = [member(i) for i in (1, 2, 3, 4)]
    courses = [course(1, 1, [1, 2]), course(2, 1, [3, 4])]
    config = ScheduleConfig(weeks=range(1, 2), weekdays=[1], blocks=[1, 2, 3],
                            per_slot=1, max_per_week=2, max_per_day=1)
    result = generate_schedule(members, courses, config)

    busy = build_busy_map(members, courses)
    ctx = ScheduleContext(members, busy, set(), config)
    ctx.load(result.assignments)
    diags = diagnose_gaps(members, busy, None, result.assignments, config)

    for d in diags:
        reason_counts = Counter(ctx.reason(m.id, d.week, d.weekday, d.block) for m in members)
        assert d.blocked_by_course == reason_counts[REASON_COURSE]
        assert d.blocked_by_week_cap == reason_counts[REASON_WEEK]
        # 每天上限与「该时段已排」互斥：缺口时段的已排人数 < per_slot，
        # 因此 reason 里不会出现「该时段已值班」
        assert d.blocked_by_day_cap == reason_counts[REASON_DAY]


def test_gap_dataframe_columns() -> None:
    """缺口表展示列：周次/星期/时段/主要原因/无课人数"""
    members = [member(i) for i in (1, 2)]
    courses = [course(m.id, 1, [1, 2]) for m in members]
    config = ScheduleConfig(weeks=range(1, 2), weekdays=[1], blocks=[1], per_slot=1)
    result = generate_schedule(members, courses, config)
    diags = diagnose_gaps(members, build_busy_map(members, courses), None,
                          result.assignments, config)

    df = build_gap_df(diags)
    assert list(df.columns) == ["周次", "星期", "时段", "主要原因", "无课人数"]
    assert len(df) == len(result.gaps)
    assert df.iloc[0]["周次"] == "第1周" and df.iloc[0]["主要原因"] == GAP_NO_FREE
    assert build_gap_df([]).empty, "无缺口时也应返回带列名的空表"
    assert capacity_advice({}, 1) == "", "无缺口时不应给出建议"


# --------------------------------------------------------------------------- #
# 解析 -> 排班 链路：单双周必须真正影响可用性
# --------------------------------------------------------------------------- #

def test_odd_week_course_blocks_only_odd_weeks() -> None:
    """单周课程只占用单周：双周同一时段应可值班（修复前该课程会被整条丢弃）"""
    from pathlib import Path

    from duty_system.database import Database
    from duty_system.parser import ParsedSchedule, parse_cell
    from duty_system.scheduler import ScheduleConfig, build_busy_map, generate_schedule
    from duty_system.parser import BLOCK_SESSIONS

    odd_course = parse_cell("\n课程A\n张三(讲师)\n1-16单周([周])[01-02节]\n教1\n", 1)
    assert odd_course and odd_course[0].week_list == [1, 3, 5, 7, 9, 11, 13, 15]

    import tempfile

    db = Database(Path(tempfile.mkdtemp()) / "t.db")
    schedule = ParsedSchedule(name="测试甲", student_id="1", term="t", class_name="C",
                              major="M", department="D", file_name="f")
    schedule.courses = odd_course
    db.upsert_member(schedule)
    members = db.list_members()
    courses = db.get_courses()
    busy = build_busy_map(members, courses)
    mid = members[0].id

    assert all((1, 1, sec) in busy[mid] for sec in BLOCK_SESSIONS[1]), "第1周（单周）应有课"
    assert not any((2, 1, sec) in busy[mid] for sec in BLOCK_SESSIONS[1]), "第2周（双周）应空闲"

    config = ScheduleConfig(weeks=range(1, 5), weekdays=[1], blocks=[1],
                            per_slot=1, max_per_week=5)
    result = generate_schedule(members, courses, config)
    weeks = sorted(a.week for a in result.assignments if a.member_id == mid)
    assert 1 not in weeks, "单周有课的时段不应安排值班"
    assert {2, 4} <= set(weeks), f"双周应可值班，实际排到 {weeks}"


def test_real_sample_flows_into_scheduling() -> None:
    """真实样例课表入库后可直接参与排班，且不存在课程冲突"""
    import tempfile
    from pathlib import Path

    from duty_system.database import Database
    from duty_system.parser import BLOCK_SESSIONS, parse_schedule_path
    from duty_system.scheduler import ScheduleConfig, build_busy_map, generate_schedule

    sample = (Path(__file__).parent / "samples" / "desensitized"
              / "学生个人课表_9999800598.xls")
    db = Database(Path(tempfile.mkdtemp()) / "real.db")
    parsed = parse_schedule_path(sample)
    db.upsert_member(parsed)
    members = db.list_members()
    courses = db.get_courses()
    assert members[0].course_count == len(parsed.courses)

    config = ScheduleConfig(weeks=range(1, 19), per_slot=1, max_per_week=3, max_per_day=1)
    result = generate_schedule(members, courses, config)
    busy = build_busy_map(members, courses)
    for a in result.assignments:
        for sec in BLOCK_SESSIONS[a.block]:
            assert (a.week, a.weekday, sec) not in busy.get(a.member_id, ()), \
                f"{a.member_name} 第{a.week}周{a.weekday} 与课程冲突"
