"""排班算法测试：约束单一入口、缺口修复、稀缺度排序、可复现性"""

from __future__ import annotations

import random
from collections import Counter

from duty_system.database import CourseRecord, Member
from duty_system.parser import BLOCK_SESSIONS
from duty_system.scheduler import (
    REASON_COURSE, REASON_DAY, REASON_LEAVE, REASON_SLOT, REASON_WEEK,
    ScheduleConfig, ScheduleContext, build_busy_map, compute_gaps, generate_schedule,
    rebuild_member_stats, replacement_candidates,
)


def make_members(n: int) -> list[Member]:
    return [Member(id=i, name=f"成员{i}", student_id=str(i), term="", class_name="",
                   major="", department="", file_name="") for i in range(1, n + 1)]


def make_courses(members: list[Member], per_member: int = 30, weeks: int = 18,
                 seed: int = 11) -> list[CourseRecord]:
    """构造随机课表：每人若干门课，覆盖不同星期/节次，制造忙闲差异"""
    rng = random.Random(seed)
    courses: list[CourseRecord] = []
    for m in members:
        for i in range(per_member):
            courses.append(CourseRecord(
                id=len(courses) + 1, member_id=m.id, course_name=f"课程{i}", teacher="",
                weekday=rng.randint(1, 7),
                week_list=sorted(rng.sample(range(1, weeks + 1), k=rng.randint(3, weeks))),
                session_list=sorted(rng.sample(range(1, 12), k=2)),
                location="", weeks_text="", sessions_text=""))
    return courses


def test_reason_single_source_of_truth() -> None:
    """自动排班与手动微调必须给出完全一致的约束结论"""
    members = make_members(12)
    courses = make_courses(members, seed=5)
    config = ScheduleConfig(weeks=range(1, 4), per_slot=1, max_per_week=3, max_per_day=1)
    result = generate_schedule(members, courses, config)

    busy = build_busy_map(members, courses)
    ctx = ScheduleContext(members, busy, set(), config)
    ctx.load(result.assignments)

    # 每个候选的判定都与 replacement_candidates 一致（后者供界面微调使用）
    for w in config.weeks:
        for d in config.weekdays:
            for b in config.blocks:
                cands = {m.id: reason for m, reason in replacement_candidates(
                    members, busy, set(), result.assignments, w, d, b,
                    config.max_per_week, config.max_per_day)}
                for m in members:
                    expected = ctx.reason(m.id, w, d, b)
                    if m.id in cands:
                        actual = cands[m.id] or None
                        assert actual == expected, \
                            f"{m.name} 第{w}周{d}-{b}: 微调判定 {actual} != 排班判定 {expected}"

    # 已排班次不违反任何硬约束
    for a in result.assignments:
        for s in BLOCK_SESSIONS[a.block]:
            assert (a.week, a.weekday, s) not in busy.get(a.member_id, ()), "排班与课程冲突"

    # assign / unassign 完全对称
    before = (dict(ctx.total), dict(ctx.week_cnt), dict(ctx.day_cnt), dict(ctx.slot_cnt))
    a = result.assignments[0]
    ctx.unassign(a.member_id, a.week, a.weekday, a.block)
    assert ctx.total[a.member_id] == before[0].get(a.member_id, 0) - 1
    ctx.assign(a.member_id, a.week, a.weekday, a.block)
    assert ctx.total[a.member_id] == before[0].get(a.member_id, 0), "assign/unassign 不对称"


def test_gap_repair_reduces_gaps() -> None:
    """缺口修复必须真正减少缺口，且不破坏任何硬约束与已排结果"""
    members = make_members(20)
    courses = make_courses(members, seed=3)
    config = ScheduleConfig(weeks=range(1, 19), per_slot=2, max_per_week=3, max_per_day=1)

    off = generate_schedule(members, courses, config, repair=False)
    on = generate_schedule(members, courses, config, repair=True)

    assert len(on.gaps) <= len(off.gaps), "修复后缺口不应变多"
    assert on.repaired == len(off.gaps) - len(on.gaps), \
        f"repaired={on.repaired} 与实际减少的缺口数不符"
    assert on.repaired > 0, "该场景（人手紧张）应能修复出缺口"

    busy = build_busy_map(members, courses)
    week_cnt = Counter((a.member_id, a.week) for a in on.assignments)
    day_cnt = Counter((a.member_id, a.week, a.weekday) for a in on.assignments)
    slot_cnt = Counter((a.week, a.weekday, a.block) for a in on.assignments)
    assert all(v <= config.max_per_week for v in week_cnt.values()), "修复后超出每周上限"
    assert all(v <= config.max_per_day for v in day_cnt.values()), "修复后超出每天上限"
    assert all(v <= config.per_slot for v in slot_cnt.values()), "修复后时段人数超额"
    for a in on.assignments:
        for s in BLOCK_SESSIONS[a.block]:
            assert (a.week, a.weekday, s) not in busy.get(a.member_id, ()), "修复后出现课程冲突"


def test_gaps_are_consistent_with_assignments() -> None:
    """缺口口径 = 排班范围内人数不足 per_slot 的时段（生成/恢复/微调共用）"""
    members = make_members(6)          # 人少 -> 必然有缺口
    courses = make_courses(members, seed=8)
    config = ScheduleConfig(weeks=range(1, 4), per_slot=2, max_per_week=3, max_per_day=1)
    result = generate_schedule(members, courses, config)

    cnt = Counter((a.week, a.weekday, a.block) for a in result.assignments)
    expected = sorted((w, d, b) for w in config.weeks for d in config.weekdays
                      for b in config.blocks if cnt[(w, d, b)] < config.per_slot)
    assert result.gaps == expected, "缺口与实际排班人数不一致"
    assert result.gaps == compute_gaps(result.assignments, config), "compute_gaps 口径不一致"

    # 成员/课程/参数完全相同的两次生成结果一致（同种子可复现）
    again = generate_schedule(members, courses, config)
    key = lambda a: (a.week, a.weekday, a.block, a.member_id)  # noqa: E731
    assert sorted(key(a) for a in again.assignments) == sorted(key(a) for a in result.assignments), \
        "固定种子应可复现同一份排班"


def test_scarcity_ordering_improves_coverage() -> None:
    """次级排序键（按静态可用人数先排抢手时段）应减少缺口"""
    members = make_members(20)
    courses = make_courses(members, seed=3)
    config = ScheduleConfig(weeks=range(1, 19), per_slot=2, max_per_week=3, max_per_day=1)

    def generate_with_scarcity(enabled: bool) -> int:
        """enabled=False 时把稀缺度键屏蔽掉，其余逻辑完全一致"""
        import duty_system.scheduler as sched

        original = sched._availability_counts
        try:
            if not enabled:
                sched._availability_counts = (
                    lambda members, busy, leave_set, week, weekdays, blocks:
                    {slot: 0 for slot in ((d, b) for d in weekdays for b in blocks)})
            return len(generate_schedule(members, courses, config).gaps)
        finally:
            sched._availability_counts = original

    gaps_plain = generate_with_scarcity(False)
    gaps_scarce = generate_with_scarcity(True)
    assert gaps_scarce <= gaps_plain, \
        f"稀缺度排序不应让覆盖变差（{gaps_plain} -> {gaps_scarce}）"


def test_repair_respects_leave_and_week_bounds() -> None:
    """修复阶段同样不能违反请假与每周上限"""
    members = make_members(8)
    courses = make_courses(members, seed=13, per_member=26)
    leave_set = {(members[0].id, 1, 1), (members[1].id, 1, 2)}

    class FakeLeave:
        def __init__(self, member_id, week, weekday):
            self.member_id, self.week, self.weekday, self.reason = member_id, week, weekday, ""

    config = ScheduleConfig(weeks=range(1, 3), per_slot=2, max_per_week=3, max_per_day=1)
    result = generate_schedule(members, courses, config,
                               leaves=[FakeLeave(*x) for x in leave_set])
    for a in result.assignments:
        assert (a.member_id, a.week, a.weekday) not in leave_set, "修复后安排了请假成员"
    week_cnt = Counter((a.member_id, a.week) for a in result.assignments)
    assert all(v <= config.max_per_week for v in week_cnt.values()), "修复后超出每周上限"
    stats = rebuild_member_stats(members, result.assignments)
    assert sum(s["total"] for s in stats.values()) == len(result.assignments)


def test_reason_precedence_is_stable() -> None:
    """约束判定顺序固定：课程 > 请假 > 该时段已排 > 每天上限 > 每周上限

    界面微调列表按该顺序展示不可用原因，顺序变化会改变用户看到的提示文案。
    """
    members = make_members(5)
    m1, m2, m3, m4, m5 = members
    busy = {m1.id: {(1, 1, 1)}}                      # m1 该时段第 1 节有课
    config = ScheduleConfig(weeks=range(1, 2), per_slot=1, max_per_week=1, max_per_day=1)
    ctx = ScheduleContext(members, busy, {(m2.id, 1, 1)}, config)   # m2 请假
    ctx.assign(m3.id, 1, 1, 5)      # m3 同一天的另一个时段 -> 每天上限
    ctx.assign(m4.id, 1, 2, 1)      # m4 本周已值 1 次 -> 每周上限
    ctx.assign(m1.id, 1, 3, 1)      # m1 同时「有课 + 已满周上限」

    assert ctx.reason(m1.id, 1, 1, 1) == REASON_COURSE, "课程冲突应优先于其他原因"
    assert ctx.reason(m2.id, 1, 1, 1) == REASON_LEAVE, "请假判定错误"
    assert ctx.reason(m3.id, 1, 1, 1) == REASON_DAY, "当天已值班应被拒"
    assert ctx.reason(m4.id, 1, 1, 1) == REASON_WEEK, "每周上限应被拒"
    assert ctx.reason(m5.id, 1, 1, 1) is None, "无冲突成员应可安排"

    # 同一时段的重复安排单独给出「该时段已值班」，便于微调界面区分
    ctx.assign(m5.id, 1, 1, 1)
    assert ctx.reason(m5.id, 1, 1, 1) == REASON_SLOT, "重复安排同一时段应被拒"
    allowed = {None, REASON_COURSE, REASON_DAY, REASON_LEAVE, REASON_SLOT, REASON_WEEK}
    assert {ctx.reason(m.id, 1, 1, 1) for m in members} <= allowed, "出现未定义的原因"
