"""值班排班算法

硬约束：
  - 值班时刻内（时段块的每一节）不能有任何课程；
  - 每人每天最多值一次（max_per_day，默认 1）；
  - 每人每周值班不超过 max_per_week 次；
  - 请假/临时占用的 (成员, 周, 星期) 不安排值班。
软目标：按 累计总次数 -> 本周次数 -> 当天次数 三级均衡贪心分配，
        随机种子固定可复现，平手时随机打散。
"""

from __future__ import annotations

import random
from collections import Counter, defaultdict
from dataclasses import dataclass, field

from .database import Assignment, CourseRecord, Leave, Member
from .parser import BLOCK_LABELS, BLOCK_SESSIONS, WEEKDAY_LABELS


@dataclass
class ScheduleConfig:
    weeks: range = range(1, 19)        # 值班周范围，如 range(1, 19)
    weekdays: list[int] = field(default_factory=lambda: [1, 2, 3, 4, 5])  # 值班星期
    blocks: list[int] = field(default_factory=lambda: [1, 2, 3, 4, 5])   # 值班时段块
    per_slot: int = 1                   # 每个时段需要值班人数
    max_per_week: int = 3               # 每人每周值班上限
    max_per_day: int = 1                # 每人每天值班上限（默认每天只值一次）
    seed: int = 42                      # 随机种子（平手时打破平衡，可复现）


@dataclass
class ScheduleResult:
    assignments: list[Assignment] = field(default_factory=list)
    gaps: list[tuple[int, int, int]] = field(default_factory=list)  # (周, 星期, 时段) 无人可用
    member_stats: dict[int, dict] = field(default_factory=dict)     # id -> {name, total, weeks}

    @property
    def balanced_spread(self) -> int:
        """成员总值班次数的极差，越小越均衡"""
        totals = [s["total"] for s in self.member_stats.values()]
        return (max(totals) - min(totals)) if totals else 0


def build_busy_map(members: list[Member], courses: list[CourseRecord]) -> dict[int, set]:
    """member_id -> {(周, 星期, 节次)} 忙时集合（课表节次粒度）"""
    member_ids = {m.id for m in members}
    busy: dict[int, set] = defaultdict(set)
    for c in courses:
        if c.member_id not in member_ids:
            continue
        for week in c.week_list:
            for session in c.session_list:
                busy[c.member_id].add((week, c.weekday, session))
    return busy


def generate_schedule(
    members: list[Member],
    courses: list[CourseRecord],
    config: ScheduleConfig,
    leaves: list[Leave] | None = None,
    base_assignments: list[Assignment] | None = None,
) -> ScheduleResult:
    """base_assignments：排班范围外的已有安排（按周增量模式），
    作为总次数均衡基数计入，且与新排班合并进返回结果。"""
    busy = build_busy_map(members, courses)
    leave_set = {(l.member_id, l.week, l.weekday) for l in (leaves or [])}
    rng = random.Random(config.seed)
    assigned_slots: set[tuple] = set()  # 已占用的 (成员, 周, 星期, 时段)

    total: Counter = Counter()
    week_cnt: Counter = Counter()   # (member_id, week) -> n
    day_cnt: Counter = Counter()    # (member_id, week, weekday) -> n

    result = ScheduleResult()
    for a in base_assignments or []:
        result.assignments.append(a)
        total[a.member_id] += 1
        week_cnt[(a.member_id, a.week)] += 1
        day_cnt[(a.member_id, a.week, a.weekday)] += 1
        assigned_slots.add((a.member_id, a.week, a.weekday, a.block))

    def available(member_id: int, week: int, weekday: int, block: int) -> bool:
        return all(
            (week, weekday, s) not in busy.get(member_id, ())
            for s in BLOCK_SESSIONS[block]
        )

    for week in config.weeks:
        # 本周任务池（每时段 per_slot 人）
        pending = [
            (weekday, block)
            for _ in range(config.per_slot)
            for weekday in config.weekdays
            for block in config.blocks
        ]
        # 覆盖均衡：优先安排当前已覆盖最少的星期 / 时段，
        # 避免顺序处理时前几个星期耗尽每周上限、周尾整天空缺
        day_fills: Counter = Counter()
        block_fills: Counter = Counter()
        while pending:
            pending.sort(key=lambda t: (day_fills[t[0]], block_fills[t[1]], rng.random()))
            weekday, block = pending.pop(0)
            candidates = [
                m for m in members
                if available(m.id, week, weekday, block)
                and (m.id, week, weekday) not in leave_set
                and week_cnt[(m.id, week)] < config.max_per_week
                and day_cnt[(m.id, week, weekday)] < config.max_per_day
                and (m.id, week, weekday, block) not in assigned_slots
            ]
            if not candidates:
                result.gaps.append((week, weekday, block))
                continue
            # 三级均衡：总次数 -> 本周次数 -> 当天次数 -> 随机打散
            candidates.sort(key=lambda m: (
                total[m.id],
                week_cnt[(m.id, week)],
                day_cnt[(m.id, week, weekday)],
                rng.random(),
            ))
            chosen = candidates[0]
            total[chosen.id] += 1
            week_cnt[(chosen.id, week)] += 1
            day_cnt[(chosen.id, week, weekday)] += 1
            assigned_slots.add((chosen.id, week, weekday, block))
            day_fills[weekday] += 1
            block_fills[block] += 1
            result.assignments.append(Assignment(
                week=week, weekday=weekday, block=block,
                member_id=chosen.id, member_name=chosen.name,
            ))

    result.gaps.sort()
    result.member_stats = rebuild_member_stats(members, result.assignments)
    return result


def rebuild_member_stats(
    members: list[Member],
    assignments: list[Assignment],
) -> dict[int, dict]:
    """按 assignments 重算各成员统计（手动微调后复用）"""
    total = Counter(a.member_id for a in assignments)
    stats: dict[int, dict] = {}
    for m in members:
        stats[m.id] = {
            "name": m.name, "total": total[m.id],
            "weeks": sorted({a.week for a in assignments if a.member_id == m.id}),
        }
    return stats


def replacement_candidates(
    members: list[Member],
    busy: dict[int, set],
    leave_set: set[tuple[int, int, int]],
    assignments: list[Assignment],
    week: int,
    weekday: int,
    block: int,
    max_per_week: int,
    max_per_day: int,
) -> list[tuple[Member, str]]:
    """手动微调候选：返回 (成员, 不可用原因)，原因为空串表示可值班。
    排除该时段已在岗的成员；校验课程冲突、请假、每天/每周上限。"""
    current = {a.member_id for a in assignments
               if a.week == week and a.weekday == weekday and a.block == block}
    day_cnt = Counter(a.member_id for a in assignments
                      if a.week == week and a.weekday == weekday and a.block != block)
    week_cnt = Counter(a.member_id for a in assignments if a.week == week)

    out: list[tuple[Member, str]] = []
    for m in members:
        if m.id in current:
            continue
        if any((week, weekday, s) in busy.get(m.id, ()) for s in BLOCK_SESSIONS[block]):
            out.append((m, "该时段有课"))
        elif (m.id, week, weekday) in leave_set:
            out.append((m, "请假"))
        elif day_cnt[m.id] >= max_per_day:
            out.append((m, "当天已值班"))
        elif week_cnt[m.id] + 1 > max_per_week:
            out.append((m, "本周已达上限"))
        else:
            out.append((m, ""))
    return out
