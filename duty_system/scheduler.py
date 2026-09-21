"""值班排班算法

硬约束：
  - 值班时刻内（时段块的每一节）不能有任何课程；
  - 每人每天最多值一次（max_per_day，默认 1）；
  - 每人每周值班不超过 max_per_week 次；
  - 长期特殊安排占用的 (成员, 周, 星期, 节次) 不安排值班；
  - 请假/临时占用的 (成员, 周, 星期) 不安排值班。
软目标：按 累计总次数 -> 本周次数 -> 当天次数 三级均衡贪心分配，
        随机种子固定可复现，平手时随机打散。
覆盖优化：贪心结束后对缺口做「链式挪动」修复（增广路，深度可配）——
        把已排成员挪到缺人时段、再回填其原时段，消除单遍贪心的局部最优空缺。

约束判定集中在本模块的 ScheduleContext.reason()：自动排班、手动微调候选、
界面校验共用同一实现，新增约束只需改一处。
"""

from __future__ import annotations

import random
from collections import Counter, defaultdict
from collections.abc import Callable
from dataclasses import dataclass, field

from .database import Assignment, CourseRecord, Leave, Member, SpecialArrangement
from .parser import BLOCK_SESSIONS, WHOLE_WEEK_SESSIONS, WHOLE_WEEK_WEEKDAY

# 不可用原因：算法与界面共用同一组文案，避免两处漂移
REASON_COURSE = "该时段有课"
REASON_SPECIAL = "其他安排"
REASON_LEAVE = "请假"
REASON_DAY = "当天已值班"
REASON_WEEK = "本周已达上限"
REASON_SLOT = "该时段已值班"
# 软约束：手动微调时可「知情越限」，其余原因均为硬约束
SOFT_REASONS = frozenset({REASON_WEEK})

REPAIR_DEPTH = 2          # 缺口修复的最大链式深度（1=仅同天挪动一次）
REPAIR_BUDGET_BASE = 200  # 缺口修复的基础尝试预算
REPAIR_BUDGET_PER_GAP = 10


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
    gaps: list[tuple[int, int, int]] = field(default_factory=list)  # (周, 星期, 时段) 人手不足
    member_stats: dict[int, dict] = field(default_factory=dict)     # id -> {name, total, weeks}
    repaired: int = 0                                                # 缺口修复成功填补的时段数

    @property
    def balanced_spread(self) -> int:
        """成员总值班次数的极差，越小越均衡"""
        totals = [s["total"] for s in self.member_stats.values()]
        return (max(totals) - min(totals)) if totals else 0


class ScheduleContext:
    """排班/微调共用的约束与计数上下文。

    只做增量计数，assign 与 unassign 完全对称，便于缺口修复时试排与回滚。
    """

    def __init__(
        self,
        members: list[Member],
        busy: dict[int, set],
        leave_set: set[tuple[int, int, int]] | None,
        config: ScheduleConfig,
        special_set: set[tuple[int, int, int, int]] | None = None,
    ) -> None:
        self.members = members
        self.busy = busy
        self.leave_set = leave_set or set()
        self.special_set = special_set or set()
        self.config = config
        self.assigned: set[tuple[int, int, int, int]] = set()  # (mid, week, day, block)
        self.total: Counter = Counter()
        self.week_cnt: Counter = Counter()                     # (mid, week) -> n
        self.day_cnt: Counter = Counter()                      # (mid, week, day) -> n
        self.slot_cnt: Counter = Counter()                     # (week, day, block) -> n
        self._day_slots: dict[tuple[int, int, int], set[int]] = defaultdict(set)

    # ---------- 计数 ----------

    def load(self, assignments: list[Assignment]) -> None:
        for a in assignments:
            self.assign(a.member_id, a.week, a.weekday, a.block)

    def load_assignments(self, assignments: list[Assignment]) -> None:
        """别名，语义与 load 一致（供微调候选构建使用）"""
        self.load(assignments)

    def assign(self, member_id: int, week: int, weekday: int, block: int) -> None:
        key = (member_id, week, weekday, block)
        if key in self.assigned:
            return
        self.assigned.add(key)
        self.total[member_id] += 1
        self.week_cnt[(member_id, week)] += 1
        self.day_cnt[(member_id, week, weekday)] += 1
        self.slot_cnt[(week, weekday, block)] += 1
        self._day_slots[(member_id, week, weekday)].add(block)

    def unassign(self, member_id: int, week: int, weekday: int, block: int) -> None:
        key = (member_id, week, weekday, block)
        if key not in self.assigned:
            return
        self.assigned.discard(key)
        self.total[member_id] -= 1
        self.week_cnt[(member_id, week)] -= 1
        self.day_cnt[(member_id, week, weekday)] -= 1
        self.slot_cnt[(week, weekday, block)] -= 1
        day_key = (member_id, week, weekday)
        self._day_slots[day_key].discard(block)
        if not self._day_slots[day_key]:
            self._day_slots.pop(day_key, None)

    def blocks_of(self, member_id: int, week: int, weekday: int) -> set[int]:
        """该成员当天已排的时段集合"""
        return set(self._day_slots.get((member_id, week, weekday), ()))

    # ---------- 约束（唯一判定入口） ----------

    def reason(self, member_id: int, week: int, weekday: int, block: int) -> str | None:
        """返回不可用原因；None 表示可以安排到该 (周, 星期, 时段)。"""
        if any((week, weekday, s) in self.busy.get(member_id, ())
               for s in BLOCK_SESSIONS[block]):
            return REASON_COURSE
        if any((member_id, week, weekday, s) in self.special_set
               for s in BLOCK_SESSIONS[block]):
            return REASON_SPECIAL
        if (member_id, week, weekday) in self.leave_set:
            return REASON_LEAVE
        if (member_id, week, weekday, block) in self.assigned:
            return REASON_SLOT
        if self.day_cnt[(member_id, week, weekday)] >= self.config.max_per_day:
            return REASON_DAY
        if self.week_cnt[(member_id, week)] >= self.config.max_per_week:
            return REASON_WEEK
        return None

    def eligible(self, member_id: int, week: int, weekday: int, block: int) -> bool:
        return self.reason(member_id, week, weekday, block) is None


def build_busy_map(members: list[Member], courses: list[CourseRecord]) -> dict[int, set]:
    """member_id -> {(周, 星期, 节次)} 忙时集合（课表节次粒度）

    展开量 = 课程数 x 周次 x 节次，是排班与甘特图的热路径：
    内层循环里先把目标集合与标量取到局部变量，避免每次 add 都做字典索引。
    调用方（界面）应缓存结果，课表不变时无需重建。
    """
    member_ids = {m.id for m in members}
    busy: dict[int, set] = defaultdict(set)
    for c in courses:
        if c.member_id not in member_ids:
            continue
        target = busy[c.member_id]
        weekday = c.weekday
        sessions = c.session_list
        if weekday == WHOLE_WEEK_WEEKDAY:
            # 整周集中安排（军训/思政实践）：没有星期与节次，整周都不可值班
            days = range(1, 8)
            sessions = WHOLE_WEEK_SESSIONS
        else:
            days = (weekday,)
        for week in c.week_list:
            for day in days:
                for session in sessions:
                    target.add((week, day, session))
    return busy


def build_special_busy_map(
    members: list[Member],
    arrangements: list[SpecialArrangement] | None,
) -> dict[int, set]:
    """长期特殊安排 -> member_id: {(周, 星期, 节次)} 忙时集合。"""
    member_ids = {m.id for m in members}
    busy: dict[int, set] = defaultdict(set)
    for a in arrangements or []:
        if a.member_id not in member_ids:
            continue
        target = busy[a.member_id]
        for week in a.week_list:
            for session in a.session_list:
                target.add((week, a.weekday, session))
    return busy


def compute_gaps(
    assignments: list[Assignment],
    config: ScheduleConfig,
    is_off: Callable[[int, int], bool] | None = None,
    is_class: Callable[[int, int], bool] | None = None,
) -> list[tuple[int, int, int]]:
    """按「排班范围内人数不足 per_slot 的时段」重算缺口。
    is_off 为真表示该逻辑日没有真实日期；is_class 为真表示周末补课日，
    即使对应逻辑星期未勾选也应自动加入缺口网格。"""
    cnt = Counter((a.week, a.weekday, a.block) for a in assignments)
    grid = (
        (w, d, b)
        for w in config.weeks
        for d in _active_weekdays(config, w, is_class)
        for b in config.blocks
        if is_off is None or not is_off(w, d)
    )
    return sorted(s for s in grid if cnt[s] < max(config.per_slot, 1))


def _active_weekdays(
    config: ScheduleConfig,
    week: int,
    is_class: Callable[[int, int], bool] | None = None,
) -> list[int]:
    """本周任务星期 = 勾选星期 + 自动纳入的补课日。"""
    weekdays = set(config.weekdays)
    if is_class is not None:
        weekdays.update(
            weekday for weekday in range(1, 8) if is_class(week, weekday))
    return sorted(weekdays)


def _best_candidate(
    ctx: ScheduleContext,
    members: list[Member],
    week: int,
    weekday: int,
    block: int,
    rng: random.Random,
    exclude: frozenset[int] | set[int] = frozenset(),
) -> Member | None:
    """三级均衡挑选：总次数 -> 本周次数 -> 当天次数 -> 随机打散"""
    best: Member | None = None
    best_key: tuple | None = None
    for m in members:
        if m.id in exclude:
            continue
        if ctx.reason(m.id, week, weekday, block) is not None:
            continue
        key = (ctx.total[m.id], ctx.week_cnt[(m.id, week)],
               ctx.day_cnt[(m.id, week, weekday)], rng.random())
        if best_key is None or key < best_key:
            best, best_key = m, key
    return best


def _try_fill(
    ctx: ScheduleContext,
    members: list[Member],
    slot: tuple[int, int, int],
    rng: random.Random,
    depth: int,
    trail: frozenset[tuple[int, int, int]],
    exclude: frozenset[int],
) -> bool:
    """尝试填满 slot：先直接安排；不行则把当天已排成员挪过来，递归回填其原时段。"""
    week, weekday, block = slot
    if ctx.slot_cnt[slot] >= ctx.config.per_slot:
        return False
    if slot in trail:
        return False
    trail = trail | {slot}

    chosen = _best_candidate(ctx, members, week, weekday, block, rng, exclude)
    if chosen is not None:
        ctx.assign(chosen.id, week, weekday, block)
        return True
    if depth <= 0:
        return False

    # 链式挪动：候选被「当天已值班 / 本周已达上限」挡住时，把他挪到本时段再回填原时段
    for m in members:
        if m.id in exclude:
            continue
        reason = ctx.reason(m.id, week, weekday, block)
        if reason not in (REASON_DAY, REASON_WEEK):
            continue
        for other_block in sorted(ctx.blocks_of(m.id, week, weekday)):
            other = (week, weekday, other_block)
            if other in trail or ctx.slot_cnt[other] <= 0:
                continue
            ctx.unassign(m.id, week, weekday, other_block)
            moved = ctx.reason(m.id, week, weekday, block) is None
            if moved:
                ctx.assign(m.id, week, weekday, block)
                if _try_fill(ctx, members, other, rng, depth - 1, trail,
                             exclude | {m.id}):
                    return True
                ctx.unassign(m.id, week, weekday, block)
            ctx.assign(m.id, week, weekday, other_block)  # 回滚
    return False


def repair_gaps(
    ctx: ScheduleContext,
    members: list[Member],
    config: ScheduleConfig,
    gap_slots: list[tuple[int, int, int]],
    is_off: Callable[[int, int], bool] | None = None,
    is_class: Callable[[int, int], bool] | None = None,
) -> int:
    """对缺口做链式挪动修复，返回修复成功的时段数。

    仅在本周仍有余量（有人周上限未满且当天有可用时段）时尝试，
    避免容量本身不足时做无意义搜索。
    """
    if not gap_slots:
        return 0
    rng = random.Random(config.seed ^ 0x5EED)  # 修复阶段用独立随机流，不动主流程
    budget = REPAIR_BUDGET_BASE + REPAIR_BUDGET_PER_GAP * len(gap_slots)
    has_spare: dict[int, bool] = {}
    fixed = 0
    for slot in gap_slots:
        if budget <= 0:
            break
        week = slot[0]
        if week not in has_spare:
            has_spare[week] = _week_has_spare_capacity(
                ctx, members, week, is_off, is_class)
        if not has_spare[week]:
            continue
        if ctx.slot_cnt[slot] >= config.per_slot:
            continue
        budget -= 1
        if _try_fill(ctx, members, slot, rng, REPAIR_DEPTH, frozenset(), frozenset()):
            fixed += 1
    return fixed


# 缺口原因：让界面能区分「排不了」（课程/特殊安排/请假冲突）与「排不下」（上限不足），
# 后者可以通过提高每周/每天上限或增加成员解决。
GAP_NO_FREE = "成员均有课、其他安排或请假"
GAP_DAY_CAP = "受每天上限限制"
GAP_WEEK_CAP = "受每周上限限制"
GAP_MIXED_CAP = "受每天/每周上限限制"
GAP_ELIGIBLE = "仍有空闲成员未被排入"


@dataclass
class GapDiagnosis:
    """单个缺人时段的诊断结果"""

    week: int
    weekday: int
    block: int
    cause: str
    free_members: int = 0          # 该时段无课、无特殊安排且未请假的成员数
    blocked_by_course: int = 0
    blocked_by_special: int = 0
    blocked_by_leave: int = 0
    blocked_by_day_cap: int = 0
    blocked_by_week_cap: int = 0
    eligible: int = 0              # 现在仍可安排的人数（>0 说明排漏了）

    @property
    def slot(self) -> tuple[int, int, int]:
        return (self.week, self.weekday, self.block)


def diagnose_gaps(
    members: list[Member],
    busy: dict[int, set],
    leaves: list[Leave] | None,
    assignments: list[Assignment],
    config: ScheduleConfig,
    special_arrangements: list[SpecialArrangement] | None = None,
    is_off: Callable[[int, int], bool] | None = None,
    is_class: Callable[[int, int], bool] | None = None,
) -> list[GapDiagnosis]:
    """逐条分析缺口成因：谁被课程/特殊安排/请假挡住、谁只是被上限挡住。

    与排班共用 ScheduleContext.reason()，因此「诊断说可排」与「算法认为可排」
    永远一致，不会出现界面解释与算法行为不符。
    """
    leave_set = {(l.member_id, l.week, l.weekday) for l in (leaves or [])}
    if is_off is not None:
        assignments = [
            a for a in assignments if not is_off(a.week, a.weekday)
        ]
    special_busy = build_special_busy_map(members, special_arrangements)
    special_set = {
        (mid, week, weekday, session)
        for mid, sessions in special_busy.items()
        for week, weekday, session in sessions
    }
    ctx = ScheduleContext(members, busy, leave_set, config, special_set)
    ctx.load(assignments)

    out: list[GapDiagnosis] = []
    for week, weekday, block in compute_gaps(
            assignments, config, is_off, is_class):
        d = GapDiagnosis(week=week, weekday=weekday, block=block, cause=GAP_NO_FREE)
        for m in members:
            if any((week, weekday, s) in busy.get(m.id, ())
                   for s in BLOCK_SESSIONS[block]):
                d.blocked_by_course += 1
                continue
            if any((m.id, week, weekday, s) in special_set
                   for s in BLOCK_SESSIONS[block]):
                d.blocked_by_special += 1
                continue
            if (m.id, week, weekday) in leave_set:
                d.blocked_by_leave += 1
                continue
            d.free_members += 1
            if ctx.day_cnt[(m.id, week, weekday)] >= config.max_per_day:
                d.blocked_by_day_cap += 1
            elif ctx.week_cnt[(m.id, week)] >= config.max_per_week:
                d.blocked_by_week_cap += 1
            else:
                d.eligible += 1

        if d.eligible:
            d.cause = GAP_ELIGIBLE
        elif d.free_members == 0:
            d.cause = GAP_NO_FREE
        elif d.blocked_by_day_cap and d.blocked_by_week_cap:
            d.cause = GAP_MIXED_CAP
        elif d.blocked_by_day_cap:
            d.cause = GAP_DAY_CAP
        elif d.blocked_by_week_cap:
            d.cause = GAP_WEEK_CAP
        else:
            d.cause = GAP_NO_FREE
        out.append(d)
    return out


def summarize_gap_causes(diagnoses: list[GapDiagnosis]) -> dict[str, int]:
    """按原因汇总缺口时段数（供界面一行说明）"""
    return dict(Counter(d.cause for d in diagnoses))


def capacity_advice(summary: dict[str, int], per_slot: int) -> str:
    """根据缺口成因给出可操作建议；没有缺口时返回空串"""
    if not summary:
        return ""
    hard = summary.get(GAP_NO_FREE, 0)
    capped = sum(summary.get(k, 0) for k in (GAP_DAY_CAP, GAP_WEEK_CAP, GAP_MIXED_CAP))
    missed = summary.get(GAP_ELIGIBLE, 0)
    parts: list[str] = []
    if hard:
        parts.append(f"{hard} 个时段成员普遍有课、其他安排或请假（只能靠增加成员解决）")
    if capped:
        parts.append(f"{capped} 个时段还有无课成员、只是达到每天/每周上限"
                     f"（提高上限或增加成员即可排满）")
    if missed:
        parts.append(f"{missed} 个时段仍有可用成员（可重新生成或手动补排）")
    return "；".join(parts) + "。"


def _week_has_spare_capacity(
    ctx: ScheduleContext,
    members: list[Member],
    week: int,
    is_off: Callable[[int, int], bool] | None = None,
    is_class: Callable[[int, int], bool] | None = None,
) -> bool:
    """本周是否还有人「周上限未满 且 存在可用时段」——否则修无可修"""
    cfg = ctx.config
    for m in members:
        if ctx.week_cnt[(m.id, week)] >= cfg.max_per_week:
            continue
        for d in _active_weekdays(cfg, week, is_class):
            if is_off is not None and is_off(week, d):
                continue
            if ctx.day_cnt[(m.id, week, d)] >= cfg.max_per_day:
                continue
            if any(ctx.reason(m.id, week, d, b) is None for b in cfg.blocks):
                return True
    return False


def _availability_counts(
    members: list[Member],
    busy: dict[int, set],
    leave_set: set[tuple[int, int, int]],
    week: int,
    weekdays: list[int],
    blocks: list[int],
) -> dict[tuple[int, int], int]:
    """第 week 周各 (星期, 时段) 的静态可用人数（传入的忙时含课程/特殊安排）。

    作为贪心的次级排序键：同样的覆盖度下优先处理「可选人少」的时段，
    把抢手的时段先占住，实测显著减少单遍贪心留下的缺口。每个时段只算一次。
    """
    free_of: dict[int, set[tuple[int, int]]] = {}
    for m in members:
        free_slots = set()
        for d in weekdays:
            if (m.id, week, d) in leave_set:
                continue
            busy_days = busy.get(m.id, ())
            for b in blocks:
                if all((week, d, s) not in busy_days for s in BLOCK_SESSIONS[b]):
                    free_slots.add((d, b))
        free_of[m.id] = free_slots
    return {slot: sum(1 for fs in free_of.values() if slot in fs)
            for slot in ((d, b) for d in weekdays for b in blocks)}


def generate_schedule(
    members: list[Member],
    courses: list[CourseRecord],
    config: ScheduleConfig,
    leaves: list[Leave] | None = None,
    base_assignments: list[Assignment] | None = None,
    repair: bool = True,
    special_arrangements: list[SpecialArrangement] | None = None,
    is_off: Callable[[int, int], bool] | None = None,
    is_class: Callable[[int, int], bool] | None = None,
) -> ScheduleResult:
    """base_assignments：排班范围外的已有安排（按周增量模式），
    作为总次数均衡基数计入，且与新排班合并进返回结果。
    is_off：逻辑日放假谓词，命中的日期不进入任务网格或均衡基数。
    is_class：周末补课日谓词，命中的逻辑日即使未勾选也会自动加入任务网格。
    repair=False 可关闭缺口修复（用于对比/测试）。"""
    busy = build_busy_map(members, courses)
    special_busy = build_special_busy_map(members, special_arrangements)
    special_set = {
        (mid, week, weekday, session)
        for mid, sessions in special_busy.items()
        for week, weekday, session in sessions
    }
    scarcity_busy = busy
    if special_busy:
        scarcity_busy = {
            m.id: busy.get(m.id, set()) | special_busy.get(m.id, set())
            for m in members
        }
    leave_set = {(l.member_id, l.week, l.weekday) for l in (leaves or [])}
    rng = random.Random(config.seed)
    ctx = ScheduleContext(members, busy, leave_set, config, special_set)

    result = ScheduleResult()
    result.assignments.extend(
        a for a in (base_assignments or [])
        if is_off is None or not is_off(a.week, a.weekday)
    )
    ctx.load(result.assignments)

    for week in config.weeks:
        weekdays = [
            weekday for weekday in _active_weekdays(config, week, is_class)
            if is_off is None or not is_off(week, weekday)
        ]
        if not weekdays:
            continue
        # 本周任务池（每时段 per_slot 人）
        pending = [
            (weekday, block)
            for _ in range(config.per_slot)
            for weekday in weekdays
            for block in config.blocks
        ]
        scarcity = _availability_counts(members, scarcity_busy, leave_set, week,
                                        weekdays, list(config.blocks))
        # 每条任务只抽一次随机数作为平手排序键，主排序仍是「覆盖均衡」
        keys = {i: rng.random() for i in range(len(pending))}
        keys_idx = list(range(len(pending)))
        # 覆盖均衡：优先安排当前已覆盖最少的星期 / 时段，
        # 避免顺序处理时前几个星期耗尽每周上限、周尾整天空缺
        day_fills: Counter = Counter()
        block_fills: Counter = Counter()
        while keys_idx:
            keys_idx.sort(key=lambda i: (day_fills[pending[i][0]],
                                         block_fills[pending[i][1]],
                                         scarcity[pending[i]],
                                         keys[i]))
            idx = keys_idx.pop(0)
            weekday, block = pending[idx]
            chosen = _best_candidate(ctx, members, week, weekday, block, rng)
            if chosen is None:
                continue  # 记入缺口待修复，最后统一重算
            ctx.assign(chosen.id, week, weekday, block)
            day_fills[weekday] += 1
            block_fills[block] += 1
            result.assignments.append(Assignment(
                week=week, weekday=weekday, block=block,
                member_id=chosen.id, member_name=chosen.name,
            ))

    result.gaps = compute_gaps(result.assignments, config, is_off, is_class)
    repaired = 0
    if repair and result.gaps:
        repaired = repair_gaps(
            ctx, members, config, result.gaps, is_off, is_class)
    if repaired:
        # 输出统一按 (周, 星期, 时段, 成员) 排序，与数据库读取顺序一致，
        # 保证「生成 -> 落库 -> 恢复」三处明细顺序稳定可对比
        names = {m.id: m.name for m in members}
        result.assignments = [
            Assignment(week=w, weekday=d, block=b, member_id=mid,
                       member_name=names.get(mid, ""))
            for (mid, w, d, b) in sorted(ctx.assigned, key=lambda t: (t[1], t[2], t[3], t[0]))
        ]
    result.repaired = repaired
    result.gaps = compute_gaps(result.assignments, config, is_off, is_class)
    result.member_stats = rebuild_member_stats(members, result.assignments)
    return result


def _name_of(members: list[Member], member_id: int) -> str:
    for m in members:
        if m.id == member_id:
            return m.name
    return ""


def rebuild_member_stats(
    members: list[Member],
    assignments: list[Assignment],
) -> dict[int, dict]:
    """按 assignments 重算各成员统计（手动微调后复用）。
    一次遍历完成，复杂度 O(排班数) 而非 O(成员数 x 排班数)。"""
    total = Counter(a.member_id for a in assignments)
    weeks: dict[int, set[int]] = defaultdict(set)
    for a in assignments:
        weeks[a.member_id].add(a.week)
    return {
        m.id: {"name": m.name, "total": total[m.id], "weeks": sorted(weeks[m.id])}
        for m in members
    }


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
    special_set: set[tuple[int, int, int, int]] | None = None,
) -> list[tuple[Member, str]]:
    """手动微调候选：返回 (成员, 不可用原因)，原因为空串表示可值班。
    排除该时段已在岗的成员；与自动排班共用 ScheduleContext.reason() 判定。"""
    config = ScheduleConfig(
        weeks=range(week, week + 1), weekdays=[weekday], blocks=[block],
        max_per_week=max_per_week, max_per_day=max_per_day,
    )
    ctx = ScheduleContext(members, busy, leave_set, config, special_set)
    ctx.load(assignments)
    current = {a.member_id for a in assignments
               if (a.week, a.weekday, a.block) == (week, weekday, block)}
    out: list[tuple[Member, str]] = []
    for m in members:
        if m.id in current:
            continue
        out.append((m, ctx.reason(m.id, week, weekday, block) or ""))
    return out
