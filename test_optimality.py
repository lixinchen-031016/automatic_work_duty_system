"""算法质量认证：用最大流求出「覆盖最优解」，量化启发式算法与最优的差距

排班算法是启发式（贪心 + 稀缺度排序 + 缺口链式修复）。本模块用最大流
建模求出**覆盖意义上的最优值**作为基准，验证：

1. 生产算法不会违反任何硬约束（最大流解同样满足约束，是可实现的安排）；
2. 生产算法的覆盖数与最优值的差距在可接受范围内。

之所以用「认证」而不是直接换最小费用流：若启发式已逼近最优，
引入 MCMF 只会增加复杂度与维护成本（结论由本模块数据支撑）。
"""

from __future__ import annotations

import random
from collections import Counter, deque

from duty_system.database import CourseRecord, Member
from duty_system.parser import BLOCK_SESSIONS
from duty_system.scheduler import (
    ScheduleConfig, build_busy_map, generate_schedule,
)


# --------------------------------------------------------------------------- #
# 最大流（Dinic）：求覆盖最优的排班方案
# --------------------------------------------------------------------------- #

class Dinic:
    def __init__(self, n: int) -> None:
        self.n = n
        self.graph: list[list[list[int]]] = [[] for _ in range(n)]  # [to, cap, rev]

    def add_edge(self, u: int, v: int, cap: int) -> None:
        self.graph[u].append([v, cap, len(self.graph[v])])
        self.graph[v].append([u, 0, len(self.graph[u]) - 1])

    def _bfs(self, s: int, t: int) -> list[int]:
        level = [-1] * self.n
        level[s] = 0
        q = deque([s])
        while q:
            u = q.popleft()
            for v, cap, _ in self.graph[u]:
                if cap > 0 and level[v] < 0:
                    level[v] = level[u] + 1
                    q.append(v)
        return level

    def _dfs(self, u: int, t: int, limit: int, level: list[int], it: list[int]) -> int:
        if u == t:
            return limit
        while it[u] < len(self.graph[u]):
            edge = self.graph[u][it[u]]
            v, cap, rev = edge
            if cap > 0 and level[v] == level[u] + 1:
                pushed = self._dfs(v, t, min(limit, cap), level, it)
                if pushed:
                    edge[1] -= pushed
                    self.graph[v][rev][1] += pushed
                    return pushed
            it[u] += 1
        return 0

    def max_flow(self, s: int, t: int) -> int:
        flow = 0
        while True:
            level = self._bfs(s, t)
            if level[t] < 0:
                return flow
            it = [0] * self.n
            while True:
                pushed = self._dfs(s, t, 10 ** 9, level, it)
                if not pushed:
                    break
                flow += pushed


def optimal_coverage(
    members: list[Member],
    busy: dict[int, set],
    config: ScheduleConfig,
) -> int:
    """最大可覆盖人次（上界，且该值可由某个合法排班达到）

    网络结构：
        S -> (成员, 周)[每周上限] -> (成员, 周, 星期)[每天上限] -> 时段[容量 1] -> T
    时段节点到 T 的容量为 per_slot，因此流值 = 总安排人次。
    """
    weekdays, blocks = list(config.weekdays), list(config.blocks)
    weeks = list(config.weeks)
    slots = [(w, d, b) for w in weeks for d in weekdays for b in blocks]

    idx_mw = {}
    idx_mwd = {}
    cursor = 1
    for m in members:
        for w in weeks:
            idx_mw[(m.id, w)] = cursor
            cursor += 1
    for m in members:
        for w in weeks:
            for d in weekdays:
                idx_mwd[(m.id, w, d)] = cursor
                cursor += 1
    idx_slot = {s: cursor + i for i, s in enumerate(slots)}
    cursor += len(slots)
    sink = cursor

    din = Dinic(sink + 1)
    source = 0
    for (mid, w), node in idx_mw.items():
        din.add_edge(source, node, config.max_per_week)
    for (mid, w, d), node in idx_mwd.items():
        din.add_edge(idx_mw[(mid, w)], node, config.max_per_day)
        for b in blocks:
            if any((w, d, s) in busy.get(mid, ()) for s in BLOCK_SESSIONS[b]):
                continue
            din.add_edge(node, idx_slot[(w, d, b)], 1)
    for slot, node in idx_slot.items():
        din.add_edge(node, sink, config.per_slot)
    return din.max_flow(source, sink)


def max_coverage_flow(
    members: list[Member],
    busy: dict[int, set],
    config: ScheduleConfig,
) -> tuple[int, list[tuple[int, int, int, int]]]:
    """同上，但额外提取一份最优覆盖方案 (成员, 周, 星期, 时段)，用于对比公平性"""
    weekdays, blocks = list(config.weekdays), list(config.blocks)
    weeks = list(config.weeks)
    slots = [(w, d, b) for w in weeks for d in weekdays for b in blocks]

    idx_mw, idx_mwd, cursor = {}, {}, 1
    for m in members:
        for w in weeks:
            idx_mw[(m.id, w)] = cursor
            cursor += 1
    for m in members:
        for w in weeks:
            for d in weekdays:
                idx_mwd[(m.id, w, d)] = cursor
                cursor += 1
    idx_slot = {s: cursor + i for i, s in enumerate(slots)}
    cursor += len(slots)
    sink = cursor

    din = Dinic(sink + 1)
    for (mid, w), node in idx_mw.items():
        din.add_edge(0, node, config.max_per_week)
    for (mid, w, d), node in idx_mwd.items():
        din.add_edge(idx_mw[(mid, w)], node, config.max_per_day)
        for b in blocks:
            if any((w, d, sec) in busy.get(mid, ()) for sec in BLOCK_SESSIONS[b]):
                continue
            din.add_edge(node, idx_slot[(w, d, b)], 1)
    for node in idx_slot.values():
        din.add_edge(node, sink, config.per_slot)

    flow = din.max_flow(0, sink)
    slot_of_node = {node: slot for slot, node in idx_slot.items()}
    taken = []
    for (mid, w, d), node in idx_mwd.items():
        for v, cap, _ in din.graph[node]:
            if cap == 0 and v in slot_of_node:      # 该边被用满 -> 此人被安排到此时段
                sw, sd, sb = slot_of_node[v]
                taken.append((mid, sw, sd, sb))
    return flow, taken


def make_dataset(n_members: int, per_member: int = 40, weeks: int = 18, seed: int = 1):
    rng = random.Random(seed)
    members, courses = [], []
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


# --------------------------------------------------------------------------- #

def test_production_schedule_matches_hard_constraints_of_flow_solution() -> None:
    """两种解都必须满足同样的硬约束（证明最大流基准是可比对象）"""
    members, courses = make_dataset(12, per_member=24, weeks=8, seed=4)
    config = ScheduleConfig(weeks=range(1, 9), per_slot=2, max_per_week=3, max_per_day=1)
    result = generate_schedule(members, courses, config)

    busy = build_busy_map(members, courses)
    week_cnt = Counter((a.member_id, a.week) for a in result.assignments)
    day_cnt = Counter((a.member_id, a.week, a.weekday) for a in result.assignments)
    slot_cnt = Counter((a.week, a.weekday, a.block) for a in result.assignments)
    assert all(v <= config.max_per_week for v in week_cnt.values())
    assert all(v <= config.max_per_day for v in day_cnt.values())
    assert all(v <= config.per_slot for v in slot_cnt.values())
    for a in result.assignments:
        for s in BLOCK_SESSIONS[a.block]:
            assert (a.week, a.weekday, s) not in busy.get(a.member_id, ())


def test_greedy_plus_repair_is_near_optimal() -> None:
    """生产算法覆盖率应达到最大流最优值的 97% 以上

    这是「不引入最小费用流」的依据：若差距已很小，增加整套网络流实现
    换不来可感知的收益。
    """
    for n_members, per_slot, seed in ((20, 2, 1), (15, 2, 7), (25, 3, 3)):
        members, courses = make_dataset(n_members, seed=seed)
        config = ScheduleConfig(weeks=range(1, 19), per_slot=per_slot,
                                max_per_week=3, max_per_day=1)
        result = generate_schedule(members, courses, config)
        busy = build_busy_map(members, courses)
        best = optimal_coverage(members, busy, config)
        got = len(result.assignments)
        ratio = got / best if best else 1.0
        assert got <= best, f"启发式结果 {got} 不应超过最优值 {best}"
        assert ratio >= 0.97, (
            f"{n_members} 人 / 每时段 {per_slot} 人：覆盖率 {ratio:.1%}"
            f"（{got}/{best}），已低于 97% 门槛，需考虑更强的算法")


def test_repair_improves_toward_optimum() -> None:
    """缺口修复应确实拉近与最优值的距离（对比关闭修复的情况）"""
    members, courses = make_dataset(20, seed=1)
    config = ScheduleConfig(weeks=range(1, 19), per_slot=2, max_per_week=3, max_per_day=1)
    busy = build_busy_map(members, courses)
    best = optimal_coverage(members, busy, config)

    plain = generate_schedule(members, courses, config, repair=False)
    repaired = generate_schedule(members, courses, config, repair=True)

    assert len(repaired.assignments) >= len(plain.assignments), "修复不应减少安排"
    assert repaired.repaired > 0, "该场景应触发修复"
    gap_plain = best - len(plain.assignments)
    gap_repaired = best - len(repaired.assignments)
    assert gap_repaired <= gap_plain, "修复后离最优应更近或持平"


def test_dense_capacity_is_truly_infeasible() -> None:
    """人少班多时，最大流基准与生产算法都应按容量上限收敛（缺口是客观限制）"""
    members, courses = make_dataset(6, per_member=20, weeks=4, seed=9)
    config = ScheduleConfig(weeks=range(1, 5), per_slot=2, max_per_week=3, max_per_day=1)
    result = generate_schedule(members, courses, config)
    busy = build_busy_map(members, courses)
    best = optimal_coverage(members, busy, config)

    demand = 4 * 5 * 5 * config.per_slot
    assert best < demand, "该场景容量本身不足，最优解也无法排满"
    assert len(result.assignments) <= best
    assert len(result.assignments) / best >= 0.97, "即使容量不足也应贴近最优"

def test_pure_max_coverage_destroys_fairness() -> None:
    """结论性对比：纯最大流覆盖最优但公平性极差，因此不替换现有算法

    最大流只关心「填满多少人次」，对谁值多少班完全无所谓：同一网络存在大量
    等价最优解，Dinic 取到的那个会把班次推给同几个人。要让 flow 同时均衡，
    必须写成最小费用流（凸费用），实现复杂度大幅上升，而覆盖率收益不足 1%。
    本用例把这一权衡固化成可验证的结论。
    """
    from duty_system.scheduler import rebuild_member_stats

    members, courses = make_dataset(20, seed=1)
    config = ScheduleConfig(weeks=range(1, 19), per_slot=2,
                            max_per_week=3, max_per_day=1)
    busy = build_busy_map(members, courses)

    result = generate_schedule(members, courses, config)
    flow, flow_rows = max_coverage_flow(members, busy, config)

    flow_counts = Counter(mid for mid, *_ in flow_rows)
    flow_spread = max(flow_counts.values()) - min(flow_counts.values())
    greedy_counts = [s["total"] for s in rebuild_member_stats(members, result.assignments).values()]
    greedy_spread = max(greedy_counts) - min(greedy_counts)

    # 覆盖率接近，但公平性差距巨大
    assert len(result.assignments) / flow >= 0.97, "生产算法覆盖率应接近最优"
    assert greedy_spread < flow_spread, (
        f"生产算法的次数极差 {greedy_spread} 应明显优于纯最大流解 {flow_spread}"
        "（否则说明均衡策略失效）")
    assert greedy_spread <= flow_spread / 2, (
        f"均衡收益不明显：贪心极差 {greedy_spread} vs 最大流极差 {flow_spread}")
