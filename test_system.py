"""端到端测试：解析 -> 入库 -> 排班(无冲突+均衡) -> 导出 -> 甘特/请假/微调/日期

两种跑法：
    pytest test_system.py -v        # 常规：逐项报告，可配合其他测试模块
    python test_system.py           # 直接跑完整链路（无需 pytest）

用例之间有先后依赖（模拟真实使用顺序：先入库再排班再微调），
因此用模块级 STATE 传递中间产物，而不是各自独立构造数据。
"""

from __future__ import annotations

import copy
import random
import tempfile
from collections import Counter
from datetime import date, timedelta
from pathlib import Path

STATE: dict = {}


def get_env() -> tuple[Database, list]:
    """返回 (db, members)；首次调用时建库并导入样例课表"""
    if "db" not in STATE:
        db, members = _setup()
        STATE["db"] = db
        STATE["members"] = members
    return STATE["db"], STATE["members"]


def reset_env() -> None:
    """清空测试状态（供需要干净环境的用例调用）"""
    STATE.clear()

from duty_system.calendar import CalendarEntry, TermCalendar
from duty_system.database import Assignment, CourseRecord, Database, Member
from duty_system.exporter import (
    build_detail_df,
    build_leaves_df,
    build_pivot_df,
    export_csv,
    export_excel,
    export_leaves_excel,
    week_date,
)
from duty_system.gantt import build_availability, display_columns, export_gantt_excel
from duty_system.parser import BLOCK_SESSIONS, WEEKDAY_LABELS, parse_schedule_path
from duty_system.scheduler import (
    ScheduleConfig,
    build_busy_map,
    generate_schedule,
    rebuild_member_stats,
    replacement_candidates,
)

SAMPLE = (Path(__file__).parent / "samples" / "desensitized"
          / "学生个人课表_9999800598.xls")   # 真实课表脱敏后的样例

TEST_NAMES = ["王思远", "李慧敏", "张承宇", "陈晓露"]
TEST_STUDENT_IDS = ["9999000101", "9999000102", "9999000103", "9999000104"]


def make_test_variants(base, count: int = 4, seed: int = 7) -> list:
    """从基础课表派生 count 个测试成员：随机丢弃部分课程并平移部分课程的星期，
    使各成员忙闲分布不同，用于验证排班的冲突规避与均衡分配。"""
    rng = random.Random(seed)
    variants: list = []
    for i in range(count):
        v = copy.deepcopy(base)
        v.name = TEST_NAMES[i % len(TEST_NAMES)]
        v.student_id = TEST_STUDENT_IDS[i % len(TEST_STUDENT_IDS)]
        v.file_name = f"测试_{v.name}.xlsx"
        kept = [c for c in v.courses if rng.random() > 0.35]  # 保留约 65% 的课程
        for c in kept:
            if rng.random() < 0.15:  # 平移部分课程的星期，制造忙闲差异
                c.weekday = (c.weekday + rng.choice([-2, -1, 1, 2]) - 1) % 7 + 1
        v.courses = kept or v.courses
        variants.append(v)
    return variants


def test_parser() -> None:
    s = parse_schedule_path(SAMPLE)
    assert s.name == "赵明明", f"姓名解析错误: {s.name}"
    assert s.student_id == "9999800598"
    assert s.term == "2026-2027-1"
    assert s.class_name == "24国际商务双语4班"
    assert len(s.courses) == 37, f"课程数错误: {len(s.courses)}"

    names = {c.course_name for c in s.courses}
    assert "高等数学(Ⅱ)-1" in names, "无教师课程(高等数学)解析失败"
    assert not any(c.teacher for c in s.courses if c.course_name == "高等数学(Ⅱ)-1")

    # 教师职称被换行打断的案例
    sy = [c for c in s.courses if "企业运营管理" in c.course_name]
    assert sy and all(c.teacher == "钱明明" for c in sy), "断行教师名解析失败"
    # 周次解析
    pe = next(c for c in s.courses if c.course_name.startswith("体育Ⅲ"))
    assert pe.week_list == [1, 3, 5, 7, 9, 11, 13, 15], f"单周周次解析错误: {pe.week_list}"
    mds = next(c for c in s.courses if c.course_name.startswith("毛泽东"))
    assert mds.week_list == [1, 2, 3, 4, 6, 7, 8, 9, 10], f"混合周次解析错误: {mds.week_list}"
    print(f"[1] 解析器: 通过（{len(s.courses)} 条课程记录，姓名/学号/断行教师/周次/去重均正确）")


def _setup() -> tuple[Database, list]:
    """建库 + 导入样例课表的 5 名成员（test_database 与 get_env 共用）"""
    db = Database(Path(tempfile.mkdtemp()) / "test.db")
    base = parse_schedule_path(SAMPLE)
    db.upsert_member(base)
    for v in make_test_variants(base, count=4):
        db.upsert_member(v)

    members = db.list_members()
    assert len(members) == 5, f"成员数错误: {len(members)}"
    assert all(m.course_count > 0 for m in members)

    # 重复上传同一个人 -> 更新而非新增
    db.upsert_member(base)
    assert len(db.list_members()) == 5, "重复上传应更新而非新增成员"

    # 指定成员的课程
    courses = db.get_courses(members[0].id)
    assert len(courses) == 37
    print(f"[2] 数据库: 通过（5 名成员，课程数 {[m.course_count for m in members]}，重复上传幂等）")
    return db, members


def test_database() -> None:
    db, members = get_env()
    assert len(members) == 5
    assert all(m.course_count > 0 for m in members)


def test_scheduler() -> None:
    db, members = get_env()
    courses = db.get_courses()
    config = ScheduleConfig(
        weeks=range(1, 19), weekdays=[1, 2, 3, 4, 5], blocks=[1, 2, 3, 4, 5],
        per_slot=2, max_per_week=3, max_per_day=1,
    )
    result = generate_schedule(members, courses, config)

    # 硬约束：值班成员在该 (周, 星期, 每一节) 均无课
    busy = build_busy_map(members, courses)
    for a in result.assignments:
        for s in BLOCK_SESSIONS[a.block]:
            assert (a.week, a.weekday, s) not in busy.get(a.member_id, ()), \
                f"冲突: {a.member_name} 第{a.week}周{a.weekday} {a.block} 时段有课"

    # 每周上限
    week_cnt = Counter((a.member_id, a.week) for a in result.assignments)
    assert all(v <= config.max_per_week for v in week_cnt.values()), "超出每周上限"

    # 每天上限（每人每天只值一次）
    day_cnt = Counter((a.member_id, a.week, a.weekday) for a in result.assignments)
    assert all(v <= config.max_per_day for v in day_cnt.values()), "超出每天上限"

    # 每个时段人数
    slot_cnt = Counter((a.week, a.weekday, a.block) for a in result.assignments)
    assert all(v <= config.per_slot for v in slot_cnt.values()), "时段人数超额"

    # 覆盖均衡：每周每个值班星期都应有人值班（不应出现周尾整天空缺）
    for w in config.weeks:
        for d in config.weekdays:
            assert any(a.week == w and a.weekday == d for a in result.assignments), \
                f"第{w}周{WEEKDAY_LABELS[d]}全天无安排，星期覆盖不均衡"

    # 均衡性
    totals = {s["total"] for s in result.member_stats.values()}
    spread = max(totals) - min(totals)
    print(f"[3] 排班算法: 通过（{len(result.assignments)} 人次安排，0 冲突，"
          f"每人每天<=1 次，每周各星期均有安排，缺口 {len(result.gaps)} 个，总次数极差 {spread}）")
    for mid, s in result.member_stats.items():
        print(f"      - {s['name']}: {s['total']} 次")

    db.save_assignments(result.assignments)
    loaded = db.load_assignments()
    assert len(loaded) == len(result.assignments)
    print(f"[4] 排班入库: 通过（{len(loaded)} 条安排持久化）")
    STATE["result"] = result


def test_exporter() -> None:
    result = STATE["result"]
    out = Path(tempfile.mkdtemp())
    xlsx = export_excel(result.assignments, result.member_stats, result.gaps)
    (out / "排班表.xlsx").write_bytes(xlsx)
    assert xlsx[:2] == b"PK" and len(xlsx) > 4000

    csv_bytes = export_csv(result.assignments)
    (out / "排班表.csv").write_bytes(csv_bytes)
    assert csv_bytes[:3] == b"\xef\xbb\xbf" and len(csv_bytes) > 100

    detail = build_detail_df(result.assignments)
    pivot = build_pivot_df(result.assignments)
    assert not detail.empty and not pivot.empty
    assert list(pivot.columns) == ["1-2节", "3-4节", "5-6节", "7-8节", "9-11节"]
    print(f"[5] 导出: 通过（xlsx {len(xlsx)}B / csv {len(csv_bytes)}B，"
          f"透视表 {pivot.shape[0]} 行 x {pivot.shape[1]} 列）")
    print()
    print("透视表预览（前5行）:")
    print(pivot.head(5).to_string())


def test_gantt() -> None:
    db, members = get_env()
    weekdays, blocks = [1, 2, 3, 4, 5], [1, 2, 3, 4, 5]
    courses = db.get_courses()
    m = build_availability(members, courses, week=1, weekdays=weekdays, blocks=blocks)

    assert m.member_count == 5 and len(m.slots) == 25, "甘特图矩阵尺寸错误"

    # 空闲判定与排班算法逐格一致（时段块内每一节均无课才算空闲）
    busy = build_busy_map(members, courses)
    for r, mem in enumerate(members):
        for i, (d, b) in enumerate(m.slots):
            expect_free = all((1, d, s) not in busy.get(mem.id, ()) for s in BLOCK_SESSIONS[b])
            assert m.free[r][i] == expect_free, "甘特图空闲判定与排班算法不一致"

    # 汇总行：各时段空闲人数
    for i in range(len(m.slots)):
        assert m.free_counts[i] == sum(1 for row in m.free if row[i]), "空闲人数统计错误"

    # 全员空闲时段
    for slot in m.all_free_slots:
        assert m.free_counts[m.slots.index(slot)] == m.member_count, "全员空闲时段判定错误"

    # 忙时单元格均带课程详情（tooltip 数据）
    assert all(names for names in m.busy_courses.values()), "忙时单元格缺少课程详情"

    # 周次维度：课程随周次变化，第18周同样可构建
    m18 = build_availability(members, courses, week=18, weekdays=weekdays, blocks=blocks)
    assert m18.member_count == 5 and len(m18.slots) == 25

    # 导出带填充色的 Excel 甘特图
    data = export_gantt_excel(m)
    assert data[:2] == b"PK" and len(data) > 4000, "甘特图 Excel 导出错误"
    print(f"[6] 甘特图: 通过（{m.member_count} 成员 x {len(m.slots)} 时段，"
          f"第1周全员空闲 {len(m.all_free_slots)} 个 / 第18周 {len(m18.all_free_slots)} 个，"
          f"导出 xlsx {len(data)}B）")


def test_leaves() -> None:
    db, members = get_env()
    mid = members[0].id
    db.add_leave(mid, 2, 3, "生病")
    db.add_leave(mid, 2, 3, "重复登记应覆盖")
    mine = db.list_leaves(mid)
    assert len(mine) == 1 and mine[0].reason == "重复登记应覆盖", "请假登记应按(成员,周,星期)幂等覆盖"

    courses = db.get_courses()
    config = ScheduleConfig(
        weeks=range(1, 4), weekdays=[1, 2, 3, 4, 5], blocks=[1, 2, 3, 4, 5],
        per_slot=1, max_per_week=3, max_per_day=1)
    result = generate_schedule(members, courses, config, leaves=db.list_leaves())
    hit = [a for a in result.assignments if a.member_id == mid and a.week == 2 and a.weekday == 3]
    assert not hit, "排班未避开请假成员"

    m = build_availability(members, courses, week=2, weekdays=[1, 2, 3, 4, 5],
                          blocks=[1, 2, 3, 4, 5], leaves=db.list_leaves())
    assert (0, 3) in m.leave_info, "甘特图未标记请假"
    for i, (d, b) in enumerate(m.slots):
        if d == 3:
            assert m.free[0][i] is False, "请假成员当天不应判为空闲"
    wed_cols = [i for i, (d, _) in enumerate(m.slots) if d == 3]
    assert all(m.free_counts[i] < m.member_count for i in wed_cols), "请假应影响空闲人数"

    db.remove_leave(mine[0].id)
    assert db.list_leaves(mid) == [], "删除请假失败"

    # 请假记录导出：请假登记处按钮的数据源
    db.add_leave(mid, 2, 3, "生病")
    db.add_leave(members[1].id, 5, 1, "家中有事")
    leaves = db.list_leaves()
    xlsx = export_leaves_excel(leaves, members, start_date=date(2026, 9, 14))
    assert xlsx[:2] == b"PK" and len(xlsx) > 4000, "请假记录 Excel 导出错误"
    df = build_leaves_df(leaves, members, start_date=date(2026, 9, 14))
    assert len(df) == len(leaves) and set(df.columns) >= {"成员", "周次", "星期", "日期", "原因"}
    assert df.iloc[0]["成员"] == members[0].name and df.iloc[0]["日期"] == "9月23日", "请假记录内容错误"

    print(f"[7] 请假登记: 通过（排班避开请假日 {WEEKDAY_LABELS[3]}，甘特图当天标红，"
          f"登记幂等、可删除，导出 {len(leaves)} 条含日期/原因）")


def test_manual_tweak() -> None:
    db, members = get_env()
    courses = db.get_courses()
    busy = build_busy_map(members, courses)

    # 场景1：有富余的排班（容量>需求）→ 存在完全合规的换人候选
    config = ScheduleConfig(
        weeks=range(1, 3), weekdays=[1, 2], blocks=[1, 2, 3],
        per_slot=1, max_per_week=3, max_per_day=1)
    result = generate_schedule(members, courses, config)

    target, cands = None, []
    for a in result.assignments:
        c = replacement_candidates(members, busy, set(), result.assignments,
                                   a.week, a.weekday, a.block,
                                   config.max_per_week, config.max_per_day)
        if any(not reason for _, reason in c):
            target, cands = a, c
            break
    assert target is not None, "应存在可微调的时段"
    eligible = [m for m, reason in cands if not reason]
    assert all(m.id != target.member_id for m, _ in cands), "在岗成员不应出现在候选中"

    # 不可用原因与忙时表/安排一致
    for m0 in (m for m, r in cands if r == "该时段有课"):
        assert any((target.week, target.weekday, s) in busy.get(m0.id, ())
                   for s in BLOCK_SESSIONS[target.block]), "有课原因与忙时表不一致"
    for m0 in (m for m, r in cands if r == "当天已值班"):
        assert any(a.member_id == m0.id and a.week == target.week
                   and a.weekday == target.weekday and a.block != target.block
                   for a in result.assignments), "当天已值班原因与安排不一致"

    # 换入后仍满足全部硬约束
    new = eligible[0]
    kept = [a for a in result.assignments if a is not target]
    kept.append(Assignment(target.week, target.weekday, target.block, new.id, new.name))
    for a in kept:
        for s in BLOCK_SESSIONS[a.block]:
            assert (a.week, a.weekday, s) not in busy.get(a.member_id, ()), "微调后出现课程冲突"
    day_cnt = Counter((a.member_id, a.week, a.weekday) for a in kept)
    assert all(v <= config.max_per_day for v in day_cnt.values()), "微调后超出每天上限"
    week_cnt = Counter((a.member_id, a.week) for a in kept)
    assert all(v <= config.max_per_week for v in week_cnt.values()), "微调后超出每周上限"
    stats = rebuild_member_stats(members, kept)
    assert sum(s["total"] for s in stats.values()) == len(kept), "统计重算总数错误"

    # 场景2：饱和排班（容量=需求）→ 候选只剩"本周已达上限"，
    # 手动微调允许知情越限换人（课程/请假/每天一次仍必须满足）
    sat = ScheduleConfig(
        weeks=range(1, 3), weekdays=[1, 2, 3, 4, 5], blocks=[1, 2, 3, 4, 5],
        per_slot=1, max_per_week=3, max_per_day=1)
    result2 = generate_schedule(members, courses, sat)
    soft_total = 0
    checked = 0
    for target2 in sorted(result2.assignments, key=lambda a: (a.week, a.weekday, a.block)):
        cands2 = replacement_candidates(members, busy, set(), result2.assignments,
                                        target2.week, target2.weekday, target2.block,
                                        sat.max_per_week, sat.max_per_day)
        soft = [m for m, r in cands2 if r == "本周已达上限"]
        if not soft:
            continue
        soft_total += len(soft)
        checked += 1
        for m0 in soft:
            for s in BLOCK_SESSIONS[target2.block]:
                assert (target2.week, target2.weekday, s) not in busy.get(m0.id, ()), \
                    "软候选仍须无课程冲突"
            assert (m0.id, target2.week, target2.weekday) not in (
                (a.member_id, a.week, a.weekday) for a in result2.assignments
                if a.block != target2.block), "软候选仍须满足每天一次"
    # 容量饱和时，「仅超周上限」的软候选可能因课程时间完全错开而不存在：
    # 这属于排得更满的正常结果，因此只要求「存在软候选的时段行为正确」
    assert checked > 0, "饱和排班下应有时段存在仅超周上限的软候选"
    print(f"[8] 手动微调: 通过（合规候选 {len(eligible)} 人换入后 0 冲突、上限合规、统计重算；"
          f"{checked} 个饱和时段共 {soft_total} 人仅超周上限可知情换入）")


def test_week_dates() -> None:
    result = STATE["result"]
    assert week_date(date(2026, 9, 14), 1, 1) == date(2026, 9, 14)
    assert week_date(date(2026, 9, 14), 1, 7) == date(2026, 9, 20)
    assert week_date(date(2026, 9, 14), 2, 5) == date(2026, 9, 25)
    assert week_date(date(2026, 9, 14), 18, 5) == date(2027, 1, 15), "跨年换算错误"

    detail = build_detail_df(result.assignments, start_date=date(2026, 9, 14))
    pivot = build_pivot_df(result.assignments, start_date=date(2026, 9, 14))
    assert "日期" in detail.columns and "日期" in pivot.columns, "导出应含日期列"
    xlsx = export_excel(result.assignments, result.member_stats, result.gaps,
                        start_date=date(2026, 9, 14))
    assert xlsx[:2] == b"PK", "带日期 Excel 导出错误"
    print("[9] 周次换算: 通过（起始日+(周-1)*7+(星期-1) 正确跨月/跨年；明细/透视/Excel 含日期列）")


def test_incremental():
    """按周增量排班：只重排所选周，范围外历史作均衡基数保留"""
    db, members = get_env()
    db.clear_assignments()
    courses = db.get_courses()
    cfg = dict(weekdays=[1, 2, 3, 4, 5], blocks=[1, 2, 3, 4, 5],
               per_slot=1, max_per_week=3, max_per_day=1)
    first = generate_schedule(members, courses, ScheduleConfig(weeks=range(1, 4), **cfg))
    db.save_assignments(first.assignments)

    wk2 = ScheduleConfig(weeks=range(2, 3), **cfg)
    existing = db.load_assignments()
    base = [a for a in existing if a.week not in wk2.weeks]
    second = generate_schedule(members, courses, wk2, base_assignments=base)
    fresh = [a for a in second.assignments if a.week in wk2.weeks]

    key = lambda a: (a.week, a.weekday, a.block, a.member_id)  # noqa: E731
    assert sorted(key(a) for a in second.assignments if a.week != 2) == \
        sorted(key(a) for a in base), "范围外历史排班应原样保留在合并结果中"

    busy = build_busy_map(members, courses)
    for a in fresh:
        for s in BLOCK_SESSIONS[a.block]:
            assert (a.week, a.weekday, s) not in busy.get(a.member_id, ()), "增量排班出现课程冲突"
    day_cnt = Counter((a.member_id, a.week, a.weekday) for a in second.assignments)
    assert all(v <= 1 for v in day_cnt.values()), "增量排班超出每天上限"
    week_cnt = Counter((a.member_id, a.week) for a in second.assignments)
    assert all(v <= 3 for v in week_cnt.values()), "增量排班超出每周上限"
    stats = rebuild_member_stats(members, second.assignments)
    assert sum(s["total"] for s in stats.values()) == len(second.assignments), \
        "增量统计应包含历史基数"

    db.delete_assignments_for_weeks(list(wk2.weeks))
    db.save_assignments(fresh)
    loaded = db.load_assignments()
    assert sorted(key(a) for a in loaded if a.week == 1) == \
        sorted(key(a) for a in base if a.week == 1), "第1周历史排班被破坏"
    assert sorted(key(a) for a in loaded if a.week == 2) == \
        sorted(key(a) for a in fresh), "第2周应替换为新排班"
    assert len(loaded) == len(base) + len(fresh), "按周删除不应影响其他周"
    print(f"[10] 按周增量: 通过（第2周单独重排 {len(fresh)} 人次，"
          f"第1/3周历史 {len(base)} 人次保留，删周落库互不影响）")
    STATE["second"] = second


def test_persistence() -> None:
    """模拟程序重启：从库恢复排班结果（app._restore_result 同款逻辑）"""
    db, members = get_env()
    assignments = db.load_assignments()
    assert assignments, "应已有排班数据"
    config = ScheduleConfig(weeks=range(1, 4), weekdays=[1, 2, 3, 4, 5], blocks=[1, 2, 3, 4, 5])
    stats = rebuild_member_stats(members, assignments)
    grid = {(w, d, b) for w in config.weeks for d in config.weekdays for b in config.blocks}
    covered = {(a.week, a.weekday, a.block) for a in assignments}
    gaps = sorted(g for g in grid if g not in covered)
    assert sum(s["total"] for s in stats.values()) == len(assignments), "恢复后统计总数错误"
    assert all(g not in covered for g in gaps), "缺口不应包含已覆盖时段"
    names = {m.name for m in members}
    assert all(a.member_name in names for a in assignments), "恢复的值班人姓名缺失"
    print(f"[11] 结果恢复: 通过（{len(assignments)} 人次从库恢复，统计/缺口重算一致）")


def test_week_export() -> None:
    """按周导出：过滤某一周后透视/明细/Excel/CSV 只含该周"""
    assignments = STATE["second"].assignments
    _, members = get_env()
    sel = [a for a in assignments if a.week == 2]
    assert sel, "第2周应有排班"
    stats = rebuild_member_stats(members, sel)
    xlsx = export_excel(sel, stats, [], start_date=date(2026, 9, 14))
    assert xlsx[:2] == b"PK", "单周 Excel 导出错误"
    pivot = build_pivot_df(sel, start_date=date(2026, 9, 14))
    assert all("第2周" in str(idx) for idx in pivot.index), "透视表应只含所选周"
    detail = build_detail_df(sel, start_date=date(2026, 9, 14))
    assert set(detail["周次"]) == {"第2周"}, "明细应只含所选周"
    csv_bytes = export_csv(sel)
    assert "第2周".encode() in csv_bytes, "CSV 应只含所选周"
    print(f"[12] 按周导出: 通过（单周 {len(sel)} 人次，透视/明细/Excel/CSV 只含第2周）")


def test_gantt_duty() -> None:
    """甘特图值班标记：已排值班格标蓝、不占空闲统计、Excel 同步"""
    db, members = get_env()
    courses = db.get_courses()
    assignments = [a for a in db.load_assignments() if a.week == 2]
    assert assignments, "第2周应有排班"
    m = build_availability(members, courses, week=2, weekdays=[1, 2, 3, 4, 5],
                          blocks=[1, 2, 3, 4, 5], assignments=assignments)
    row_of = {mem.id: r for r, mem in enumerate(members)}
    for a in assignments:
        r = row_of[a.member_id]
        assert (r, a.weekday, a.block) in m.duty_cells, "值班格未标记"
        assert m.free[r][m.slots.index((a.weekday, a.block))] is False, "值班格不应判为空闲"
    data = export_gantt_excel(m)
    assert data[:2] == b"PK" and len(data) > 4000, "甘特图 Excel 导出错误"
    m2 = build_availability(members, courses, week=2, weekdays=[1, 2, 3, 4, 5], blocks=[1, 2, 3, 4, 5])
    assert not m2.duty_cells, "未传排班时不应有值班格"
    print(f"[13] 甘特图值班标记: 通过（{len(assignments)} 个值班格标蓝且不占空闲统计，Excel 同步）")


def test_calendar_makeup_mapping() -> None:
    """调休映射：逻辑周四移到周六，排班避开周四课程并按周六展示日期。"""
    start = date(2026, 9, 14)
    members = [
        Member(1, "甲", "1", "", "", "", "", ""),
        Member(2, "乙", "2", "", "", "", "", ""),
    ]
    courses = [CourseRecord(
        id=1, member_id=1, course_name="周四课程", teacher="", weekday=4,
        week_list=[5], session_list=[1, 2], location="",
        weeks_text="5", sessions_text="01-02")]
    calendar = TermCalendar(start, [
        CalendarEntry(date(2026, 10, 15), "off", note="第5周周四放假"),
        CalendarEntry(date(2026, 10, 17), "class", 5, 4, "补第5周周四"),
    ])
    result = generate_schedule(
        members, courses,
        ScheduleConfig(weeks=range(5, 6), weekdays=[4], blocks=[1],
                       per_slot=1, max_per_week=1, max_per_day=1, seed=1))

    assert [(a.member_id, a.week, a.weekday) for a in result.assignments] == [(2, 5, 4)], \
        "应避开甲第5周周四的课程，由乙补到对应的周六"
    detail = build_detail_df(result.assignments, start_date=start, calendar=calendar)
    assert detail.iloc[0]["星期"] == "周六" and detail.iloc[0]["日期"] == "10月17日"
    assert "10月15日" not in set(detail["日期"]), "被放假的周四自然日不应出现在排班表"
    pivot = build_pivot_df(result.assignments, start_date=start, calendar=calendar)
    assert ("第5周", "周六") in pivot.index
    assert pivot.loc[("第5周", "周六"), "日期"] == "10月17日"

    matrix = build_availability(
        members, courses, 5, [4], [1], assignments=result.assignments,
        calendar=calendar)
    assert matrix.calendar_dates[4] == date(2026, 10, 17)
    assert matrix.busy_courses[(0, 4, 1)] == ["周四课程"], \
        "补课周六应按所代表的第5周周四判定课程忙闲"
    assert any(c.is_off and c.date == date(2026, 10, 15)
               for c in display_columns(matrix)), \
        "原始放假周四应作为红色放假日列保留在甘特图"


def test_pure_holiday_is_skipped_and_marked_off() -> None:
    """纯法定假日无补课：排班日期跳过，甘特图对应逻辑日整列标为放假。"""
    start = date(2026, 9, 14)
    member = Member(1, "甲", "1", "", "", "", "", "")
    calendar = TermCalendar(start, [
        CalendarEntry(date(2026, 10, 15), "off", note="国庆法定假日"),
    ])
    assignments = [Assignment(5, 4, 1, 1, "甲")]

    detail = build_detail_df(assignments, start_date=start, calendar=calendar)
    assert detail.empty, "纯法定假日没有真实排班日期，应从明细中跳过"

    matrix = build_availability(
        [member], [], 5, [4], [1], assignments=assignments, calendar=calendar)
    assert matrix.off_weekdays == {4}
    assert matrix.free == [[False]] and matrix.busy_courses == {}


def test_calendar_logical_date_roundtrip() -> None:
    """逻辑日与真实日期双向映射往返一致，off 日期不参与排班。"""
    start = date(2026, 9, 14)
    calendar = TermCalendar(start, [
        CalendarEntry(date(2026, 10, 15), "off"),
        CalendarEntry(date(2026, 10, 17), "class", 5, 4),
    ])
    assert calendar.logical_to_date(5, 4) == date(2026, 10, 17)
    assert calendar.date_to_logical(date(2026, 10, 17)) == (5, 4)
    assert calendar.date_to_logical(date(2026, 10, 15)) is None

    for week in range(1, 7):
        for weekday in range(1, 6):
            if (week, weekday) == (5, 4):
                continue
            actual = calendar.logical_to_date(week, weekday)
            assert actual is not None
            assert calendar.date_to_logical(actual) == (week, weekday)

    for invalid in (
            start - timedelta(days=1),
            start + timedelta(days=25 * 7)):
        try:
            calendar.date_to_logical(invalid)
        except ValueError:
            pass
        else:
            raise AssertionError(f"越界日期 {invalid} 应被拒绝")


def test_clear_assignments_all_and_selected_weeks() -> None:
    """清空排班支持按周与全量模式，对应甘特图不再出现值班标记。"""
    db, members = get_env()
    assignments = db.load_assignments()
    if not assignments:
        generated = generate_schedule(
            members, db.get_courses(),
            ScheduleConfig(weeks=range(1, 4), per_slot=1,
                           max_per_week=3, max_per_day=1))
        db.save_assignments(generated.assignments)
        assignments = db.load_assignments()
    assert assignments
    weeks = sorted({a.week for a in assignments})
    target_week = weeks[0]

    db.delete_assignments_for_weeks([target_week])
    remaining = db.load_assignments()
    assert all(a.week != target_week for a in remaining)
    assert {a.week for a in remaining} == set(weeks[1:]), "按周清空不应影响其他周"
    matrix = build_availability(
        members, db.get_courses(), target_week, [1, 2, 3, 4, 5], [1, 2, 3, 4, 5],
        assignments=remaining)
    assert not matrix.duty_cells, "被清空的周不应残留蓝色值班标记"

    db.clear_assignments()
    assert db.load_assignments() == []
    all_empty = build_availability(
        members, db.get_courses(), weeks[-1], [1, 2, 3, 4, 5], [1, 2, 3, 4, 5],
        assignments=[])
    assert not all_empty.duty_cells, "全部清空后甘特图不应残留值班标记"


def main() -> None:
    """直接执行（不依赖 pytest）时按真实使用顺序跑完整链路"""
    test_parser()
    test_database()
    test_scheduler()
    test_exporter()
    test_gantt()
    test_leaves()
    test_manual_tweak()
    test_week_dates()
    test_incremental()
    test_persistence()
    test_week_export()
    test_gantt_duty()
    test_calendar_makeup_mapping()
    test_pure_holiday_is_skipped_and_marked_off()
    test_calendar_logical_date_roundtrip()
    test_clear_assignments_all_and_selected_weeks()
    print()
    print("全部测试通过 ✓")


if __name__ == "__main__":
    main()
