"""端到端测试：解析 -> 入库 -> 排班(无冲突+均衡) -> 导出 -> 甘特/请假/微调/日期"""

from __future__ import annotations

import copy
import random
import tempfile
from collections import Counter
from datetime import date
from pathlib import Path

from duty_system.database import Assignment, Database
from duty_system.exporter import (
    build_detail_df, build_pivot_df, export_csv, export_excel, week_date,
)
from duty_system.gantt import build_availability, export_gantt_excel
from duty_system.parser import BLOCK_SESSIONS, WEEKDAY_LABELS, parse_schedule_path
from duty_system.scheduler import (
    ScheduleConfig, build_busy_map, generate_schedule, rebuild_member_stats,
    replacement_candidates,
)

SAMPLE = Path(__file__).parent / "samples" / "学生个人课表_2307724110.xls"

TEST_NAMES = ["王思远", "李慧敏", "张承宇", "陈晓露"]
TEST_STUDENT_IDS = ["2307724101", "2307724102", "2307724103", "2307724104"]


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
    assert s.name == "刘雨昂", f"姓名解析错误: {s.name}"
    assert s.student_id == "2307724110"
    assert s.term == "2026-2027-1"
    assert s.class_name == "24国际商务双语4班"
    assert len(s.courses) == 37, f"课程数错误: {len(s.courses)}"

    names = {c.course_name for c in s.courses}
    assert "高等数学(Ⅱ)-1" in names, "无教师课程(高等数学)解析失败"
    assert not any(c.teacher for c in s.courses if c.course_name == "高等数学(Ⅱ)-1")

    # 教师职称被换行打断的案例
    sy = [c for c in s.courses if "企业运营管理" in c.course_name]
    assert sy and all(c.teacher == "唐心智" for c in sy), "断行教师名解析失败"
    # 周次解析
    pe = next(c for c in s.courses if c.course_name.startswith("体育Ⅲ"))
    assert pe.week_list == [1, 3, 5, 7, 9, 11, 13, 15], f"单周周次解析错误: {pe.week_list}"
    mds = next(c for c in s.courses if c.course_name.startswith("毛泽东"))
    assert mds.week_list == [1, 2, 3, 4, 6, 7, 8, 9, 10], f"混合周次解析错误: {mds.week_list}"
    print(f"[1] 解析器: 通过（{len(s.courses)} 条课程记录，姓名/学号/断行教师/周次/去重均正确）")


def test_database() -> tuple[Database, list]:
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


def test_scheduler(db: Database, members: list) -> None:
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
    return result


def test_exporter(result) -> None:
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


def test_gantt(db: Database, members: list) -> None:
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


def test_leaves(db: Database, members: list) -> None:
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
    print(f"[7] 请假登记: 通过（排班避开请假日 {WEEKDAY_LABELS[3]}，甘特图当天标红，"
          "登记幂等、可删除）")


def test_manual_tweak(db: Database, members: list) -> None:
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

    # 场景2：饱和排班（容量=需求，默认配置即如此）→ 剩余候选仅"本周已达上限"，
    # 手动微调允许知情越限换人（课程/请假/每天一次仍必须满足）
    sat = ScheduleConfig(
        weeks=range(1, 3), weekdays=[1, 2, 3, 4, 5], blocks=[1, 2, 3, 4, 5],
        per_slot=1, max_per_week=3, max_per_day=1)
    result2 = generate_schedule(members, courses, sat)
    target2 = result2.assignments[0]
    cands2 = replacement_candidates(members, busy, set(), result2.assignments,
                                    target2.week, target2.weekday, target2.block,
                                    sat.max_per_week, sat.max_per_day)
    soft = [m for m, r in cands2 if r == "本周已达上限"]
    assert soft, "饱和排班应存在仅超周上限的可换候选"
    for m0 in (m for m, r in cands2 if r in ("", "本周已达上限")):
        for s in BLOCK_SESSIONS[target2.block]:
            assert (target2.week, target2.weekday, s) not in busy.get(m0.id, ()), "软候选仍须无课程冲突"
        assert (m0.id, target2.week, target2.weekday) not in (
            (a.member_id, a.week, a.weekday) for a in result2.assignments
            if a.block != target2.block), "软候选仍须满足每天一次"
    print(f"[8] 手动微调: 通过（合规候选 {len(eligible)} 人换入后 0 冲突、上限合规、统计重算；"
          f"饱和排班下 {len(soft)} 人仅超周上限可知情换入）")


def test_week_dates(result) -> None:
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


if __name__ == "__main__":
    test_parser()
    db, members = test_database()
    result = test_scheduler(db, members)
    test_exporter(result)
    test_gantt(db, members)
    test_leaves(db, members)
    test_manual_tweak(db, members)
    test_week_dates(result)
    print()
    print("全部测试通过 ✓")
