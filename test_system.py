"""端到端测试：解析 -> 入库 -> 排班(无冲突+均衡) -> 导出"""

from __future__ import annotations

import copy
import random
import tempfile
from collections import Counter
from pathlib import Path

from duty_system.database import Database
from duty_system.exporter import build_detail_df, build_pivot_df, export_csv, export_excel
from duty_system.gantt import build_availability, export_gantt_excel
from duty_system.parser import BLOCK_SESSIONS, WEEKDAY_LABELS, parse_schedule_path
from duty_system.scheduler import ScheduleConfig, build_busy_map, generate_schedule

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


if __name__ == "__main__":
    test_parser()
    db, members = test_database()
    result = test_scheduler(db, members)
    test_exporter(result)
    test_gantt(db, members)
    print()
    print("全部测试通过 ✓")
