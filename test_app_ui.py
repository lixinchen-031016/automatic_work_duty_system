"""界面层测试（offscreen）：数据缓存、后台排班、甘特表复用、参数持久化、手动微调落库

说明：这些用例通过 QT_QPA_PLATFORM=offscreen 在无显示器环境运行，
      QSettings 由 conftest 重定向到临时目录，不会影响用户真实配置。
"""

from __future__ import annotations

import random
import sqlite3
import time
from pathlib import Path

import pytest

pytest.importorskip("PySide6.QtWidgets")
from PySide6.QtCore import Qt  # noqa: E402
from PySide6.QtWidgets import QMessageBox  # noqa: E402

import app as appmod  # noqa: E402
from duty_system.database import Member  # noqa: E402
from duty_system.parser import BLOCK_SESSIONS, Course, ParsedSchedule  # noqa: E402
from duty_system.scheduler import (  # noqa: E402
    compute_gaps, replacement_candidates, build_busy_map,
)


@pytest.fixture(autouse=True)
def no_modal_dialogs(monkeypatch):
    """测试中不允许弹模态框（无人点击会直接卡死）"""
    calls: list[tuple] = []
    monkeypatch.setattr(QMessageBox, "information",
                        staticmethod(lambda *a, **k: calls.append(("info", a[1], a[2]))))
    monkeypatch.setattr(QMessageBox, "warning",
                        staticmethod(lambda *a, **k: calls.append(("warn", a[1], a[2]))))
    monkeypatch.setattr(QMessageBox, "question",
                        staticmethod(lambda *a, **k: (calls.append(("ask", a[1], a[2])),
                                                      QMessageBox.No)[1]))
    return calls


@pytest.fixture
def window(qt_app, settings, tmp_path: Path):
    win = appmod.MainWindow(db_path=tmp_path / "ui.db")
    yield win
    win.close()
    win.deleteLater()
    qt_app.processEvents()


def seed_members(win, count: int = 12, seed: int = 7) -> None:
    """写入 count 名成员的随机课表"""
    rng = random.Random(seed)
    for i in range(count):
        win.db.upsert_member(ParsedSchedule(
            name=f"成员{i:02d}", student_id=str(1000 + i), term="2026-2027-1",
            class_name="测试班", major="专业", department="学院", file_name=f"f{i}.xls",
            courses=[Course(course_name=f"课程{c}", teacher="教师",
                            weekday=rng.randint(1, 7), weeks_text="1-18",
                            week_list=list(range(1, 19)), sessions_text="01-02",
                            session_list=sorted(rng.sample(range(1, 12), 2)))
                     for c in range(24)]))
    win.invalidate_cache()
    win.refresh_members()


def clear_members(win) -> None:
    """清空成员（含级联课程/排班/请假），用于验证表格结构变化"""
    for m in win.members():
        win.db.delete_member(m.id)
    win.invalidate_cache()
    win.refresh_members()


def wait_idle(win, qt_app, timeout: float = 30.0) -> bool:
    """等待后台任务结束（期间保持事件循环转动）"""
    deadline = time.time() + timeout
    while win.busy and time.time() < deadline:
        qt_app.processEvents()
        time.sleep(0.005)
    qt_app.processEvents()
    return not win.busy


def run_generate(win, qt_app, **widgets):
    for name, value in widgets.items():
        getattr(win, name).setValue(value)
    win.generate()
    assert wait_idle(win, qt_app), "排班后台任务超时未结束"


# --------------------------------------------------------------------------- #

def test_settings_roundtrip(window, settings) -> None:
    """参数读写在同一份声明表里，保存 -> 修改 -> 恢复应完全一致"""
    window.week_from.setValue(3)
    window.week_to.setValue(9)
    window.per_slot.setValue(2)
    window.max_week.setValue(4)
    window.max_day.setValue(2)
    window.seed.setValue(123)
    window.gantt_week.setValue(5)
    window.term_start.setDate(window.term_start.date().fromString("2026-09-14", "yyyy-MM-dd"))
    window.weekday_checks[6].setChecked(True)
    window.weekday_checks[1].setChecked(False)
    window.block_checks[5].setChecked(False)
    window._save_settings()

    window.week_from.setValue(1)
    window.per_slot.setValue(1)
    window.seed.setValue(0)
    window.weekday_checks[6].setChecked(False)
    window.weekday_checks[1].setChecked(True)
    window.block_checks[5].setChecked(True)
    window._load_settings()

    assert (window.week_from.value(), window.week_to.value()) == (3, 9)
    assert (window.per_slot.value(), window.max_week.value(), window.max_day.value()) == (2, 4, 2)
    assert window.seed.value() == 123 and window.gantt_week.value() == 5
    assert window.term_start.date().toString("yyyy-MM-dd") == "2026-09-14"
    assert window.weekday_checks[6].isChecked() and not window.weekday_checks[1].isChecked()
    assert not window.block_checks[5].isChecked()

    # 新增参数只需在 _param_specs 登记，键名自动生成
    keys = {key for _kind, _w, key, _d in window._param_specs()}
    assert {"week_from", "term_start", "weekday_6", "block_5"} <= keys


def test_last_config_persisted_and_reused_on_restart(qt_app, settings, tmp_path, monkeypatch) -> None:
    """重启后缺口按「生成时参数」计算，而不是当前界面参数"""
    db_path = tmp_path / "restore.db"
    win = appmod.MainWindow(db_path=db_path)
    seed_members(win, 6)
    run_generate(win, qt_app, week_from=1, week_to=2, per_slot=1)
    saved = win._last_config
    assert saved is not None
    win.close()
    win.deleteLater()
    qt_app.processEvents()

    # 用户把界面参数改成差别很大的值（每周 3 人、1-18 周）
    win2 = appmod.MainWindow(db_path=db_path)
    win2.week_to.setValue(18)
    win2.per_slot.setValue(3)
    win2._restore_result()

    assert win2.result is not None
    assert win2._last_config is not None, "应恢复上次生成时的参数"
    assert win2._last_config.weeks == saved.weeks
    assert win2._last_config.per_slot == 1
    expected = compute_gaps(win2.result.assignments, saved)
    assert win2.result.gaps == expected, "缺口应按保存的参数计算，而非当前界面参数"
    win2.close()
    win2.deleteLater()


def test_last_config_is_per_database(qt_app, settings, tmp_path) -> None:
    """切换数据库后，各自沿用自己那次的排班参数"""
    path_a, path_b = tmp_path / "a.db", tmp_path / "b.db"
    win_a = appmod.MainWindow(db_path=path_a)
    seed_members(win_a, 6)
    run_generate(win_a, qt_app, week_from=1, week_to=2, per_slot=1)

    win_b = appmod.MainWindow(db_path=path_b)
    seed_members(win_b, 6, seed=42)
    run_generate(win_b, qt_app, week_from=1, week_to=4, per_slot=2)

    assert win_a._last_config.weeks == range(1, 3)
    assert win_b._last_config.weeks == range(1, 5)
    assert win_a._last_config.per_slot == 1 and win_b._last_config.per_slot == 2

    # 重新打开 A 库：应恢复 A 自己的参数，而不是最近一次生成（B 库）的参数
    win_a.close(); win_a.deleteLater()
    win_a2 = appmod.MainWindow(db_path=path_a)
    win_a2._restore_result()
    assert win_a2._last_config.weeks == range(1, 3), "A 库应恢复自己的周范围"
    assert win_a2._last_config.per_slot == 1, "A 库应恢复自己的每时段人数"
    assert win_a2.result.gaps == compute_gaps(win_a2.result.assignments, win_a2._last_config)
    for w in (win_a2, win_b):
        w.close(); w.deleteLater()


def test_cache_avoids_repeated_full_queries(window, monkeypatch) -> None:
    """成员/课程/请假只查一次，写操作后失效"""
    seed_members(window, 5)
    calls = {"members": 0, "courses": 0, "leaves": 0, "specials": 0}

    def counting(name, original):
        def wrapper(*args, **kwargs):
            calls[name] += 1
            return original(*args, **kwargs)
        return wrapper

    monkeypatch.setattr(window.db, "list_members", counting("members", window.db.list_members))
    monkeypatch.setattr(window.db, "get_courses", counting("courses", window.db.get_courses))
    monkeypatch.setattr(window.db, "list_leaves", counting("leaves", window.db.list_leaves))
    monkeypatch.setattr(window.db, "list_special_arrangements",
                        counting("specials", window.db.list_special_arrangements))
    window.invalidate_cache()

    for _ in range(3):
        window.members()
        window.courses()
        window.leaves()
        window.specials()
    assert calls == {"members": 1, "courses": 1, "leaves": 1, "specials": 1}, \
        f"缓存未生效，实际查询次数 {calls}"

    window.invalidate_cache()
    window.members()
    assert calls["members"] == 2, "失效后应重新查询"

    # 登记请假会写库 -> 缓存必须失效，否则界面读到旧数据
    window.invalidate_cache()
    window.leaves()
    window.db.add_leave(window.members()[0].id, 1, 1, "事假")
    window.invalidate_cache()
    assert len(window.leaves()) == 1, "写操作后应读到最新请假"


def test_async_generate_updates_ui_and_db(window, qt_app) -> None:
    """生成排班走后台线程：不阻塞（busy 标志）、结果落库、界面同步刷新"""
    seed_members(window, 12)
    run_generate(window, qt_app, week_from=1, week_to=4, per_slot=1)

    result = window.result
    assert result is not None and result.assignments
    assert not window.busy, "任务结束后应恢复空闲状态"
    assert window.btn_generate.isEnabled(), "任务结束后应重新启用生成按钮"
    assert window.tabs.currentIndex() == 0, "生成后应切到排班表页"

    stored = window.db.load_assignments()
    key = lambda a: (a.week, a.weekday, a.block, a.member_id)  # noqa: E731
    assert sorted(key(a) for a in stored) == sorted(key(a) for a in result.assignments), \
        "落库结果应与计算结果一致"
    assert window.gantt_table.rowCount() == 13, "甘特图应已刷新（成员数 + 汇总行）"
    assert window.btn_export_xlsx.isEnabled() and window.btn_export_png.isEnabled()
    assert "已重新排班" in window.statusBar().currentMessage()

    # 相同参数的第二次生成：未变化的安排应保持原行（不整周删除重建）
    conn = sqlite3.connect(window.db.path)
    before = {r[0]: r[1] for r in conn.execute("SELECT id, member_id FROM duty_assignments")}
    conn.close()
    run_generate(window, qt_app, week_from=1, week_to=4, per_slot=1)
    conn = sqlite3.connect(window.db.path)
    after = {r[0]: r[1] for r in conn.execute("SELECT id, member_id FROM duty_assignments")}
    conn.close()
    assert before == after, "重复生成相同排班不应重写数据行"


def test_gantt_table_reuses_items(window, qt_app) -> None:
    """甘特图刷新复用已有单元格对象，避免数千个 item 反复重建"""
    seed_members(window, 8)
    window.gantt_week.setValue(1)
    window.refresh_gantt()
    first = window.gantt_table.item(0, 0)
    assert first is not None
    text_week1 = first.text()

    # 换一周刷新：应复用同一批 item 对象，只更新内容/颜色
    window.gantt_week.setValue(3)
    window.refresh_gantt()
    again = window.gantt_table.item(0, 0)
    assert again is first, "刷新甘特图不应重建单元格对象"

    # 结构变化（成员数变化）时才允许重建
    assert isinstance(text_week1, str)
    clear_members(window)
    seed_members(window, 3, seed=99)
    window.refresh_gantt()
    assert window.gantt_table.rowCount() == 4, "成员数变化后应重建表格结构"
    assert window.gantt_table.item(0, 0) is not None
    assert window.gantt_table.rowCount() == len(window.members()) + 1


def test_manual_tweak_writes_only_target_slot(window, qt_app) -> None:
    """手动微调只重写目标槽位，其他行保持原样"""
    seed_members(window, 10)
    run_generate(window, qt_app, week_from=1, week_to=3, per_slot=1)

    members = window.members()
    busy = build_busy_map(members, window.courses())
    target = None
    for a in sorted(window.result.assignments, key=lambda x: (x.week, x.weekday, x.block)):
        cands = replacement_candidates(members, busy, set(), window.result.assignments,
                                       a.week, a.weekday, a.block, 3, 1)
        if any(not reason for _m, reason in cands):
            target = a
            break
    assert target is not None, "应存在可微调的目标时段"

    conn = sqlite3.connect(window.db.path)
    before = {(r[0], r[1]): r[2] for r in
              conn.execute("SELECT id, member_id, week FROM duty_assignments")}
    conn.close()

    new_member = next(m for m, reason in replacement_candidates(
        members, busy, set(), window.result.assignments, target.week, target.weekday,
        target.block, 3, 1) if not reason)
    window._apply_tweak(target.week, target.weekday, target.block, target.member_id,
                        new_member.id)
    qt_app.processEvents()

    conn = sqlite3.connect(window.db.path)
    after = {(r[0], r[1]): r[2] for r in
             conn.execute("SELECT id, member_id, week FROM duty_assignments")}
    conn.close()
    unchanged = {k: v for k, v in before.items() if k[1] != target.member_id}
    assert all(after.get(k) == v for k, v in unchanged.items()), "其他时段的行不应被改写"

    slot_rows = {(a.member_id) for a in window.db.load_assignments()
                 if (a.week, a.weekday, a.block) == (target.week, target.weekday, target.block)}
    assert slot_rows == {new_member.id}, "目标槽位应只保留换入的成员"
    assert all(a.member_id != target.member_id for a in window.db.load_assignments()
               if (a.week, a.weekday, a.block) == (target.week, target.weekday, target.block))
    assert window.result.gaps == compute_gaps(window.result.assignments, window._last_config)


def test_stale_marking_and_empty_state(window) -> None:
    """参数变化标记结果过期；无结果时给出引导文案"""
    window.refresh_members()
    assert "请上传" in window.summary_label.text() or "开始使用" in window.summary_label.text()

    seed_members(window, 3)
    window.refresh_members()
    window.result = None
    window._update_empty_state()
    assert "3 名成员" in window.summary_label.text()

    # 有结果后改参数 -> 提示过期
    window.result = appmod.ScheduleResult()
    window._stale = False
    window._mark_stale()
    assert "过期" in window.summary_label.text()

def test_busy_map_is_cached_and_reused(window, monkeypatch) -> None:
    """忙时表展开成本高，甘特图刷新不应重建它（改由缓存提供）"""
    seed_members(window, 8)
    import duty_system.gantt as gantt

    calls = {"n": 0}
    original = appmod.build_busy_map

    def counting(*args, **kwargs):
        calls["n"] += 1
        return original(*args, **kwargs)

    # app 与 gantt 都是「from ... import」，两处引用都要替换才能观察到调用
    monkeypatch.setattr(appmod, "build_busy_map", counting)
    monkeypatch.setattr(gantt, "build_busy_map", counting)
    window.invalidate_cache()

    window.refresh_gantt()
    first = calls["n"]
    assert first >= 1, "首次刷新应构建忙时表"

    for _ in range(3):
        window.refresh_gantt()
    assert calls["n"] == first, f"后续刷新应复用缓存，实际又构建了 {calls['n'] - first} 次"

    window.invalidate_cache()
    window.refresh_gantt()
    assert calls["n"] == first + 1, "课表变更后应重建忙时表"


def test_fill_table_reuses_cells_and_content_is_correct(window, qt_app) -> None:
    """通用表格填充复用单元格，且内容与 DataFrame 一致（含中文/数字/占位符）"""
    import pandas as pd

    import app as appmod

    table = window.pivot_table
    df1 = pd.DataFrame({"周次": ["第1周", "第1周"], "星期": ["周一", "周二"],
                        "1-2节": ["甲", "—"], "3-4节": ["—", "乙"]})
    appmod.fill_table(table, df1)
    assert table.rowCount() == 2 and table.columnCount() == 4
    first_cell = table.item(0, 0)
    assert first_cell.text() == "第1周"
    assert table.item(0, 2).text() == "甲"

    # 同样的尺寸再填一次：单元格对象复用，内容按新数据更新
    df2 = pd.DataFrame({"周次": ["第2周", "第2周"], "星期": ["周一", "周二"],
                        "1-2节": ["丙", "丁"], "3-4节": ["戊", "—"]})
    appmod.fill_table(table, df2)
    assert table.item(0, 0) is first_cell, "尺寸未变时应复用单元格对象"
    assert table.item(0, 0).text() == "第2周" and table.item(0, 2).text() == "丙"

    # 尺寸变化时重建结构，不残留旧数据
    appmod.fill_table(table, df1.iloc[:1, :3])
    assert (table.rowCount(), table.columnCount()) == (1, 3)
    assert table.item(0, 2).text() == "甲" and table.item(1, 0) is None

def test_tweak_candidate_list_flags(window) -> None:
    """微调候选项规则：硬约束原因不可选，仅超每周上限可「知情越限」"""
    from PySide6.QtWidgets import QListWidget

    member_hard = Member(id=1, name="有课的", student_id="1", term="", class_name="",
                                major="", department="", file_name="")
    member_soft = Member(id=2, name="满周上限的", student_id="2", term="",
                                class_name="", major="", department="", file_name="")
    member_ok = Member(id=3, name="可换的", student_id="3", term="", class_name="",
                              major="", department="", file_name="")
    cands = [(member_hard, "该时段有课"), (member_soft, "本周已达上限"), (member_ok, "")]

    lst = QListWidget()
    window._fill_tweak_candidates(lst, cands, target_mid=9)
    labels = [lst.item(i).text() for i in range(lst.count())]
    assert labels[0].startswith("（移除该值班人）"), "应提供「移除该值班人」选项"

    items = {lst.item(i).data(Qt.UserRole): lst.item(i) for i in range(lst.count())}
    assert items[member_hard.id].flags() & Qt.ItemIsEnabled
    assert not (items[member_hard.id].flags() & Qt.ItemIsSelectable), \
        "课程冲突等硬约束原因不可被选中"
    assert items[member_soft.id].flags() & Qt.ItemIsSelectable, "本周已达上限属软约束，可选"
    assert "超过每周上限" in items[member_soft.id].toolTip(), "软约束应有知情提示"
    assert items[member_ok.id].text() == member_ok.name and items[member_ok.id].flags() & Qt.ItemIsSelectable

    # 纯新增（无当前值班人）时不显示「移除」项
    lst2 = QListWidget()
    window._fill_tweak_candidates(lst2, [], target_mid=-1)
    assert lst2.count() == 1 and "暂无其他成员" in lst2.item(0).text()

def snapshot_gantt(win) -> list[tuple]:
    """把甘特图全部单元格的可见状态拍下来，用于比对「缓存后」与「全量重建」是否一致"""
    t = win.gantt_table
    out = []
    for r in range(t.rowCount()):
        row = []
        for c in range(t.columnCount()):
            item = t.item(r, c)
            if item is None:
                row.append(None)
            else:
                brush = item.foreground()
                fg = None if brush.style() == Qt.NoBrush else brush.color().name()
                row.append((item.text(), item.background().color().name(), fg, item.toolTip()))
        out.append(tuple(row))
    return out


def test_gantt_cell_cache_never_shows_stale_state(window, qt_app) -> None:
    """单元格状态缓存不能导致显示与数据不一致

    缓存按 (行, 列) 记录上次写入的状态并跳过无变化的单元格 —— 一旦判断错误
    就会显示过期颜色。这里用「反复换周 + 改筛选后回到原状态」的方式，
    与完全重建（清空缓存）的结果逐格比对。
    """
    seed_members(window, 10)
    run_generate(window, qt_app, week_from=1, week_to=4, per_slot=1)

    window.gantt_week.setValue(1)
    window.refresh_gantt()
    expected = snapshot_gantt(window)

    # 来回切换周次与筛选条件，让缓存经历大量状态变化
    for week in (2, 3, 4, 1, 3, 1):
        window.gantt_week.setValue(week)
        window.refresh_gantt()
    window.weekday_checks[6].setChecked(True)
    window.refresh_gantt()
    window.weekday_checks[6].setChecked(False)
    window.refresh_gantt()
    window.gantt_week.setValue(1)
    window.refresh_gantt()

    after_cache = snapshot_gantt(window)
    assert after_cache == expected, "重复刷新后显示与首次渲染不一致（缓存未失效）"

    # 强制全量重建（清空缓存 + 结构重置）后应得到完全相同的结果
    window._gantt_cell_state = {}
    window.gantt_table.clearContents()
    window.gantt_table.setRowCount(0)
    window.refresh_gantt()
    assert snapshot_gantt(window) == expected, "全量重建结果与缓存路径不一致"


def test_gantt_renders_expected_colors_and_tooltips(window, qt_app) -> None:
    """颜色语义与 tooltip 内容：空闲/有课/值班/请假/全员空闲"""
    seed_members(window, 6)
    run_generate(window, qt_app, week_from=1, week_to=2, per_slot=1)
    window.gantt_week.setValue(1)
    window.refresh_gantt()

    m = window.gantt_matrix
    t = window.gantt_table
    assert m is not None

    # 抽查若干单元格：状态必须与矩阵一致
    checked = {"free": 0, "busy": 0, "duty": 0}
    for r in range(m.member_count):
        for i, (d, b) in enumerate(m.slots):
            item = t.item(r, i)
            if (r, d, b) in m.duty_cells:
                assert item.text() == "值" and "已排值班" in item.toolTip()
                checked["duty"] += 1
            elif m.free[r][i]:
                assert item.text() == "" and "空闲" in item.toolTip()
                checked["free"] += 1
            else:
                assert item.text() == "课" and m.busy_courses.get((r, d, b))
                checked["busy"] += 1
    assert checked["duty"] > 0 and checked["free"] > 0 and checked["busy"] > 0

    # 汇总行
    row = m.member_count
    for i in range(len(m.slots)):
        assert t.item(row, i).text() == f"{m.free_counts[i]}/{m.member_count}"

    # 请假：整行标「假」并带上原因
    window.gantt_week.setValue(1)
    member_id = window.members()[0].id
    window.db.add_leave(member_id, 1, 2, "生病")
    window.invalidate_cache()
    window.refresh_gantt()
    m2 = window.gantt_matrix
    row_of = {mem.id: r for r, mem in enumerate(window.members())}
    r0 = row_of[member_id]
    for i, (d, _b) in enumerate(m2.slots):
        if d == 2:
            assert t.item(r0, i).text() == "假" and "生病" in t.item(r0, i).toolTip()


def test_gantt_renders_long_term_special_arrangement(window, qt_app) -> None:
    """长期特殊安排应在甘特图中显示为紫色「其他安排」并计入忙时。"""
    seed_members(window, 3)
    member_id = window.members()[0].id
    window.db.add_special_arrangement(
        member_id, 1, 3, 1, list(BLOCK_SESSIONS[1]), "固定实习")
    window.invalidate_cache()
    window.gantt_week.setValue(2)
    window.refresh_gantt()

    matrix = window.gantt_matrix
    assert matrix is not None
    row = next(r for r, m in enumerate(window.members()) if m.id == member_id)
    col = matrix.slots.index((1, 1))
    assert matrix.free[row][col] is False
    assert matrix.special_info[(row, 1, 1)] == ["固定实习"]
    item = window.gantt_table.item(row, col)
    assert item.text() == "特" and "其他安排" in item.toolTip()
    assert item.background().color().name() == "#e5d8ff"
