"""花名册管理页：数据库增删改查、匹配查找/筛选，以及界面交互。

界面用例在 QT_QPA_PLATFORM=offscreen 下运行，模态框全部被替换，
因此可在无显示器环境完整验证一遍流程。
"""

from __future__ import annotations

from pathlib import Path

import pytest

from duty_system.database import Database
from duty_system.parser import Course, ParsedSchedule
from duty_system.roster import (
    UNKNOWN_STUDIO,
    RosterEntry,
    filter_roster_entries,
    match_roster_entry,
    roster_leader_conflicts,
    validate_roster_entries,
)

ENTRY_ID_ROLE = 0x0100  # Qt.UserRole


def make_db(tmp_path: Path) -> Database:
    return Database(tmp_path / "roster_editor.db")


def seed_roster(db: Database) -> list[RosterEntry]:
    """三条记录：一名部长、一名副部长、一条缺学号的重名记录。"""
    db.add_roster_entry(RosterEntry(
        studio="微信工作室", position="部长", name="张三", student_id="20250001",
        phone="13800000000", college_major="计算机学院软件工程"))
    db.add_roster_entry(RosterEntry(
        studio="设计工作室", position="副部长", name="李四", student_id="20250002",
        phone="13900000000", college_major="设计学院视觉传达"))
    db.add_roster_entry(RosterEntry(
        studio="设计工作室", name="李四", college_major="设计学院视觉传达"))
    return db.list_roster_entries()


def member_schedule(name: str, student_id: str) -> ParsedSchedule:
    return ParsedSchedule(
        name=name, student_id=student_id, term="2026-2027-1", class_name="测试班",
        major="专业", department="学院", file_name=f"{student_id}.xls",
        courses=[Course(
            course_name="课程", teacher="教师", weekday=1, weeks_text="1-18",
            week_list=list(range(1, 19)), sessions_text="01-02",
            session_list=[1, 2])])


# --------------------------------------------------------------------------- #
# 数据库层：增删改查
# --------------------------------------------------------------------------- #

def test_roster_entry_crud_normalizes_and_keeps_ids(tmp_path: Path) -> None:
    db = make_db(tmp_path)
    first_id = db.add_roster_entry(RosterEntry(
        studio="  微信工作室 ", position=" 部长 ", name=" 张三 ",
        student_id="20250001.0", phone=" 13800000000 "))
    rows = db.list_roster_entries()
    assert [row.entry_id for row in rows] == [first_id]
    assert rows[0].studio == "微信工作室"
    assert rows[0].name == "张三"
    assert rows[0].student_id == "20250001"
    assert rows[0].phone == "13800000000"

    second_id = db.add_roster_entry(RosterEntry(studio="", name="李四"))
    assert db.list_roster_entries()[1].studio == UNKNOWN_STUDIO

    db.update_roster_entry(second_id, RosterEntry(
        studio="设计工作室", position="副部长", name="李四", student_id="20250002"))
    updated = db.list_roster_entries()[1]
    assert (updated.entry_id, updated.studio, updated.position) == (
        second_id, "设计工作室", "副部长")
    assert updated.student_id == "20250002"

    with pytest.raises(ValueError, match="不存在"):
        db.update_roster_entry(9999, RosterEntry(studio="设计工作室", name="李四"))

    assert db.delete_roster_entries([]) == 0
    assert db.delete_roster_entries([first_id, 9999]) == 1
    assert [row.name for row in db.list_roster_entries()] == ["李四"]


def test_restore_roster_entries_keeps_order_and_ids(tmp_path: Path) -> None:
    db = make_db(tmp_path)
    rows = seed_roster(db)
    reversed_rows = list(reversed(rows))
    db.restore_roster_entries(reversed_rows)
    # 整体写回保留原 id（列表按 id 升序读取，因此顺序不变）
    assert [row.entry_id for row in db.list_roster_entries()] == [
        row.entry_id for row in rows]

    db.restore_roster_entries([
        *db.list_roster_entries(),
        RosterEntry(studio="运营工作室", position="高级顾问", name="王五"),
    ])
    restored = db.list_roster_entries()
    assert [row.name for row in restored] == ["张三", "李四", "李四", "王五"]
    assert restored[-1].entry_id > max(row.entry_id for row in rows)


def test_roster_edit_syncs_unlocked_member_studio_and_phone(tmp_path: Path) -> None:
    db = make_db(tmp_path)
    member_id = db.upsert_member(member_schedule("张三", "20250001"))
    assert db.get_member(member_id).studios == [UNKNOWN_STUDIO]

    db.add_roster_entry(RosterEntry(
        studio="微信工作室", position="部长", name="张三",
        student_id="20250001", phone="13800000000"))
    member = db.get_member(member_id)
    assert member.studios == ["微信工作室"]
    assert member.studio_positions == {"微信工作室": "部长"}
    assert member.phone == "13800000000"

    entry_id = db.list_roster_entries()[0].entry_id
    db.update_roster_entry(entry_id, RosterEntry(
        studio="设计工作室", position="成员", name="张三",
        student_id="20250001", phone="13800000000"))
    assert db.get_member(member_id).studios == ["设计工作室"]
    db.delete_roster_entries([entry_id])
    assert db.get_member(member_id).studios == [UNKNOWN_STUDIO]
    assert db.get_member(member_id).phone == "13800000000"


def test_roster_edit_keeps_manually_locked_member_studio(tmp_path: Path) -> None:
    db = make_db(tmp_path)
    member_id = db.upsert_member(member_schedule("张三", "20250001"))
    db.set_member_studios(member_id, ["博Q工作室"], {"博Q工作室": "成员"})

    db.add_roster_entry(RosterEntry(
        studio="微信工作室", position="部长", name="张三", student_id="20250001"))
    assert db.get_member(member_id).studios == ["博Q工作室"]


# --------------------------------------------------------------------------- #
# 匹配查找 / 筛选 / 校验
# --------------------------------------------------------------------------- #

def test_match_roster_entry_scopes_modes_and_normalization() -> None:
    entry = RosterEntry(
        studio="微信工作室", position="部长", name="张三", student_id="20250001",
        phone="13800000000", college_major="计算机学院软件工程")

    assert match_roster_entry(entry, "２０２５ ０００１")
    assert match_roster_entry(entry, "2025", scope="student_id", mode="prefix")
    assert not match_roster_entry(entry, "2025", scope="student_id", mode="exact")
    assert match_roster_entry(entry, "20250001", scope="student_id", mode="exact")
    assert not match_roster_entry(entry, "张三", scope="student_id")
    assert match_roster_entry(entry, "张三", scope="name")
    assert match_roster_entry(entry, "软件工程", scope="college_major")
    assert match_roster_entry(entry, "   ")


def test_filter_roster_entries_by_filters_and_status(tmp_path: Path) -> None:
    db = make_db(tmp_path)
    rows = seed_roster(db)

    assert len(filter_roster_entries(rows, studio="设计工作室")) == 2
    assert len(filter_roster_entries(
        rows, studio="设计工作室", position="副部长")) == 1
    assert [row.name for row in filter_roster_entries(
        rows, status="missing_student_id")] == ["李四"]
    assert [row.name for row in filter_roster_entries(
        rows, status="missing_phone")] == ["李四"]
    duplicates = filter_roster_entries(rows, status="duplicate")
    assert len(duplicates) == 2
    assert {row.studio for row in duplicates} == {"设计工作室"}
    assert filter_roster_entries(rows, status="leader_conflict") == []

    rows.append(RosterEntry(studio="微信工作室", position="部长", name="王五"))
    assert len(filter_roster_entries(rows, status="leader_conflict")) == 2
    assert len(filter_roster_entries(
        rows, query="李四", mode="exact", scope="name")) == 2


def test_validate_roster_entries_reports_problems() -> None:
    clean = [RosterEntry(studio="微信工作室", position="部长", name="张三")]
    assert validate_roster_entries(clean) == []
    assert roster_leader_conflicts(clean) == {}

    text = "\n".join(validate_roster_entries([
        *clean,
        RosterEntry(studio="微信工作室", position="部长", name="张三"),
        RosterEntry(studio="微信工作室", name=""),
    ]))
    assert "姓名为空" in text
    assert "微信工作室 指定了 2 名部长" in text
    assert "同工作室重名" in text


# --------------------------------------------------------------------------- #
# 界面：匹配查找、筛选、增删改、撤销重做
# --------------------------------------------------------------------------- #

@pytest.fixture(autouse=True)
def no_modal_dialogs(monkeypatch):
    """离屏运行时不允许弹出真实模态框。"""
    pytest.importorskip("PySide6.QtWidgets")
    from PySide6.QtWidgets import QMessageBox

    monkeypatch.setattr(QMessageBox, "information",
                        staticmethod(lambda *a, **k: None))
    monkeypatch.setattr(QMessageBox, "warning",
                        staticmethod(lambda *a, **k: None))
    monkeypatch.setattr(QMessageBox, "question",
                        staticmethod(lambda *a, **k: QMessageBox.Yes))


def make_editor(tmp_path: Path, **kwargs):
    from duty_system.roster_editor import RosterEditorDialog

    db = make_db(tmp_path)
    seed_roster(db)
    return db, RosterEditorDialog(db=db, **kwargs)


def test_editor_lists_roster_with_problem_marks(qt_app, tmp_path: Path) -> None:
    _db, dialog = make_editor(tmp_path)
    assert dialog.table.rowCount() == 3
    assert dialog.table.columnCount() == 7
    assert dialog.table.item(1, 0).text().startswith("⚠")
    assert "同工作室重名" in dialog.status_label.text()
    assert dialog.table.item(0, 0).text() == "微信工作室"


def test_editor_matching_search_highlights_and_filters(qt_app, tmp_path: Path) -> None:
    _db, dialog = make_editor(tmp_path)

    dialog.search_edit.setText("２０２５ ０００２")
    assert [entry.name for entry in dialog._visible] == ["李四"]
    assert len(dialog._visible) == 1
    # 命中字段用加粗 + 蓝色前景强调（主题 QSS 会吞掉模型背景色）
    assert dialog.table.item(0, 3).foreground().color().name() == "#1d4ed8"
    assert dialog.table.item(0, 3).font().bold() is True
    assert dialog.table.item(0, 2).font().bold() is False
    assert dialog.table.item(0, 3).text() == "20250002"

    dialog.search_edit.setText("李四")
    dialog.mode_combo.setCurrentIndex(dialog.mode_combo.findData("exact"))
    assert len(dialog._visible) == 2
    dialog.scope_combo.setCurrentIndex(dialog.scope_combo.findData("student_id"))
    assert dialog._visible == []

    dialog.clear_filters()
    assert len(dialog._visible) == 3
    dialog.studio_filter.setCurrentIndex(
        dialog.studio_filter.findData("设计工作室"))
    assert len(dialog._visible) == 2
    dialog.clear_filters()
    dialog.status_filter.setCurrentIndex(
        dialog.status_filter.findData("missing_student_id"))
    assert [entry.name for entry in dialog._visible] == ["李四"]

    dialog.clear_filters()
    dialog.search_edit.setText("张三")
    dialog.focus_first_match()
    assert dialog._selected_entries()[0].name == "张三"


def test_editor_delete_key_is_scoped_to_table(qt_app, tmp_path: Path) -> None:
    """Delete 键在表格里删除记录，在搜索框里只做文本编辑。"""
    from PySide6.QtCore import Qt
    from PySide6.QtTest import QTest

    _db, dialog = make_editor(tmp_path)
    dialog.show()
    qt_app.processEvents()

    dialog.search_edit.setFocus()
    dialog.search_edit.setText("李四")
    dialog.search_edit.setCursorPosition(0)
    QTest.keyClick(dialog.search_edit, Qt.Key_Delete)
    qt_app.processEvents()
    assert dialog.search_edit.text() == "四"
    assert len(dialog.entries) == 3

    dialog.clear_filters()
    dialog.table.setFocus()
    dialog.table.selectRow(0)
    QTest.keyClick(dialog.table, Qt.Key_Delete)
    qt_app.processEvents()
    assert len(dialog.entries) == 2


def test_editor_resets_studio_filter_when_option_disappears(
    qt_app, tmp_path: Path,
) -> None:
    """筛选项被删光后自动回到「全部」，避免下拉与表格内容不一致。"""
    _db, dialog = make_editor(tmp_path)
    dialog.studio_filter.setCurrentIndex(
        dialog.studio_filter.findData("设计工作室"))
    assert len(dialog._visible) == 2

    dialog.table.selectAll()
    assert len(dialog._selected_entries()) == 2
    dialog.delete_selected()

    assert dialog.studio_filter.currentData() == ""
    assert len(dialog.entries) == 1
    assert len(dialog._visible) == 1


def test_editor_add_edit_delete_with_undo_redo(
    qt_app, tmp_path: Path, monkeypatch,
) -> None:
    from PySide6.QtWidgets import QDialog

    from duty_system import roster_editor

    messages: list[str] = []
    db, dialog = make_editor(tmp_path, on_changed=messages.append)

    def fill_form(studio: str, name: str, student_id: str, position: str):
        def fake_exec(self) -> int:
            self.studio_combo.setCurrentText(studio)
            self._sync_positions(studio)
            self.position_combo.setCurrentIndex(
                max(self.position_combo.findData(position), 0))
            self.name_edit.setText(name)
            self.student_id_edit.setText(student_id)
            self.phone_edit.setText("13900000000")
            self.college_major_edit.setText("管理学院")
            return QDialog.Accepted
        return fake_exec

    monkeypatch.setattr(
        roster_editor.RosterEntryDialog, "exec",
        fill_form("运营工作室", "王五", "20250009", "高级顾问"))
    dialog.add_entry()
    assert [entry.name for entry in dialog.entries] == [
        "张三", "李四", "李四", "王五"]
    added = dialog.entries[-1]
    assert (added.studio, added.position, added.student_id) == (
        "运营工作室", "高级顾问", "20250009")
    assert messages[-1].startswith("已新增花名册记录：王五")

    target_id = added.entry_id
    for row in range(dialog.table.rowCount()):
        if int(dialog.table.item(row, 0).data(ENTRY_ID_ROLE)) == target_id:
            dialog.table.selectRow(row)
            break
    assert [entry.name for entry in dialog._selected_entries()] == ["王五"]

    monkeypatch.setattr(
        roster_editor.RosterEntryDialog, "exec",
        fill_form("设计工作室", "王五", "20250009", "成员"))
    dialog.edit_selected()
    edited = next(entry for entry in dialog.entries if entry.entry_id == target_id)
    assert (edited.studio, edited.position) == ("设计工作室", "成员")
    assert "已更新花名册记录：王五" in messages[-1]

    dialog.delete_selected()
    assert len(dialog.entries) == 3
    assert "已删除 1 条花名册记录" in messages[-1]
    dialog.undo_change()
    assert len(dialog.entries) == 4
    assert any(entry.entry_id == target_id for entry in dialog.entries)
    dialog.redo_change()
    assert len(dialog.entries) == 3

    dialog.table.selectRow(0)
    original = dialog._selected_entries()[0]
    dialog.duplicate_selected()
    copies = [
        entry for entry in dialog.entries
        if entry.name == original.name and entry.entry_id != original.entry_id
    ]
    assert copies and copies[0].source_file == ""
    assert len(dialog.entries) == 4

    assert [entry.name for entry in db.list_roster_entries()] == [
        entry.name for entry in dialog.entries]


def test_editor_blocks_empty_name_and_wires_import_export(
    qt_app, tmp_path: Path, monkeypatch,
) -> None:
    from PySide6.QtWidgets import QDialog, QMessageBox

    from duty_system import roster_editor

    calls: list[str] = []
    db = make_db(tmp_path)
    seed_roster(db)
    imported: list = []
    dialog = roster_editor.RosterEditorDialog(
        db=db,
        on_changed=calls.append,
        import_callback=lambda on_done: (imported.append(on_done), on_done()),
        export_callback=lambda: calls.append("导出"),
    )

    monkeypatch.setattr(QMessageBox, "warning",
                        staticmethod(lambda *a, **k: None))
    assert not dialog._confirm_proposed([
        *dialog.entries, RosterEntry(studio="微信工作室", name="")])

    assert dialog.btn_import.isEnabled() and dialog.btn_export.isEnabled()
    monkeypatch.setattr(roster_editor.RosterEntryDialog, "exec",
                        lambda self: QDialog.Rejected)
    dialog.request_import()
    assert imported and callable(imported[0])
    db.add_roster_entry(RosterEntry(studio="微信工作室", name="新同学"))
    imported[0]()
    assert len(dialog.entries) == 4
    assert dialog._undo_stack == []
    dialog.request_export()
    assert calls[-1] == "导出"


# --------------------------------------------------------------------------- #
# 主窗口集成：文件菜单入口 + 缓存刷新
# --------------------------------------------------------------------------- #

def test_file_menu_exposes_roster_editor_action(
    qt_app, tmp_path: Path, monkeypatch,
) -> None:
    import app as appmod

    window = appmod.MainWindow(db_path=tmp_path / "ui.db")
    try:
        file_menu = window.menuBar().actions()[0].menu()
        labels = [action.text() for action in file_menu.actions()]
        assert "花名册管理…" in labels
        assert labels.index("花名册管理…") > labels.index("导出完整花名册…")

        opened: dict = {}

        class FakeDialog:
            def __init__(self, parent=None, **kwargs) -> None:
                opened.update(kwargs)
                opened["parent"] = parent

            def exec(self) -> int:
                opened["exec"] = True
                return 0

        monkeypatch.setattr(appmod, "RosterEditorDialog", FakeDialog)
        window.action_manage_roster.trigger()
        assert opened["exec"] is True
        assert opened["db"] is window.db
        assert callable(opened["on_changed"])
        assert callable(opened["import_callback"])
        assert callable(opened["export_callback"])
        assert "花名册管理已关闭" in window.statusBar().currentMessage()
    finally:
        window.close()
        window.deleteLater()
        qt_app.processEvents()


def test_editor_change_refreshes_main_window_members(
    qt_app, tmp_path: Path, monkeypatch,
) -> None:
    from PySide6.QtWidgets import QDialog

    import app as appmod
    from duty_system import roster_editor

    window = appmod.MainWindow(db_path=tmp_path / "ui.db")
    try:
        window.db.upsert_member(member_schedule("张三", "20250001"))
        window.invalidate_cache()
        window.refresh_members()
        assert window.members()[0].studios == [UNKNOWN_STUDIO]

        dialog = roster_editor.RosterEditorDialog(
            parent=window, db=window.db,
            on_changed=window._on_roster_editor_changed,
            import_callback=window.import_roster,
            export_callback=window.export_roster,
        )

        def fake_exec(self) -> int:
            self.studio_combo.setCurrentText("微信工作室")
            self._sync_positions("微信工作室")
            self.position_combo.setCurrentIndex(
                max(self.position_combo.findData("部长"), 0))
            self.name_edit.setText("张三")
            self.student_id_edit.setText("20250001")
            self.phone_edit.setText("13800000000")
            return QDialog.Accepted

        monkeypatch.setattr(roster_editor.RosterEntryDialog, "exec", fake_exec)
        dialog.add_entry()

        # 模态期间只失效缓存并打标记，避免反复重建整个主界面
        assert window._roster_changed is True
        assert "已新增花名册记录：张三" in dialog.status_label.text()
        member = window.members()[0]
        assert member.studios == ["微信工作室"]
        assert member.phone == "13800000000"
    finally:
        window.close()
        window.deleteLater()
        qt_app.processEvents()


def test_open_roster_editor_from_menu_refreshes_after_close(
    qt_app, tmp_path: Path, monkeypatch,
) -> None:
    """完整走一遍「文件 → 花名册管理…」：新增记录 → 关闭页面 → 主界面刷新。"""
    from PySide6.QtWidgets import QDialog

    import app as appmod
    from duty_system import roster_editor

    window = appmod.MainWindow(db_path=tmp_path / "ui.db")
    opened: dict = {}
    try:
        window.db.upsert_member(member_schedule("张三", "20250001"))
        window.invalidate_cache()
        window.refresh_members()
        assert window.members()[0].studios == [UNKNOWN_STUDIO]

        def fake_entry_exec(self) -> int:
            self.studio_combo.setCurrentText("微信工作室")
            self._sync_positions("微信工作室")
            self.position_combo.setCurrentIndex(
                max(self.position_combo.findData("部长"), 0))
            self.name_edit.setText("张三")
            self.student_id_edit.setText("20250001")
            return QDialog.Accepted

        def fake_editor_exec(self) -> int:
            opened["dialog"] = self
            self.add_entry()
            return QDialog.Accepted

        monkeypatch.setattr(roster_editor.RosterEntryDialog, "exec", fake_entry_exec)
        monkeypatch.setattr(roster_editor.RosterEditorDialog, "exec", fake_editor_exec)

        window.action_manage_roster.trigger()

        assert opened["dialog"].entries[0].name == "张三"
        assert window.members()[0].studios == ["微信工作室"]
        assert "已新增花名册记录：张三" in window.statusBar().currentMessage()
        assert "1 条花名册记录" in window.statusBar().currentMessage()
    finally:
        window.close()
        window.deleteLater()
        qt_app.processEvents()


def test_closing_roster_editor_without_change_keeps_ui(
    qt_app, tmp_path: Path, monkeypatch,
) -> None:
    import app as appmod

    window = appmod.MainWindow(db_path=tmp_path / "ui.db")
    try:
        window.db.add_roster_entry(RosterEntry(
            studio="微信工作室", position="部长", name="张三",
            student_id="20250001"))
        window.invalidate_cache()
        window.refresh_members()

        def fake_exec(self) -> int:  # 只打开又关闭，不做任何写入
            assert self.table.rowCount() == 1
            return 0

        monkeypatch.setattr(appmod.RosterEditorDialog, "exec", fake_exec)
        window.open_roster_editor()
        assert window._roster_changed is False
        assert "花名册未修改" in window.statusBar().currentMessage()
    finally:
        window.close()
        window.deleteLater()
        qt_app.processEvents()


def test_import_roster_ignores_qt_bool_argument(
    qt_app, tmp_path: Path, monkeypatch,
) -> None:
    """Qt 的 triggered / clicked 会带 bool 参数，不能被当成"导入完成回调"。"""
    from PySide6.QtWidgets import QFileDialog

    import app as appmod

    window = appmod.MainWindow(db_path=tmp_path / "ui.db")
    try:
        monkeypatch.setattr(QFileDialog, "getOpenFileName",
                            staticmethod(lambda *a, **k: ("", "")))
        done: list[bool] = []

        window.import_roster(False)  # 模拟 Qt 传入的 checked 参数
        assert done == []
        window.action_import_roster.trigger()
        window.btn_import_roster.click()
        assert window.roster_entries() == []

        window.import_roster(lambda: done.append(True))
        assert done == [True]
    finally:
        window.close()
        window.deleteLater()
        qt_app.processEvents()
