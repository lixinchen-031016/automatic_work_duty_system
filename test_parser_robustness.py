"""课表解析健壮性测试

基准数据是 `samples/desensitized/` 下的三份样例（真实教务系统导出后脱敏）：
它们是**真实文件的等长字节替换版本**——课程名/周次/节次/地点/课程数与原始
文件逐字段一致，只把姓名与学号换成了虚构值，因此既能放进仓库供 CI 使用，
又保留了真实布局（合并单元格、SST 共享、换行方式）带来的解析难度。

修复解析器时必须保证这三份课表的基线不变，再用变体用例覆盖各校常见写法
（单双周、表头写法、全角字符、空行异常、合并单元格）。

原始文件与脱敏样例的对应关系由 `tools/desensitize_samples.py` 维护；
本地有原始文件时可跑 `python tools/desensitize_samples.py --check` 校验一致性。
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from duty_system.parser import (
    parse_cell, parse_grid, parse_schedule_file,
    parse_schedule_path, parse_sessions, parse_weekday_header, parse_weeks,
)

SAMPLES_DIR = Path(__file__).parent / "samples" / "desensitized"
SAMPLE = SAMPLES_DIR / "学生个人课表_9999800598.xls"          # 原始 37 门，国际商务

# 样例课表的黄金基线：修复解析器时这些值必须保持不变
# （姓名/学号为脱敏后的虚构值，课程数据与真实文件完全一致）
BASELINE = {
    "name": "赵明明",
    "student_id": "9999800598",
    "term": "2026-2027-1",
    "class_name": "24国际商务双语4班",
    "major": "国际商务双语",
    "department": "管理工程学院",
    "course_count": 37,
}


# --------------------------------------------------------------------------- #
# 基线：真实样例课表
# --------------------------------------------------------------------------- #

def test_sample_schedule_baseline() -> None:
    """真实样例课表：元信息与课程数必须与基线一致"""
    s = parse_schedule_path(SAMPLE)
    assert s.name == BASELINE["name"]
    assert s.student_id == BASELINE["student_id"]
    assert s.term == BASELINE["term"]
    assert s.class_name == BASELINE["class_name"]
    assert s.major == BASELINE["major"]
    assert s.department == BASELINE["department"]
    assert len(s.courses) == BASELINE["course_count"], \
        f"课程数从 {BASELINE['course_count']} 变成 {len(s.courses)}"
    assert not s.warnings, f"不应有解析告警: {s.warnings}"


def test_sample_schedule_field_level_snapshot() -> None:
    """逐字段校验关键课程：课程名/教师/周次/节次/地点都不许漂移"""
    s = parse_schedule_path(SAMPLE)
    by_name = {}
    for c in s.courses:
        by_name.setdefault(c.course_name, []).append(c)

    # 教师名被换行打断的案例（原解析器的已知难点）
    sy = by_name.get("企业运营管理综合模拟实验", [])
    assert sy and all(c.teacher == "钱明明" for c in sy)
    assert all(c.location for c in sy), "该课程应解析出地点"

    # 单周周次（真实数据里用逗号列举，不是"单周"字样）
    pe = next(c for c in s.courses if c.course_name.startswith("体育Ⅲ"))
    assert pe.week_list == [1, 3, 5, 7, 9, 11, 13, 15]
    assert pe.teacher == "王明明"

    # 混合周次 1-4,6-10
    mds = next(c for c in s.courses if c.course_name.startswith("毛泽东"))
    assert mds.week_list == [1, 2, 3, 4, 6, 7, 8, 9, 10]

    # 无教师课程（原文件里确实没有教师行）
    gaoshu = [c for c in s.courses if c.course_name.startswith("高等数学")]
    assert gaoshu and all(c.teacher == "" for c in gaoshu)

    # 罗马数字与全角括号必须原样保留（不能被字符归一化改掉）
    names = set(by_name)
    assert any(n.startswith("体育Ⅲ") for n in names), f"体育Ⅲ 丢失: {names}"
    assert any(n.startswith("体育Ⅴ") for n in names), f"体育Ⅴ 丢失: {names}"
    assert "高等数学(Ⅱ)-1" in names, f"高等数学(Ⅱ)-1 丢失: {names}"
    assert any("（西肆三楼）" in c.location for c in s.courses), \
        "地点中的全角括号应原样保留"


def test_sample_schedule_is_stable_across_reparse() -> None:
    """同一文件重复解析结果完全一致（无字典序/随机性依赖）"""
    a = parse_schedule_path(SAMPLE)
    b = parse_schedule_path(SAMPLE)
    assert sorted(c.key() for c in a.courses) == sorted(c.key() for c in b.courses)
    assert [c.week_list for c in a.courses] == [c.week_list for c in b.courses]


def test_sample_week_lists_match_weeks_text() -> None:
    """每条记录的 week_list 都必须由 weeks_text 推导得来且非空"""
    s = parse_schedule_path(SAMPLE)
    for c in s.courses:
        assert c.week_list, f"{c.course_name} 周次为空"
        assert c.session_list, f"{c.course_name} 节次为空"
        assert c.weekday in range(1, 8)
        assert parse_weeks(c.weeks_text, c.week_mode) == c.week_list or \
            parse_weeks(c.weeks_text.replace("单周", "").replace("双周", ""), c.week_mode) == c.week_list


# --------------------------------------------------------------------------- #
# 周次与节次解析
# --------------------------------------------------------------------------- #

@pytest.mark.parametrize("text,expected", [
    ("18", [18]),
    ("1-6", [1, 2, 3, 4, 5, 6]),
    ("1,3,5,7", [1, 3, 5, 7]),
    ("1-4,6-10", [1, 2, 3, 4, 6, 7, 8, 9, 10]),
    ("１８", [18]),                 # 全角数字
    ("1，3，5", [1, 3, 5]),          # 全角逗号
    ("1、3、5", [1, 3, 5]),          # 顿号
    ("16-1", list(range(1, 17))),   # 区间反写（自动交换上下限）
])
def test_parse_weeks_variants(text: str, expected: list[int]) -> None:
    assert parse_weeks(text) == expected


@pytest.mark.parametrize("text,mode,expected", [
    ("1-16", "单周", [1, 3, 5, 7, 9, 11, 13, 15]),
    ("1-16", "双周", [2, 4, 6, 8, 10, 12, 14, 16]),
    ("1-16", "奇周", [1, 3, 5, 7, 9, 11, 13, 15]),
    ("1-16", "偶周", [2, 4, 6, 8, 10, 12, 14, 16]),
    ("1-16", "单", [1, 3, 5, 7, 9, 11, 13, 15]),
    ("1-16", "", list(range(1, 17))),
])
def test_parse_weeks_with_mode(text: str, mode: str, expected: list[int]) -> None:
    assert parse_weeks(text, mode) == expected


@pytest.mark.parametrize("text,expected", [
    ("01-02-03-04", [1, 2, 3, 4]),
    ("01,02", [1, 2]),
    ("01、02", [1, 2]),
    ("09-10", [9, 10]),
    ("０１-０２", [1, 2]),
])
def test_parse_sessions_variants(text: str, expected: list[int]) -> None:
    assert parse_sessions(text) == expected


@pytest.mark.parametrize("header,expected", [
    ("星期一", 1), ("星期日", 7), ("周一", 1), ("周1", 1), ("周三", 3),
    ("礼拜一", 1), ("礼拜天", 7), ("星期7", 7), ("星期天", 7),
    ("周一 ", 1), ("\u3000周二", 2), ("教材", 0), ("", 0),
])
def test_parse_weekday_header(header: str, expected: int) -> None:
    assert parse_weekday_header(header) == expected


# --------------------------------------------------------------------------- #
# 单元格解析：真实课表里常见的各种写法
# --------------------------------------------------------------------------- #

def cell(lines: list[str]) -> str:
    return "\n" + "\n".join(lines) + "\n"


def test_cell_week_only_and_even_are_kept() -> None:
    """单双周课程必须被解析出来（修复前会被整条丢弃）"""
    single = parse_cell(cell(["企业运营管理", "钱明明(教授)", "1-16单周([周])[01-02节]", "德五楼3412"]), 1)
    assert len(single) == 1
    c = single[0]
    assert c.course_name == "企业运营管理" and c.teacher == "钱明明"
    assert c.week_list == [1, 3, 5, 7, 9, 11, 13, 15]
    assert c.week_mode == "odd"
    assert "单周" in c.weeks_text, "界面展示的周次文本应体现单周"

    double = parse_cell(cell(["高等数学", "李四(讲师)", "1-16双周([周])[03-04节]", "教1"]), 2)
    assert double[0].week_list == [2, 4, 6, 8, 10, 12, 14, 16]
    assert double[0].week_mode == "even"


@pytest.mark.parametrize("week_token", [
    "1-16([周])[01-02节]",
    "1-16周([周])[01-02节]",
    "第1-16周([周])[01-02节]",
    "1-16[01-02节]",
    "１６([周])[０１-０２节]",
    "1-16([周])[01,02节]",
    "1-16([周])[01、02节]",
    "1-16（[周]）［01-02节］",
])
def test_cell_week_session_writing_variants(week_token: str) -> None:
    """各种周次/节次书写方式都应解析出「课程 + 周次 + 节次」"""
    courses = parse_cell(cell(["测试课程", "张三(讲师)", week_token, "教室A"]), 3)
    assert len(courses) == 1, f"{week_token} 解析失败"
    c = courses[0]
    assert c.course_name == "测试课程"
    assert c.teacher == "张三"
    assert c.session_list == [1, 2]
    assert c.week_list, "周次不应为空"
    assert c.location == "教室A"


def test_cell_fullwidth_teacher_parens() -> None:
    """教师职称用全角括号时也应能识别"""
    courses = parse_cell(cell(["测试课程", "张三（副教授）", "1-16([周])[01-02节]", "教室"]), 1)
    assert courses[0].teacher == "张三", f"实际 {courses[0].teacher!r}"
    assert "（副教授）" not in courses[0].course_name


def test_cell_multiple_courses_in_one_cell() -> None:
    """一个单元格含多门课程（样例课表里最多 4 门）"""
    text = cell([
        "课程甲", "教师甲(讲师)", "1-8([周])[01-02节]", "教室甲",
        "", "课程乙", "教师乙(教授)", "9-16单周([周])[01-02节]", "教室乙",
        "", "课程丙", "教师丙(助教)", "18([周])[01-02节]",
    ])
    courses = parse_cell(text, 5)
    names = {c.course_name for c in courses}
    assert names == {"课程甲", "课程乙", "课程丙"}, f"实际 {names}"
    by_name = {c.course_name: c for c in courses}
    assert by_name["课程乙"].week_list == [9, 11, 13, 15]
    assert by_name["课程丙"].teacher == "教师丙"
    assert by_name["课程丙"].location == ""


def test_cell_without_blank_line_between_courses() -> None:
    """课程之间没有空行时，也要按「周次行」正确切分，不能并成一门"""
    text = "\n课程甲\n教师甲(讲师)\n1-8([周])[01-02节]\n教室甲\n课程乙\n教师乙(教授)\n9-16([周])[01-02节]\n教室乙\n"
    courses = parse_cell(text, 4)
    names = [c.course_name for c in courses]
    assert "课程甲" in names and "课程乙" in names, f"实际 {names}"


def test_cell_name_split_by_stray_blank_line() -> None:
    """课程名与周次行之间被插入空行（样例中的「高等数学」就是这种结构）"""
    text = "\n高等数学(Ⅱ)-1\n\n5-16([周])[09-10节]\n零号教室021\n\n企业运营管理\n钱明明(教授)\n18([周])[09-10节]\n德五楼3412\n"
    courses = parse_cell(text, 3)
    by_name = {c.course_name: c for c in courses}
    assert "高等数学(Ⅱ)-1" in by_name, f"实际 {list(by_name)}"
    assert by_name["高等数学(Ⅱ)-1"].week_list == [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
    assert by_name["高等数学(Ⅱ)-1"].location == "零号教室021"
    assert by_name["企业运营管理"].teacher == "钱明明"


def test_cell_garbage_and_empty_cells() -> None:
    """空单元格、纯空白、无周次行的文本都不应抛出异常或产出脏数据"""
    assert parse_cell("", 1) == []
    assert parse_cell("\n\n\n", 1) == []
    assert parse_cell("只是说明文字", 1) == []
    assert parse_cell("\n课程名\n教师(讲师)\n", 1) == [], "缺周次行时不应产出课程"


# --------------------------------------------------------------------------- #
# 网格 / 文件层
# --------------------------------------------------------------------------- #

def make_grid(header: list[str]) -> list[list[str]]:
    return [
        ["成都工业学院 张三 学生个人课表"],
        ["学年学期：2026-2027-1 班级：24测试班 专业：国际商务 院系：管理工程学院"],
        [""] + header,
        ["1-2节(8:30-10:05)", "高等数学\n李四(教授)\n1-16([周])[01-02节]\n教1", "", "", "", ""],
        ["3-4节(10:20-11:55)", "", "英语\n王五(讲师)\n1-16单周([周])[03-04节]\n教2", "", "", ""],
    ]


@pytest.mark.parametrize("header", [
    ["周一", "周二", "周三", "周四", "周五"],
    ["星期一", "星期二", "星期三", "星期四", "星期五"],
    ["周1", "周2", "周3", "周4", "周5"],
    ["礼拜一", "礼拜二", "礼拜三", "礼拜四", "礼拜五"],
])
def test_grid_reads_all_header_styles(header: list[str]) -> None:
    """表头写法不同时整份课表都要能读出来（修复前「周一」写法读不出任何课程）"""
    res = parse_grid(make_grid(header), "学生个人课表_9999800598.xls")
    names = {c.course_name for c in res.courses}
    assert names == {"高等数学", "英语"}, f"实际 {names}"
    en = next(c for c in res.courses if c.course_name == "英语")
    assert en.weekday == 2 and en.week_list == [1, 3, 5, 7, 9, 11, 13, 15]
    assert res.name == "张三" and res.class_name == "24测试班"


def test_grid_header_row_ignores_body_mentions() -> None:
    """正文里出现「周一」等字样不应被误判成表头行"""
    grid = make_grid(["周一", "周二", "周三", "周四", "周五"])
    grid.append(["7-8节", "说明：周一至周五上课", "", "", "", ""])
    res = parse_grid(grid, "f.xls")
    assert len(res.courses) == 2, "表头行定位错误会丢课"


def test_grid_expands_merged_header_cells() -> None:
    """合并单元格：只有左上角有值时，应把值填充到整个合并区域

    真实 .xls 里标题行/备注行就是这种结构（merged_cells 为 (0,1,0,8)），
    不填充会漏读被合并覆盖的格子。
    """
    from duty_system.parser import _expand_merged

    grid = [
        ["成都工业学院 张三 学生个人课表", "", "", "", ""],
        ["1-2节", "课程甲\n教师甲\n1-4([周])[01-02节]", "", "", ""],
    ]
    merged = [(0, 1, 0, 5)]      # A1:E1 合并，只有 grid[0][0] 有值
    _expand_merged(grid, merged)
    assert all(grid[0][c] == grid[0][0] for c in range(5)), "合并区域应被填满"

    # 合并区域为空时不覆盖已有内容（避免把空值刷进别的格子）
    grid2 = [["", "保留", ""]]
    _expand_merged(grid2, [(0, 1, 0, 3)])
    assert grid2[0][1] == "保留"

    # 越界范围不应抛异常
    _expand_merged(grid2, [(0, 99, 0, 99)])


def test_parse_schedule_file_rejects_unknown_extension() -> None:
    with pytest.raises(ValueError, match="不支持的文件格式"):
        parse_schedule_file(b"x", "课表.csv")


def test_parse_schedule_file_roundtrip_bytes() -> None:
    """按字节流解析（上传路径）与按路径解析结果一致"""
    a = parse_schedule_path(SAMPLE)
    b = parse_schedule_file(SAMPLE.read_bytes(), SAMPLE.name)
    assert len(a.courses) == len(b.courses)
    assert sorted(c.key() for c in a.courses) == sorted(c.key() for c in b.courses)
    assert b.student_id == "9999800598"


# --------------------------------------------------------------------------- #
# 文件层：格式识别与容错
# --------------------------------------------------------------------------- #

def _make_xlsx_bytes(tmp_path: Path) -> bytes:
    """生成一份最小可用的 .xlsx 课表"""
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "成都工业学院 李四 学生个人课表"
    ws["A2"] = "学年学期：2026-2027-1 班级：24测试班 专业：国际商务 院系：管理工程学院"
    for i, day in enumerate(["星期一", "星期二", "星期三"], start=1):
        ws.cell(row=3, column=i + 1, value=day)
    ws["A4"] = "1-2节(8:30-10:05)"
    ws["B4"] = "高等数学\n王五(教授)\n1-16单周([周])[01-02节]\n教1"
    path = tmp_path / "tmp.xlsx"
    wb.save(path)
    return path.read_bytes()


def test_format_sniffing_ignores_extension(tmp_path: Path) -> None:
    """按内容而非扩展名识别格式：教务系统导出的文件经常被改名"""
    from duty_system.parser import _sniff_format

    xls_bytes = SAMPLE.read_bytes()
    xlsx_bytes = _make_xlsx_bytes(tmp_path)
    assert _sniff_format(xls_bytes) == "xls"
    assert _sniff_format(xlsx_bytes) == "xlsx"
    assert _sniff_format(b"garbage") == ""

    # xlsx 内容 + .xls 扩展名（修复前会直接抛 XLRDError，整份课表读不出来）
    r1 = parse_schedule_file(xlsx_bytes, "学生个人课表_2307724999.xls")
    assert [c.course_name for c in r1.courses] == ["高等数学"]
    assert r1.name == "李四"

    # xls 内容 + .xlsx 扩展名
    r2 = parse_schedule_file(xls_bytes, "学生个人课表_9999800598.xlsx")
    assert len(r2.courses) == BASELINE["course_count"]
    assert r2.student_id == "9999800598"


def test_corrupt_file_raises_clear_error(tmp_path: Path) -> None:
    """损坏/加密文件应给出明确错误，而不是底层库的怪异异常"""
    bad = tmp_path / "损坏.xls"
    bad.write_bytes(b"this is not an excel file")
    with pytest.raises(ValueError, match="无法读取该课表文件"):
        parse_schedule_file(bad.read_bytes(), bad.name)


def test_xls_fallback_when_formatting_info_unsupported(monkeypatch) -> None:
    """formatting_info 不被支持时应降级重试，而不是让整份课表读取失败"""
    from duty_system import parser as parser_mod

    calls: list[bool] = []
    real = parser_mod._read_xls_grid

    def flaky(file_bytes, with_formatting):
        calls.append(with_formatting)
        if with_formatting:
            raise NotImplementedError("formatting_info not supported")
        return real(file_bytes, with_formatting=False)

    monkeypatch.setattr(parser_mod, "_read_xls_grid", flaky)
    result = parser_mod._grid_from_xls(SAMPLE.read_bytes())
    assert calls == [True, False], "应先试带样式模式，失败后降级"
    assert result, "降级后仍应读出网格"


def test_merged_cells_do_not_break_sample(tmp_path: Path) -> None:
    """样例课表含 3 处合并区域，解析结果不应受影响"""
    import xlrd

    wb = xlrd.open_workbook(str(SAMPLE), formatting_info=True)
    assert wb.sheet_by_index(0).merged_cells, "该样例应存在合并区域"
    assert len(parse_schedule_path(SAMPLE).courses) == BASELINE["course_count"]


# --------------------------------------------------------------------------- #
# 真实样例二三：25材科 / 24电信（新格式：多教师、文件名带后缀）
# --------------------------------------------------------------------------- #

SAMPLE_2403 = SAMPLES_DIR / "学生个人课表_9999832478.xls"      # 原始 57 门，电信
SAMPLE_2502 = SAMPLES_DIR / "学生个人课表_9999453245(1).xls"   # 原始 49 门，材科（文件名带后缀）

REAL_SAMPLES = [
    pytest.param(SAMPLE, "赵明明", "9999800598", "24国际商务双语4班", 37, id="国际商务"),
    pytest.param(SAMPLE_2403, "钱明明", "9999832478", "24电信3班", 57, id="电信"),
    pytest.param(SAMPLE_2502, "郑明明", "9999453245", "25材科4班", 49, id="材科"),
]


def _require(path: Path) -> Path:
    if not path.exists():
        pytest.skip(f"缺少样例课表 {path.name}（先跑 tools/desensitize_samples.py 生成）")
    return path


@pytest.mark.parametrize("path,name,student_id,class_name,count", REAL_SAMPLES)
def test_real_samples_meta_and_count(path, name, student_id, class_name, count) -> None:
    """三份样例课表的元信息与课程数都必须稳定（学生姓名/学号/班级不得错位）"""
    s = parse_schedule_path(_require(path))
    assert s.name == name
    assert s.student_id == student_id, f"学号提取失败：{s.student_id!r}"
    assert s.class_name == class_name
    assert s.term == "2026-2027-1"
    assert len(s.courses) == count, f"课程数 {len(s.courses)} != {count}"
    assert not s.warnings


@pytest.mark.parametrize("path", [SAMPLE, SAMPLE_2403, SAMPLE_2502],
                         ids=["国际商务", "电信", "材科"])
def test_real_samples_note_row_agrees_with_parsed_courses(path: Path) -> None:
    """用教务系统自己生成的末尾「备注行」反查解析结果

    备注行形如 `：课程名 教师 周次周;课程名 教师,教师 周次周;`，
    是文件内自带的权威课程清单：解析出的课程名/教师/周次必须覆盖它，
    任何一条对不上都说明该课程的读取有偏差。
    """
    import xlrd

    path = _require(path)
    sheet = xlrd.open_workbook(str(path), formatting_info=True).sheet_by_index(0)
    from duty_system.parser import normalize_cell

    note = normalize_cell(sheet.cell_value(sheet.nrows - 1, 1)).lstrip("：:")
    assert "周" in note, "该样例应带备注行"

    parsed = parse_schedule_path(path)
    by_name: dict[str, dict[str, set]] = {}
    for c in parsed.courses:
        bucket = by_name.setdefault(c.course_name, {"weeks": set(), "teachers": set()})
        bucket["weeks"] |= set(c.week_list)
        if c.teacher:
            bucket["teachers"] |= {t for t in re.split(r"[,，、]", c.teacher) if t}

    for item in note.split(";"):
        item = item.strip()
        if not item:
            continue
        parts = item.split()
        assert len(parts) >= 3, f"备注行格式异常：{item!r}"
        course, teachers, weeks_token = " ".join(parts[:-2]), parts[-2], parts[-1]
        key = course if course in by_name else next(
            (k for k in by_name if k.startswith(course) or course.startswith(k)), None)
        assert key is not None, f"备注行里的课程未解析出来：{course!r}"

        want_teachers = {t for t in re.split(r"[,，、]", teachers) if t}
        want_weeks = set(parse_weeks(weeks_token.rstrip("周")))
        assert want_weeks <= by_name[key]["weeks"], (
            f"{course} 周次缺失：备注 {sorted(want_weeks)} vs 解析 {sorted(by_name[key]['weeks'])}")
        assert want_teachers <= by_name[key]["teachers"], (
            f"{course} 教师缺失：备注 {sorted(want_teachers)} vs 解析 {sorted(by_name[key]['teachers'])}")


@pytest.mark.parametrize("week_token", ["2,4,6,8,10,12,14,16", "1,3,7,9,11,13,15"])
def test_real_sample_nonuniform_week_lists(week_token: str) -> None:
    """真实课表里体育课的周次是不规则列举（如 1,3,7,9），不能当成等差数列处理"""
    assert parse_weeks(week_token) == [int(x) for x in week_token.split(",")]


def test_real_sample_fullwidth_course_parens_preserved() -> None:
    """课程名里的全角括号/罗马数字必须原样保留（不能被字符归一化改写）"""
    s = parse_schedule_path(_require(SAMPLE_2502))
    names = {c.course_name for c in s.courses}
    assert "材料科学与工程基础（I）" in names, f"全角括号课程名丢失：{sorted(names)}"
    assert "高等数学(Ⅰ)-1" in names, f"罗马数字课程名丢失：{sorted(names)}"


def test_real_sample_shared_room_multiple_experiments() -> None:
    """同一门课不同周次在不同地点（大学物理实验-2）不能被去重合并"""
    s = parse_schedule_path(_require(SAMPLE_2502))
    exps = [c for c in s.courses if c.course_name == "大学物理实验-2"]
    weeks_to_room = {tuple(c.week_list): c.location for c in exps}
    assert len({c.location for c in exps}) >= 4, f"实验地点被合并：{weeks_to_room}"
    for c in exps:
        assert len(c.week_list) == 1, f"单周实验被合并：{c.week_list}"


# --------------------------------------------------------------------------- #
# 教师字段：多人授课与括号被换行打断
# --------------------------------------------------------------------------- #

def test_cell_multiple_teachers_on_one_line() -> None:
    """多人授课（郑明(副教授),周明明(讲师)）不能把第二位教师并进课程名"""
    courses = parse_cell(cell([
        "电工电子技术B", "郑明(副教授),周明明(讲师)",
        "7-11([周])[03-04节]", "甲工楼A座6205",
    ]), 2)
    assert len(courses) == 1
    c = courses[0]
    assert c.course_name == "电工电子技术B", f"课程名被教师名污染：{c.course_name!r}"
    assert c.teacher == "郑明,周明明", f"教师解析错误：{c.teacher!r}"


@pytest.mark.parametrize("sep", [",", "，", "、", " ,", ", "])
def test_cell_multiple_teachers_separator_variants(sep: str) -> None:
    """多人授课的分隔符（半角/全角逗号、顿号）都应支持"""
    courses = parse_cell(cell([
        "测试课程", f"张三(讲师){sep}李四(教授)",
        "1-8([周])[01-02节]", "教室",
    ]), 1)
    c = courses[0]
    assert c.course_name == "测试课程", f"实际 {c.course_name!r}"
    assert c.teacher == "张三,李四", f"实际 {c.teacher!r}"


def test_cell_teacher_parens_split_by_newline() -> None:
    """教师职称的括号被换行打断（钱明明\\n(教授)）也要能识别

    模块文档声明该情形由「合并后匹配」覆盖，但合并会引入空格，
    若正则不容忍空格则教师名会被并进课程名。
    """
    courses = parse_cell("\n测试课程\n钱明明\n(教授)\n1-16([周])[01-02节]\n教室A\n", 1)
    assert len(courses) == 1
    c = courses[0]
    assert c.course_name == "测试课程", f"实际 {c.course_name!r}"
    assert c.teacher == "钱明明", f"实际 {c.teacher!r}"


def test_cell_course_subtitle_is_not_treated_as_teacher() -> None:
    """课程副标题 (PD8-1) 不是职称，必须留在课程名里"""
    courses = parse_cell(cell([
        "毛泽东思想和中国特色社会主义理论体系概论 (PD8-1)",
        "吴明明(副教授)", "1-10([周])[05-06节]", "允明楼1424",
    ]), 2)
    c = courses[0]
    assert c.course_name == "毛泽东思想和中国特色社会主义理论体系概论 (PD8-1)"
    assert c.teacher == "吴明明"


# --------------------------------------------------------------------------- #
# 学号提取：文件名带后缀括号
# --------------------------------------------------------------------------- #

@pytest.mark.parametrize("file_name,expected", [
    ("学生个人课表_9999800598.xls", "9999800598"),
    ("学生个人课表_9999800598.xlsx", "9999800598"),
    ("学生个人课表_9999453245(1).xls", "9999453245"),
    ("学生个人课表_9999453245(1)(2).xls", "9999453245"),
    ("课表_9999453245（1）.xls", "9999453245"),
    ("课表_9999453245 .xls", "9999453245"),
    ("没有学号.xls", ""),
])
def test_student_id_extraction_from_file_name(file_name: str, expected: str) -> None:
    """学号取文件名中的数字；带 (1) 之类后缀时也不能丢"""
    grid = [["成都工业学院 张三 学生个人课表"]]
    assert parse_grid(grid, file_name).student_id == expected
