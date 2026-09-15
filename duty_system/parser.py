"""课表文件解析模块

支持解析教务系统导出的"学生个人课表"(.xls / .xlsx)文件。

文件结构（以成都工业学院课表为例）：
    行0: "成都工业学院 张三 学生个人课表"          -> 标题（含学校、姓名）
    行1: "学年学期：2026-2027-1 班级：xx 专业：xx ..." -> 元信息
    行2: 星期一 ~ 星期日                            -> 表头
    行3-7: 节次行（1-2节/3-4节/5-6节/7-8节/9-11节） -> 每列单元格内含课程
    末行: 备注说明

单元格内多门课程以空行分隔，每门课程为多行文本：
    课程名（可能跨行，含 (板块x) 副标题）
    教师(职称)          <- 可能缺失，职称括号可能被换行打断
    周次([周])[节次节]  <- 如 18([周])[01-02-03-04节]
    地点                <- 可能缺失（如无地点或网课编号）
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from pathlib import Path

WEEKDAY_MAP = {
    "星期一": 1, "星期二": 2, "星期三": 3, "星期四": 4,
    "星期五": 5, "星期六": 6, "星期日": 7, "周日": 7,
}
WEEKDAY_LABELS = {1: "周一", 2: "周二", 3: "周三", 4: "周四", 5: "周五", 6: "周六", 7: "周日"}

# 值班时段块：与课表节次行对应
BLOCK_SESSIONS = {1: (1, 2), 2: (3, 4), 3: (5, 6), 4: (7, 8), 5: (9, 10, 11)}
BLOCK_LABELS = {
    1: "1-2节 08:30-10:05",
    2: "3-4节 10:20-11:55",
    3: "5-6节 14:00-15:35",
    4: "7-8节 15:50-17:25",
    5: "9-11节 18:50-21:15",
}

# 周次节次行，如 "18([周])[01-02-03-04节]"、"1,3,5([周])[03-04节]"
WEEK_SESSION_RE = re.compile(r"^([\d,\-]+)\s*(?:\(\[周\]\)\s*)?\[([\d\-]+)节\]$")

# 教师行，如 "唐心智(教授)"；职称括号可能被换行打断，合并后再匹配
TEACHER_RE = re.compile(
    r"(?:^|\s)([\u4e00-\u9fa5·]{2,4})"
    r"\((?:教授|副教授|讲师|助教|未评级|研究员|高级工程师|工程师|实验师|高级实验师|外教)\)"
)

# 节次行（第0列），如 "1-2节(8:30-10:05)\n(01,02)\n08:30-10:05"
PERIOD_ROW_RE = re.compile(r"^(\d+)\s*-\s*(\d+)节")

TITLE_KEYWORD = "学生个人课表"

META_PATTERNS = {
    "term": re.compile(r"学年学期[:：]\s*(\S+)"),
    "class_name": re.compile(r"班级[:：]\s*(\S+)"),
    "major": re.compile(r"专业[:：]\s*(\S+)"),
    "department": re.compile(r"院系[:：]\s*(\S+)"),
}

STUDENT_ID_FILE_RE = re.compile(r"_(\d{6,})\.(?:xls|xlsx)$", re.IGNORECASE)


@dataclass
class Course:
    """一门课程的上课安排（周次 x 节次展开前的原始记录）"""
    course_name: str
    teacher: str = ""
    weekday: int = 0
    weeks_text: str = ""
    week_list: list[int] = field(default_factory=list)
    sessions_text: str = ""
    session_list: list[int] = field(default_factory=list)
    location: str = ""

    def key(self) -> tuple:
        return (self.course_name, self.teacher, self.weekday,
                self.weeks_text, self.sessions_text, self.location)


@dataclass
class ParsedSchedule:
    """一份课表文件的解析结果"""
    name: str = "未知成员"
    student_id: str = ""
    term: str = ""
    class_name: str = ""
    major: str = ""
    department: str = ""
    file_name: str = ""
    courses: list[Course] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def parse_weeks(text: str) -> list[int]:
    """解析周次文本：'18' / '1-6' / '1,3,5,7' / '1-4,6-10' -> 周次列表"""
    weeks: set[int] = set()
    for part in text.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, _, b = part.partition("-")
            try:
                lo, hi = int(a), int(b)
            except ValueError:
                continue
            weeks.update(range(lo, hi + 1))
        elif part.isdigit():
            weeks.add(int(part))
    return sorted(weeks)


def parse_sessions(text: str) -> list[int]:
    """解析节次文本：'01-02-03-04' -> [1, 2, 3, 4]"""
    return sorted({int(x) for x in text.split("-") if x.strip().isdigit()})


def _split_name_teacher(name_part: str) -> tuple[str, str]:
    """从合并后的名称文本中拆出课程名与教师。

    name_part 形如 '思想政治理论课实践教学 官芯如(副教授)'，
    教师匹配取最后一个（课程名副标题如 (PD8-1) 不含职称关键字，不会误匹配）。
    """
    teacher = ""
    matches = list(TEACHER_RE.finditer(name_part))
    if matches:
        m = matches[-1]
        teacher = m.group(1)
        course_name = (name_part[: m.start()] + " " + name_part[m.end():])
        course_name = re.sub(r"\s+", " ", course_name).strip(" -")
    else:
        course_name = re.sub(r"\s+", " ", name_part).strip(" -")
    return course_name, teacher


def parse_cell(text: str, weekday: int) -> list[Course]:
    """解析一个课程单元格，返回其中的课程记录列表"""
    courses: list[Course] = []
    text = (text or "").strip()
    if not text:
        return courses

    # 单元格内多门课程以空行分隔
    segments = [s for s in re.split(r"\n\s*\n", text) if s.strip()]
    pending_name: str | None = None  # 课程名与周次行之间被空行截断时的暂存名

    for seg in segments:
        lines = [ln.strip() for ln in seg.split("\n") if ln.strip()]
        if not lines:
            continue

        ws_idx = next(
            (i for i, ln in enumerate(lines) if WEEK_SESSION_RE.match(ln)), None
        )
        if ws_idx is None:
            # 无周次节次行的段落：多半是"高等数学(Ⅱ)-1"这类课程名被空行截断，暂存
            joined = " ".join(lines)
            if pending_name is None:
                pending_name = joined
            else:
                pending_name += " " + joined
            continue

        if ws_idx == 0:
            # 段落以周次行开头：课程名在上一段（空行截断），如 "高等数学(Ⅱ)-1"
            name_part = pending_name or ""
            pending_name = None
        else:
            name_part = " ".join(lines[:ws_idx])

        m = WEEK_SESSION_RE.match(lines[ws_idx])
        weeks_text, sessions_text = m.group(1), m.group(2)
        location = " ".join(lines[ws_idx + 1:]).strip()

        course_name, teacher = _split_name_teacher(name_part)
        week_list = parse_weeks(weeks_text)
        session_list = parse_sessions(sessions_text)

        if not course_name or not week_list or not session_list:
            continue
        courses.append(Course(
            course_name=course_name, teacher=teacher, weekday=weekday,
            weeks_text=weeks_text, week_list=week_list,
            sessions_text=sessions_text, session_list=session_list,
            location=location,
        ))
    return courses


def _grid_from_xls(file_bytes: bytes) -> list[list[str]]:
    import xlrd
    wb = xlrd.open_workbook(file_contents=file_bytes)
    sheet = wb.sheet_by_index(0)
    return [
        [str(sheet.cell_value(r, c)) for c in range(sheet.ncols)]
        for r in range(sheet.nrows)
    ]


def _grid_from_xlsx(file_bytes: bytes) -> list[list[str]]:
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    ws = wb.active
    grid = [[("" if cell.value is None else str(cell.value)) for cell in row]
            for row in ws.iter_rows()]
    wb.close()
    return grid


def _extract_person_name(grid: list[list[str]]) -> str:
    """从标题行 '成都工业学院 刘雨昂 学生个人课表' 提取姓名"""
    for row in grid:
        for cell in row:
            if TITLE_KEYWORD in cell:
                prefix = cell.split(TITLE_KEYWORD)[0].strip()
                tokens = prefix.split()
                if tokens:
                    return tokens[-1]
                return ""
    return ""


def _extract_meta(grid: list[list[str]]) -> dict:
    meta = {}
    for row in grid:
        joined = " ".join(str(c) for c in row if str(c).strip())
        if not any(p.search(joined) for p in META_PATTERNS.values()):
            continue
        for key, pat in META_PATTERNS.items():
            m = pat.search(joined)
            if m:
                meta[key] = m.group(1)
    return meta


def _find_header_row(grid: list[list[str]]) -> tuple[int, dict[int, int]]:
    """定位星期表头行，返回 (行号, {列号: 星期几})"""
    for r, row in enumerate(grid):
        cols = {}
        for c, cell in enumerate(row):
            v = str(cell).strip()
            if v in WEEKDAY_MAP:
                cols[c] = WEEKDAY_MAP[v]
        if len(cols) >= 5:
            return r, cols
    return -1, {}


def parse_grid(grid: list[list[str]], file_name: str = "") -> ParsedSchedule:
    """解析课表二维网格，返回结构化结果"""
    result = ParsedSchedule(file_name=Path(file_name).name if file_name else "")

    result.name = _extract_person_name(grid)
    meta = _extract_meta(grid)
    result.term = meta.get("term", "")
    result.class_name = meta.get("class_name", "")
    result.major = meta.get("major", "")
    result.department = meta.get("department", "")

    m = STUDENT_ID_FILE_RE.search(Path(file_name).name if file_name else "")
    if m:
        result.student_id = m.group(1)

    header_row, weekday_cols = _find_header_row(grid)
    if header_row < 0:
        result.warnings.append("未找到星期表头行，文件格式可能不兼容")
        return result

    seen: set[tuple] = set()
    for r in range(header_row + 1, len(grid)):
        row = grid[r]
        if not row:
            continue
        first = str(row[0]).strip() if row else ""
        if not PERIOD_ROW_RE.match(first):
            continue  # 跳过备注等非节次行
        for c, weekday in weekday_cols.items():
            if c >= len(row):
                continue
            cell = str(row[c])
            for course in parse_cell(cell, weekday):
                if course.key() in seen:
                    continue  # 同一门跨节次课程会在多个节次行重复出现
                seen.add(course.key())
                result.courses.append(course)

    if not result.name:
        result.warnings.append("未能从标题行提取姓名")
    return result


def parse_schedule_file(file_bytes: bytes, file_name: str) -> ParsedSchedule:
    """解析上传的课表文件（.xls / .xlsx 均可），file_bytes 为文件内容字节流"""
    ext = Path(file_name).suffix.lower()
    if ext == ".xls":
        grid = _grid_from_xls(file_bytes)
    elif ext == ".xlsx":
        grid = _grid_from_xlsx(file_bytes)
    else:
        raise ValueError(f"不支持的文件格式：{ext}（仅支持 .xls / .xlsx）")
    return parse_grid(grid, file_name=file_name)


def parse_schedule_path(path: str | Path) -> ParsedSchedule:
    """从本地文件路径解析课表"""
    path = Path(path)
    return parse_schedule_file(path.read_bytes(), path.name)
