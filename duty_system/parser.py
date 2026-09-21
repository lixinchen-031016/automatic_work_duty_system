"""课表文件解析模块

支持解析教务系统导出的"学生个人课表"(.xls / .xlsx)文件。

文件结构（以成都工业学院课表为例）：
    行0: "成都工业学院 张三 学生个人课表"          -> 标题（含学校、姓名）
    行1: "学年学期：2026-2027-1 班级：xx 专业：xx ..." -> 元信息
    行2: 星期一 ~ 星期日                            -> 表头
    行3-7: 节次行（1-2节/3-4节/5-6节/7-8节/9-11节） -> 每列单元格内含课程
    末行: 备注说明（：课程 教师 周次周;…）

单元格内多门课程以空行分隔，每门课程为多行文本：
    课程名（可能跨行，含 (板块x) 副标题）
    教师(职称)          <- 可能缺失，职称括号可能被换行打断
    周次([周])[节次节]  <- 如 18([周])[01-02-03-04节]
    地点                <- 可能缺失（如无地点或网课编号）

解析健壮性（均由 samples/ 下九份真实课表与变体用例覆盖）：
  - 表头识别兼容 星期一 / 周一 / 周1 / 礼拜一 / 星期7；
  - 周次兼容 单双周（1-16单周 / 1-16双周 / 1-16(单) / 奇偶周），
    以及 "1-16周"、"第1-16周" 等带字写法；
  - 节次分隔兼容 "-"、","、"、"、"~"，并兼容全角括号与全角数字；
  - 教师职称括号兼容中文全角括号（钱明明（教授）），也容忍括号被换行
    打断（钱明明\n(教授)）；职称按「白名单 + 兜底」识别，教务系统里
    不常见的写法（如「其他中级」）也能正确拆出教师名；
  - 多人授课按 郑明(副教授),周明明(讲师) 解析为 "郑明,周明明"，
    不会把第二位教师并进课程名；
  - 学号取文件名中最长的一段连续数字，兼容教务系统的 (1) 后缀；
  - 单元格按空行切分课程，但空行后仅剩周次行的情况会回退到「逐行」
    切分，避免课程名与周次行被空行拆散后丢失课程；
  - 合并单元格会把值填充到整个合并区域，避免「表头只出现在第一列」
    等导出差异导致整份课表读不出来。

备注行里还会出现课表网格中没有的课程（军训、思政实践、认识实习这类集中实践），
它们不占具体节次、但整周占用，因此单独放在 `ParsedSchedule.whole_week_courses`，
weekday 记为 `WHOLE_WEEK_WEEKDAY`，供排班时整天避开。

`samples/` 下九份真实课表（国际商务 37 门 / 材控 57 门 / 电信 57 门 / 机器人 39 门 /
材科 49 门 / 计科 45 门 / 物流管理 40 门 / 机械电子 30 门 / 电信 36 门）是
回归基线：每份都要求课程数、姓名、学号、班级不变，且解析结果能覆盖文件
末尾备注行列出的全部课程名/教师/周次。
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from pathlib import Path

# 星期表头：兼容 星期一 / 周一 / 周1 / 礼拜一 / 星期7 / 星期一（周） 等写法
WEEKDAY_MAP = {
    "星期一": 1, "星期二": 2, "星期三": 3, "星期四": 4,
    "星期五": 5, "星期六": 6, "星期日": 7,
}
WEEKDAY_LABELS = {1: "周一", 2: "周二", 3: "周三", 4: "周四", 5: "周五", 6: "周六", 7: "周日"}

# 星期表头归一化用：周一 / 周1 / 星期1 / 礼拜一 ...
WEEKDAY_ALIAS_RE = re.compile(
    r"^(?:星期|周|礼拜)([一二三四五六日天1-7])$")
WEEKDAY_DIGIT = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "日": 7, "天": 7}

# 值班时段块：与课表节次行对应
BLOCK_SESSIONS = {1: (1, 2), 2: (3, 4), 3: (5, 6), 4: (7, 8), 5: (9, 10, 11)}
BLOCK_LABELS = {
    1: "1-2节 08:30-10:05",
    2: "3-4节 10:20-11:55",
    3: "5-6节 14:00-15:35",
    4: "7-8节 15:50-17:25",
    5: "9-11节 18:50-21:15",
}

# 整周占用的星期取值：备注行里的集中实践（军训 / 思政实践 / 认识实习）不排
# 具体节次，只占满整周，网格里没有它对应的星期列。
WHOLE_WEEK_WEEKDAY = 0
# 整周占用要覆盖的节次（各时段块的并集，即 1-11 节）
WHOLE_WEEK_SESSIONS = tuple(sorted({s for block in BLOCK_SESSIONS.values() for s in block}))

# 周次行，如 "18([周])[01-02-03-04节]"、"1,3,5([周])[03-04节]"
# 周次部分允许 单双周/奇偶周、周字后缀、第字前缀；节次分隔允许 - , 、 ~
WEEK_SESSION_RE = re.compile(
    r"^第?(?P<weeks>[\d,\-、\s]+?)"                       # 周次：1-16,6 / 1,3,5
    r"\s*(?:周)?\s*"                                        # 允许 "1-16周"
    r"(?P<mode>单周|双周|单数周|双数周|奇周|偶周|\(单\)|\(双\)|（单）|（双）)?"
    r"\s*(?:\(\[周\]\)|\[周\]|（\[周\]）)?\s*"              # 允许省略或半/全角 [周]
    r"\[(?P<sessions>[\d,\-、~～\s]+?)节\]"                 # 节次：01-02 / 01,02
    r"\s*$"
)

# 节次行（第0列），如 "1-2节(8:30-10:05)\n(01,02)\n08:30-10:05"
PERIOD_ROW_RE = re.compile(r"^(\d+)\s*-\s*(\d+)节")

# 教师行，如 "钱明明(教授)"；兼容全角括号，职称括号可能被换行打断（合并后匹配）。
# 姓名按 2-8 字收：少数民族姓名（阿依古丽·买买提）会超过 4 字。
# 职称字典由教务系统维护且会扩充，白名单之外还有「其他中级」这类写法
# （真实课表里出现过），漏掉会让教师被并进课程名，因此白名单之外再加兜底规则。
TEACHER_TITLES = (
    "教授", "副教授", "讲师", "助教", "未评级", "研究员", "副研究员",
    "高级工程师", "工程师", "实验师", "高级实验师", "外教", "讲师（外聘）",
    "其他中级", "其他初级", "其他副高", "其他正高", "实验员", "助理研究员",
)
# 兜底：职称文本必须以 级/师/授/员/研/教 收尾。
# 课程副标题（(PD03-3) / (板块1羽毛球) / (通识课) / (五) / (25PJ04) / (PD8-1)）
# 都不以这些字收尾，因此不会被误当成职称——这条规则在 5 份真实课表的
# 71 条教师行上验证过，零误伤。
_TEACHER_TITLE_FALLBACK = r"(?:其他)?[\u4e00-\u9fa5A-Za-z]{0,6}(?:级|师|授|员|研|教)"
_TEACHER_TITLE_ALT = (
    "(?:"
    + "|".join(map(re.escape, sorted(TEACHER_TITLES, key=len, reverse=True)))
    + "|" + _TEACHER_TITLE_FALLBACK + ")"
)
_TEACHER_NAME = r"[\u4e00-\u9fa5·]{2,8}"
# 单个教师："钱明明(教授)"；姓名与括号之间允许空白（职称括号被换行打断后
# 行会被合并成 "钱明明 (教授)"）
TEACHER_RE = re.compile(
    rf"(?:^|\s)({_TEACHER_NAME})\s*[(（]\s*(?:{_TEACHER_TITLE_ALT})\s*[)）]"
)
# 单个教师 token（不带前导锚点），用于从已匹配的教师列表里逐个抽出姓名
_TEACHER_TOKEN_RE = re.compile(
    rf"({_TEACHER_NAME})\s*[(（]\s*(?:{_TEACHER_TITLE_ALT})\s*[)）]"
)
# 多人授课："郑明(副教授),周明明(讲师)" / "张三(讲师)、李四(教授)" /
# 空格分隔的 "张三(讲师) 李四(教授)"。
# 只有当前一项确实带职称括号时才会吞掉后面的分隔符，因此不会把
# 不带职称的课程名误当成教师。
TEACHER_LIST_RE = re.compile(
    rf"{_TEACHER_NAME}\s*[(（]\s*(?:{_TEACHER_TITLE_ALT})\s*[)）]"
    rf"(?:\s*[,，、]\s*{_TEACHER_NAME}\s*[(（]\s*(?:{_TEACHER_TITLE_ALT})\s*[)）]"
    rf"|\s+{_TEACHER_NAME}\s*[(（]\s*(?:{_TEACHER_TITLE_ALT})\s*[)）])*"
)

TITLE_KEYWORD = "学生个人课表"

META_PATTERNS = {
    "term": re.compile(r"学年学期[:：]\s*(\S+)"),
    "class_name": re.compile(r"班级[:：]\s*(\S+)"),
    "major": re.compile(r"专业[:：]\s*(\S+)"),
    "department": re.compile(r"院系[:：]\s*(\S+)"),
}

# 学号：取文件名里长度 >= 6 的连续数字。不锚定扩展名，因为教务系统导出的
# 文件名常带后缀（学生个人课表_9999453245(1).xls），锚定后会导致学号为空；
# 多段数字时取最长的一段，避免把日期（20260914）当成学号。
STUDENT_ID_FILE_RE = re.compile(r"(?<!\d)(\d{6,})(?!\d)")

# 各校在标题行里可能用的占位写法
UNKNOWN_NAMES = {"", "未知成员", "未知", "学生", "同学"}


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
    # 单双周标记：normal=每周 / odd=单周 / even=双周（供界面展示，不改变周次列表）
    week_mode: str = "normal"
    # 是否来自备注行的整周集中安排（军训/思政实践/认识实习）：
    # 这类记录没有星期与节次，weekday == WHOLE_WEEK_WEEKDAY
    whole_week: bool = False

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
    # 备注行里有、课表网格里没有的整周集中安排（军训/思政实践/认识实习）。
    # 单独成表而不是塞进 courses：它们没有星期/节次，混进去会让「按星期展示
    # 课程」和「课程数统计」都失真。
    whole_week_courses: list[Course] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


# 只做「无歧义」的字符归一化：全角数字 -> 半角、全角方括号 -> 半角方括号。
# 刻意不使用 unicodedata.normalize("NFKC")：它会把罗马数字与全角括号一起改掉
# （体育Ⅲ -> 体育III、桌球室（西肆三楼） -> 桌球室(西肆三楼)），
# 而课程名/地点必须与教务系统原样一致，否则界面显示会对不上。
_FULLWIDTH_DIGITS = str.maketrans("０１２３４５６７８９", "0123456789")
_FULLWIDTH_BRACKETS = str.maketrans({"［": "[", "］": "]"})

_WS_AND_INVISIBLE = [
    ("\r\n", "\n"), ("\r", "\n"),
    ("\u200b", ""), ("\ufeff", ""), ("\xa0", " "),
]


def normalize_cell(text: str) -> str:
    """归一化单元格文本：统一换行、去零宽字符、全角数字与方括号转半角。

    Excel 导出的文本里经常混入全角数字（１８）与 CRLF、零宽空格，
    不归一化会导致周次/节次正则匹配不上，课程被整条丢弃。
    """
    if not text:
        return ""
    text = str(text)
    for old, new in _WS_AND_INVISIBLE:
        text = text.replace(old, new)
    return text.translate(_FULLWIDTH_DIGITS).translate(_FULLWIDTH_BRACKETS)


def normalize_header(text: str) -> str:
    """归一化表头文本（去空白、全角转半角），便于匹配星期列"""
    return re.sub(r"\s+", "", normalize_cell(text))


def parse_weekday_header(text: str) -> int:
    """把表头文本解析成星期几（1-7），无法识别返回 0"""
    value = normalize_header(text)
    if not value:
        return 0
    if value in WEEKDAY_MAP:
        return WEEKDAY_MAP[value]
    m = WEEKDAY_ALIAS_RE.match(value)
    if m:
        token = m.group(1)
        return int(token) if token.isdigit() else WEEKDAY_DIGIT[token]
    return 0


def parse_weeks(text: str, mode: str = "") -> list[int]:
    """解析周次文本：'18' / '1-6' / '1,3,5,7' / '1-4,6-10' -> 周次列表

    mode 为单双周标记时，会在区间内只保留对应的周：
    '1-16单周' -> 1,3,5,...,15；'1-16双周' -> 2,4,...,16。
    """
    weeks: set[int] = set()
    for part in re.split(r"[,，、]", text):
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(\d+)\s*[-~～]\s*(\d+)$", part)
        if m:
            lo, hi = int(m.group(1)), int(m.group(2))
            if lo > hi:
                lo, hi = hi, lo
            weeks.update(range(lo, hi + 1))
            continue
        if part.isdigit():
            weeks.add(int(part))
    return keep_by_week_mode(sorted(weeks), mode)


def keep_by_week_mode(weeks: list[int], mode: str) -> list[int]:
    """按单双周过滤周次列表（mode 为空/normal 时原样返回）"""
    if not mode:
        return weeks
    if "单" in mode or "奇" in mode:
        return [w for w in weeks if w % 2 == 1]
    if "双" in mode or "偶" in mode:
        return [w for w in weeks if w % 2 == 0]
    return weeks


def parse_sessions(text: str) -> list[int]:
    """解析节次文本：'01-02-03-04' / '01,02' / '01、02' -> [1, 2, 3, 4]"""
    return sorted({int(x) for x in re.split(r"[-~～,，、]", text) if x.strip().isdigit()})


def _split_name_teacher(name_part: str) -> tuple[str, str]:
    """从合并后的名称文本中拆出课程名与教师。

    name_part 形如 '思想政治理论课实践教学 孙明明(副教授)'，
    教师匹配取最后一个（课程名副标题如 (PD8-1) 不含职称关键字，不会误匹配）。
    """
    teacher = ""
    # 取「第一个」教师而不是最后一个：课程名在前、教师在后，取第一个才能
    # 保证多人授课时第二位教师不会被当成课程名的一部分
    # （真实样例：电工电子技术B 郑明(副教授),周明明(讲师)）。
    m = TEACHER_LIST_RE.search(name_part)
    if m:
        teachers = _TEACHER_TOKEN_RE.findall(m.group(0))
        teacher = ",".join(teachers)
        course_name = (name_part[: m.start()] + " " + name_part[m.end():])
        course_name = re.sub(r"\s+", " ", course_name).strip(" -,，、")
    else:
        course_name = re.sub(r"\s+", " ", name_part).strip(" -")
    return course_name, teacher



def _split_segments(text: str) -> list[list[str]]:
    """把单元格文本切成「课程段落」，每个段落是若干非空行。

    主要按空行切分；若某段里出现了第二个「周次行」，说明空行并没有把课程
    分开（例如课程名与周次行之间被插入空行），此时退回逐行切分并重新聚合，
    保证不会因为空行位置异常而丢课。
    """
    raw_segments = [s for s in re.split(r"\n\s*\n", text) if s.strip()]
    ws_re = WEEK_SESSION_RE

    segments: list[list[str]] = []
    for seg in raw_segments:
        lines = [ln.strip() for ln in seg.split("\n") if ln.strip()]
        if not lines:
            continue
        week_line_idx = [i for i, ln in enumerate(lines) if ws_re.match(normalize_cell(ln))]
        if len(week_line_idx) <= 1:
            segments.append(lines)
            continue
        # 一段里有多个周次行 -> 空行不可靠，改为逐行切分后按「周次行」聚合
        segments.extend(_split_lines_by_week_line(lines))
    return segments


# 地点行的常见特征：楼/室/场/馆等字样，或含房间号、网课编号
_LOCATION_HINT_RE = re.compile(
    r"(楼|室|场|馆|厅|教室|实验室|中心|操场|田径|智慧树|网课)"
    r"|[A-Za-z]{2,}\s*\d+"      # 如 ZHSWK03
    r"|\d{3,}"                    # 如 1205 / 021
)


def _looks_like_location(line: str) -> bool:
    return bool(_LOCATION_HINT_RE.search(line))


def _split_lines_by_week_line(lines: list[str]) -> list[list[str]]:
    """逐行扫描：把「课程名(+教师) + 周次行 + 地点」聚成一个课程段落。

    仅在单元格内课程之间没有空行时需要逐行判断。难点是周次行之后的那一行
    既可能是本门课程的地点，也可能是下一门课程的课程名开头：
      - 地点（德五楼3412 / 允明楼1236 / 智慧树ZHSWK03）带有房间号或楼室场字样；
      - 课程名一般是文字，不含这些特征。
    因此用地点特征词判定，判不出来就当作下一门课程的开头，
    宁可少解析一个地点，也不要把地点并进课程名（会污染整门课的显示）。
    """
    segments: list[list[str]] = []
    pending: list[str] = []
    location_open = False          # 刚读完周次行，正在等可能的地点行
    for line in lines:
        if WEEK_SESSION_RE.match(normalize_cell(line)):
            segments.append(pending + [line])
            pending = []
            location_open = True
            continue
        if location_open:
            location_open = False
            if _looks_like_location(line):
                segments[-1].append(line)
                continue
        pending.append(line)
    if pending:
        if segments:
            segments[-1].extend(pending)   # 尾部无周次行的行归入上一段
        else:
            segments.append(pending)
    return segments


def parse_cell(text: str, weekday: int) -> list[Course]:
    """解析一个课程单元格，返回其中的课程记录列表"""
    courses: list[Course] = []
    text = normalize_cell(text)
    if not text.strip():
        return courses

    pending_name: str | None = None  # 课程名与周次行之间被空行截断时的暂存名

    for lines in _split_segments(text):
        if not lines:
            continue

        ws_idx = next(
            (i for i, ln in enumerate(lines)
             if WEEK_SESSION_RE.match(normalize_cell(ln))), None
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

        m = WEEK_SESSION_RE.match(normalize_cell(lines[ws_idx]))
        weeks_text, sessions_text = m.group("weeks").strip(), m.group("sessions").strip()
        mode = (m.group("mode") or "").strip("()（）")
        location = " ".join(lines[ws_idx + 1:]).strip()

        course_name, teacher = _split_name_teacher(name_part)
        week_list = parse_weeks(weeks_text, mode)
        session_list = parse_sessions(sessions_text)

        if not course_name or not week_list or not session_list:
            continue
        week_mode = _mode_label(mode)
        courses.append(Course(
            course_name=course_name, teacher=teacher, weekday=weekday,
            # weeks_text 保留「单周/双周」字样：界面直接展示它，
            # 只显示 "1-16" 会让人误以为每周都上课
            weeks_text=_weeks_label(weeks_text, week_mode),
            week_list=week_list,
            sessions_text=sessions_text, session_list=session_list,
            location=location, week_mode=week_mode,
        ))
    return courses


def _mode_label(mode: str) -> str:
    """单双周标记 -> normal/odd/even"""
    if "单" in mode or "奇" in mode:
        return "odd"
    if "双" in mode or "偶" in mode:
        return "even"
    return "normal"


def _weeks_label(weeks_text: str, week_mode: str) -> str:
    """界面展示用的周次文本，单双周时补上字样（如 '1-16单周'）"""
    suffix = {"odd": "单周", "even": "双周"}.get(week_mode, "")
    return f"{weeks_text}{suffix}"


def _expand_merged(grid: list[list[str]], spans: list[tuple[int, int, int, int]]) -> list[list[str]]:
    """把合并单元格的值填满整个区域。

    Excel 里合并区域只有左上角有值，其余为空——若表头/节次列被合并，
    直接按格子读会丢列。spans 为 (r1, r2, c1, c2)（半开区间）。
    """
    for r1, r2, c1, c2 in spans:
        if not (r1 < len(grid) and c1 < len(grid[r1])):
            continue
        value = grid[r1][c1]
        if not str(value).strip():
            continue
        for r in range(r1, min(r2, len(grid))):
            row = grid[r]
            for c in range(c1, min(c2, len(row))):
                if not str(row[c]).strip():
                    row[c] = value
    return grid


def _read_xls_grid(file_bytes: bytes, with_formatting: bool) -> tuple[list[list[str]], list]:
    """按指定模式读取 .xls，返回 (网格, 合并区域)"""
    import xlrd
    wb = xlrd.open_workbook(file_contents=file_bytes, formatting_info=with_formatting)
    sheet = wb.sheet_by_index(0)
    grid = [
        [normalize_cell(sheet.cell_value(r, c)) for c in range(sheet.ncols)]
        for r in range(sheet.nrows)
    ]
    return grid, list(getattr(sheet, "merged_cells", []) or [])


def _grid_from_xls(file_bytes: bytes) -> list[list[str]]:
    """读取 .xls。

    优先带 formatting_info 打开，这样可以拿到 merged_cells 并把合并单元格
    的值填满整块；但该模式对部分文件不被支持（例如把 xlsx 改名成 .xls、
    或带密码保护的工作簿会抛 NotImplementedError/XLRDError），
    此时降级为普通模式读取——宁可少解析合并区域，也不能让整份课表读不出来。
    """
    try:
        grid, spans = _read_xls_grid(file_bytes, with_formatting=True)
    except Exception:  # noqa: BLE001 - 任何 xlrd 打开失败都降级重试
        grid, spans = _read_xls_grid(file_bytes, with_formatting=False)
    return _expand_merged(grid, spans)


def _grid_from_xlsx(file_bytes: bytes) -> list[list[str]]:
    """读取 .xlsx。

    用 read_only=False 是为了拿 merged_cells（只读模式不支持合并区域）；
    成员课表体量很小（几十行），普通模式的内存开销可以忽略。
    """
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    try:
        ws = wb.active
        grid = [[normalize_cell("" if cell.value is None else cell.value) for cell in row]
                for row in ws.iter_rows()]
        spans = [(rng.min_row - 1, rng.max_row, rng.min_col - 1, rng.max_col)
                 for rng in ws.merged_cells.ranges]
        return _expand_merged(grid, spans)
    finally:
        wb.close()


# 备注行条目：`课程名[ 教师] 周次周`，如
#   "大学军事技能训练  11-12周"（无教师，课程名与周次之间是两个空格）
#   "物流前沿讲座Ⅰ 景乔松,罗霞,王宁 18周"
#   "思想政治理论课实践教学 王丽茹 21周"
_NOTE_WEEKS_RE = re.compile(r"^[\d,，、~～\-\s]+周$")
_NOTE_TEACHER_RE = re.compile(r"^[\u4e00-\u9fa5·]{2,8}(?:[,，、][\u4e00-\u9fa5·]{2,8})*$")


def _grid_course_matching(key: str, grid_names) -> str | None:
    """key 是否指向网格里已有的课程（同名，或互为前缀）

    互为前缀是为了兼容副标题差异：备注行写「体育Ⅰ」，网格里是
    「体育Ⅰ (板块1A1乒乓球)」——这类条目属于网格里已有的课，不该当成
    新增的整周安排。
    """
    if not key:
        return None
    if key in grid_names:
        return key
    for name in grid_names:
        if name.startswith(key) or key.startswith(name):
            return name
    return None


def _unique_grid_schedule(
    matched_name: str,
    teacher: str,
    grid_courses: list[Course],
) -> Course | None:
    """找备注课程唯一可继承的网格星期/节次；多套安排时返回 None。"""
    candidates = [c for c in grid_courses if c.course_name == matched_name]
    if teacher:
        teacher_matches = [
            c for c in candidates
            if c.teacher and teacher == c.teacher.strip()
        ]
        if teacher_matches:
            candidates = teacher_matches
    schedules: dict[tuple[int, tuple[int, ...]], Course] = {}
    for course in candidates:
        if course.weekday == WHOLE_WEEK_WEEKDAY or not course.session_list:
            continue
        schedules.setdefault(
            (course.weekday, tuple(course.session_list)), course)
    if len(schedules) != 1:
        return None
    return next(iter(schedules.values()))


def find_note_text(grid: list[list[str]]) -> str:
    """找出教务系统写在末尾的「备注」单元格文本

    备注单元格以全角/半角冒号开头，内容是 `课程名 教师 周次周;` 清单。
    课表正文里也有以冒号开头的内容（节次列是时间，不会），因此额外要求
    至少有一条「以周次结尾」的条目才认定为备注。
    """
    best = ""
    for row in grid:
        for cell in row:
            text = normalize_cell(cell).strip()
            if not text or text[0] not in "：:":
                continue
            body = text.lstrip("：:").strip()
            if len(body) <= len(best):
                continue
            items = [x.strip() for x in body.split(";") if x.strip()]
            if not any(_NOTE_WEEKS_RE.match(x.split()[-1]) for x in items if x.split()):
                continue
            best = body
    return best


def _compress_weeks(weeks: list[int]) -> str:
    """周次列表 -> 紧凑文本： [1,2,3,5] -> '1-3,5'（与 parse_weeks 互逆）"""
    parts: list[str] = []
    run_start = prev = None
    for w in sorted(set(weeks)):
        if run_start is None:
            run_start = prev = w
            continue
        if w == prev + 1:
            prev = w
            continue
        parts.append(str(run_start) if run_start == prev else f"{run_start}-{prev}")
        run_start = prev = w
    if run_start is not None:
        parts.append(str(run_start) if run_start == prev else f"{run_start}-{prev}")
    return ",".join(parts)


def parse_note_courses(note_text: str, grid_courses) -> tuple[list[Course], list[str]]:
    """解析备注行，返回 (整周集中安排, 警告信息)

    备注行是教务系统自己生成的课程清单，比课表网格更全：网格只画得出「有星期
    有节次」的课，军训、思政实践这类集中实践没有具体节次，只会出现在备注里；
    也有课程的某几周只在备注里出现（网格漏画）。

    处理规则：
    * 若课程名能在网格中匹配到唯一的星期/节次，则用网格时间补全备注缺失周次；
    * 若网格里完全没有该课程，或同名课程存在多套时间而无法判断，才按
      「整周避让」处理，并给出警告说明具体星期未知。

    无法定位的整周记录单独放进 `ParsedSchedule.whole_week_courses`，
    weekday 记为 WHOLE_WEEK_WEEKDAY。

    备注行里大部分条目其实是网格已有课程的变更记录（教师换人、周次调整），
    这类直接跳过——网格更权威，它带星期/节次/地点。
    """
    if not note_text:
        return [], []

    # 网格覆盖情况：课程名 -> 已覆盖周次
    coverage: dict[str, set[int]] = {}
    for c in grid_courses:
        coverage.setdefault(c.course_name, set()).update(c.week_list)

    # key -> (未覆盖周次, 依据说明)；同一门课会分多条写在备注里，合并后再提示一次
    whole_merged: dict[tuple[str, str], tuple[set[int], str]] = {}
    supplemental: dict[
        tuple[str, str, int, str, tuple[int, ...], str], set[int]
    ] = {}
    for item in (x.strip() for x in note_text.split(";")):
        if not item:
            continue
        tokens = item.split()
        if len(tokens) < 2:
            continue
        weeks_token = tokens[-1]
        if not _NOTE_WEEKS_RE.match(weeks_token):
            continue
        head = tokens[:-1]
        # 教师是排在周次前的一段姓名清单；没有教师时（"军训  11-12周"）整段都是课程名
        teacher = head[-1] if len(head) >= 2 and _NOTE_TEACHER_RE.match(head[-1]) else ""
        name = " ".join(head[:-1]) if teacher else " ".join(head)
        weeks = parse_weeks(weeks_token[:-1])
        if not name or not weeks:
            continue

        matched = _grid_course_matching(name, coverage)
        if matched is None:
            uncovered = set(weeks)
            reason = "课表中没有该课程的网格安排"
        else:
            # 用网格里的课程名（备注行常省略副标题，如「体育Ⅰ」对「体育Ⅰ (板块1A1乒乓球)」）
            name = matched
            uncovered = set(weeks) - coverage[matched]
            reason = f"课表网格只画了第 {_compress_weeks(sorted(coverage[matched]))} 周"
        if not uncovered:
            continue

        # 同课程且网格时间唯一时，直接继承星期/节次并补全周次；只有确实
        # 无法定位时间的条目才降级为整周避让。
        if matched is not None:
            schedule = _unique_grid_schedule(matched, teacher, grid_courses)
            if schedule is not None:
                if not teacher or teacher == schedule.teacher.strip():
                    merged_weeks = sorted(set(schedule.week_list) | uncovered)
                    schedule.week_list = merged_weeks
                    schedule.weeks_text = _weeks_label(
                        _compress_weeks(merged_weeks), "normal")
                    schedule.week_mode = "normal"
                    coverage[matched].update(uncovered)
                else:
                    key = (
                        matched, teacher, schedule.weekday,
                        schedule.sessions_text, tuple(schedule.session_list),
                        schedule.location,
                    )
                    supplemental.setdefault(key, set()).update(uncovered)
                continue

        prev = whole_merged.get((name, teacher))
        whole_merged[(name, teacher)] = (
            (prev[0] | uncovered) if prev else set(uncovered), reason)

    warnings = [
        f"{name} 第 {_compress_weeks(sorted(weeks))} 周"
        + (f"（{teacher}）" if teacher else "")
        + f"来自备注行，{reason}，具体上课时间未知；排班时按整周避让。"
        for (name, teacher), (weeks, reason) in whole_merged.items()
    ]
    courses = [
        Course(
            course_name=name, teacher=teacher, weekday=WHOLE_WEEK_WEEKDAY,
            weeks_text=_weeks_label(_compress_weeks(sorted(weeks)), "normal"),
            week_list=sorted(weeks), sessions_text="", session_list=[],
            location="", whole_week=True,
        )
        for (name, teacher), (weeks, _) in whole_merged.items()
    ]
    for (name, teacher, weekday, sessions_text, sessions, location), weeks in supplemental.items():
        grid_courses.append(Course(
            course_name=name, teacher=teacher, weekday=weekday,
            weeks_text=_weeks_label(_compress_weeks(sorted(weeks)), "normal"),
            week_list=sorted(weeks), sessions_text=sessions_text,
            session_list=list(sessions), location=location,
        ))
    return courses, warnings


def _extract_person_name(grid: list[list[str]]) -> str:
    """从标题行 '成都工业学院 赵明明 学生个人课表' 提取姓名"""
    for row in grid:
        for cell in row:
            if TITLE_KEYWORD in cell:
                prefix = cell.split(TITLE_KEYWORD)[0].strip()
                tokens = prefix.split()
                if tokens:
                    return tokens[-1]
                return ""
    return ""


def _extract_student_id(file_name: str) -> str:
    """从文件名里提取学号。

    教务系统导出的文件名形态很多：`学生个人课表_9999800598.xls`、
    `学生个人课表_9999453245(1).xls`、`课表_9999453245（1）.xls`。
    早期实现把正则锚定在 `.xls$` 上，导致带 `(1)` 后缀的文件学号为空；
    这里改为取文件名中最长的一段连续数字（>= 6 位），多段时取最长，
    这样既兼容各种后缀，也不会把打印日期误当成学号。
    """
    if not file_name:
        return ""
    candidates = STUDENT_ID_FILE_RE.findall(file_name)
    if not candidates:
        return ""
    return max(candidates, key=len)


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
    """定位星期表头行，返回 (行号, {列号: 星期几})

    取「识别出的星期列最多」的那一行：课表正文里出现的「周一周二…」
    字样可能被误判为表头，按覆盖数取最大可避免选错行。
    """
    best_row, best_cols = -1, {}
    for r, row in enumerate(grid):
        cols = {}
        for c, cell in enumerate(row):
            weekday = parse_weekday_header(cell)
            if weekday:
                cols[c] = weekday
        if len(cols) > len(best_cols):
            best_row, best_cols = r, cols
    if len(best_cols) < 2:
        return -1, {}
    return best_row, best_cols


def parse_grid(grid: list[list[str]], file_name: str = "") -> ParsedSchedule:
    """解析课表二维网格，返回结构化结果"""
    result = ParsedSchedule(file_name=Path(file_name).name if file_name else "")

    result.name = _extract_person_name(grid)
    meta = _extract_meta(grid)
    result.term = meta.get("term", "")
    result.class_name = meta.get("class_name", "")
    result.major = meta.get("major", "")
    result.department = meta.get("department", "")

    result.student_id = _extract_student_id(Path(file_name).name if file_name else "")

    header_row, weekday_cols = _find_header_row(grid)
    if header_row < 0:
        result.warnings.append("未找到星期表头行，文件格式可能不兼容")
        return result

    seen: set[tuple] = set()
    for r in range(header_row + 1, len(grid)):
        row = grid[r]
        if not row:
            continue
        first = normalize_cell(row[0]).strip() if row else ""
        if not PERIOD_ROW_RE.match(first):
            continue  # 跳过备注等非节次行
        for c, weekday in weekday_cols.items():
            if c >= len(row):
                continue
            for course in parse_cell(row[c], weekday):
                if course.key() in seen:
                    continue  # 同一门跨节次课程会在多个节次行重复出现
                seen.add(course.key())
                result.courses.append(course)

    if not result.name:
        result.warnings.append("未能从标题行提取姓名")

    note_text = find_note_text(grid)
    result.whole_week_courses, note_warnings = parse_note_courses(note_text, result.courses)
    result.warnings.extend(note_warnings)
    return result


def _sniff_format(file_bytes: bytes) -> str:
    """按文件内容判断真实格式，返回 'xlsx' / 'xls' / ''

    不能只看扩展名：教务系统导出的文件经常被改名（.xls 里装的是 xlsx，
    或反之），此时按扩展名选择解析器会直接报错、整份课表读不出来。
    xlsx 是 zip 容器（PK\x03\x04），xls 是 OLE2 复合文档（D0 CF 11 E0）。
    """
    if file_bytes[:4] == b"PK\x03\x04":
        return "xlsx"
    if file_bytes[:4] == b"\xd0\xcf\x11\xe0":
        return "xls"
    return ""


def parse_schedule_file(file_bytes: bytes, file_name: str) -> ParsedSchedule:
    """解析上传的课表文件（.xls / .xlsx 均可），file_bytes 为文件内容字节流"""
    ext = Path(file_name).suffix.lower()
    sniffed = _sniff_format(file_bytes)
    order: list[str] = []
    if sniffed:
        order.append(sniffed)           # 内容优先：扩展名不可信
    if ext == ".xls":
        order.append("xls")
    elif ext == ".xlsx":
        order.append("xlsx")
    else:
        raise ValueError(f"不支持的文件格式：{ext}（仅支持 .xls / .xlsx）")
    order.extend(f for f in ("xlsx", "xls") if f not in order)   # 兜底再试另一种

    errors: list[str] = []
    for fmt in order:
        try:
            grid = _grid_from_xlsx(file_bytes) if fmt == "xlsx" else _grid_from_xls(file_bytes)
        except Exception as exc:  # noqa: BLE001 - 换另一种格式再试
            errors.append(f"{fmt}: {exc}")
            continue
        return parse_grid(grid, file_name=file_name)

    raise ValueError(
        "无法读取该课表文件（已尝试 " + "、".join(order) + " 两种格式）。"
        "请确认文件未被加密或损坏：" + "；".join(errors))


def parse_schedule_path(path: str | Path) -> ParsedSchedule:
    """从本地文件路径解析课表"""
    path = Path(path)
    return parse_schedule_file(path.read_bytes(), path.name)
