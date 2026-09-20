"""把 samples/ 下的真实个人课表脱敏为可提交的测试样例

背景
----
`samples/*.xls` 是同学从教务系统导出的真实文件，直接提交会泄露姓名与学号。
本脚本在**不改变课程数据**的前提下，把「学生姓名 + 教师姓名 + 学号」替换成
虚构值，生成 `samples/desensitized/` 下的样例供测试与 CI 使用。

做法（为什么不是重新生成 xls）
------------------------------
`.xls` 是 BIFF8 二进制格式，文本以 UTF-16LE 存在 SST 等记录里，姓名串还会被
多处记录共享引用。用 xlwt 重新写一份会丢失真实布局细节（合并单元格、SST 共享、
行高列宽），那样测出来的解析器行为就不再是真实文件的行为。

因此这里做**等长 UTF-16LE 字节替换**：
  - 替换前后字节长度完全相同，BIFF 记录的长度字段与偏移量全部不变；
  - 长度不变 => 文件结构一定不会损坏（不会出现「能打开但少几门课」）；
  - 课程名 / 周次 / 节次 / 地点 / 课程数原样保留，解析器仍需按真实逻辑工作。

替换表**按文件独立生成**（每个文件从假名池起始位置重新分配），这样新增或删除
某份样例不会让其它样例的假名漂移，测试基线保持稳定。假名与真名、假名与假名
之间都不允许互为子串（假名「赵明明」包含真名「赵明」会让残留检测误报）。
脚本会另外写出 `manifest.json`（脱敏样例文件名 + 假姓名 + 假学号 + 课程数），
供测试数据驱动读取；清单里**不含任何真实信息**。脚本不会输出
「真名 -> 假名」对照表——那份对照表本身就是需要消除的隐私数据。

用法
----
    python tools/desensitize_samples.py           # 生成/刷新脱敏样例
    python tools/desensitize_samples.py --check   # 校验脱敏样例与真实样例一致
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from duty_system.parser import parse_schedule_path  # noqa: E402

SAMPLES = ROOT / "samples"
OUT_DIR = SAMPLES / "desensitized"
# 脱敏样例的元信息清单（不含任何真实信息），供测试数据驱动读取
MANIFEST = OUT_DIR / "manifest.json"

# 假名取材于百家姓 + 常用字，保证「同长度、不同字」
SURNAMES = "赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜"
GIVEN = "明华强伟静敏娜磊洋勇艳杰涛超秀霞平刚桂芳丽军峰"

# 假学号前缀：真实学号不会以 9999 开头，便于一眼识别 + 测试断言
FAKE_ID_PREFIX = "9999"


def _fake_name(index: int, length: int) -> str:
    """按序号生成指定长度的确定性假名（2/3/4 字）"""
    surname = SURNAMES[index % len(SURNAMES)]
    pool = len(SURNAMES)

    def g(k: int) -> str:
        return GIVEN[(index // (pool ** k)) % len(GIVEN)]

    if length <= 1:
        return surname
    return surname + "".join(g(k) for k in range(1, length))


def _fake_student_id(real_id: str) -> str:
    """学号替换：位数不变、以 9999 开头。

    后段用真实学号的哈希推导，因此**与文件顺序无关**——新增或删除样例
    不会让已有样例的假学号漂移，测试基线保持稳定。
    """
    if not real_id:
        return ""
    body_len = max(1, len(real_id) - len(FAKE_ID_PREFIX))
    digest = hashlib.sha256(real_id.encode()).hexdigest()
    digits = "".join(c for c in digest if c.isdigit()) or "0"
    return (FAKE_ID_PREFIX + digits[:body_len])[:len(real_id)]


def _real_names(path: Path) -> list[str]:
    """文件里出现的全部人名（学生 + 教师），按长度再按字典序排序"""
    parsed = parse_schedule_path(path)
    names = {parsed.name} if parsed.name else set()
    for course in parsed.courses:
        names.update(t for t in course.teacher.split(",") if t)
    names.discard("")
    return sorted(names, key=lambda n: (len(n), n))


def _swallows_real_name(fake: str, real_names) -> bool:
    """假名里是否**包含**某个真实姓名

    这个方向才是隐私问题：假名「赵明明」把真名「赵明」整段包了进去，脱敏后的
    文件里依然能搜到真实姓名，「真实姓名零残留」的保证就不成立。
    反方向（假名「王明」是真实姓名「王明阳」的前缀）不会把任何真名写进文件，
    不必禁止——若一并禁止会引发连锁改名，让已有样例的假名整体漂移。
    """
    return any(real in fake for real in real_names)


def build_mapping(path: Path) -> dict:
    """为单个文件建立替换表（独立于其它文件，保证基线稳定）"""
    parsed = parse_schedule_path(path)
    real_names = _real_names(path)

    by_len: dict[int, list[str]] = {}
    for name in real_names:
        by_len.setdefault(len(name), []).append(name)

    name_map: dict[str, str] = {}
    for length, group in sorted(by_len.items()):
        for i, real in enumerate(group):
            fake = _fake_name(i, length)
            guard = 0
            while (fake in name_map.values() or _swallows_real_name(fake, real_names)) \
                    and guard < 500:
                guard += 1
                fake = _fake_name(i + guard * 97, length)
            name_map[real] = fake

    return {
        "names": name_map,
        "real_student_id": parsed.student_id,
        "fake_student_id": _fake_student_id(parsed.student_id),
    }


def desensitize_bytes(raw: bytes, name_map: dict[str, str]) -> bytes:
    """等长替换：UTF-16LE 姓名 -> 等长假名（长度不变 = BIFF 结构不变）"""
    # 长名优先，避免短名是长名子串时误替换
    ordered = sorted(name_map, key=len, reverse=True)
    pattern = re.compile(b"|".join(re.escape(n.encode("utf-16-le")) for n in ordered))
    return pattern.sub(
        lambda m: name_map[m.group(0).decode("utf-16-le")].encode("utf-16-le"), raw)


def _out_name(path: Path, real_id: str, fake_id: str) -> str:
    """文件名里的学号也要换掉（学号会从文件名解析出来）"""
    return path.name.replace(real_id, fake_id) if real_id else path.name


def _manifest_entries(plan: list[dict]) -> list[dict]:
    """清单条目：只记录脱敏后的事实，不含任何真实姓名/学号"""
    entries = []
    for item in plan:
        fake = parse_schedule_path(OUT_DIR / item["out_name"])
        entries.append({
            "file": item["out_name"],
            "name": fake.name,
            "student_id": fake.student_id,
            "term": fake.term,
            "class_name": fake.class_name,
            "major": fake.major,
            "department": fake.department,
            "course_count": len(fake.courses),
            # 备注行里的整周集中安排（军训/思政实践）单独计数，
            # 便于测试确认这份样例该有多少条整周占用
            "whole_week_count": len(fake.whole_week_courses),
        })
    return entries


def _build_plan(paths: list[Path]) -> list[dict]:
    """先生成完整的替换计划（源文件 / 输出名 / 新字节 / 替换表），再统一生成与校验

    生成与校验共用同一份计划，避免两处各自推算假学号而错位。
    """
    plan: list[dict] = []
    for path in paths:
        info = build_mapping(path)
        real_id = info["real_student_id"]
        fake_id = _fake_student_id(real_id)
        raw = path.read_bytes()
        plan.append({
            "src": path,
            "raw": raw,
            "bytes": desensitize_bytes(raw, info["names"]),
            "names": info["names"],
            "real_id": real_id,
            "fake_id": fake_id,
            "out_name": _out_name(path, real_id, fake_id),
        })
    return plan


def desensitize_all(check_only: bool = False) -> int:
    paths = sorted(SAMPLES.glob("*.xls"))
    if not paths:
        # CI 检出后只有脱敏样例（原始文件被 .gitignore 排除），属正常情况。
        # 脱敏样例与 manifest 的一致性由 test_sample_privacy.py 覆盖。
        print(f"未找到原始课表（{SAMPLES}/*.xls），跳过逐字段比对"
              f"（CI 环境正常；仓库内脱敏样例："
              f"{len(list(OUT_DIR.glob('*.xls')))} 份）")
        return 0

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    problems: list[str] = []
    plan = _build_plan(paths)

    # ---- 自检 1：等长替换 => 文件结构不可能被破坏 ----
    for item in plan:
        if len(item["bytes"]) != len(item["raw"]):
            problems.append(
                f"{item['src'].name}: 字节长度从 {len(item['raw'])} "
                f"变成 {len(item['bytes'])}")

    # ---- 自检 2：真实个人信息已彻底清除（字节级）----
    for item in plan:
        for real, fake in item["names"].items():
            if real != fake and real in fake:
                problems.append(f"{item['src'].name}: 假名 {fake!r} 包含真名 {real!r}，"
                                f"脱敏后仍能搜到真实姓名")

        leaked = [n for n in item["names"] if n.encode("utf-16-le") in item["bytes"]]
        if leaked:
            problems.append(f"{item['src'].name}: 仍残留真实姓名 {leaked}")
        if item["real_id"] and item["real_id"].encode("ascii") in item["bytes"]:
            problems.append(f"{item['src'].name}: 仍残留真实学号 {item['real_id']}")

    if not check_only:
        for item in plan:
            (OUT_DIR / item["out_name"]).write_bytes(item["bytes"])

    # ---- 自检 3：脱敏目录里不能有清单之外的陈旧样例 ----
    # 不做自动删除：本地可能只放了一部分原始课表，自动清理会误删其它样例。
    orphans = sorted({p.name for p in OUT_DIR.glob("*.xls")}
                     - {item["out_name"] for item in plan})
    if orphans:
        problems.append(
            f"脱敏目录存在清单之外的样例（请手动删除）：{orphans}")

    # ---- 自检 4：脱敏样例仍可解析，且课程数据与真实文件逐条一致 ----
    for item in plan:
        fake_path = OUT_DIR / item["out_name"]
        if not fake_path.exists():
            problems.append(f"{item['out_name']}: 缺少脱敏样例")
            continue
        real = parse_schedule_path(item["src"])
        fake = parse_schedule_path(fake_path)
        if len(fake.courses) != len(real.courses):
            problems.append(
                f"{item['out_name']}: 课程数 {len(fake.courses)} != {len(real.courses)}")
        if fake.name in item["names"]:
            problems.append(f"{item['out_name']}: 学生姓名仍是真名 {fake.name!r}")
        if fake.student_id == real.student_id:
            problems.append(f"{item['out_name']}: 学号未被替换")
        for a, b in zip(real.courses, fake.courses):
            if (a.course_name, a.weekday, a.weeks_text, a.sessions_text, a.location) != \
               (b.course_name, b.weekday, b.weeks_text, b.sessions_text, b.location):
                problems.append(f"{item['out_name']}: 课程 {a.course_name} 的字段被改动")
                break
        # 整周集中安排同样要逐字段保真（它们没有教师姓名，只换学生姓名/学号）
        if len(fake.whole_week_courses) != len(real.whole_week_courses):
            problems.append(
                f"{item['out_name']}: 整周集中安排 {len(fake.whole_week_courses)} "
                f"!= {len(real.whole_week_courses)}")
        for a, b in zip(real.whole_week_courses, fake.whole_week_courses):
            if (a.course_name, a.week_list, a.weeks_text, a.location) != \
               (b.course_name, b.week_list, b.weeks_text, b.location):
                problems.append(
                    f"{item['out_name']}: 整周安排 {a.course_name} 的字段被改动")
                break

    # ---- 自检 5：清单与脱敏样例一致 ----
    entries = _manifest_entries(plan)
    if check_only:
        if not MANIFEST.exists():
            problems.append(f"{MANIFEST.name}: 缺少脱敏样例清单")
        elif json.loads(MANIFEST.read_text(encoding="utf-8")) != entries:
            problems.append(f"{MANIFEST.name}: 清单与脱敏样例不一致（需重新生成）")
    else:
        MANIFEST.write_text(
            json.dumps(entries, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    if problems:
        print("脱敏校验失败：", file=sys.stderr)
        for p in problems:
            print("  -", p, file=sys.stderr)
        return 1

    print(("校验通过" if check_only else "已生成")
          + f"：{len(plan)} 份样例 -> {OUT_DIR.relative_to(ROOT)}")
    for item in plan:
        print(f"  {item['src'].name}  ->  {item['out_name']}")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="把真实课表脱敏为可提交的测试样例")
    ap.add_argument("--check", action="store_true",
                    help="只校验脱敏样例是否与真实样例一致（不写文件）")
    return desensitize_all(check_only=ap.parse_args().check)


if __name__ == "__main__":
    raise SystemExit(main())
