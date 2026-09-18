"""样例课表的隐私与完整性门禁

`samples/desensitized/` 下的样例是真实教务系统课表的**脱敏副本**，会随仓库提交。
本文件保证两件事：

1. 仓库里的样例不含真实个人信息（姓名/学号必须都是虚构值）；
2. 样例本身没有被手改坏——课程数、元信息与 `manifest` 一致，
   且解析器仍能按真实逻辑读出全部课程。

原始文件只留在本地（`.gitignore` 忽略 `samples/*.xls`），
用 `python tools/desensitize_samples.py` 生成脱敏副本；
本地有原始文件时加 `--check` 可校验副本与原始文件逐字段一致。
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from duty_system.parser import parse_schedule_path

SAMPLES = Path(__file__).parent / "samples"
DESENSITIZED = SAMPLES / "desensitized"
MANIFEST = DESENSITIZED / "manifest.json"

# 脱敏学号统一以 9999 开头，真实学号不会长这样
FAKE_ID_PREFIX = "9999"


def _manifest() -> list[dict]:
    if not MANIFEST.exists():
        pytest.skip(f"缺少 {MANIFEST.name}（先跑 tools/desensitize_samples.py）")
    data = json.loads(MANIFEST.read_text(encoding="utf-8"))
    assert isinstance(data, list) and data, "manifest 应为非空列表"
    return data


def _manifest_ids() -> list:
    """收集期（collection）不能调用 pytest.skip，因此这里用安全退避"""
    if not MANIFEST.exists():
        return []
    return json.loads(MANIFEST.read_text(encoding="utf-8"))


def test_desensitized_samples_are_committed() -> None:
    """脱敏样例必须随仓库提交：CI 检出后不能依赖本地才有的原始文件"""
    assert DESENSITIZED.is_dir(), "缺少 samples/desensitized 目录"
    files = sorted(DESENSITIZED.glob("*.xls"))
    assert len(files) >= 3, f"脱敏样例不足：{[f.name for f in files]}"


@pytest.mark.parametrize("entry", _manifest_ids(),
                         ids=lambda e: e["file"] if isinstance(e, dict) else "")
def test_manifest_entry_matches_sample(entry: dict) -> None:
    """manifest 记录必须与样例实际解析结果一致（防止手改样例或忘记刷新 manifest）"""
    path = DESENSITIZED / entry["file"]
    assert path.exists(), f"样例不存在：{entry['file']}"
    s = parse_schedule_path(path)
    assert s.name == entry["name"], f"姓名 {s.name!r} != manifest {entry['name']!r}"
    assert s.student_id == entry["student_id"]
    assert s.term == entry["term"]
    assert s.class_name == entry["class_name"]
    assert s.major == entry["major"]
    assert s.department == entry["department"]
    assert len(s.courses) == entry["course_count"], \
        f"课程数 {len(s.courses)} != manifest {entry['course_count']}"
    assert not s.warnings


@pytest.mark.parametrize("path", sorted(DESENSITIZED.glob("*.xls")),
                         ids=lambda p: p.name if isinstance(p, Path) else "")
def test_no_real_pii_in_sample(path: Path) -> None:
    """样例里不能出现真实个人信息

    判据（不依赖「知道真名是什么」，因此对新增样例同样有效）：
      - 学号必须是以 9999 开头的虚构号；
      - 姓名字段与 manifest 一致，且不是常见真实姓名写法（由脱敏脚本保证）。
    """
    s = parse_schedule_path(path)
    assert s.student_id.startswith(FAKE_ID_PREFIX), \
        f"学号 {s.student_id!r} 不像脱敏号（应以 {FAKE_ID_PREFIX} 开头）"

    known = {e["file"]: e for e in _manifest()}
    if path.name in known:
        assert s.name == known[path.name]["name"]
        assert s.student_id == known[path.name]["student_id"]


def test_samples_keep_real_course_data() -> None:
    """脱敏只改姓名/学号，课程数据必须保持真实（课程名/地点/周次都在）"""
    for path in sorted(DESENSITIZED.glob("*.xls")):
        s = parse_schedule_path(path)
        assert len(s.courses) >= 30, f"{path.name} 课程数异常少：{len(s.courses)}"
        # 真实课表的关键特征：有地点、有单双周式周次、有跨节次课程
        assert any(c.location for c in s.courses), f"{path.name} 没有解析出任何地点"
        assert any(len(c.session_list) >= 4 for c in s.courses), \
            f"{path.name} 没有跨节次（实验/实习类）课程"
        for c in s.courses:
            assert c.weekday in range(1, 8)
            assert c.week_list, f"{c.course_name} 周次为空"
            assert c.session_list, f"{c.course_name} 节次为空"


def test_real_samples_are_gitignored_if_present() -> None:
    """真实课表（若本地存在）不应被 git 跟踪

    用 .gitignore 规则判断：`samples/*.xls` 必须被忽略，
    同时 `samples/desensitized/` 不能被忽略（否则 CI 拿不到样例）。
    """
    ignore = SAMPLES.parent / ".gitignore"
    text = ignore.read_text(encoding="utf-8")
    assert "samples/*.xls" in text, "原始课表应被 .gitignore 忽略"
    assert "!samples/desensitized/" in text, "脱敏样例必须可提交（不被忽略）"
