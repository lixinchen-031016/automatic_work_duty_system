"""GitHub Actions 工作流静态校验

背景：`build.yml` 里有一段 `run: |` 的 shell 脚本曾漏写一个闭合双引号，
本地一切正常、推送后才在 CI 的「安装 Qt 运行库」步骤炸掉：

    unexpected EOF while looking for matching `"'

这类问题本地 `pytest` 完全发现不了（pytest 不读 yaml），但代价是一次失败的
CI 运行 + 一次修复提交。本文件把工作流也纳入测试门禁：

  - YAML 必须能解析（缩进/引号把 YAML 写坏时立刻失败）；
  - 每个 bash 步骤的脚本必须通过 `bash -n` 语法检查（等价于 CI 解释器的
    第一道解析，能在本地复现 "unexpected EOF" 这类错误）；
  - 引号成对、续行符不落在脚本末尾（历史上出问题的正是这两种形态）。
"""

from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

import pytest

yaml = pytest.importorskip("yaml")

WORKFLOW = Path(__file__).parent / ".github" / "workflows" / "build.yml"


def _workflow() -> dict:
    assert WORKFLOW.exists(), f"缺少工作流文件 {WORKFLOW}"
    data = yaml.safe_load(WORKFLOW.read_text(encoding="utf-8"))
    assert isinstance(data, dict), "工作流顶层应是映射"
    return data


def _windows_matrix_entries() -> list[dict]:
    """返回 Windows 构建矩阵项，并约束其结构可被测试继续校验。"""
    matrix = _workflow()["jobs"]["build-windows"]["strategy"]["matrix"]["include"]
    assert isinstance(matrix, list) and matrix, "Windows 构建矩阵不能为空"
    assert all(isinstance(entry, dict) for entry in matrix), "Windows 构建矩阵项应是映射"
    return matrix


def _bash_steps() -> list[tuple[str, str, str]]:
    """收集所有 bash 步骤，返回 (job, step 名, 脚本) 列表"""
    steps: list[tuple[str, str, str]] = []
    for job, spec in _workflow()["jobs"].items():
        for step in spec.get("steps", []):
            script = step.get("run")
            if not script:
                continue
            shell = step.get("shell", "bash")
            if shell not in ("bash", None):
                continue          # pwsh 等其它 shell 另有语法，跳过
            steps.append((job, step.get("name", "<未命名>"), script))
    assert steps, "没有解析出任何 bash 步骤，工作流结构可能已变"
    return steps


def test_workflow_yaml_is_parseable() -> None:
    """YAML 必须能解析且包含预期的作业"""
    jobs = _workflow()["jobs"]
    assert {"test", "build-windows", "build-macos"} <= set(jobs), f"作业缺失：{list(jobs)}"
    assert "needs" in jobs["build-windows"], "构建作业应依赖测试门禁"


def test_windows_build_matrix_covers_x64_and_arm64() -> None:
    """Windows 必须同时构建 x64 与 ARM64 原生产物。"""
    entries = _windows_matrix_entries()
    by_arch = {entry["arch"]: entry for entry in entries}

    assert set(by_arch) == {"x86_64", "arm64"}, f"Windows 架构不完整：{set(by_arch)}"
    assert by_arch["x86_64"]["runner"] == "windows-latest"
    assert by_arch["x86_64"]["python_arch"] == "x64"
    assert by_arch["x86_64"]["expected_machine"] == "AMD64"
    assert by_arch["arm64"]["runner"] == "windows-11-arm"
    assert by_arch["arm64"]["python_arch"] == "arm64"
    assert by_arch["arm64"]["expected_machine"] == "ARM64"

    artifact_names = {entry["artifact_name"] for entry in entries}
    asset_names = {entry["asset_name"] for entry in entries}
    assert len(artifact_names) == len(entries), "Windows Artifact 名称必须按架构唯一"
    assert len(asset_names) == len(entries), "Windows 发布文件名必须按架构唯一"

    setup = next(
        step for step in _workflow()["jobs"]["build-windows"]["steps"]
        if step.get("uses", "").startswith("actions/setup-python@")
    )
    assert setup["with"]["architecture"] == "${{ matrix.python_arch }}", \
        "setup-python 必须显式使用矩阵架构，避免在 ARM runner 上误装 x64 Python"


def test_release_contains_windows_arm64_artifact() -> None:
    """ARM64 构建产物必须进入 GitHub Release，而不是只留在 Actions Artifacts。"""
    deploy_steps = _workflow()["jobs"]["deploy"]["steps"]
    download_names = {
        step["with"]["name"]
        for step in deploy_steps
        if step.get("uses", "").startswith("actions/download-artifact@")
    }
    assert "windows-arm64" in download_names, "发布作业未下载 Windows ARM64 产物"

    release_script = next(
        step["run"] for step in deploy_steps
        if step.get("run", "").startswith("gh release create")
    )
    assert "artifacts/windows-arm64/" in release_script, "Release 未附带 Windows ARM64 产物"
    assert "${{ env.APP_NAME }}-windows-arm64.exe" in release_script, \
        "Release 中的 Windows ARM64 文件缺少清晰架构标识"


@pytest.mark.parametrize("job,name,script", _bash_steps(),
                         ids=lambda v: v if isinstance(v, str) else "")
def test_bash_steps_are_syntactically_valid(job: str, name: str, script: str) -> None:
    """每个 bash 步骤都要通过 `bash -n`（解析但不执行）

    这正是漏写引号会失败的检查：CI 当时的报错
    `unexpected EOF while looking for matching '"'` 在本地就能被这一步抓到。
    """
    bash = shutil.which("bash")
    if not bash:
        pytest.skip("环境没有 bash，跳过脚本语法检查")
    proc = subprocess.run(
        [bash, "-n"], input=script, text=True, capture_output=True, check=False)
    assert proc.returncode == 0, (
        f"{job} / {name} 的 shell 脚本语法错误：\n{proc.stderr.strip()}\n--- 脚本 ---\n{script}")


def test_quotes_are_balanced_in_bash_steps() -> None:
    """shell 脚本里的双引号必须成对（漏写引号是最容易犯且最难看出的一类错）"""
    for job, name, script in _bash_steps():
        # 去掉 GitHub 表达式后再数，避免 ${{ }} 里的引号干扰
        import re

        code = re.sub(r"\$\{\{[^}]*\}\}", "", script)
        for lineno, line in enumerate(code.split("\n"), 1):
            if line.lstrip().startswith("#"):
                continue
            assert line.count('"') % 2 == 0, (
                f"{job} / {name} 第 {lineno} 行双引号不成对：{line!r}")


def test_no_dangling_line_continuation() -> None:
    """脚本不能以续行符 `\\` 结尾：会把最后一行与文件末尾粘在一起"""
    for job, name, script in _bash_steps():
        assert not script.rstrip("\n").endswith("\\"), \
            f"{job} / {name} 的脚本以续行符结尾，疑似被截断"


def test_qt_install_step_does_not_block_pipeline() -> None:
    """安装 Qt 运行库失败不应阻断流水线（把问题留给明确的测试失败）"""
    qt = [s for j, n, s in _bash_steps() if "Qt" in n]
    assert qt, "未找到安装 Qt 运行库的步骤"
    joined = "\n".join(qt)
    assert "|| echo" in joined or "|| true" in joined, \
        "apt-get install 失败应被容忍，否则网络抖动会直接让流水线挂掉"


def test_test_step_uses_offscreen_qt() -> None:
    """界面测试必须用 offscreen 后端，否则无显示器的 runner 上会失败"""
    for job, spec in _workflow()["jobs"].items():
        for step in spec.get("steps", []):
            if step.get("run") and "pytest" in step["run"]:
                env = step.get("env", {})
                assert env.get("QT_QPA_PLATFORM") == "offscreen", \
                    f"{job} 的测试步骤缺少 QT_QPA_PLATFORM=offscreen"
                return
    pytest.fail("工作流里没有找到运行 pytest 的步骤")


def test_brand_assets_are_bundled_into_builds() -> None:
    """品牌标识必须随 PyInstaller 打包，避免安装版顶部栏缺少 Logo。"""
    root = Path(__file__).parent
    assert (root / "assets" / "chengdu_technological_university_logo_mark.png").exists()
    assert (root / "assets" / "chengdu_technological_university_logo.jpg").exists()

    jobs = _workflow()["jobs"]
    windows_scripts = "\n".join(
        step.get("run", "") for step in jobs["build-windows"]["steps"])
    macos_scripts = "\n".join(
        step.get("run", "") for step in jobs["build-macos"]["steps"])
    assert '--add-data "assets;assets"' in windows_scripts
    assert '--add-data "assets:assets"' in macos_scripts
