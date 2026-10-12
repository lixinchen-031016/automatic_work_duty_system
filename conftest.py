"""pytest 公共配置

- 把仓库根加入 sys.path，测试可直接 import duty_system / app；
- Qt 一律用 offscreen 后端，保证 CI 与无显示器环境可运行界面测试；
- QSettings 重定向到临时目录，避免测试污染用户真实配置（注册表 / plist）。
"""

from __future__ import annotations

import gc
import os
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

# 依赖 QApplication / QWidget 的 fixture：用到它们的用例会被自动标记为 ui
_QT_FIXTURES = {"qt_app", "window", "settings", "editor_factory"}


def pytest_collection_modifyitems(items) -> None:
    """给所有依赖 Qt 的用例自动打上 `ui` 标记。

    - `pytest -m ui`：只跑界面用例；
    - `pytest -m "not ui"`：只跑后端用例。

    CI 门禁使用 `-m "not ui"` 跳过无头界面测试：部分 Linux + offscreen 环境会在
    用例全部通过后的解释器退出阶段触发进程级 SIGABRT（exit 134），且没有 Python
    调用栈可供定位。自建 QApplication（不使用上述 fixture）的用例需显式写
    `@pytest.mark.ui`，以便同样被覆盖。
    """
    ui_marker = pytest.mark.ui
    for item in items:
        if _QT_FIXTURES & set(getattr(item, "fixturenames", ())):
            item.add_marker(ui_marker)


@pytest.fixture(scope="session")
def qt_app(tmp_path_factory):
    """会话级 QApplication，并把 QSettings 隔离到临时目录"""
    QtCore = pytest.importorskip("PySide6.QtCore")
    QtWidgets = pytest.importorskip("PySide6.QtWidgets")

    settings_dir = tmp_path_factory.mktemp("qsettings")
    fmt = QtCore.QSettings.IniFormat
    QtCore.QSettings.setDefaultFormat(fmt)
    QtCore.QSettings.setPath(fmt, QtCore.QSettings.UserScope, str(settings_dir))
    QtCore.QSettings.setPath(fmt, QtCore.QSettings.SystemScope, str(settings_dir))

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    yield app
    # 会话结束前先回收所有残留的顶层窗口，再让 QApplication 被回收：
    # 若 Qt 对象留到解释器退出阶段才析构（此时 QApplication 已销毁），
    # 部分平台（如 Linux + offscreen）会直接 abort 导致整个测试进程失败。
    for widget in app.topLevelWidgets():
        try:
            widget.close()
            widget.setParent(None)
            widget.deleteLater()
        except RuntimeError:
            # 包装器对应的 C++ 对象已被 Qt 删除，跳过即可
            continue
    app.processEvents()
    # 主动回收引用环（对话框 <-> 信号 <-> 绑定方法），确保 Qt 对象在
    # QApplication 仍然存活时被销毁，而不是拖到解释器退出阶段
    gc.collect()


@pytest.fixture
def settings(qt_app):
    """每个用例独立的空 QSettings（用完即清）"""
    from PySide6.QtCore import QSettings

    s = QSettings()
    yield s
    for key in s.allKeys():
        s.remove(key)
    s.sync()
