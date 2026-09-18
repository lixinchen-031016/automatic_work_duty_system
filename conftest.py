"""pytest 公共配置

- 把仓库根加入 sys.path，测试可直接 import duty_system / app；
- Qt 一律用 offscreen 后端，保证 CI 与无显示器环境可运行界面测试；
- QSettings 重定向到临时目录，避免测试污染用户真实配置（注册表 / plist）。
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


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


@pytest.fixture
def settings(qt_app):
    """每个用例独立的空 QSettings（用完即清）"""
    from PySide6.QtCore import QSettings

    s = QSettings()
    yield s
    for key in s.allKeys():
        s.remove(key)
    s.sync()
