"""mkdocs-macros 钩子：从 changelog 自动读取最新协议版本号。

文档中使用 `{{ protocol_version }}` 占位即可自动同步。唯一来源是
`docs/appendix/changelog.md` 顶部最新版本条目（`## [X.Y.Z]` 格式）。
"""

from __future__ import annotations

import re
from pathlib import Path

_PROJECT_ROOT = Path(__file__).resolve().parents[2]
_CHANGELOG = _PROJECT_ROOT / "docs" / "appendix" / "changelog.md"
_VERSION_RE = re.compile(r"^##\s*\[(\d+\.\d+\.\d+(?:[\w.-]+)?)\]", re.MULTILINE)


def _read_latest_version() -> str:
    if not _CHANGELOG.is_file():
        return "unknown"
    match = _VERSION_RE.search(_CHANGELOG.read_text(encoding="utf-8"))
    return match.group(1) if match else "unknown"


def define_env(env) -> None:  # type: ignore[no-untyped-def]
    """mkdocs-macros 插件入口，在构建开始时注册变量。"""
    env.variables["protocol_version"] = _read_latest_version()
