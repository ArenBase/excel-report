"""版本一致性用例：防止 AI 改漏版本号（docs/03 开发方案 §3.2）。"""

from __future__ import annotations

import re
import tomllib
from pathlib import Path

import pytest

from excel_report import __version__

ROOT = Path(__file__).resolve().parent.parent


def test_version_is_semver() -> None:
    assert re.fullmatch(r"\d+\.\d+\.\d+", __version__), f"版本号不符合 MAJOR.MINOR.PATCH: {__version__}"


def test_pyproject_version_matches_package() -> None:
    data = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))

    assert data["project"]["version"] == __version__, (
        "pyproject.toml 的 version 与 __init__.py 的 __version__ 不一致"
    )


def test_changelog_mentions_current_version() -> None:
    changelog = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")

    assert __version__ in changelog, f"CHANGELOG.md 未记录当前版本 {__version__}"


def test_requires_python_is_at_least_310() -> None:
    """考核要求 Python >= 3.10；Ruff 也据此推断 target-version。"""
    data = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    requires = data["project"]["requires-python"]

    assert "3.10" in requires or "3.11" in requires or "3.12" in requires


@pytest.mark.parametrize("dependency", ["pandas", "openpyxl"])
def test_runtime_dependencies_are_within_allowed_list(dependency: str) -> None:
    """依赖收敛审查（C-09 / R-05）：运行时依赖必须在考核允许清单内。"""
    data = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    deps = " ".join(data["project"]["dependencies"])

    assert dependency in deps


def test_no_forbidden_runtime_dependencies() -> None:
    """不得引入 xlsxwriter / polars / pandera 等考题未列出的运行依赖。"""
    data = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    deps = " ".join(data["project"]["dependencies"]).lower()

    for forbidden in ("xlsxwriter", "polars", "pandera", "great-expectations", "fastapi"):
        assert forbidden not in deps, f"运行时依赖中不应出现 {forbidden}"


def test_source_imports_stay_within_allowed_libraries() -> None:
    """静态审查源码的第三方 import，确保没有越界依赖。"""
    allowed_third_party = {"pandas", "openpyxl", "numpy", "dateutil"}
    stdlib_ok = {
        "__future__",
        "argparse",
        "dataclasses",
        "datetime",
        "os",
        "pathlib",
        "re",
        "shutil",
        "sys",
        "tempfile",
        "traceback",
        "unicodedata",
        "warnings",
        "typing",
        "zipfile",
    }
    pattern = re.compile(r"^\s*(?:from|import)\s+([A-Za-z_][\w]*)", re.MULTILINE)

    src = ROOT / "src" / "excel_report"
    for path in sorted(src.glob("*.py")):
        found = set(pattern.findall(path.read_text(encoding="utf-8")))
        unexpected = found - allowed_third_party - stdlib_ok - {"excel_report"}
        assert not unexpected, f"{path.name} 引入了未批准的模块: {sorted(unexpected)}"
