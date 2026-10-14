"""pytest 共享夹具与工具函数。

测试策略（docs/03-开发方案与技术选型.md §4）：
- 单元测试直接构造 DataFrame，**不落盘**（快、稳定）；
- 集成测试用 ``tmp_path`` 写真实 xlsx 再读回；
- 断言必须是**真断言**：比对具体数值，而不是只断言"没抛异常"。
"""

from __future__ import annotations

import hashlib
import sys
from pathlib import Path
from typing import Any

import pandas as pd
import pytest

# src layout：让测试在未安装包的情况下也能 import excel_report
_ROOT = Path(__file__).resolve().parent.parent
_SRC = _ROOT / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from excel_report.loader import ALL_COLUMNS  # noqa: E402


def make_raw(records: list[dict[str, Any]]) -> pd.DataFrame:
    """由 dict 列表构造"原始表"DataFrame（列顺序固定，缺失列补 NA）。"""
    frame = pd.DataFrame(records)
    for column in ALL_COLUMNS:
        if column not in frame.columns:
            frame[column] = pd.NA
    return frame.loc[:, ALL_COLUMNS]


def write_xlsx(path: Path, records: list[dict[str, Any]], *, columns: list[str] | None = None) -> Path:
    """把记录写成真实 xlsx 文件（集成测试用）。"""
    frame = pd.DataFrame(records)
    if columns is not None:
        frame = frame.loc[:, columns]
    path.parent.mkdir(parents=True, exist_ok=True)
    frame.to_excel(path, index=False, engine="openpyxl")
    return path


#: 用例 TC-01 的标准 3 行数据：便于手算核对
NORMAL_RECORDS: list[dict[str, Any]] = [
    {"日期": "2024-01-05", "销售员": "张三", "产品": "A", "数量": 2, "单价": 10.00, "地区": "华东"},
    {"日期": "2024-01-20", "销售员": "张三", "产品": "B", "数量": 1, "单价": 30.00, "地区": "华北"},
    {"日期": "2024-02-03", "销售员": "李四", "产品": "C", "数量": 4, "单价": 25.00, "地区": "华东"},
]

#: TC-01 的期望值（人工手算）
EXPECTED_TOTAL = 150.00
EXPECTED_BY_PERSON = {"张三": 50.00, "李四": 100.00}
EXPECTED_BY_REGION = {"华东": 120.00, "华北": 30.00}
EXPECTED_MONTHLY = {"2024-01": 50.00, "2024-02": 100.00}


@pytest.fixture
def normal_raw() -> pd.DataFrame:
    """TC-01 的原始 DataFrame。"""
    return make_raw(NORMAL_RECORDS)


@pytest.fixture
def normal_xlsx(tmp_path: Path) -> Path:
    """TC-01 的原始 xlsx 文件。"""
    return write_xlsx(tmp_path / "sales_raw.xlsx", NORMAL_RECORDS)


@pytest.fixture(autouse=True)
def _isolate_cwd(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """所有测试在临时工作目录下运行，避免误写工程目录。"""
    monkeypatch.chdir(tmp_path)


def sha256_of(path: Path) -> str:
    """计算文件 SHA256（用于只读/幂等断言）。"""
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()
