"""读取原始 Excel 并做表头/列校验。

**本模块是本工程唯一读取源文件的地方，且全程只读，不含任何写入 API。**

读取检查顺序（顺序本身即规格，见 docs/01-需求澄清清单.md §4）::

    1. 路径是否存在            -> InputFileNotFoundError (退出码 2)
    2. 是否能作为 Excel 打开   -> InvalidWorkbookError  (退出码 5)
    3. 选表：第一个工作表；若为空则回退到第一个非空表
    4. 表是否完全无数据行      -> EmptyDataError        (退出码 3)
    5. 列名 strip 归一化       (容忍 " 数量 " 这类表头)
    6. 必需列是否齐全          -> MissingColumnsError   (退出码 4)
"""

from __future__ import annotations

import warnings
from pathlib import Path

import pandas as pd

from excel_report.errors import (
    EmptyDataError,
    InputFileNotFoundError,
    InvalidWorkbookError,
    MissingColumnsError,
)

__all__ = ["REQUIRED_COLUMNS", "KEY_COLUMNS", "ALL_COLUMNS", "load_raw", "normalize_columns"]

#: 原始表的全部业务列（顺序即报表中的列顺序）
ALL_COLUMNS: list[str] = ["日期", "销售员", "产品", "数量", "单价", "地区"]

#: 必需列：缺任何一列都无法计算统计口径（C-04）
REQUIRED_COLUMNS: list[str] = ["日期", "销售员", "数量", "单价", "地区"]

#: 关键字段（D-01 决策）：任一缺失则该行被删除；"产品" 不属于关键字段
KEY_COLUMNS: list[str] = ["日期", "销售员", "数量", "单价", "地区"]

#: "产品" 缺失时的占位值（D-02 决策：保留该行，保证销售额守恒）
PRODUCT_FILLNA = "未填写"


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """把列名的首尾空白去掉（C-05）。

    pandas 不会自动做这件事，表头写成 ``" 数量 "`` 时 ``df["数量"]`` 会 KeyError。
    仅做 ``strip``，不做模糊匹配——模糊匹配会掩盖真实的缺列错误（见调研报告 §3）。
    """
    renamed = {col: (col.strip() if isinstance(col, str) else col) for col in df.columns}
    # 归一化后可能出现重复列名，保留第一个即可
    df = df.rename(columns=renamed)
    if df.columns.duplicated().any():
        dupes = sorted({str(c) for c in df.columns[df.columns.duplicated()]})
        warnings.warn(f"表头归一化后出现重复列名，已保留第一列: {', '.join(dupes)}", stacklevel=2)
        df = df.loc[:, ~df.columns.duplicated()]
    return df


def _pick_sheet(book: pd.ExcelFile) -> tuple[str, pd.DataFrame]:
    """选取工作表：优先第一个；若第一个为空则取第一个非空表。"""
    names = list(book.sheet_names)
    if not names:
        raise EmptyDataError("工作簿中没有任何工作表。")

    first = names[0]
    df = book.parse(first)
    if not df.empty:
        return first, df

    for name in names[1:]:
        candidate = book.parse(name)
        if not candidate.empty:
            warnings.warn(
                f"第一个工作表 '{first}' 为空，已回退使用第一个非空工作表 '{name}'。",
                stacklevel=2,
            )
            return name, candidate

    raise EmptyDataError(
        f"文件中共有 {len(names)} 个工作表，但没有任何数据行（已检查: {', '.join(names)}）。",
        hint="请确认导出的原始表确实包含数据行。",
    )


def load_raw(path: str | Path) -> pd.DataFrame:
    """只读加载原始销售表，返回已完成列名校验的 DataFrame。

    Args:
        path: 原始 ``.xlsx`` 文件路径。

    Returns:
        原始数据的 DataFrame（**未做任何清洗**，保留脏数据以便统计剔除数量）。

    Raises:
        InputFileNotFoundError: 文件不存在。
        InvalidWorkbookError: 文件不是有效的 Excel 工作簿（损坏、0 字节、实为 CSV 等）。
        EmptyDataError: 有工作表但没有任何数据行。
        MissingColumnsError: 缺少必需列。
    """
    file_path = Path(path)

    if not file_path.exists():
        raise InputFileNotFoundError(
            f"输入文件不存在: {file_path}",
            hint="请用 --input 指定正确的 sales_raw.xlsx 路径。",
        )
    if file_path.is_dir():
        raise InputFileNotFoundError(
            f"输入路径是一个目录，不是文件: {file_path}",
            hint="请用 --input 指定具体的 .xlsx 文件。",
        )

    try:
        book = pd.ExcelFile(file_path, engine="openpyxl")
    except Exception as exc:  # noqa: BLE001 - 统一转译为领域异常，避免裸 traceback
        raise InvalidWorkbookError(
            f"无法作为 Excel 文件打开: {file_path}\n       底层错误: {type(exc).__name__}: {exc}",
            hint="请确认文件是完整的 .xlsx 工作簿（0 字节文件或改名后的 CSV 都会触发此错误）。",
        ) from exc

    with book:
        sheet_name, df = _pick_sheet(book)

    df = normalize_columns(df)

    actual = [str(c) for c in df.columns]
    missing = [c for c in REQUIRED_COLUMNS if c not in actual]
    if missing:
        raise MissingColumnsError(missing, actual)

    # 统一保留原始业务列（原始表若多出无关列则忽略，避免影响去重语义）
    keep = [c for c in ALL_COLUMNS if c in df.columns]
    return df.loc[:, keep].copy()
