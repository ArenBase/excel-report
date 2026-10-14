"""清洗用例 TC-05 ~ TC-16（docs/02-报表用例.md §2）。

全部为纯函数测试：直接构造 DataFrame，不落盘。
"""

from __future__ import annotations

import pandas as pd
import pytest

from conftest import make_raw
from excel_report.cleaning import SALES_AMOUNT_COLUMN, clean

BASE = {"日期": "2024-01-05", "销售员": "张三", "产品": "A", "数量": 2, "单价": 10.00, "地区": "华东"}


def test_clean_normal_rows_are_kept() -> None:
    """正常行全部保留，并生成派生列 销售额 = 数量 × 单价。"""
    detail, stats = clean(make_raw([BASE]))

    assert len(detail) == 1
    assert stats.cleaned_rows == 1
    assert stats.dropped_total == 0
    assert detail[SALES_AMOUNT_COLUMN].iloc[0] == 20.00
    assert detail["日期"].iloc[0] == pd.Timestamp("2024-01-05")


# --- TC-05 全字段重复行只保留一条 ---------------------------------------------
def test_clean_drops_duplicate_rows_keeps_one() -> None:
    detail, stats = clean(make_raw([BASE, dict(BASE), dict(BASE)]))

    assert len(detail) == 1, "完全重复的 3 行应只保留 1 行"
    assert stats.dropped_duplicate == 2
    assert detail[SALES_AMOUNT_COLUMN].sum() == 20.00, "重复行不得让销售额翻倍"


def test_clean_different_region_is_not_duplicate() -> None:
    """不同地区的同额订单不得被误判为重复（验证"全字段"去重语义）。"""
    detail, stats = clean(make_raw([BASE, {**BASE, "地区": "华北"}]))

    assert len(detail) == 2
    assert stats.dropped_duplicate == 0


# --- TC-06 数量 = 0 与 数量 = -5 均被剔除 -------------------------------------
@pytest.mark.parametrize("bad_quantity", [0, -5, -1])
def test_clean_drops_nonpositive_quantity(bad_quantity: int) -> None:
    detail, stats = clean(make_raw([BASE, {**BASE, "数量": bad_quantity, "产品": "B"}]))

    assert len(detail) == 1
    assert stats.dropped_nonpositive == 1


# --- TC-07 边界：数量 = 1、单价 = 0.01（最小合法正值） ------------------------
def test_clean_boundary_min_positive_values_kept() -> None:
    detail, stats = clean(make_raw([{**BASE, "数量": 1, "单价": 0.01}]))

    assert len(detail) == 1, "数量=1 与 单价=0.01 都是 >0 的合法值，必须保留"
    assert stats.dropped_nonpositive == 0
    assert detail[SALES_AMOUNT_COLUMN].iloc[0] == 0.01


# --- TC-08 仅 1 行有效数据 ----------------------------------------------------
def test_clean_single_row() -> None:
    detail, stats = clean(make_raw([BASE]))

    assert stats.cleaned_rows == 1
    assert stats.raw_rows == 1


# --- TC-09 日期写法多样性（同一月份合并） -------------------------------------
def test_clean_parses_multiple_date_formats() -> None:
    records = [
        {**BASE, "日期": "2024-03-01", "产品": "A"},
        {**BASE, "日期": "2024/3/15", "产品": "B"},
        {**BASE, "日期": "2024-03-31", "产品": "C"},
    ]
    detail, _ = clean(make_raw(records))

    assert len(detail) == 3, "三种日期写法都应解析成功"
    months = detail["日期"].dt.to_period("M").astype(str).unique().tolist()
    assert months == ["2024-03"], "同一月份的写法必须归并到同一个月"


# --- TC-11 关键字段"空白字符串"视为缺失 ---------------------------------------
@pytest.mark.parametrize("column", ["日期", "销售员", "数量", "单价", "地区"])
def test_clean_blank_string_counts_as_missing(column: str) -> None:
    """⚠️ pandas 的 dropna 不识别 "   "，这里断言我们补上了该判定。"""
    detail, stats = clean(make_raw([{**BASE, column: "   "}]))

    assert len(detail) == 0, f"{column}='   ' 应被视为缺失并剔除"
    assert stats.dropped_missing == 1, "必须计入 dropped_missing（而非 dropped_invalid）"


def test_clean_none_counts_as_missing() -> None:
    detail, stats = clean(make_raw([{**BASE, "销售员": None}]))

    assert len(detail) == 0
    assert stats.dropped_missing == 1


# --- TC-12 列名含首尾空白自动归一化 -------------------------------------------
def test_normalize_columns_strips_whitespace() -> None:
    frame = pd.DataFrame(
        [
            {
                " 日期 ": "2024-01-05",
                "销售员": "张三",
                "产品": "A",
                "数量": 2,
                "单价": 10.0,
                " 地区 ": "华东",
            }
        ]
    )
    from excel_report.loader import normalize_columns

    normalized = normalize_columns(frame)
    detail, stats = clean(normalized)

    assert len(detail) == 1
    assert stats.cleaned_rows == 1


# --- TC-13 产品列缺失（非关键字段） -------------------------------------------
@pytest.mark.parametrize("empty_product", [None, "", "   "])
def test_clean_missing_product_is_filled_and_row_kept(empty_product: object) -> None:
    detail, stats = clean(make_raw([{**BASE, "产品": empty_product}]))

    assert len(detail) == 1, "产品不是关键字段，该行必须保留（销售额守恒）"
    assert detail["产品"].iloc[0] == "未填写"
    assert stats.product_filled == 1
    assert detail[SALES_AMOUNT_COLUMN].iloc[0] == 20.00


# --- TC-14 非法数值文本 -------------------------------------------------------
@pytest.mark.parametrize(
    ("column", "bad_value"),
    [("数量", "待定"), ("数量", "abc"), ("单价", "面议"), ("单价", "N/A")],
)
def test_clean_invalid_numeric_text_is_dropped(column: str, bad_value: str) -> None:
    detail, stats = clean(make_raw([{**BASE, column: bad_value}]))

    assert len(detail) == 0, "非法数值行应被剔除"
    assert stats.dropped_invalid == 1, "必须计入 dropped_invalid（与缺失区分开）"


# --- TC-15 非法日期 -----------------------------------------------------------
@pytest.mark.parametrize("bad_date", ["2024-13-45", "not-a-date", "2024/02/30"])
def test_clean_invalid_date_is_dropped(bad_date: str) -> None:
    detail, stats = clean(make_raw([{**BASE, "日期": bad_date}]))

    assert len(detail) == 0
    assert stats.dropped_invalid == 1


# --- TC-16 Excel 序列日期被正确解析 -------------------------------------------
@pytest.mark.parametrize(
    ("serial", "expected"),
    [(45292, "2024-01-01"), (45658, "2025-01-01"), (1, "1900-01-01"), (61, "1900-03-01")],
)
def test_clean_excel_serial_date(serial: int, expected: str) -> None:
    detail, stats = clean(make_raw([{**BASE, "日期": serial}]))

    assert stats.dropped_invalid == 0, "Excel 序列日期是合法日期，不应被剔除"
    assert len(detail) == 1
    assert detail["日期"].iloc[0] == pd.Timestamp(expected)


def test_excel_serial_matches_openpyxl_reference() -> None:
    """与 openpyxl 官方实现逐点比对，确保 1900 闰年缺陷处理一致。"""
    from openpyxl.utils.datetime import from_excel

    from excel_report.cleaning import to_excel_serial_datetime

    mismatches = [
        serial
        for serial in range(1, 73416)
        if from_excel(serial) != to_excel_serial_datetime(serial).to_pydatetime()
    ]
    assert mismatches == [], f"与 openpyxl 不一致的序列号: {mismatches[:10]}"


# --- TC-26 清洗守恒不变量 -----------------------------------------------------
def test_clean_conservation_invariant() -> None:
    """原始行数 == 清洗后行数 + 各剔除原因之和（一次性抓住"某行被静默吞掉"）。"""
    records = [
        BASE,  # 合法
        {**BASE, "产品": "B", "数量": 3},  # 合法
        dict(BASE),  # 重复
        {**BASE, "销售员": None, "产品": "C"},  # 关键字段缺失
        {**BASE, "数量": 0, "产品": "D"},  # 数量<=0
        {**BASE, "数量": "待定", "产品": "E"},  # 非法数值
        {**BASE, "日期": "2024-13-45", "产品": "F"},  # 非法日期
        {**BASE, "地区": "   ", "产品": "G"},  # 空白串缺失
    ]
    detail, stats = clean(make_raw(records))

    assert stats.raw_rows == len(records)
    assert stats.raw_rows == stats.cleaned_rows + stats.dropped_total
    assert stats.dropped_total == 6
    assert stats.cleaned_rows == 2
