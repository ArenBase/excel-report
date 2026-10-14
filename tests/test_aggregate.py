"""统计用例 TC-02 ~ TC-04、TC-09、TC-10、TC-23（docs/02-报表用例.md）。"""

from __future__ import annotations

import pandas as pd
import pytest

from conftest import (
    EXPECTED_BY_PERSON,
    EXPECTED_BY_REGION,
    EXPECTED_MONTHLY,
    make_raw,
    normal_raw,
)
from excel_report import aggregate
from excel_report.cleaning import SALES_AMOUNT_COLUMN, clean

BASE = {"日期": "2024-01-05", "销售员": "张三", "产品": "A", "数量": 2, "单价": 10.00, "地区": "华东"}


def _detail(records: list[dict[str, object]]) -> pd.DataFrame:
    detail, _ = clean(make_raw(records))
    return detail


# --- TC-01 汇总正确性（与人工手算一致） ---------------------------------------
def test_by_salesperson_matches_manual_calculation(normal_raw: pd.DataFrame) -> None:
    detail, _ = clean(normal_raw)
    result = aggregate.by_salesperson(detail)

    actual = dict(zip(result["销售员"], result["销售额"]))
    assert actual == EXPECTED_BY_PERSON
    assert dict(zip(result["销售员"], result["订单行数"])) == {"张三": 2, "李四": 1}


def test_by_region_matches_manual_calculation(normal_raw: pd.DataFrame) -> None:
    detail, _ = clean(normal_raw)
    result = aggregate.by_region(detail)

    actual = dict(zip(result["地区"], result["销售额"]))
    assert actual == EXPECTED_BY_REGION


def test_total_sales_matches_manual_calculation(normal_raw: pd.DataFrame) -> None:
    detail, _ = clean(normal_raw)

    assert aggregate.total_sales(detail) == 150.00


def test_monthly_totals_matches_manual_calculation(normal_raw: pd.DataFrame) -> None:
    detail, _ = clean(normal_raw)
    result = aggregate.monthly_totals(detail)

    actual = dict(zip(result["月份"], result["销售额"]))
    assert actual == EXPECTED_MONTHLY


# --- TC-02 口径：逐行相乘后汇总 != 汇总数量 × 汇总单价 ------------------------
def test_sales_amount_is_per_row_product() -> None:
    """3×10 + 2×20 = 70；错误算法（(3+2)×平均单价 或 5×10）会得到别的值。"""
    detail = _detail(
        [
            {**BASE, "数量": 3, "单价": 10.00, "产品": "A"},
            {**BASE, "数量": 2, "单价": 20.00, "产品": "B"},
        ]
    )
    result = aggregate.by_salesperson(detail)

    assert result[SALES_AMOUNT_COLUMN].iloc[0] == 70.00


# --- TC-03 销售额占比合计 = 100% ---------------------------------------------
def test_share_sums_to_100(normal_raw: pd.DataFrame) -> None:
    detail, _ = clean(normal_raw)

    for frame in (aggregate.by_salesperson(detail), aggregate.by_region(detail)):
        assert frame["销售额占比"].sum() == pytest.approx(100.00, abs=0.01)


# --- TC-04 排序确定性：销售额降序，并列按名称升序 -----------------------------
def test_sort_is_sales_desc_then_name_asc() -> None:
    """张三与李四并列 100.00 时，必须按名称升序。

    名称升序 = **Unicode 码点序**（本用例刻意先断言码点序，让期望值自证，而非靠记忆）：
    张 U+5F20 < 李 U+674E < 王 U+738B，因此期望顺序为 张三、李四、王五。
    """
    assert "张三" < "李四" < "王五", "前提校验：按码点序 张 < 李 < 王"

    detail = _detail(
        [
            {**BASE, "销售员": "李四", "数量": 1, "单价": 100.00, "产品": "A"},
            {**BASE, "销售员": "张三", "数量": 1, "单价": 100.00, "产品": "B"},
            {**BASE, "销售员": "王五", "数量": 1, "单价": 50.00, "产品": "C"},
        ]
    )
    result = aggregate.by_salesperson(detail)

    assert result["销售员"].tolist() == ["张三", "李四", "王五"], (
        "并列 100.00 的两项必须按名称升序（张 U+5F20 < 李 U+674E）"
    )
    assert result[SALES_AMOUNT_COLUMN].tolist() == [100.00, 100.00, 50.00], "销售额必须降序"


def test_tie_break_order_is_stable_and_documented() -> None:
    """锁定"并列按名称升序"的口径：必须是 Unicode 码点序，且跨两次运行稳定。

    这条用例存在的原因：开发过程中曾把期望值写成 ``["李四", "张三"]``（凭直觉认为
    "李"应排在"张"前面），被本用例抓出。正确的码点序是
    张 U+5F20 < 李 U+674E < 王 U+738B，因此期望值以 ``sorted()`` 自证，避免再次凭记忆写错。
    """
    names = ["王五", "张三", "李四"]

    assert sorted(names) == ["张三", "李四", "王五"], "前提校验：码点序"

    detail = _detail([{**BASE, "销售员": name, "数量": 1, "单价": 50.00} for name in names])
    first = aggregate.by_salesperson(detail)
    second = aggregate.by_salesperson(detail)

    assert first["销售员"].tolist() == sorted(names), "三者销售额并列，应按名称码点升序"
    assert first["销售员"].tolist() == second["销售员"].tolist(), "并列排序必须跨运行稳定"


def test_sort_is_deterministic_across_runs() -> None:
    """同一输入两次统计结果完全一致（C-08 确定性）。"""
    detail = _detail(
        [
            {**BASE, "销售员": "王五", "数量": 1, "单价": 100.00, "产品": "A"},
            {**BASE, "销售员": "李四", "数量": 1, "单价": 100.00, "产品": "B"},
            {**BASE, "销售员": "张三", "数量": 1, "单价": 100.00, "产品": "C"},
        ]
    )
    first = aggregate.by_salesperson(detail)
    second = aggregate.by_salesperson(detail)

    pd.testing.assert_frame_equal(first, second)


# --- TC-09 月度分组：同一月份写法归并 -----------------------------------------
def test_monthly_merges_same_month() -> None:
    detail = _detail(
        [
            {**BASE, "日期": "2024-03-01", "数量": 2, "单价": 100.0, "产品": "A"},
            {**BASE, "日期": "2024/3/15", "数量": 3, "单价": 50.0, "产品": "B"},
            {**BASE, "日期": "2024-03-31", "数量": 1, "单价": 200.0, "产品": "C"},
        ]
    )
    result = aggregate.monthly_totals(detail)

    assert result["月份"].tolist() == ["2024-03"]
    assert result[SALES_AMOUNT_COLUMN].iloc[0] == 550.00


# --- TC-10 跨年边界不合并 -----------------------------------------------------
def test_monthly_does_not_merge_across_year() -> None:
    detail = _detail(
        [
            {**BASE, "日期": "2023-12-31", "数量": 4, "单价": 25.0, "产品": "A"},
            {**BASE, "日期": "2024-01-01", "数量": 4, "单价": 25.0, "产品": "B"},
        ]
    )
    result = aggregate.monthly_totals(detail)

    assert result["月份"].tolist() == ["2023-12", "2024-01"]
    assert result[SALES_AMOUNT_COLUMN].tolist() == [100.00, 100.00]


# --- TC-23 总销售额为 0 时不除零崩溃 ------------------------------------------
def test_share_no_zero_division() -> None:
    """直接调用 with_share 传入销售额全为 0 的分组，断言返回 0.00 而非抛异常。"""
    grouped = pd.DataFrame({"销售员": ["张三", "李四"], "销售额": [0.0, 0.0]})

    result = aggregate.with_share(grouped, "销售员")

    assert result["销售额占比"].tolist() == [0.0, 0.0]


def test_share_empty_frame() -> None:
    grouped = pd.DataFrame({"销售员": [], "销售额": []})

    result = aggregate.with_share(grouped, "销售员")

    assert result.empty
    assert "销售额占比" in result.columns


def test_aggregates_on_empty_detail_return_empty_frames() -> None:
    empty = pd.DataFrame(columns=["日期", "销售员", "产品", "数量", "单价", "地区", SALES_AMOUNT_COLUMN])

    assert aggregate.by_salesperson(empty).empty
    assert aggregate.by_region(empty).empty
    assert aggregate.monthly_totals(empty).empty
    assert aggregate.total_sales(empty) == 0.0


# --- 金额精度：输出保留 2 位小数 ----------------------------------------------
def test_amounts_are_rounded_to_two_decimals() -> None:
    detail = _detail([{**BASE, "数量": 3, "单价": 33.333}])
    result = aggregate.by_salesperson(detail)

    assert detail[SALES_AMOUNT_COLUMN].iloc[0] == 100.00
    assert result[SALES_AMOUNT_COLUMN].iloc[0] == 100.00
    assert result["销售额占比"].iloc[0] == 100.00


# --- TC-27 三方对账 ----------------------------------------------------------
def test_three_way_reconciliation(normal_raw: pd.DataFrame) -> None:
    detail, _ = clean(normal_raw)

    detail_sum = round(float(detail[SALES_AMOUNT_COLUMN].sum()), 2)
    person_sum = round(float(aggregate.by_salesperson(detail)[SALES_AMOUNT_COLUMN].sum()), 2)
    region_sum = round(float(aggregate.by_region(detail)[SALES_AMOUNT_COLUMN].sum()), 2)

    assert detail_sum == person_sum == region_sum == aggregate.total_sales(detail)
