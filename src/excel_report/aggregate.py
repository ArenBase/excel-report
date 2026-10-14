"""统计汇总（纯函数，不触碰文件系统）。

统计口径（docs/01-需求澄清清单.md §3.2）::

    销售额       = 数量 × 单价，逐行计算后再汇总
    金额精度     中间不四舍五入，输出保留 2 位小数
    销售额占比   该组销售额 ÷ 清洗后总销售额 × 100%，保留 2 位小数；总额为 0 时记 0.00（不除零）
    排序         销售额降序；并列时按名称升序（保证结果确定性，便于人工核对与测试断言）
    订单行数     清洗后明细中该组的行数（1 行 = 1 笔销售记录）
"""

from __future__ import annotations

import pandas as pd

from excel_report.cleaning import SALES_AMOUNT_COLUMN

__all__ = [
    "by_salesperson",
    "by_region",
    "monthly_totals",
    "overview",
    "total_sales",
    "with_share",
]


def total_sales(detail: pd.DataFrame) -> float:
    """清洗后明细的总销售额（保留 2 位小数）。"""
    if detail.empty:
        return 0.0
    return round(float(detail[SALES_AMOUNT_COLUMN].sum()), 2)


def with_share(grouped: pd.DataFrame, key_column: str) -> pd.DataFrame:
    """为分组结果补充 ``销售额占比`` 列。

    除零保护（用例 TC-23）：总销售额为 0 时占比一律记 ``0.00``，
    而不是抛 :class:`ZeroDivisionError`。

    Args:
        grouped: 含 ``key_column`` 与 ``销售额`` 两列的分组结果。
        key_column: 分组键列名。

    Returns:
        追加了 ``销售额占比`` 的新 DataFrame。
    """
    result = grouped.copy()
    total = float(result[SALES_AMOUNT_COLUMN].sum())
    if total:
        share = (result[SALES_AMOUNT_COLUMN] / total * 100).round(2)
    else:
        share = pd.Series(0.0, index=result.index)
    result["销售额占比"] = share
    return result


def _group_by(detail: pd.DataFrame, key_column: str) -> pd.DataFrame:
    """按单列分组汇总，并按 (销售额降序, 名称升序) 排序。"""
    columns = [key_column, "订单行数", "总数量", SALES_AMOUNT_COLUMN]
    if detail.empty:
        return pd.DataFrame(columns=columns + ["销售额占比"])

    grouped = (
        detail.groupby(key_column, dropna=False, observed=True)
        .agg(
            订单行数=(SALES_AMOUNT_COLUMN, "size"),
            总数量=("数量", "sum"),
            **{SALES_AMOUNT_COLUMN: (SALES_AMOUNT_COLUMN, "sum")},
        )
        .reset_index()
    )
    grouped[SALES_AMOUNT_COLUMN] = grouped[SALES_AMOUNT_COLUMN].round(2)
    grouped["总数量"] = grouped["总数量"].astype("Int64")
    # 排序契约：销售额降序，并列时按名称升序（Unicode 码点序）—— 保证输出确定性（C-08 / TC-04）
    grouped = grouped.sort_values(
        by=[SALES_AMOUNT_COLUMN, key_column], ascending=[False, True], kind="stable"
    ).reset_index(drop=True)
    return with_share(grouped, key_column)


def by_salesperson(detail: pd.DataFrame) -> pd.DataFrame:
    """按销售员汇总销售额（用例 TC-01 / TC-02）。"""
    return _group_by(detail, "销售员")


def by_region(detail: pd.DataFrame) -> pd.DataFrame:
    """按地区汇总销售额（用例 TC-01）。"""
    return _group_by(detail, "地区")


def monthly_totals(detail: pd.DataFrame) -> pd.DataFrame:
    """按"年-月"汇总销售额（用例 TC-09 / TC-10）。

    Returns:
        含 ``月份``（``YYYY-MM`` 字符串）、``订单行数``、``销售额`` 的 DataFrame，按月份升序。
    """
    columns = ["月份", "订单行数", SALES_AMOUNT_COLUMN]
    if detail.empty:
        return pd.DataFrame(columns=columns)

    work = detail.copy()
    # 直接用 dt.to_period("M") 转字符串，避免 strftime 在不同平台上对 1970 前的行为差异
    work["月份"] = work["日期"].dt.to_period("M").astype(str)

    monthly = (
        work.groupby("月份", observed=True)
        .agg(
            订单行数=(SALES_AMOUNT_COLUMN, "size"),
            **{SALES_AMOUNT_COLUMN: (SALES_AMOUNT_COLUMN, "sum")},
        )
        .reset_index()
    )
    monthly[SALES_AMOUNT_COLUMN] = monthly[SALES_AMOUNT_COLUMN].round(2)
    return monthly.sort_values("月份", kind="stable").reset_index(drop=True).loc[:, columns]


def overview(detail: pd.DataFrame, stats: object) -> pd.DataFrame:
    """构造 "总览" 工作表：全局 KPI + 数据质量报告 + 月度总额。

    Args:
        detail: 清洗后的明细。
        stats: :class:`excel_report.cleaning.CleaningStats` 实例。

    Returns:
        含 ``指标`` / ``数值`` 两列的 DataFrame。
    """
    rows: list[tuple[str, object]] = [
        ("统计口径·销售额", "数量 × 单价（逐行计算后汇总）"),
        ("统计口径·金额精度", "保留 2 位小数"),
    ]
    rows.extend(stats.as_records())  # type: ignore[attr-defined]

    rows.append(("总销售额", total_sales(detail)))

    monthly = monthly_totals(detail)
    if monthly.empty:
        rows.append(("月度总额", "(无数据)"))
        rows.append(("月份数", 0))
    else:
        rows.append(("月份数", int(len(monthly))))
        # 最早 / 最晚月份，便于人工核对时间范围
        rows.append(("数据起始月份", str(monthly["月份"].iloc[0])))
        rows.append(("数据结束月份", str(monthly["月份"].iloc[-1])))

    result = pd.DataFrame(rows, columns=["指标", "数值"])

    # 月度明细追加在下方（"月份" 列留空以保持两列结构）
    if not monthly.empty:
        spacer = pd.DataFrame([{"指标": "", "数值": ""}])
        header = pd.DataFrame([{"指标": "── 月度销售额明细 ──", "数值": ""}])
        monthly_block = pd.DataFrame(
            {
                "指标": monthly["月份"],
                "数值": monthly[SALES_AMOUNT_COLUMN],
            }
        )
        result = pd.concat([result, spacer, header, monthly_block], ignore_index=True)

    # 结构化空白行（间隔行、区块标题行）会产生 NaN，输出时应显示为空单元格而非 "NaN"
    return result.fillna("")
