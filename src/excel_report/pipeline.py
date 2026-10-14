"""流程编排：读取 → 清洗 → 统计 → 写出。

本模块只负责"把各层串起来"，**不含任何业务规则**。

关于"异常时不生成部分报表"（决策 D-05）
----------------------------------------
这里**不靠 try/except 回滚**，而是靠**结构**保证：所有可能抛异常的步骤
（读文件、清洗、统计）都排在 :func:`excel_report.report.write_report` **之前**。
因此只要开始写文件，就说明四个工作表的数据已全部就绪，不会留下半个报表。
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from excel_report import aggregate, cleaning, loader, report
from excel_report.cleaning import CleaningStats
from excel_report.errors import NoValidDataError, SamePathError

__all__ = ["ReportResult", "run", "build_frames"]


@dataclass(frozen=True)
class ReportResult:
    """一次报表生成的完整结果，供 CLI 打印摘要与测试断言使用。"""

    input_path: Path
    output_path: Path
    detail: pd.DataFrame
    by_person: pd.DataFrame
    by_region: pd.DataFrame
    overview: pd.DataFrame
    monthly: pd.DataFrame
    stats: CleaningStats

    @property
    def total_sales(self) -> float:
        """总销售额。"""
        return aggregate.total_sales(self.detail)


def build_frames(
    raw: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, CleaningStats]:
    """由原始数据构造四个工作表的 DataFrame 与清洗统计（不触碰文件系统）。

    Returns:
        ``(detail, by_person, by_region, overview, monthly, stats)``
    """
    detail, stats = cleaning.clean(raw)

    if detail.empty:
        reasons = "；".join(
            f"{name}={value}" for name, value in stats.as_records() if name.startswith("剔除")
        )
        raise NoValidDataError(
            f"清洗后没有任何有效数据行（原始 {stats.raw_rows} 行全部被剔除）。\n       {reasons}",
            hint="请检查原始表：是否关键字段整列为空、数量是否全部 <= 0，或存在大量重复行。",
        )

    by_person = aggregate.by_salesperson(detail)
    by_region = aggregate.by_region(detail)
    monthly = aggregate.monthly_totals(detail)
    overview = aggregate.overview(detail, stats)
    return detail, by_person, by_region, overview, monthly, stats


def run(input_path: str | Path, output_path: str | Path) -> ReportResult:
    """执行完整流程并生成报表。

    Args:
        input_path: 原始 ``sales_raw.xlsx`` 路径。
        output_path: 输出 ``report.xlsx`` 路径。

    Returns:
        含四个工作表数据与统计信息的 :class:`ReportResult`。

    Raises:
        SamePathError: 输入输出指向同一文件（防覆盖原始数据，R-01）。
        ReportError: 其余各类业务异常（见 errors 模块）。
    """
    in_path = Path(input_path)
    out_path = Path(output_path)

    # R-01 硬防线：绝不允许把报表写到原始文件上。
    # 用 absolute() 而非 resolve()，这样输出文件尚不存在时也能比对同一路径。
    if in_path.absolute() == out_path.absolute():
        raise SamePathError(
            f"输出路径与输入路径相同: {in_path}\n       拒绝执行以免覆盖原始数据。",
            hint="请用 --output 指定一个不同的输出文件（如 output/report.xlsx）。",
        )
    if in_path.exists() and out_path.exists() and in_path.resolve() == out_path.resolve():
        raise SamePathError(
            f"输出路径与输入路径指向同一文件: {in_path.resolve()}\n       拒绝执行以免覆盖原始数据。",
            hint="请用 --output 指定一个不同的输出文件（如 output/report.xlsx）。",
        )

    raw = loader.load_raw(in_path)
    detail, by_person, by_region, overview, monthly, stats = build_frames(raw)

    # 走到这里说明全部数据已就绪，可以安全写文件
    written = report.write_report(
        out_path,
        detail=detail,
        by_person=by_person,
        by_region=by_region,
        overview_frame=overview,
    )

    return ReportResult(
        input_path=in_path,
        output_path=written,
        detail=detail,
        by_person=by_person,
        by_region=by_region,
        overview=overview,
        monthly=monthly,
        stats=stats,
    )
