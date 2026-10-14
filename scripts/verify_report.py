"""核对生成的 report.xlsx 与清洗/统计口径（人工验收辅助脚本）。

用法::

    python scripts/verify_report.py --report output/report.xlsx --raw data/raw/sales_raw.xlsx
"""

from __future__ import annotations

import argparse
import os
import sys

import pandas as pd

if not __package__:
    _ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    _SRC = os.path.join(_ROOT, "src")
    if _SRC not in sys.path:
        sys.path.insert(0, _SRC)

from excel_report import aggregate, loader  # noqa: E402
from excel_report.report import SHEET_ORDER  # noqa: E402

MONEY = "{:,.2f}"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="核对 report.xlsx 的四表与统计口径")
    parser.add_argument("--report", default=os.path.join("output", "report.xlsx"))
    parser.add_argument("--raw", default=os.path.join("data", "raw", "sales_raw.xlsx"))
    args = parser.parse_args(argv)

    book = pd.ExcelFile(args.report, engine="openpyxl")
    print("=== 工作表清单 ===")
    print("实际:", book.sheet_names)
    print("期望:", list(SHEET_ORDER))
    assert book.sheet_names == list(SHEET_ORDER), "工作表名称或顺序不符"

    detail = book.parse("明细清洗后")
    by_person = book.parse("按销售员")
    by_region = book.parse("按地区")
    overview = book.parse("总览")

    print("\n=== 明细清洗后（前 5 行） ===")
    print(detail.head().to_string(index=False))

    print("\n=== 按销售员 ===")
    print(by_person.to_string(index=False))

    print("\n=== 按地区 ===")
    print(by_region.to_string(index=False))

    print("\n=== 总览 ===")
    # 空单元格读回后是 NaN，仅影响打印观感，显示为空白
    print(overview.fillna("").to_string(index=False))

    print("\n=== 三方对账（应完全相等） ===")
    d = round(float(detail["销售额"].sum()), 2)
    p = round(float(by_person["销售额"].sum()), 2)
    r = round(float(by_region["销售额"].sum()), 2)
    print(f"明细合计      = {MONEY.format(d)}")
    print(f"按销售员合计  = {MONEY.format(p)}")
    print(f"按地区合计    = {MONEY.format(r)}")
    assert d == p == r, "三方合计不一致"
    print("对账通过 ✅")

    print("\n=== 占比合计（应为 100.00） ===")
    print(f"按销售员占比合计 = {by_person['销售额占比'].sum():.2f}")
    print(f"按地区占比合计   = {by_region['销售额占比'].sum():.2f}")

    print("\n=== 原始文件只读校验（SHA256） ===")
    import hashlib

    def sha256(path: str) -> str:
        digest = hashlib.sha256()
        with open(path, "rb") as handle:
            for chunk in iter(lambda: handle.read(65536), b""):
                digest.update(chunk)
        return digest.hexdigest()

    print(f"原始文件 {args.raw}\n  SHA256 = {sha256(args.raw)}")

    print("\n=== 独立复算（不依赖生成的报表） ===")
    raw = loader.load_raw(args.raw)
    from excel_report import cleaning

    recomputed, stats = cleaning.clean(raw)
    print(f"独立复算总销售额 = {MONEY.format(aggregate.total_sales(recomputed))}")
    assert round(aggregate.total_sales(recomputed), 2) == d, "独立复算与报表不一致"
    print(f"独立复算清洗行数 = {stats.cleaned_rows}")
    print("复算通过 ✅")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
