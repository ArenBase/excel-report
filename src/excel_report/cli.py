"""命令行入口：参数解析 + 摘要打印 + 异常到退出码的映射。

**本模块是本工程唯一处理退出码的地方。**

退出码约定见 :mod:`excel_report.errors`。库层只抛异常、不打印；
CLI 层把异常转译为"中文可读提示 + 退出码"，因此用户永远不会看到裸 traceback。
"""

from __future__ import annotations

import argparse
import sys
import traceback
from pathlib import Path

from excel_report import __version__
from excel_report.errors import ReportError
from excel_report.pipeline import ReportResult, run

__all__ = ["build_parser", "main", "format_summary"]

DEFAULT_INPUT = "data/raw/sales_raw.xlsx"
DEFAULT_OUTPUT = "output/report.xlsx"


def build_parser() -> argparse.ArgumentParser:
    """构造命令行参数解析器。"""
    parser = argparse.ArgumentParser(
        prog="excel_report",
        description="读取 sales_raw.xlsx，清洗统计后生成含四个工作表的 report.xlsx。",
        epilog=(
            "退出码: 0 成功 | 1 未预期异常 | 2 文件不存在 | 3 无有效数据 "
            "| 4 缺少必需列 | 5 非有效 Excel | 6 输入输出路径相同"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "-i",
        "--input",
        default=DEFAULT_INPUT,
        help=f"原始销售表路径（默认: {DEFAULT_INPUT}）",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=DEFAULT_OUTPUT,
        help=f"报表输出路径（默认: {DEFAULT_OUTPUT}）",
    )
    parser.add_argument(
        "-q",
        "--quiet",
        action="store_true",
        help="静默模式：仅在出错时输出",
    )
    parser.add_argument(
        "-V",
        "--version",
        action="version",
        version=f"excel_report {__version__}",
    )
    return parser


def format_summary(result: ReportResult) -> str:
    """渲染运行摘要（格式见 docs/01-需求澄清清单.md §3.3）。"""
    stats = result.stats
    lines = [
        f"[OK] 读取: {result.input_path}  (原始 {stats.raw_rows} 行)",
        f"[OK] 清洗后: {stats.cleaned_rows} 行  (剔除 {stats.dropped_total} 行)",
        f"     - 关键字段缺失 : {stats.dropped_missing} 行",
        f"     - 数量<=0      : {stats.dropped_nonpositive} 行",
        f"     - 全字段重复   : {stats.dropped_duplicate} 行",
        f"     - 非法数值/日期: {stats.dropped_invalid} 行",
    ]
    if stats.product_filled:
        lines.append(f"     - 产品填充'未填写': {stats.product_filled} 行")
    lines.append(f"[OK] 总销售额: {result.total_sales:,.2f}")
    if not result.monthly.empty:
        first = result.monthly["月份"].iloc[0]
        last = result.monthly["月份"].iloc[-1]
        lines.append(f"[OK] 月度区间: {first} ~ {last}  (共 {len(result.monthly)} 个月)")
    lines.append(f"[OK] 报表已生成: {result.output_path} (4 个工作表)")
    lines.append(
        "     工作表: " + " / ".join(["明细清洗后", "按销售员", "按地区", "总览"])
    )
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    """程序主入口。

    Args:
        argv: 命令行参数列表；``None`` 表示取 ``sys.argv[1:]``。

    Returns:
        进程退出码。
    """
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        result = run(Path(args.input), Path(args.output))
    except ReportError as exc:
        print(exc.format_for_user(), file=sys.stderr)
        return exc.exit_code
    except KeyboardInterrupt:
        print("\n[中断] 用户取消了本次运行。", file=sys.stderr)
        return 130
    except Exception as exc:  # noqa: BLE001 - 兜底：绝不向用户抛裸 traceback
        print(
            f"[错误] 未预期的异常: {type(exc).__name__}: {exc}\n"
            "       这是本工具的缺陷，请保留以下调用栈以便排查。",
            file=sys.stderr,
        )
        traceback.print_exc(file=sys.stderr)
        return 1

    if not args.quiet:
        print(format_summary(result))
    return 0
