"""生成样例脏数据 sales_raw.xlsx。

按 docs/01-需求澄清清单.md §2.2 的脏数据清单构造，涵盖：
空值、全字段重复行、负数量/零数量、非法数值文本、非法日期、Excel 序列日期。

**本脚本是"造数据工具"，不属于报表脚本本体**；报告脚本本身绝不写 data/raw 以外的位置。

用法::

    python scripts/make_sample_data.py                       # 默认写入 data/raw/sales_raw.xlsx
    python scripts/make_sample_data.py --rows 500 --seed 7   # 指定规模与随机种子
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

import pandas as pd

if not __package__:
    _ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    _SRC = os.path.join(_ROOT, "src")
    if _SRC not in sys.path:
        sys.path.insert(0, _SRC)

from excel_report.loader import ALL_COLUMNS  # noqa: E402
from excel_report.report import freeze_workbook_timestamps  # noqa: E402

DEFAULT_OUTPUT = os.path.join("data", "raw", "sales_raw.xlsx")

SALESPERSONS = ["张三", "李四", "王五", "赵六", "钱七"]
PRODUCTS = ["无线耳机", "机械键盘", "显示器", "鼠标垫", "摄像头"]
REGIONS = ["华东", "华北", "华南", "西南"]


def build_sample(rows: int = 120, seed: int = 42) -> pd.DataFrame:
    """构造含脏数据的样例数据集。

    Args:
        rows: 期望的"干净"行数（实际总行数会因注入脏数据而更多）。
        seed: 随机种子，保证样例可复现。

    Returns:
        含脏数据的 DataFrame，列顺序为 ``ALL_COLUMNS``。
    """
    import random

    rng = random.Random(seed)

    records: list[dict[str, object]] = []
    for _ in range(rows):
        records.append(
            {
                "日期": f"2024-{rng.randint(1, 6):02d}-{rng.randint(1, 28):02d}",
                "销售员": rng.choice(SALESPERSONS),
                "产品": rng.choice(PRODUCTS),
                "数量": rng.randint(1, 20),
                "单价": round(rng.uniform(19.9, 899.0), 2),
                "地区": rng.choice(REGIONS),
            }
        )

    dirty: list[dict[str, object]] = []

    # 脏数据 1：关键字段为空（4 行）
    dirty.append({**records[0], "销售员": None})
    dirty.append({**records[1], "日期": None})
    dirty.append({**records[2], "数量": None})
    dirty.append({**records[3], "地区": "   "})  # 纯空格：pandas 的 dropna 识别不了

    # 脏数据 2：负数量 / 零数量（3 行）
    dirty.append({**records[4], "数量": -5})
    dirty.append({**records[5], "数量": 0})
    dirty.append({**records[6], "数量": -1})

    # 脏数据 3：非法数值文本（2 行）
    dirty.append({**records[7], "数量": "待定"})
    dirty.append({**records[8], "单价": "面议"})

    # 脏数据 4：非法日期（1 行）
    dirty.append({**records[9], "日期": "2024-13-45"})

    # 脏数据 5：Excel 序列日期（1 行，45292 == 2024-01-01）
    dirty.append({**records[10], "日期": 45292})

    # 脏数据 6：产品缺失（非关键字段，应被保留并填"未填写"）（2 行）
    dirty.append({**records[11], "产品": None})
    dirty.append({**records[12], "产品": ""})

    # 脏数据 7：全字段重复（把前 5 条干净记录各复制一份）（5 行）
    dirty.extend(dict(record) for record in records[:5])

    # 脏数据 8：日期写法多样性（同一月份三种写法）（3 行）
    dirty.append({**records[13], "日期": "2024/3/1", "数量": 2, "单价": 100.0})
    dirty.append({**records[14], "日期": "2024/3/15", "数量": 3, "单价": 50.0})
    dirty.append({**records[15], "日期": "2024-03-31", "数量": 1, "单价": 200.0})

    # 脏数据 9：跨年边界（2 行）
    dirty.append({**records[16], "日期": "2023-12-31", "数量": 4, "单价": 25.0})
    dirty.append({**records[17], "日期": "2024-01-01", "数量": 4, "单价": 25.0})

    # 脏数据 10：表头列名前后带空格（由参数控制，默认不加，供专门用例手工构造）
    all_records = records + dirty
    rng.shuffle(all_records)
    return pd.DataFrame(all_records, columns=ALL_COLUMNS)


def main(argv: list[str] | None = None) -> int:
    """写出样例文件并返回退出码。"""
    parser = argparse.ArgumentParser(description="生成样例脏数据 sales_raw.xlsx")
    parser.add_argument("-o", "--output", default=DEFAULT_OUTPUT, help="输出路径")
    parser.add_argument("--rows", type=int, default=120, help="干净行数（默认 120）")
    parser.add_argument("--seed", type=int, default=42, help="随机种子（默认 42）")
    args = parser.parse_args(argv)

    frame = build_sample(rows=args.rows, seed=args.seed)
    output = os.path.abspath(args.output)
    os.makedirs(os.path.dirname(output), exist_ok=True)
    frame.to_excel(output, index=False, engine="openpyxl")
    # 固定时间戳，使同一 --seed 生成的文件逐字节可复现
    freeze_workbook_timestamps(Path(output))

    print(f"[OK] 样例数据已生成: {output}")
    print(f"     总行数: {len(frame)}（其中包含人为注入的各类脏数据）")
    print(f"     列: {', '.join(frame.columns)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
