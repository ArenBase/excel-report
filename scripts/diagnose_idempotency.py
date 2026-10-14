"""诊断脚本：量化"幂等性"到底差在哪里。

用途：区分"内容不确定"（真 bug）与"文件元数据/容器不确定"（openpyxl 写入时间戳导致）。
"""

from __future__ import annotations

import os
import sys
import zipfile

import pandas as pd

if not __package__:
    _ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    _SRC = os.path.join(_ROOT, "src")
    if _SRC not in sys.path:
        sys.path.insert(0, _SRC)

from excel_report.report import SHEET_ORDER  # noqa: E402


def zip_member_hashes(path: str) -> dict[str, str]:
    """列出 xlsx（zip）内每个成员的哈希，用于定位差异来源。"""
    import hashlib

    result = {}
    with zipfile.ZipFile(path) as archive:
        for info in sorted(archive.infolist(), key=lambda i: i.filename):
            digest = hashlib.sha256(archive.read(info.filename)).hexdigest()
            result[info.filename] = digest
    return result


def main() -> int:
    first, second = sys.argv[1], sys.argv[2]

    print("=== 1. 单元格内容是否完全一致？ ===")
    all_same = True
    for sheet in SHEET_ORDER:
        a = pd.read_excel(first, sheet_name=sheet, engine="openpyxl")
        b = pd.read_excel(second, sheet_name=sheet, engine="openpyxl")
        same = a.equals(b)
        all_same &= same
        print(f"  {sheet:12} 内容一致={same}  行数 {len(a)}/{len(b)}")
    print(f"  => 内容层面的幂等性: {'通过' if all_same else '不通过'}")

    print("\n=== 2. zip 成员层面哪些不同？ ===")
    ha = zip_member_hashes(first)
    hb = zip_member_hashes(second)
    names = sorted(set(ha) | set(hb))
    for name in names:
        if ha.get(name) != hb.get(name):
            print(f"  [差异] {name}")
    print(f"  => 成员级差异数: {sum(1 for n in names if ha.get(n) != hb.get(n))} / {len(names)}")

    print("\n=== 3. 元数据（docProps）内容 ===")
    with zipfile.ZipFile(first) as archive:
        for name in ("docProps/core.xml", "docProps/app.xml"):
            if name in archive.namelist():
                print(f"  --- {name} ---")
                print("   ", archive.read(name).decode("utf-8", "replace")[:400])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
