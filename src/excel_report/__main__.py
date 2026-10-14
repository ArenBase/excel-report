"""``python -m excel_report`` 入口。

src layout 的官方代价是"必须先安装才能运行"，而本项目要求**一键运行、零安装**。
因此按 PyPA 官方文档给出的 workaround，在包未被安装时把 ``src`` 目录前插到
``sys.path``（见 research/01 调研报告 §1.2）。

运行方式::

    python -m excel_report                 # 在已 pip install -e . 的环境
    python src/excel_report --help         # 未安装时，从源码树直接运行
"""

from __future__ import annotations

import os
import sys

# src layout workaround：仅当作为脚本直接运行（__package__ 为空）时生效
if not __package__:
    _SRC_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if _SRC_DIR not in sys.path:
        sys.path.insert(0, _SRC_DIR)

from excel_report.cli import main  # noqa: E402  (必须在 sys.path 调整之后导入)

if __name__ == "__main__":
    sys.exit(main())
