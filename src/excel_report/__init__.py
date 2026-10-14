"""Excel 报表生成脚本（考题一）。

读取 sales_raw.xlsx，完成清洗与统计，输出含四个工作表的 report.xlsx。

设计要点（详见 docs/03-开发方案与技术选型.md）：
- ``loader`` 是唯一读取源文件的地方（只读，绝不写入）；
- ``report`` 是唯一写文件的地方；
- ``cleaning`` / ``aggregate`` 为纯函数，不触碰文件系统，便于单元测试；
- ``cli`` 负责把领域异常映射为进程退出码与中文可读提示。
"""

from __future__ import annotations

__all__ = ["__version__"]

# 版本号三处同步：此处、pyproject.toml、CHANGELOG.md（由 tests/test_version.py 断言）
__version__ = "1.0.0"
