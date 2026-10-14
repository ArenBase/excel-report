"""领域异常定义。

设计契约
--------
库层（loader / cleaning / aggregate / report）只抛 :class:`ReportError` 的子类；
命令行层（cli）统一把异常映射为 **退出码 + 中文可读提示**，不向用户抛裸 traceback。

``exit_code`` 作为**类属性**而不是构造参数，使异常类本身就是
"错误类型 → 退出码" 的映射表，cli 无需冗长的 isinstance 判断链。

退出码约定
----------
====  ======================================
码    含义
====  ======================================
0     成功
1     未预期异常（兜底）
2     输入文件不存在
3     无有效数据（空表 / 全部被剔除）
4     缺少必需列
5     不是有效的 Excel 文件
6     输出路径与输入路径相同（防覆盖原始文件）
====  ======================================
"""

from __future__ import annotations

__all__ = [
    "ReportError",
    "InputFileNotFoundError",
    "EmptyDataError",
    "NoValidDataError",
    "MissingColumnsError",
    "InvalidWorkbookError",
    "SamePathError",
]


class ReportError(Exception):
    """所有业务异常的基类。

    Attributes:
        exit_code: 该错误对应的进程退出码。
        hint: 面向用户的补充建议（可为空）。
    """

    exit_code: int = 1
    hint: str = ""

    def __init__(self, message: str, *, hint: str = "") -> None:
        super().__init__(message)
        self.message = message
        if hint:
            self.hint = hint

    def format_for_user(self) -> str:
        """渲染为面向用户的单条提示（不含 traceback）。"""
        text = f"[错误] {self.message}"
        if self.hint:
            text += f"\n       建议: {self.hint}"
        return text


class InputFileNotFoundError(ReportError):
    """输入文件不存在（C-02）。"""

    exit_code = 2


class EmptyDataError(ReportError):
    """文件中没有任何数据行（C-03）。"""

    exit_code = 3


class NoValidDataError(ReportError):
    """原始文件有数据行，但清洗后无任何有效数据（C-03）。"""

    exit_code = 3


class MissingColumnsError(ReportError):
    """缺少必需列（C-04）。

    Attributes:
        missing: 缺失的列名列表。
        actual: 实际读到的表头列表。
    """

    exit_code = 4

    def __init__(self, missing: list[str], actual: list[str]) -> None:
        self.missing = list(missing)
        self.actual = list(actual)
        message = "缺少必需列: " + ", ".join(self.missing)
        message += "\n       实际表头: " + (", ".join(self.actual) if self.actual else "(空)")
        super().__init__(message, hint="请检查原始表的表头拼写，或补充缺失的列后重试。")


class InvalidWorkbookError(ReportError):
    """文件不是有效的 Excel 工作簿（C-03）。"""

    exit_code = 5


class SamePathError(ReportError):
    """输出路径与输入路径相同，拒绝执行以防覆盖原始文件（R-01）。"""

    exit_code = 6
