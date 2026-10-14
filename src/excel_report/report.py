"""写出 report.xlsx 并加工样式。

**本模块是本工程唯一写文件的地方。**

实现要点（来源：research/01 调研报告 §1.4 / §1.5）
--------------------------------------------------
* 用 ``pd.ExcelWriter(..., engine="openpyxl")`` 的 **with 上下文管理器**，
  保证 ``close()`` 一定被调用，否则文件可能损坏或为空。
* 工作表名必须 ≤ 31 字符，否则 Excel 打开报错 —— 这里在 ``_check_sheet_names`` 中前置校验。
* ``index=False``：不写出 DataFrame 索引，避免报表多一列无意义序号。
* 列宽自适应时**东亚字符按 2 个字符宽计算**，否则中文列会被截断显示。
"""

from __future__ import annotations

import os
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from excel_report.errors import ReportError

__all__ = [
    "SHEET_DETAIL",
    "SHEET_BY_PERSON",
    "SHEET_BY_REGION",
    "SHEET_OVERVIEW",
    "SHEET_ORDER",
    "write_report",
    "freeze_workbook_timestamps",
]

SHEET_DETAIL = "明细清洗后"
SHEET_BY_PERSON = "按销售员"
SHEET_BY_REGION = "按地区"
SHEET_OVERVIEW = "总览"

#: 四个工作表的固定顺序（验收标准 AC-02）
SHEET_ORDER: tuple[str, ...] = (SHEET_DETAIL, SHEET_BY_PERSON, SHEET_BY_REGION, SHEET_OVERVIEW)

_EXCEL_MAX_SHEET_NAME = 31
#: Excel 单表行数上限
_EXCEL_MAX_ROWS = 1_048_576

_HEADER_FILL = PatternFill("solid", fgColor="4472C4")
_HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
_MONEY_FORMAT = "#,##0.00"
_PERCENT_FORMAT = '0.00"%"'
_THIN = Side(style="thin", color="D9D9D9")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

#: 需要按金额格式化的列
_MONEY_COLUMNS = {"销售额", "单价", "数值"}
#: 需要按百分比格式化的列
_PERCENT_COLUMNS = {"销售额占比"}

#: 固定的文档元数据时间戳。
#: openpyxl 在**保存时**会强制把 ``dcterms:modified`` 写成当前时间
#: （见 openpyxl/writer/excel.py 中 ``_save()`` 的 ``workbook.properties.modified = now()``），
#: 这会让同一输入两次运行的输出 SHA256 不同，违反 C-08（幂等/可复现）。
#: 因此在保存完成后重写 docProps/core.xml，把时间戳钉死，使输出**逐字节可复现**
#: （见 scripts/diagnose_idempotency.py 的量化验证）。
_FIXED_TIMESTAMP = datetime(2024, 1, 1, 0, 0, 0)
_FIXED_TIMESTAMP_TEXT = "2024-01-01T00:00:00Z"


def _pin_workbook_metadata(workbook) -> None:
    """固定工作簿的文档属性（对 ``created`` 有效；``modified`` 由保存后处理兜底）。"""
    props = workbook.properties
    props.created = _FIXED_TIMESTAMP
    props.modified = _FIXED_TIMESTAMP
    props.creator = "excel_report"
    props.lastModifiedBy = "excel_report"
    props.title = "销售报表"
    props.description = "由 excel_report 生成（考题一）"


def freeze_workbook_timestamps(path: Path) -> None:
    """把已写出的 xlsx 规范化为**逐字节可复现**的确定性文件。

    需要处理两处由 openpyxl 引入的非确定性（都来自"当前时间"）：

    1. ``docProps/core.xml`` 里的 ``dcterms:created`` / ``dcterms:modified``
       —— openpyxl 在 ``_save()`` 中强制写入 ``datetime.now()``；
    2. zip 每个成员的**修改时间戳**（``ZipInfo.date_time``）
       —— openpyxl 写包时使用当前时间。

    两者都会让同一输入两次运行的 SHA256 不同，违反 C-08（幂等/可复现）。
    openpyxl 没有提供关闭开关，因此在保存后重写整个 zip：内容原样复制，
    只替换时间戳，因此**不影响数据、样式或 Excel 兼容性**。
    """
    import re
    import shutil
    import tempfile
    import zipfile

    pattern = re.compile(
        r"(<dcterms:(?:created|modified)\b[^>]*>)([^<]*)(</dcterms:(?:created|modified)>)"
    )
    # ZIP 格式支持的最小时间戳，用作固定值
    fixed_zip_time = (1980, 1, 1, 0, 0, 0)

    with zipfile.ZipFile(path) as source:
        members = [(info, source.read(info.filename)) for info in source.infolist()]

    changed = False
    rewritten: list[tuple[zipfile.ZipInfo, bytes]] = []
    for info, data in members:
        if info.filename == "docProps/core.xml":
            text = data.decode("utf-8")
            new_text = pattern.sub(
                lambda m: f"{m.group(1)}{_FIXED_TIMESTAMP_TEXT}{m.group(3)}", text
            )
            if new_text != text:
                changed = True
                data = new_text.encode("utf-8")

        if info.date_time != fixed_zip_time:
            changed = True
            info = _with_fixed_timestamp(info, fixed_zip_time)
        rewritten.append((info, data))

    if not changed:
        return

    # 就地重写：先写到同目录临时文件，再替换，避免留下半成品
    handle, temp_name = tempfile.mkstemp(suffix=".xlsx", dir=str(path.parent))
    os.close(handle)
    temp_path = Path(temp_name)
    try:
        with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as target:
            for info, data in rewritten:
                target.writestr(info, data)
        shutil.move(str(temp_path), str(path))
    finally:
        temp_path.unlink(missing_ok=True)


def _with_fixed_timestamp(info, fixed_time: tuple[int, int, int, int, int, int]):
    """返回一个除时间戳外其余属性不变的新 ``ZipInfo``。"""
    import zipfile

    clone = zipfile.ZipInfo(info.filename, date_time=fixed_time)
    clone.compress_type = info.compress_type
    clone.external_attr = info.external_attr
    clone.internal_attr = info.internal_attr
    clone.create_system = info.create_system
    clone.comment = info.comment
    clone.extra = info.extra
    return clone


def _display_width(text: str) -> int:
    """计算字符串在 Excel 中的显示宽度（东亚全角字符计 2）。"""
    width = 0
    for char in text:
        width += 2 if unicodedata.east_asian_width(char) in ("W", "F") else 1
    return width


def _check_sheet_names(names: tuple[str, ...]) -> None:
    """前置校验工作表名合法（Excel 硬限制），避免生成无法打开的文件。"""
    for name in names:
        if len(name) > _EXCEL_MAX_SHEET_NAME:
            raise ReportError(
                f"工作表名超过 Excel 限制（{_EXCEL_MAX_SHEET_NAME} 字符）: {name!r}",
                hint="请缩短工作表常量 SHEET_* 的名称。",
            )
        if any(ch in name for ch in r"[]:*?/\\"):
            raise ReportError(
                f"工作表名包含 Excel 非法字符: {name!r}",
                hint=(
                    "Excel 工作表名不得包含 [ ] : * ? / \\ 中的任意字符。"
                ),
            )


def _check_row_limits(sheets: dict[str, pd.DataFrame]) -> None:
    """校验行数不超过 Excel 上限，超限时报错而非静默截断（C-12）。"""
    for name, frame in sheets.items():
        # +1 为表头行
        if len(frame) + 1 > _EXCEL_MAX_ROWS:
            raise ReportError(
                f"工作表 {name!r} 有 {len(frame)} 行，超过 Excel 单表上限 {_EXCEL_MAX_ROWS} 行。",
                hint="请先拆分原始数据，或改用 CSV 输出。",
            )


def _style_worksheet(worksheet) -> None:
    """套用表头样式、自适应列宽、数字格式、冻结首行与自动筛选。"""
    header_row = 1

    # 表头样式
    for column_index in range(1, worksheet.max_column + 1):
        cell = worksheet.cell(row=header_row, column=column_index)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _BORDER

    # 列宽自适应
    for column_index in range(1, worksheet.max_column + 1):
        letter = get_column_letter(column_index)
        header_text = worksheet.cell(row=header_row, column=column_index).value
        best = _display_width(str(header_text)) if header_text is not None else 0

        # 只采样前若干行，避免超大表计算过慢（对列宽判断已足够）
        sample_limit = min(worksheet.max_row, 500)
        for row_index in range(2, sample_limit + 1):
            value = worksheet.cell(row=row_index, column=column_index).value
            if value is None:
                continue
            if isinstance(value, float):
                text = f"{value:,.2f}"
            elif isinstance(value, int):
                text = f"{value:,}"
            else:
                text = str(value)
            best = max(best, _display_width(text))

        # +2 留出内边距；金额列额外放宽
        worksheet.column_dimensions[letter].width = min(best + 2, 60)

    # 数字格式与边框
    headers = [
        worksheet.cell(row=header_row, column=i).value for i in range(1, worksheet.max_column + 1)
    ]
    for column_index, header in enumerate(headers, start=1):
        if header in _MONEY_COLUMNS:
            number_format = _MONEY_FORMAT
        elif header in _PERCENT_COLUMNS:
            number_format = _PERCENT_FORMAT
        else:
            number_format = None
        if number_format is None:
            continue
        for row_index in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row=row_index, column=column_index)
            cell.number_format = number_format
            cell.border = _BORDER
            if header in _MONEY_COLUMNS:
                cell.alignment = Alignment(horizontal="right")

    # 冻结首行 + 自动筛选（有数据行时才加筛选）
    worksheet.freeze_panes = "A2"
    if worksheet.max_row > 1 and worksheet.max_column >= 1:
        worksheet.auto_filter.ref = worksheet.dimensions


def write_report(
    output: str | Path,
    *,
    detail: pd.DataFrame,
    by_person: pd.DataFrame,
    by_region: pd.DataFrame,
    overview_frame: pd.DataFrame,
) -> Path:
    """把四个 DataFrame 写成一个 xlsx 报表。

    Args:
        output: 输出文件路径；父目录不存在时自动创建（C-06）。
        detail: "明细清洗后" 数据。
        by_person: "按销售员" 数据。
        by_region: "按地区" 数据。
        overview_frame: "总览" 数据。

    Returns:
        实际写入的文件路径。

    Raises:
        ReportError: 工作表名非法或行数超限。
    """
    output_path = Path(output)
    sheets: dict[str, pd.DataFrame] = {
        SHEET_DETAIL: detail,
        SHEET_BY_PERSON: by_person,
        SHEET_BY_REGION: by_region,
        SHEET_OVERVIEW: overview_frame,
    }

    _check_sheet_names(SHEET_ORDER)
    _check_row_limits(sheets)

    # C-06：父目录不存在则创建
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for name in SHEET_ORDER:
            frame = sheets[name]
            frame.to_excel(writer, sheet_name=name, index=False)
            worksheet = writer.sheets[name]
            _style_worksheet(worksheet)
        # 必须在 with 退出（保存）之前固定元数据
        _pin_workbook_metadata(writer.book)

    # 保存后兜底：openpyxl 会在保存时把 modified 重写为当前时间，这里再钉死一次
    freeze_workbook_timestamps(output_path)

    return output_path
