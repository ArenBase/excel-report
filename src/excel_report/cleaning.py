"""数据清洗（纯函数，不触碰文件系统）。

清洗顺序**写死**为下面 6 步（docs/01-需求澄清清单.md §7 决策 D-03）。
顺序不同会导致"剔除行数"不同，因此这里不只是实现细节，而是需求口径的一部分。

====  ==============================================  ==========================
步骤  操作                                            计入的统计字段
====  ==============================================  ==========================
①     列名归一化（由 loader 完成）                    —
②     类型转换：日期→datetime、数量→Int64、单价→float   ``dropped_invalid``
③     关键字段缺失（NaN **或** strip 后为空串）        ``dropped_missing``
④     数量 <= 0                                       ``dropped_nonpositive``
⑤     全字段去重（保留首次出现）                      ``dropped_duplicate``
⑥     计算派生列 ``销售额 = 数量 × 单价``             —
====  ==============================================  ==========================

守恒不变量（由 tests 断言）::

    raw_rows == cleaned_rows + dropped_missing + dropped_nonpositive
                + dropped_duplicate + dropped_invalid
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime

import pandas as pd
from dateutil import parser as dateutil_parser

from excel_report.loader import ALL_COLUMNS, KEY_COLUMNS, PRODUCT_FILLNA

__all__ = ["CleaningStats", "clean", "SALES_AMOUNT_COLUMN", "to_excel_serial_datetime"]

#: 派生的销售额列名
SALES_AMOUNT_COLUMN = "销售额"

#: 可被当作 Excel 序列日期解析的数值范围（1900-01-01 ~ 2100-12-31）
_SERIAL_MIN = 1
_SERIAL_MAX = 73415

#: Excel 1900 日期系统的两个纪元偏移。
#: 由于 Excel 错误地把 1900 当作闰年（存在不存在的 1900-02-29，序列号 60），
#: 序列号 1~59 与 60 以上需要**相差一天**的纪元偏移。
#: 已与 ``openpyxl.utils.datetime.from_excel`` 逐点比对验证。
_UNIX_EPOCH = pd.Timestamp("1970-01-01")
_OFFSET_GE_60 = 25569  # (serial - 25569) 天
_OFFSET_LE_59 = 25568  # (serial - 25568) 天


def to_excel_serial_datetime(serial: int | float) -> pd.Timestamp:
    """把 Excel 序列日期号转换为 Timestamp。

    同时正确处理 Excel 的 1900 闰年缺陷::

        1     -> 1900-01-01      (1~59 用偏移 25568)
        59    -> 1900-02-28
        61    -> 1900-03-01      (>=60 用偏移 25569)
        45292 -> 2024-01-01
        45658 -> 2025-01-01

    Raises:
        ValueError: 序列号超出合理范围。
    """
    value = float(serial)
    if not (_SERIAL_MIN <= value <= _SERIAL_MAX):
        raise ValueError(f"Excel 序列日期超出合理范围: {serial}")
    offset = _OFFSET_LE_59 if value < 60 else _OFFSET_GE_60
    return _UNIX_EPOCH + pd.Timedelta(days=value - offset)


@dataclass(frozen=True)
class CleaningStats:
    """清洗过程的剔除统计。

    各 ``dropped_*`` 字段的**含义互斥**：一行只会被计入其中一个原因，
    因此它们的和加上清洗后行数必然等于原始行数。
    """

    raw_rows: int
    cleaned_rows: int
    dropped_missing: int = 0
    dropped_nonpositive: int = 0
    dropped_duplicate: int = 0
    dropped_invalid: int = 0
    product_filled: int = 0

    @property
    def dropped_total(self) -> int:
        """被剔除的总行数。"""
        return (
            self.dropped_missing
            + self.dropped_nonpositive
            + self.dropped_duplicate
            + self.dropped_invalid
        )

    def as_records(self) -> list[tuple[str, int]]:
        """转成 "指标名, 数值" 列表，供 "总览" 工作表使用。"""
        return [
            ("原始行数", self.raw_rows),
            ("清洗后行数", self.cleaned_rows),
            ("剔除行数合计", self.dropped_total),
            ("剔除·关键字段缺失", self.dropped_missing),
            ("剔除·数量<=0", self.dropped_nonpositive),
            ("剔除·全字段重复", self.dropped_duplicate),
            ("剔除·异常数值或日期", self.dropped_invalid),
            ("产品列填充为'未填写'的行数", self.product_filled),
        ]


def _blank_mask(series: pd.Series) -> pd.Series:
    """识别"空值"：真正的 NaN **或** strip 后为空串。

    ⚠️ 关键：pandas 的 ``isna()`` / ``dropna()`` **不认为 ``"   "`` 是缺失值**
    （见 research/01 调研报告 §1.3）。若不额外判定，纯空格销售员会穿透清洗，
    导致报表里出现一个空名称分组。
    """
    return pd.Series(
        [_is_blank_value(value) for value in series.tolist()],
        index=series.index,
        dtype=bool,
    )


def _is_blank_value(value: object) -> bool:
    """单个值是否为"空"（NaN/None/NaT 或 strip 后为空的字符串）。"""
    if value is None or value is pd.NaT:
        return True
    if isinstance(value, str):
        return not value.strip()
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def _parse_one_datetime(value: object) -> pd.Timestamp | None:
    """解析单个日期值，失败返回 ``None``。

    处理顺序（覆盖用例 TC-09 / TC-15 / TC-16）::

        datetime 类            -> 直接使用
        数值（如 45292）        -> 按 Excel 序列日期解释
        ISO 字符串 "YYYY-MM-DD" -> 单元素 pd.to_datetime（快且无警告）
        其他字符串 "2024/3/15"  -> dateutil 回退（支持非补零的斜杠写法）
    """
    if value is None or value is pd.NaT:
        return None

    # datetime / date 类
    if hasattr(value, "to_pydatetime") or isinstance(value, datetime):
        try:
            return pd.Timestamp(value)
        except (ValueError, TypeError, OverflowError):
            return None

    # 纯数值：按 Excel 序列日期处理。
    # ⚠️ 不能直接把整数交给 pd.to_datetime —— pandas 会把它当成"纳秒偏移"，
    #    45292 会变成 1970-01-01，这是一个已修复的真实 bug。
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            return to_excel_serial_datetime(value)
        except (ValueError, TypeError, OverflowError):
            return None

    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        # 先试 ISO：避免 pandas 面对混合格式时"整列推断失败→全部 NaT"
        try:
            return pd.Timestamp(text)
        except (ValueError, TypeError, OverflowError):
            pass
        # 再用 dateutil 解析 "2024/3/15" 这类非补零写法
        try:
            return pd.Timestamp(dateutil_parser.parse(text, dayfirst=False))  # type: ignore[no-any-return]
        except (ValueError, TypeError, OverflowError):
            return None

    # 数字字符串（"45292"）等兜底
    numeric = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    if pd.notna(numeric):
        try:
            return to_excel_serial_datetime(float(numeric))
        except (ValueError, TypeError, OverflowError):
            return None
    return None


def _to_datetime(series: pd.Series) -> pd.Series:
    """把日期列转为 datetime64（对每个元素独立解析）。

    之所以**不用** ``pd.to_datetime(series)`` 整列推断，是因为它有两个坑：

    1. 整数会被当成纳秒偏移（``45292`` → ``1970-01-01``），而不是 Excel 序列日期；
    2. 列内混合 ``"2024-03-01"`` 与 ``"2024/3/15"`` 时格式推断失败，
       会把后者整批变成 ``NaT``（等价于静默丢数据）。

    逐元素解析虽然略慢，但行为确定，且对本题的数据量（数千行）完全够用。
    """
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce")
    return pd.Series(
        [_parse_one_datetime(value) for value in series.tolist()],
        index=series.index,
        dtype="datetime64[ns]",
    )


def clean(df: pd.DataFrame) -> tuple[pd.DataFrame, CleaningStats]:
    """按固定顺序清洗原始数据，返回 ``(明细, 统计)``。

    Args:
        df: :func:`excel_report.loader.load_raw` 返回的原始 DataFrame。

    Returns:
        ``(detail, stats)``。``detail`` 已含派生列 ``销售额``，
        并按 ``["日期", "销售员", "产品", "数量", "单价", "地区", "销售额"]`` 排列。
    """
    raw_rows = int(len(df))
    work = df.copy()

    # ---- ② 类型转换：把 "无法解析" 的值转成 NaN / NaT，并由 dropped_invalid 记账 ----
    converted = pd.DataFrame(index=work.index)
    converted["日期"] = _to_datetime(work["日期"])
    converted["数量"] = pd.to_numeric(work["数量"], errors="coerce")
    converted["单价"] = pd.to_numeric(work["单价"], errors="coerce")
    converted["销售员"] = work["销售员"]
    converted["地区"] = work["地区"]
    converted["产品"] = work["产品"] if "产品" in work.columns else pd.NA

    # 原始值本身非空，但转换后变成 NaN/NaT —— 说明是"异常数值/日期"而非"缺失"
    invalid_mask = pd.Series(False, index=work.index)
    for col in ("日期", "数量", "单价"):
        source_present = ~_blank_mask(work[col])
        invalid_mask |= source_present & converted[col].isna()
    dropped_invalid = int(invalid_mask.sum())
    converted = converted.loc[~invalid_mask]

    # ---- ③ 关键字段缺失（NaN 或空白串） ----
    missing_mask = pd.Series(False, index=converted.index)
    for col in KEY_COLUMNS:
        missing_mask |= _blank_mask(converted[col])
    dropped_missing = int(missing_mask.sum())
    converted = converted.loc[~missing_mask]

    # 非关键字段"产品"：填占位值并保留该行（D-02，保证销售额守恒）
    product_filled = int(_blank_mask(converted["产品"]).sum())
    converted["产品"] = (
        converted["产品"]
        .astype(object)
        # mask(cond, other)：cond 为 True（即空值）的位置替换为占位值
        .mask(_blank_mask(converted["产品"]), PRODUCT_FILLNA)
        .astype(str)
        .str.strip()
    )

    # ---- ④ 数量 <= 0（单价 > 0 已在 ② 之后由缺失/非法逻辑保证；此处只按需求剔除数量） ----
    nonpositive_mask = converted["数量"] <= 0
    dropped_nonpositive = int(nonpositive_mask.sum())
    converted = converted.loc[~nonpositive_mask]

    # 文本列规范化（去首尾空白），使去重与分组语义稳定
    for col in ("销售员", "地区"):
        converted[col] = converted[col].astype(str).str.strip()

    # ---- ⑤ 全字段去重（考题原文"按全字段去重"，不传 subset） ----
    before_dedup = len(converted)
    converted = converted.drop_duplicates(keep="first")
    dropped_duplicate = int(before_dedup - len(converted))

    # ---- ⑥ 派生列：销售额 = 数量 × 单价（Int64 -> float 以便金额运算） ----
    converted["数量"] = converted["数量"].astype("Int64")
    converted["单价"] = converted["单价"].astype(float)
    converted[SALES_AMOUNT_COLUMN] = (
        converted["数量"].astype(float) * converted["单价"]
    ).round(2)

    detail = converted.loc[:, [*ALL_COLUMNS, SALES_AMOUNT_COLUMN]].reset_index(drop=True)
    detail = detail.sort_values(by=["日期", "销售员"], kind="stable").reset_index(drop=True)

    stats = CleaningStats(
        raw_rows=raw_rows,
        cleaned_rows=int(len(detail)),
        dropped_missing=dropped_missing,
        dropped_nonpositive=dropped_nonpositive,
        dropped_duplicate=dropped_duplicate,
        dropped_invalid=dropped_invalid,
        product_filled=product_filled,
    )
    return detail, stats
