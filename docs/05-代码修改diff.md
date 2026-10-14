# 考题一 · 代码修改轨迹（diff）

> 对应考核交付物 4 的后半部分「代码修改 diff」，以及考核要求"完整保留……代码修改轨迹"。
> 本文档记录**关键修改的前后对比**，重点是那些由验证抓出、并改变了实现方向的改动。
>
> 说明：本工程为一次性交付，未做多次发布提交；下文按"问题 → 修改前 → 修改后 → 原因 → 验证"组织。
> 各版本的功能级变更摘要见根目录 `CHANGELOG.md`。

---

## Diff 1 · 【缺陷修复】整数日期被当成纳秒偏移

**文件**：`src/excel_report/cleaning.py`
**发现方式**：`test_clean_excel_serial_date[45292-...]` 失败
**严重度**：高（静默错误，不报错但结果错）

### 修改前

```python
def _to_datetime(series: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce")

    parsed = pd.to_datetime(series, errors="coerce")

    # 对解析失败且为数值的项，尝试 Excel 序列日期
    numeric = pd.to_numeric(series, errors="coerce")
    retry = parsed.isna() & numeric.notna()
    if retry.any():
        def _convert(value):
            try:
                return to_excel_serial_datetime(value)
            except (ValueError, TypeError, OverflowError):
                return pd.NaT
        parsed.loc[retry] = numeric[retry].map(_convert)
    return pd.to_datetime(parsed, errors="coerce")
```

### 修改后

```python
def _to_datetime(series: pd.Series) -> pd.Series:
    """把日期列转为 datetime64（对每个元素独立解析）。"""
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce")
    return pd.Series(
        [_parse_one_datetime(value) for value in series.tolist()],
        index=series.index,
        dtype="datetime64[ns]",
    )


def _parse_one_datetime(value: object) -> pd.Timestamp | None:
    if value is None or value is pd.NaT:
        return None
    if hasattr(value, "to_pydatetime") or isinstance(value, datetime):
        try:
            return pd.Timestamp(value)
        except (ValueError, TypeError, OverflowError):
            return None
    # 纯数值：按 Excel 序列日期处理（不能交给 pd.to_datetime！）
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            return to_excel_serial_datetime(value)
        except (ValueError, TypeError, OverflowError):
            return None
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        try:
            return pd.Timestamp(text)          # 先试 ISO
        except (ValueError, TypeError, OverflowError):
            pass
        try:
            return pd.Timestamp(dateutil_parser.parse(text, dayfirst=False))
        except (ValueError, TypeError, OverflowError):
            return None
    ...
```

### 原因

`pd.to_datetime(45292)` **不会失败**——pandas 把整数解释为"1970-01-01 起的纳秒数"，
返回 `1970-01-01 00:00:00.000045292`。因此原先的"失败后重试"分支**永远不会触发**，
整数日期全部落进 1970 年，导致月度分组彻底错误。

### 验证

```python
@pytest.mark.parametrize(("serial", "expected"),
    [(45292, "2024-01-01"), (45658, "2025-01-01"), (1, "1900-01-01"), (61, "1900-03-01")])
def test_clean_excel_serial_date(serial, expected): ...
```
✅ 通过

---

## Diff 2 · 【缺陷修复】混合日期格式导致整批丢数据

**文件**：`src/excel_report/cleaning.py`
**发现方式**：`test_clean_parses_multiple_date_formats`、`test_monthly_merges_same_month` 失败
**严重度**：高（静默丢数据）

### 修改前

整列交给 pandas 推断格式（同 Diff 1 的 `pd.to_datetime(series, errors="coerce")`）。

### 修改后

逐元素解析，字符串**先试 ISO、再用 `dateutil` 回退**（见 Diff 1 代码）。

### 原因

实测同一列内混用格式时，pandas 的整列推断会失败并把不合格式的值**静默变成 `NaT`**：

```
>>> pd.to_datetime(pd.Series(["2024-03-01", "2024/3/15", "2024-03-31"]), errors="coerce")
[Timestamp('2024-03-01'), NaT, Timestamp('2024-03-31')]
   ^^^^^^^^^^^^^^^^^^^^^^^^^^ "2024/3/15" 被丢弃，且 pandas 只给一条 UserWarning
```

它不会抛异常，只是"这一行的日期没了"。对报表而言等于**凭空少了一笔销售额**。

### 验证

`test_clean_parses_multiple_date_formats`：三种写法必须归并到同一个 `2024-03` ✅

---

## Diff 3 · 【缺陷修复】1900 闰年缺陷区间偏差一天

**文件**：`src/excel_report/cleaning.py`
**发现方式**：人工用 `openpyxl.utils.datetime.from_excel` 做**全区间**（1~73415）逐点比对
**严重度**：中（只影响 1900-01-01 ~ 1900-02-28）

### 修改前

```python
def to_excel_serial_datetime(serial: int | float) -> pd.Timestamp:
    value = float(serial)
    if not (_SERIAL_MIN <= value <= _SERIAL_MAX):
        raise ValueError(f"Excel 序列日期超出合理范围: {serial}")
    return pd.Timestamp("1970-01-01") + pd.Timedelta(seconds=(value - 25569) * 86400)
```

### 修改后

```python
# Excel 1900 日期系统的两个纪元偏移。
# Excel 错误地把 1900 当作闰年（存在不存在的 1900-02-29，序列号 60），
# 因此序列号 1~59 与 >=60 需要**相差一天**的纪元偏移。
_UNIX_EPOCH = pd.Timestamp("1970-01-01")
_OFFSET_GE_60 = 25569
_OFFSET_LE_59 = 25568


def to_excel_serial_datetime(serial: int | float) -> pd.Timestamp:
    value = float(serial)
    if not (_SERIAL_MIN <= value <= _SERIAL_MAX):
        raise ValueError(f"Excel 序列日期超出合理范围: {serial}")
    offset = _OFFSET_LE_59 if value < 60 else _OFFSET_GE_60
    return _UNIX_EPOCH + pd.Timedelta(days=value - offset)
```

### 原因

比对结果（修改前）：

| 序列号 | openpyxl（正确） | 原实现（错误） |
|---|---|---|
| 1 | 1900-01-01 | 1899-12-31 ❌ |
| 59 | 1900-02-28 | 1900-02-27 ❌ |
| 61 | 1900-03-01 | 1900-03-01 ✅ |
| 45292 | 2024-01-01 | 2024-01-01 ✅ |

### 验证

```python
def test_excel_serial_matches_openpyxl_reference() -> None:
    from openpyxl.utils.datetime import from_excel
    mismatches = [s for s in range(1, 73416)
                  if from_excel(s) != to_excel_serial_datetime(s).to_pydatetime()]
    assert mismatches == []
```
✅ 1~73415 全区间 **0 处不一致**

> 注：AI 初版测试只采了 4 个现代日期采样点，**恰好全部落在这个缺陷区间之外**，
> 所以"测试通过"是假象。人工把采样点扩为全区间比对才暴露问题。

---

## Diff 4 · 【缺陷修复】输出不能逐字节复现

**文件**：`src/excel_report/report.py`（新增函数 + 调用）、`scripts/make_sample_data.py`
**发现方式**：`test_output_is_bytewise_idempotent` 失败
**严重度**：中（违反 C-08 幂等/可复现）

### 修改前

```python
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    for name in SHEET_ORDER:
        frame = sheets[name]
        frame.to_excel(writer, sheet_name=name, index=False)
        _style_worksheet(writer.sheets[name])

return output_path
```

### 修改后

```python
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    for name in SHEET_ORDER:
        frame = sheets[name]
        frame.to_excel(writer, sheet_name=name, index=False)
        _style_worksheet(writer.sheets[name])
    _pin_workbook_metadata(writer.book)      # created/modified/creator...

# openpyxl 会在保存时把 modified 重写为当前时间，保存后兜底再钉死一次
freeze_workbook_timestamps(output_path)

return output_path
```

其中 `freeze_workbook_timestamps()` 重写 zip，把两处"当前时间"替换为固定值：

```python
pattern = re.compile(
    r"(<dcterms:(?:created|modified)\b[^>]*>)([^<]*)(</dcterms:(?:created|modified)>)"
)
fixed_zip_time = (1980, 1, 1, 0, 0, 0)   # ZIP 格式支持的最小时间戳
```

### 原因

人工用"逐 zip 成员哈希比对"定位到差异（`scripts/diagnose_idempotency.py`）：

```
=== 2. zip 成员层面哪些不同？ ===
  [差异] docProps/core.xml
  => 成员级差异数: 1 / 12
```

进一步发现**两个**非确定性来源：

1. `docProps/core.xml` 的 `dcterms:created` / `dcterms:modified`；
2. **每个 zip 成员自身的修改时间戳**（`ZipInfo.date_time`）。

且第 1 项无法通过设置属性解决——openpyxl 在 `writer/excel.py` 的 `_save()` 中**硬编码**：

```python
# openpyxl/writer/excel.py:292
workbook.properties.modified = datetime.datetime.now(tz=datetime.timezone.utc).replace(tzinfo=None)
```

所以必须在保存后重写 zip。

### 验证

```
report 逐字节幂等: True
  SHA256 = D569124498BD54B1D16A4E7E136B9D05D8830AFCCB32F3F1069588532D668101
样例逐字节可复现: True
```

且 `test_normalized_output_is_still_a_valid_xlsx` 确保重写 zip 后
**工作表结构（4 表）与样式（冻结首行）均未丢失**，文件仍可被 Excel/openpyxl 正常打开。

---

## Diff 5 · 【需求补强】空白字符串必须视为缺失

**文件**：`src/excel_report/cleaning.py`
**来源**：调研发现（pandas `dropna` 不识别 `"   "`）+ 人工决策

### 修改前（直觉写法，未采用）

```python
df = df.dropna(subset=KEY_COLUMNS)          # "   " 会穿透，不满足 C-03
```

### 修改后

```python
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


def _blank_mask(series: pd.Series) -> pd.Series:
    """识别"空值"：真正的 NaN **或** strip 后为空串。"""
    return pd.Series([_is_blank_value(v) for v in series.tolist()],
                     index=series.index, dtype=bool)
```

### 原因

Excel 中"看起来是空的"单元格常为纯空格串（系统导出常见）。
若只用 `dropna()`，会产生一个**销售员名为空的诡异分组**，且"剔除行数"与人工核对不上。

### 验证

```python
@pytest.mark.parametrize("column", ["日期", "销售员", "数量", "单价", "地区"])
def test_clean_blank_string_counts_as_missing(column): ...
```
✅ 5 个关键字段全部覆盖

---

## Diff 6 · 【展示修正】总览表输出空单元格而非 NaN

**文件**：`src/excel_report/aggregate.py`

### 修改前

```python
result = pd.DataFrame(rows, columns=["指标", "数值"])
if not monthly.empty:
    ...
    result = pd.concat([result, spacer, header, monthly_block], ignore_index=True)
return result
```

### 修改后

```python
    result = pd.concat([result, spacer, header, monthly_block], ignore_index=True)

# 结构化空白行（间隔行、区块标题行）会产生 NaN，输出时应显示为空单元格而非 "NaN"
return result.fillna("")
```

### 原因

总览表用"间隔行 + 区块标题行"分隔 KPI 区与月度明细区，这些行的第二列为空，
DataFrame 里是 `NaN`，写进 Excel 会显示成 `NaN` 字样，观感差。

### 验证

```python
bad = [(c.coordinate, c.value) for r in ws.iter_rows() for c in r
       if isinstance(c.value, str) and c.value.lower() == "nan"]
# -> 无 ✅
```

---

## Diff 7 · 【方案否决】回退 AI 建议的 `astype(object)` 改动

**文件**：`src/excel_report/cleaning.py`、`src/excel_report/aggregate.py`
**性质**：**回退**（先加后删），记录在此以体现"人工审查纠正 AI 误判"

### AI 的建议（曾实施，后回退）

```python
# AI 认为 pandas 3.0 的 "str" dtype 排序不遵循 Unicode 码点，建议：
for col in ("销售员", "地区", "产品"):
    converted[col] = converted[col].astype(object)
# 并在 _group_by 中：
grouped[key_column] = grouped[key_column].astype(object)
```

### 实测结论（人工验证后回退）

```
A constructed dtype: str
A sort: ['张三', '李四', '王五']      <- str dtype 排序正常
B groupby dtype: str
B sort: ['张三', '李四', '王五']      <- 也正常
B->object sort: ['张三', '李四', '王五']  <- object 同样

未加 object 转换时排序顺序: ['张三', '李四', '王五']
加了 object 转换时排序顺序: ['张三', '李四', '王五']   <- 完全相同
```

### 真正的原因

失败的其实是**测试用例的期望值**，不是产品代码：

```python
# 错误（凭直觉写）
assert result["销售员"].tolist() == ["李四", "张三", "王五"]
# 正确（Unicode 码点序）：张 U+5F20 < 李 U+674E < 王 U+738B
assert result["销售员"].tolist() == ["张三", "李四", "王五"]
```

### 处理

1. **回退**所有 `astype(object)` 改动（不留无意义代码）；
2. 修正测试期望值，并让期望值由 `sorted()` **自证**，不再凭记忆写；
3. 新增 `test_tie_break_order_is_stable_and_documented` 锁定该口径。

> **教训**：AI 提出的"根因"即使听起来合理（"新版 dtype 排序行为不同"），
> 也必须**实测复现**后再采纳。否则会在正确代码上叠加无意义改动，
> 并留下一条错误的归因注释，误导后来者。

---

## Diff 8 · 【配置】依赖收敛与版本一致性

**文件**：`pyproject.toml`、`requirements.txt`、`tests/test_version.py`

### 变更

```diff
 dependencies = [
     "pandas>=2.0",
     "openpyxl>=3.1",
+    # python-dateutil 用于解析 "2024/3/15" 这类非补零日期写法；
+    # 它本身也是 pandas 的依赖
+    "python-dateutil>=2.8",
 ]
```

### 配套的自动化审查（`tests/test_version.py`）

```python
def test_no_forbidden_runtime_dependencies() -> None:
    """不得引入 xlsxwriter / polars / pandera 等考题未列出的运行依赖。"""
    for forbidden in ("xlsxwriter", "polars", "pandera", "great-expectations", "fastapi"):
        assert forbidden not in deps


def test_source_imports_stay_within_allowed_libraries() -> None:
    """静态审查源码的第三方 import，确保没有越界依赖。"""
```

### 原因

R-05 风险："AI 擅自重构/引入不必要依赖"。人工把"依赖边界"从口头约定
变成**自动化断言**，让 AI 后续任何越界引入都会立刻红灯。

---

## Diff 9 · 【架构】loader 的"无写 API"从约定变成断言

**文件**：`tests/test_invariants.py`

### 新增

```python
def test_loader_module_has_no_write_api() -> None:
    """静态审查：loader 里不得出现任何写文件的 API（结构性防线的自动化检查）。"""
    import inspect
    source = inspect.getsource(loader)
    for forbidden in ("to_excel", "ExcelWriter", "open(", ".write", "mkdir", "shutil"):
        assert forbidden not in source, f"loader 中不应出现写文件相关调用: {forbidden}"
```

### 原因

R-01（误改原始文件）是本项目最高风险。"loader 是唯一读文件处且只读"这条设计约定，
如果只写在文档里，AI 后续一次"顺手重构"就可能破坏它。
用 `inspect.getsource` 把它变成 CI 可拦截的断言，是**结构性强于纪律性**的做法。

---

## 修改清单汇总

| Diff | 类型 | 文件 | 由什么发现 |
|---|---|---|---|
| 1 | 缺陷修复（静默错误） | `cleaning.py` | TC-16 用例 |
| 2 | 缺陷修复（静默丢数据） | `cleaning.py` | TC-09/TC-10 用例 |
| 3 | 缺陷修复（边界偏差） | `cleaning.py` | 人工全区间对比 openpyxl |
| 4 | 缺陷修复（幂等性） | `report.py`、`scripts/` | TC-25 用例 + 人工逐成员比对 |
| 5 | 需求补强（边界） | `cleaning.py` | 调研发现 + 人工决策 |
| 6 | 展示修正 | `aggregate.py` | 人工肉眼核对输出 |
| 7 | **方案否决（回退）** | `cleaning.py`、`aggregate.py` | 人工实测推翻 AI 根因 |
| 8 | 配置（依赖收敛） | `pyproject.toml` 等 | R-05 风险清单 |
| 9 | 架构（约定转断言） | `test_invariants.py` | R-01 风险清单 |

> **观察**：9 项修改中，**4 项是"静默错误"**（不报错但结果错）。
> 这正说明考核强调"每步运行验证"的必要性——
> 只审代码看不出来，必须实跑并断言**具体数值**。

---

## 附：如何自行复现这些 diff 的验证

```bash
cd exam-1-excel-report

# Diff 1 & 2 & 3：日期解析
python -m pytest tests/test_cleaning.py -k "date or serial" -v

# Diff 4：幂等性
python src/excel_report -i data/raw/sales_raw.xlsx -o output/a.xlsx -q
python src/excel_report -i data/raw/sales_raw.xlsx -o output/b.xlsx -q
python scripts/diagnose_idempotency.py output/a.xlsx output/b.xlsx

# Diff 5：空白串缺失
python -m pytest tests/test_cleaning.py -k "blank" -v

# Diff 6：总览表无 NaN
python scripts/verify_report.py

# Diff 7：并列排序口径
python -m pytest tests/test_aggregate.py -k "sort or tie" -v

# Diff 8 & 9：依赖与只读边界
python -m pytest tests/test_version.py tests/test_invariants.py -v
```
