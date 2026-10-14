# 更新日志

本文件记录「考题一 · Excel 报表生成脚本」的所有版本变更。

格式参考 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，
版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

> **本项目的特殊版本约定**
> 因为"清洗口径 / 统计口径"直接决定报表数值，所以**口径变更一律视为 MAJOR 版本**。
> 例如：若把"关键字段"从 5 列改为 4 列，即使代码改动很小，也必须升 `2.0.0`。
> 版本号需在 `src/excel_report/__init__.py`、`pyproject.toml`、本文件三处保持一致
> （由 `tests/test_version.py` 自动断言）。

---

## [1.0.0] - 2026-10-14

首个交付版本。完成考题一全部要求：数据清洗、按月/销售员/地区统计、四工作表输出、异常可读提示。

### Added

- **需求澄清与设计文档**
  - `docs/01-需求澄清清单.md`：目标、输入输出、边界约束（C-01~C-12）、验收标准（AC-01~AC-09）、
    决策记录（D-01~D-07）、风险清单（R-01~R-07）
  - `docs/02-报表用例.md`：29 条用例（主流程 4 / 边界 12 / 异常 6 / 不变量 7）
  - `docs/03-开发方案与技术选型.md`：分层架构、模块接口、版本管理规则、测试策略
  - `research/01-网络深度查询调研报告.md`：anySerach 网络深度查询结果（技术规范 / 工具链 / 8 个模范项目）
- **可运行源码工程（src layout）**
  - `excel_report/errors.py`：领域异常体系，异常类自带 `exit_code`
  - `excel_report/loader.py`：**唯一读文件处**，只读加载 + 列名归一化 + 必需列校验
  - `excel_report/cleaning.py`：6 步固定顺序清洗，返回明细与剔除统计
  - `excel_report/aggregate.py`：按销售员 / 按地区 / 月度 / 总览，纯函数
  - `excel_report/report.py`：**唯一写文件处**，四工作表 + 表头样式 + 列宽自适应 + 冻结首行
  - `excel_report/pipeline.py`：流程编排，用结构（而非 try/except）保证不产出半成品报表
  - `excel_report/cli.py`：命令行入口，异常 → 退出码（0/1/2/3/4/5/6）
  - `excel_report/__main__.py`：src layout 官方 workaround，实现**零安装一键运行**
- **辅助脚本**
  - `scripts/make_sample_data.py`：生成含 10 类脏数据的样例 `sales_raw.xlsx`
  - `scripts/verify_report.py`：核对四表与统计口径、三方对账、原文件哈希
- **测试套件**：`tests/` 下 106 个用例（清洗 / 统计 / 异常 / 集成 / 不变量 / 版本）
- `pyproject.toml`（PEP 621 + Ruff + pytest 单点配置）、`requirements.txt`、`README.md`

### Fixed

开发过程中发现并修复的 2 个真实缺陷（均由测试用例抓出，体现"每步验证"的价值）：

- **整数日期被误解为纳秒偏移**：直接把整数交给 `pd.to_datetime`，`45292`（Excel 序列日期）
  会被解释成 `1970-01-01 00:00:00.000045292`。改为逐元素解析并把数值按 Excel 序列日期处理；
  同时修正了 1900 闰年缺陷区间（序列号 1~59 与 ≥60 需相差一天），
  并与 `openpyxl.utils.datetime.from_excel` 在 1~73415 全区间逐点比对通过。
- **混合日期格式导致整批丢失**：同一列里 `"2024-03-01"` 与 `"2024/3/15"` 混用时，
  pandas 的整列格式推断失败会把后者全部变成 `NaT`（等于静默丢数据）。
  改为逐元素解析（ISO 优先，`dateutil` 回退）。

### 另外修正的**测试用例缺陷**（非产品代码缺陷，如实记录）

- **并列排序的期望值写错**：`test_sort_is_sales_desc_then_name_asc` 原先断言并列时应为
  `["李四", "张三"]`，凭直觉以为"李"排在"张"前。实际 Unicode 码点序为
  张 U+5F20 < 李 U+674E < 王 U+738B，**产品代码一直都是对的，是用例写错了**。
  已把期望值改为由 `sorted()` 自证，并新增 `test_tie_break_order_is_stable_and_documented`
  锁定该口径，避免再次凭记忆写错。
  > 记录此事的意义：AI 辅助开发中，**"测试失败"不等于"产品有 bug"**，
  > 必须区分"产品缺陷"与"用例缺陷"，否则会去"修"本来正确的代码。

### Security

- 输入输出同路径检测（`SamePathError`，退出码 6），从结构上防止报表覆盖原始数据。
- `loader` 模块经 `tests/test_invariants.py` 静态断言不含任何写文件 API。
- 运行时依赖收敛为 `pandas` + `openpyxl`（+ pandas 依赖的 `python-dateutil`），
  拒绝引入考题未列出的库（由 `tests/test_version.py` 静态审查）。
- 源码不读取环境变量、不获取当前时间、不使用随机数，保证输出可复现。
