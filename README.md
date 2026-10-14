# 考题一 · Excel 报表生成脚本

> 读取月度脏数据 `sales_raw.xlsx`，完成清洗与统计，输出含**四个工作表**的 `report.xlsx`。
> 对应考核文档《Python Vibe Coding 实操》考题一（必做，40 分）。

---

## 快速开始（三步，2 分钟）

```bash
cd exam-1-excel-report

# ① 安装依赖（pandas / openpyxl / python-dateutil）
pip install -r requirements.txt

# ② 生成样例脏数据（含 10 类脏数据，便于演示）
python scripts/make_sample_data.py

# ③ 一键生成报表
python src/excel_report --input data/raw/sales_raw.xlsx --output output/report.xlsx
```

预期输出：

```
[OK] 读取: data\raw\sales_raw.xlsx  (原始 143 行)
[OK] 清洗后: 128 行  (剔除 15 行)
     - 关键字段缺失 : 4 行
     - 数量<=0      : 3 行
     - 全字段重复   : 5 行
     - 非法数值/日期: 3 行
     - 产品填充'未填写': 2 行
[OK] 总销售额: 622,141.20
[OK] 月度区间: 2023-12 ~ 2024-06  (共 7 个月)
[OK] 报表已生成: output\report.xlsx (4 个工作表)
```

> **无需 `pip install -e .` 即可运行**。本工程采用 src layout，
> 但按 PyPA 官方 workaround 在 `__main__.py` 中做了 `sys.path` 前插，
> 因此 `python src/excel_report` 可直接运行。

如果已安装为包（`pip install -e .`），也可以：

```bash
python -m excel_report --input data/raw/sales_raw.xlsx --output output/report.xlsx
excel-report --input data/raw/sales_raw.xlsx --output output/report.xlsx
```

---

## 输出说明

`report.xlsx` 含**四个工作表**（顺序固定）：

| # | 工作表 | 内容 |
|---|---|---|
| 1 | `明细清洗后` | 清洗后的逐行明细，含派生的 `销售额` 列 |
| 2 | `按销售员` | 销售员、订单行数、总数量、销售额、销售额占比 |
| 3 | `按地区` | 地区、订单行数、总数量、销售额、销售额占比 |
| 4 | `总览` | 统计口径说明 + 原始/清洗后行数 + **各剔除原因计数** + 总销售额 + **月度总额明细** |

样式：表头加粗填色、列宽按东亚字符宽度自适应、冻结首行、自动筛选、
金额列 `#,##0.00` 格式、占比列 `0.00"%"`。

---

## 数据清洗规则（顺序固定，见 `docs/01-需求澄清清单.md` §7）

```
① 列名归一化（去首尾空白）
② 类型转换：日期→datetime、数量→Int64、单价→float     → 失败计入「异常数值/日期」
③ 关键字段缺失（NaN 或 strip 后为空串）                → 计入「关键字段缺失」
④ 数量 <= 0                                          → 计入「数量<=0」
⑤ 全字段去重（保留首次出现）                          → 计入「全字段重复」
⑥ 计算派生列 销售额 = 数量 × 单价
```

**守恒律**（由测试断言）：

```
原始行数 = 清洗后行数 + 缺失 + 非正数 + 重复 + 异常数值/日期
```

### 关键字段

`日期`、`销售员`、`数量`、`单价`、`地区` 共 5 列。
**`产品` 不是关键字段** —— 缺失时填 `未填写` 并**保留该行**，避免销售额凭空消失。

---

## 统计口径

| 项目 | 口径 |
|---|---|
| 销售额 | **逐行** `数量 × 单价` 后再汇总（不是"汇总数量 × 汇总单价"） |
| 金额精度 | 输出保留 2 位小数 |
| 销售额占比 | 该组销售额 ÷ 总销售额 × 100%；总额为 0 时记 0.00（不除零崩溃） |
| 排序 | 销售额降序；**并列时按名称升序（Unicode 码点序）**，保证结果确定 |
| 订单行数 | 清洗后明细中该组的行数（1 行 = 1 笔销售记录） |

---

## 异常处理与退出码

**原始文件永不被修改**：只读加载，且输出路径与输入路径相同时直接拒绝执行。

| 退出码 | 含义 | 触发的典型场景 |
|---|---|---|
| `0` | 成功 | — |
| `1` | 未预期异常（兜底） | 本工具的缺陷 |
| `2` | 输入文件不存在 | 路径写错 |
| `3` | 无有效数据 | 空表，或所有行都被清洗剔除 |
| `4` | 缺少必需列 | 报错会**列出缺失列名与实际表头** |
| `5` | 不是有效的 Excel 文件 | 0 字节文件、或 CSV 改名成 `.xlsx` |
| `6` | 输出路径与输入路径相同 | 防覆盖原始数据 |

所有错误都输出**中文可读提示**，不会把裸 traceback 抛给用户：

```
[错误] 输入文件不存在: data\raw\sales_raw.xlsx
       建议: 请用 --input 指定正确的 sales_raw.xlsx 路径。
```

---

## 命令行参数

```
python src/excel_report [-h] [-i INPUT] [-o OUTPUT] [-q] [-V]

  -i, --input    原始销售表路径（默认 data/raw/sales_raw.xlsx）
  -o, --output   报表输出路径（默认 output/report.xlsx）
  -q, --quiet    静默模式：仅在出错时输出
  -V, --version  显示版本
```

---

## 项目结构

```
exam-1-excel-report/
├── README.md                  # 本文件
├── CHANGELOG.md               # 版本管理（Keep a Changelog）
├── pyproject.toml             # PEP 621 依赖 + Ruff + pytest 配置
├── requirements.txt           # 依赖兜底安装
├── src/excel_report/          # 源码（src layout）
│   ├── __init__.py            # 版本号（三处同步之一）
│   ├── __main__.py            # python -m / 直接运行入口（含 src layout workaround）
│   ├── errors.py              # 领域异常，异常类自带 exit_code
│   ├── loader.py              # ★唯一读文件处（只读）+ 列名校验
│   ├── cleaning.py            # 清洗（纯函数）
│   ├── aggregate.py           # 统计（纯函数）
│   ├── report.py              # ★唯一写文件处 + 样式 + 时间戳规范化
│   ├── pipeline.py            # 流程编排
│   └── cli.py                 # 命令行入口，异常 → 退出码
├── tests/                     # 109 个测试
│   ├── conftest.py            # 共享夹具与工具
│   ├── test_cleaning.py       # 清洗规则
│   ├── test_aggregate.py      # 统计口径与排序
│   ├── test_errors.py         # 异常与退出码
│   ├── test_pipeline.py       # 集成：真实 xlsx 四表与样式
│   ├── test_invariants.py     # 不变量：只读 / 幂等 / 守恒 / 对账
│   └── test_version.py        # 版本一致性与依赖边界审查
├── scripts/
│   ├── make_sample_data.py    # 生成样例脏数据
│   ├── verify_report.py       # 核对四表与统计口径
│   └── diagnose_idempotency.py# 诊断"幂等性差在哪里"
├── data/raw/                  # 原始数据（脚本只读此处）
├── output/                    # 报表输出
├── docs/                      # 交付文档（见下）
└── research/                  # 网络深度查询调研报告
```

---

## 运行测试

```bash
pip install pytest
python -m pytest -q
```

预期：**109 passed**

> 💡 **为什么用 `python -m pytest` 而不是直接敲 `pytest`？**
> 用 `--user` 或用户级安装时，`pytest.exe` 会被放到
> `%APPDATA%\Python\PythonXX\Scripts`，该目录**默认不在 PATH** 里，
> 直接敲 `pytest` 会报 `无法将"pytest"项识别为 cmdlet`。
> `python -m pytest` 由解释器直接定位模块，**不依赖 PATH**，在任何环境都能跑。
> 两种写法等价；若你的 PATH 已包含 Scripts 目录，`pytest -q` 也能用。

```bash
# 只跑某类用例
python -m pytest tests/test_errors.py -v          # 异常与退出码
python -m pytest tests/test_invariants.py -v      # 只读 / 幂等 / 守恒
python -m pytest tests/test_cleaning.py -k date   # 日期解析
```

代码规范检查（需自行安装）：

```bash
pip install ruff
ruff check .
```

---

## 核对与诊断工具

```bash
# 核对四表、三方对账、占比合计、原始文件哈希、独立复算
python scripts/verify_report.py

# 诊断"为什么两次运行的输出哈希不同"
python src/excel_report -i data/raw/sales_raw.xlsx -o output/a.xlsx -q
python src/excel_report -i data/raw/sales_raw.xlsx -o output/b.xlsx -q
python scripts/diagnose_idempotency.py output/a.xlsx output/b.xlsx
```

---

## 保证与可复现性

| 保证 | 验证方式 |
|---|---|
| 原始文件永不被修改 | 运行前后 SHA256 + mtime 双断言；`loader` 无写 API 静态审查 |
| 输出**逐字节可复现** | 同一输入两次运行 SHA256 完全相同（时间戳已规范化） |
| 统计值可人工核对 | 三方对账（明细/销售员/地区）+ 月度合计 = 总额 |
| 依赖不越界 | 仅 `pandas` + `openpyxl` + `python-dateutil`，自动化审查 |

---

## 交付文档索引

| 文档 | 内容 |
|---|---|
| [`docs/01-需求澄清清单.md`](docs/01-需求澄清清单.md) | 目标、输入输出、边界约束、验收标准、决策记录、风险清单（**交给 AI 之前先写清**） |
| [`docs/02-报表用例.md`](docs/02-报表用例.md) | 29 条用例（主流程 4 / 边界 12 / 异常 6 / 不变量 7） |
| [`docs/03-开发方案与技术选型.md`](docs/03-开发方案与技术选型.md) | 分层架构、模块接口、版本管理规则、测试策略 |
| [`docs/04-AI协作过程记录.md`](docs/04-AI协作过程记录.md) | AI 协作轨迹、关键对话摘要、人工干预节点 |
| [`docs/05-代码修改diff.md`](docs/05-代码修改diff.md) | 9 项关键修改的前后对比与根因分析 |
| [`docs/06-测试用例与运行结果.md`](docs/06-测试用例与运行结果.md) | 109 个测试的执行结果与人工核对证据 |
| [`docs/07-复盘.md`](docs/07-复盘.md) | 文字复盘：哪些交给 AI、哪些自己拍板、哪些做了人工审查 |
| [`research/01-网络深度查询调研报告.md`](research/01-网络深度查询调研报告.md) | anySerach 深度查询：技术规范、工具链、8 个模范项目 |

---

## 环境要求

- **Python ≥ 3.10**（考核要求）
- 依赖：`pandas ≥ 2.0`、`openpyxl ≥ 3.1`、`python-dateutil ≥ 2.8`

> 实测环境：Python 3.12.14 + pandas 3.0.1 + openpyxl 3.1.5 + pytest 9.1.1
