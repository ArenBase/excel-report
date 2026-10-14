"""集成用例 TC-01 ~ TC-02、TC-29：真实 xlsx 落盘 → 读回断言四表。"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from conftest import (
    EXPECTED_BY_PERSON,
    EXPECTED_BY_REGION,
    EXPECTED_MONTHLY,
    EXPECTED_TOTAL,
    NORMAL_RECORDS,
    write_xlsx,
)
from excel_report import pipeline
from excel_report.report import (
    SHEET_BY_PERSON,
    SHEET_BY_REGION,
    SHEET_DETAIL,
    SHEET_ORDER,
    SHEET_OVERVIEW,
)

ALL = ["日期", "销售员", "产品", "数量", "单价", "地区"]


def test_pipeline_generates_four_sheets_in_order(normal_xlsx: Path, tmp_path: Path) -> None:
    """TC-01 / AC-02：工作表名称与顺序必须精确匹配，不多不少。"""
    output = tmp_path / "report.xlsx"

    pipeline.run(normal_xlsx, output)

    assert pd.ExcelFile(output, engine="openpyxl").sheet_names == list(SHEET_ORDER)


def test_pipeline_detail_sheet_contents(normal_xlsx: Path, tmp_path: Path) -> None:
    """TC-01：明细表的行数与销售额逐行正确。"""
    output = tmp_path / "report.xlsx"
    pipeline.run(normal_xlsx, output)
    detail = pd.read_excel(output, sheet_name=SHEET_DETAIL, engine="openpyxl")

    assert len(detail) == len(NORMAL_RECORDS)
    assert sorted(detail["销售额"].tolist()) == [20.00, 30.00, 100.00]
    assert list(detail.columns) == [*ALL, "销售额"], "明细表列顺序含派生的销售额列"
    assert "Unnamed: 0" not in detail.columns, "不得写出 DataFrame 索引列"


def test_pipeline_summary_sheets_match_manual_calculation(normal_xlsx: Path, tmp_path: Path) -> None:
    """TC-01 / AC-04：读回的报表数值与人工手算完全一致。"""
    output = tmp_path / "report.xlsx"
    pipeline.run(normal_xlsx, output)

    by_person = pd.read_excel(output, sheet_name=SHEET_BY_PERSON, engine="openpyxl")
    by_region = pd.read_excel(output, sheet_name=SHEET_BY_REGION, engine="openpyxl")

    assert dict(zip(by_person["销售员"], by_person["销售额"])) == EXPECTED_BY_PERSON
    assert dict(zip(by_region["地区"], by_region["销售额"])) == EXPECTED_BY_REGION
    assert by_person["销售额"].sum() == EXPECTED_TOTAL
    assert by_region["销售额"].sum() == EXPECTED_TOTAL


def test_pipeline_overview_contains_total_and_monthly(normal_xlsx: Path, tmp_path: Path) -> None:
    """TC-01：总览表必须含总额与月度总额（考题原文要求"给出月度总额"）。"""
    output = tmp_path / "report.xlsx"
    pipeline.run(normal_xlsx, output)
    overview = pd.read_excel(output, sheet_name=SHEET_OVERVIEW, engine="openpyxl")

    as_dict = {
        str(row["指标"]): row["数值"]
        for _, row in overview.iterrows()
        if pd.notna(row["指标"]) and str(row["指标"]).strip()
    }
    assert list(overview.columns) == ["指标", "数值"]
    assert float(as_dict["总销售额"]) == EXPECTED_TOTAL
    assert "原始行数" in as_dict
    assert "清洗后行数" in as_dict
    for month, amount in EXPECTED_MONTHLY.items():
        assert month in as_dict, f"总览缺少月度 {month}"
        assert float(as_dict[month]) == amount


def test_pipeline_overview_reports_dropped_reasons(tmp_path: Path) -> None:
    """总览表必须能对账：原始 = 清洗后 + 各剔除原因之和。"""
    records = [
        {**NORMAL_RECORDS[0]},
        dict(NORMAL_RECORDS[0]),  # 重复
        {**NORMAL_RECORDS[1], "销售员": None},  # 缺失
        {**NORMAL_RECORDS[2], "数量": -1},  # 数量<=0
    ]
    path = write_xlsx(tmp_path / "dirty.xlsx", records, columns=ALL)
    output = tmp_path / "report.xlsx"

    pipeline.run(path, output)
    overview = pd.read_excel(output, sheet_name=SHEET_OVERVIEW, engine="openpyxl")
    as_dict = {
        str(row["指标"]): row["数值"]
        for _, row in overview.iterrows()
        if pd.notna(row["指标"]) and str(row["指标"]).strip()
    }

    raw_rows = int(as_dict["原始行数"])
    cleaned_rows = int(as_dict["清洗后行数"])
    dropped = int(as_dict["剔除行数合计"])
    reasons = sum(
        int(as_dict[key])
        for key in ("剔除·关键字段缺失", "剔除·数量<=0", "剔除·全字段重复", "剔除·异常数值或日期")
    )

    assert raw_rows == len(records)
    assert cleaned_rows + dropped == raw_rows
    assert reasons == dropped


def test_pipeline_creates_missing_output_directories(tmp_path: Path) -> None:
    """TC-28 / C-06：输出目录不存在时自动创建（含多层）。"""
    path = write_xlsx(tmp_path / "sales.xlsx", NORMAL_RECORDS, columns=ALL)
    output = tmp_path / "a" / "b" / "c" / "report.xlsx"

    result = pipeline.run(path, output)

    assert output.exists()
    assert result.output_path == output


def test_pipeline_overwrites_existing_report(tmp_path: Path) -> None:
    """TC-29 / C-07：旧报表被覆盖，且不留旧工作表。"""
    path = write_xlsx(tmp_path / "sales.xlsx", NORMAL_RECORDS, columns=ALL)
    output = tmp_path / "report.xlsx"
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame({"旧数据": [1, 2, 3]}).to_excel(writer, sheet_name="旧工作表", index=False)

    pipeline.run(path, output)

    assert pd.ExcelFile(output, engine="openpyxl").sheet_names == list(SHEET_ORDER)


def test_pipeline_handles_multiple_sheets_picks_first_non_empty(tmp_path: Path) -> None:
    """第一个工作表为空时回退到第一个非空表（loader 的选表策略）。"""
    path = tmp_path / "multi.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame(columns=ALL).to_excel(writer, sheet_name="空表", index=False)
        pd.DataFrame(NORMAL_RECORDS).to_excel(writer, sheet_name="数据", index=False)

    result = pipeline.run(path, tmp_path / "report.xlsx")

    assert result.stats.cleaned_rows == len(NORMAL_RECORDS)


def test_pipeline_ignores_extra_unrelated_columns(tmp_path: Path) -> None:
    """原始表多出无关列时不应影响统计（只保留业务列）。"""
    records = [{**record, "备注": "无关列"} for record in NORMAL_RECORDS]
    path = write_xlsx(tmp_path / "extra.xlsx", records, columns=[*ALL, "备注"])

    result = pipeline.run(path, tmp_path / "report.xlsx")

    assert list(result.detail.columns) == [*ALL, "销售额"]
    assert result.total_sales == EXPECTED_TOTAL


def test_pipeline_styles_applied(normal_xlsx: Path, tmp_path: Path) -> None:
    """样式验收：冻结首行、自动筛选、表头加粗、金额列数字格式。"""
    from openpyxl import load_workbook

    output = tmp_path / "report.xlsx"
    pipeline.run(normal_xlsx, output)

    book = load_workbook(output)
    for name in SHEET_ORDER:
        worksheet = book[name]
        assert worksheet.freeze_panes == "A2", f"{name} 未冻结首行"
        assert worksheet.auto_filter.ref is not None, f"{name} 未设置自动筛选"
        assert worksheet.cell(row=1, column=1).font.bold is True, f"{name} 表头未加粗"
        # 列宽必须大于纯表头长度，说明自适应生效
        assert worksheet.column_dimensions["A"].width > 0

    detail = book[SHEET_DETAIL]
    headers = [cell.value for cell in detail[1]]
    amount_index = headers.index("销售额") + 1
    assert detail.cell(row=2, column=amount_index).number_format == "#,##0.00"


def test_pipeline_column_width_accounts_for_cjk(normal_xlsx: Path, tmp_path: Path) -> None:
    """中文列宽按东亚字符宽度 2 计算，避免中文被截断显示。"""
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter

    output = tmp_path / "report.xlsx"
    pipeline.run(normal_xlsx, output)

    worksheet = load_workbook(output)[SHEET_BY_PERSON]
    headers = [cell.value for cell in worksheet[1]]
    index = headers.index("销售员") + 1
    letter = get_column_letter(index)

    # "销售员" 表头显示宽度 = 6，列宽至少 7（含内边距）
    assert worksheet.column_dimensions[letter].width > len("销售员")
