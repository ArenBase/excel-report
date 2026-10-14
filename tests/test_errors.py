"""异常用例 TC-17 ~ TC-22（docs/02-报表用例.md §3）。

每个用例断言两件事：**异常类型** + **退出码**，确保 CLI 行为可预期。
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from conftest import write_xlsx
from excel_report import pipeline
from excel_report.cli import main
from excel_report.errors import (
    EmptyDataError,
    InputFileNotFoundError,
    InvalidWorkbookError,
    MissingColumnsError,
    NoValidDataError,
    SamePathError,
)

ALL = ["日期", "销售员", "产品", "数量", "单价", "地区"]
ONE_ROW = [
    {"日期": "2024-01-05", "销售员": "张三", "产品": "A", "数量": 2, "单价": 10.0, "地区": "华东"}
]


def test_run_returns_result_with_four_frames(normal_xlsx: Path, tmp_path: Path) -> None:
    """基线：正常输入确实能跑通（对照组）。"""
    result = pipeline.run(normal_xlsx, tmp_path / "out" / "report.xlsx")

    assert result.output_path.exists()
    assert set(pd.ExcelFile(result.output_path).sheet_names) == {
        "明细清洗后",
        "按销售员",
        "按地区",
        "总览",
    }


# --- TC-17 输入文件不存在（退出码 2） ----------------------------------------
def test_error_input_file_not_found(tmp_path: Path) -> None:
    missing = tmp_path / "不存在.xlsx"

    with pytest.raises(InputFileNotFoundError) as excinfo:
        pipeline.run(missing, tmp_path / "report.xlsx")

    assert excinfo.value.exit_code == 2
    assert "不存在.xlsx" in str(excinfo.value), "报错信息必须包含实际查找的路径"


def test_error_input_is_directory(tmp_path: Path) -> None:
    directory = tmp_path / "adir"
    directory.mkdir()

    with pytest.raises(InputFileNotFoundError):
        pipeline.run(directory, tmp_path / "report.xlsx")


def test_cli_exit_code_2_for_missing_input(tmp_path: Path) -> None:
    code = main(["--input", str(tmp_path / "nope.xlsx"), "--output", str(tmp_path / "r.xlsx")])

    assert code == 2


# --- TC-18 空表 / 空文件（退出码 3 / 5） -------------------------------------
def test_error_empty_sheet(tmp_path: Path) -> None:
    """有表头但 0 行数据。"""
    path = tmp_path / "empty.xlsx"
    pd.DataFrame(columns=ALL).to_excel(path, index=False, engine="openpyxl")

    with pytest.raises(EmptyDataError) as excinfo:
        pipeline.run(path, tmp_path / "report.xlsx")

    assert excinfo.value.exit_code == 3
    assert "没有任何数据行" in str(excinfo.value)


def test_error_zero_byte_file(tmp_path: Path) -> None:
    """0 字节文件不是有效 Excel（退出码 5）。"""
    path = tmp_path / "broken.xlsx"
    path.write_bytes(b"")

    with pytest.raises(InvalidWorkbookError) as excinfo:
        pipeline.run(path, tmp_path / "report.xlsx")

    assert excinfo.value.exit_code == 5


def test_error_non_excel_content(tmp_path: Path) -> None:
    """实为 CSV 但被改名成 .xlsx（退出码 5）。"""
    path = tmp_path / "fake.xlsx"
    path.write_text("日期,销售员,产品,数量,单价,地区\n2024-01-05,张三,A,2,10,华东\n", encoding="utf-8")

    with pytest.raises(InvalidWorkbookError) as excinfo:
        pipeline.run(path, tmp_path / "report.xlsx")

    assert excinfo.value.exit_code == 5


def test_cli_exit_code_5_for_invalid_workbook(tmp_path: Path) -> None:
    path = tmp_path / "broken.xlsx"
    path.write_bytes(b"not an excel file at all")

    code = main(["--input", str(path), "--output", str(tmp_path / "r.xlsx")])

    assert code == 5


# --- TC-19 全部行被剔除 → 无有效数据（退出码 3，且不得生成报表） ---------------
def test_error_no_valid_data_and_no_output_written(tmp_path: Path) -> None:
    records = [
        {**ONE_ROW[0], "销售员": None},
        {**ONE_ROW[0], "数量": -3},
        {**ONE_ROW[0], "日期": "bad-date"},
    ]
    path = write_xlsx(tmp_path / "all_dirty.xlsx", records, columns=ALL)
    output = tmp_path / "report.xlsx"

    with pytest.raises(NoValidDataError) as excinfo:
        pipeline.run(path, output)

    assert excinfo.value.exit_code == 3
    assert not output.exists(), "无有效数据时不得生成（部分）报表 —— 决策 D-05"


def test_cli_exit_code_3_for_no_valid_data(tmp_path: Path) -> None:
    path = write_xlsx(tmp_path / "dirty.xlsx", [{**ONE_ROW[0], "数量": 0}], columns=ALL)

    code = main(["--input", str(path), "--output", str(tmp_path / "r.xlsx")])

    assert code == 3
    assert not (tmp_path / "r.xlsx").exists()


# --- TC-20 缺少必需列（退出码 4） --------------------------------------------
@pytest.mark.parametrize("dropped_column", ["单价", "数量", "日期", "销售员", "地区"])
def test_error_missing_required_column(dropped_column: str, tmp_path: Path) -> None:
    columns = [c for c in ALL if c != dropped_column]
    path = write_xlsx(tmp_path / "missing.xlsx", ONE_ROW, columns=columns)

    with pytest.raises(MissingColumnsError) as excinfo:
        pipeline.run(path, tmp_path / "report.xlsx")

    error = excinfo.value
    assert error.exit_code == 4
    assert dropped_column in error.missing, "报错必须列出缺失的列名"
    assert dropped_column not in error.actual, "报错必须列出实际表头"


def test_missing_product_column_is_tolerated(tmp_path: Path) -> None:
    """产品不是必需列：整列缺失时应正常出报表，且明细里补 '未填写'。"""
    columns = [c for c in ALL if c != "产品"]
    path = write_xlsx(tmp_path / "no_product.xlsx", ONE_ROW, columns=columns)

    result = pipeline.run(path, tmp_path / "report.xlsx")

    assert result.detail["产品"].tolist() == ["未填写"]


def test_cli_exit_code_4_for_missing_column(tmp_path: Path) -> None:
    columns = [c for c in ALL if c != "单价"]
    path = write_xlsx(tmp_path / "missing.xlsx", ONE_ROW, columns=columns)

    code = main(["--input", str(path), "--output", str(tmp_path / "r.xlsx")])

    assert code == 4


# --- TC-22 输出路径与输入路径相同（退出码 6，防覆写原文件） -------------------
def test_error_same_path_rejected(normal_xlsx: Path) -> None:
    with pytest.raises(SamePathError) as excinfo:
        pipeline.run(normal_xlsx, normal_xlsx)

    assert excinfo.value.exit_code == 6


def test_cli_exit_code_6_for_same_path(normal_xlsx: Path) -> None:
    code = main(["--input", str(normal_xlsx), "--output", str(normal_xlsx)])

    assert code == 6


def test_same_path_check_uses_absolute_paths(tmp_path: Path) -> None:
    """用相对路径绕道也应被拦住（输出文件尚不存在的情况）。"""
    path = write_xlsx(tmp_path / "sales.xlsx", ONE_ROW, columns=ALL)
    relative = Path("sales.xlsx")

    with pytest.raises(SamePathError):
        pipeline.run(relative, relative)

    assert path.exists()


# --- 退出码 0 与 1 -----------------------------------------------------------
def test_cli_exit_code_0_on_success(normal_xlsx: Path, tmp_path: Path, capsys) -> None:
    code = main(
        [
            "--input",
            str(normal_xlsx),
            "--output",
            str(tmp_path / "report.xlsx"),
        ]
    )

    assert code == 0
    out = capsys.readouterr().out
    assert "报表已生成" in out
    assert "总销售额: 150.00" in out


def test_cli_quiet_mode_suppresses_summary(normal_xlsx: Path, tmp_path: Path, capsys) -> None:
    code = main(
        [
            "--input",
            str(normal_xlsx),
            "--output",
            str(tmp_path / "report.xlsx"),
            "--quiet",
        ]
    )

    assert code == 0
    assert capsys.readouterr().out == ""


def test_cli_unexpected_exception_returns_1(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """兜底分支：任何未预期异常都应转成退出码 1 而不是让 traceback 裸奔。"""

    def boom(*_args: object, **_kwargs: object) -> None:
        raise RuntimeError("模拟未预期故障")

    monkeypatch.setattr("excel_report.cli.run", boom)

    code = main(["--input", "x.xlsx", "--output", "y.xlsx"])

    assert code == 1


def test_cli_error_message_is_readable_not_traceback(tmp_path: Path, capsys) -> None:
    """AC-05：面向用户的提示必须是可读中文，且不含 Traceback。"""
    code = main(["--input", str(tmp_path / "nope.xlsx"), "--output", str(tmp_path / "r.xlsx")])

    err = capsys.readouterr().err
    assert code == 2
    assert "[错误]" in err
    assert "Traceback" not in err


def test_cli_version_flag(capsys) -> None:
    from excel_report import __version__

    with pytest.raises(SystemExit) as excinfo:
        main(["--version"])

    assert excinfo.value.code == 0
    assert __version__ in capsys.readouterr().out
