# -*- coding: utf-8 -*-
"""销售单 Excel 汇总核心逻辑。"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Union

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

from config import (
    GROUP_BY_COLUMNS,
    OUTPUT_COLUMNS,
    REQUIRED_COLUMNS,
    SUM_COLUMNS,
)


@dataclass
class AggregateResult:
    """汇总结果。"""

    data: pd.DataFrame
    raw_row_count: int
    result_row_count: int


class AggregateError(Exception):
    """汇总过程中的业务错误。"""


def _validate_columns(df: pd.DataFrame) -> None:
    """校验必需列是否存在。"""
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise AggregateError(f"缺少必需列：{', '.join(missing)}")


def _convert_numeric_column(df: pd.DataFrame, column: str) -> None:
    """将指定列转为数值，并收集无效行信息。"""
    converted = pd.to_numeric(df[column], errors="coerce")
    invalid_mask = converted.isna() & df[column].notna()
    invalid_mask = invalid_mask & (df[column].astype(str).str.strip() != "")

    if invalid_mask.any():
        # Excel 行号 = DataFrame 索引 + 2（含表头）
        invalid_rows = [int(idx) + 2 for idx in df.index[invalid_mask].tolist()[:10]]
        suffix = "..." if invalid_mask.sum() > 10 else ""
        raise AggregateError(
            f"列「{column}」存在无法转换为数字的数据，"
            f"示例行号：{', '.join(map(str, invalid_rows))}{suffix}"
        )

    df[column] = converted.fillna(0)


def _format_output_values(df: pd.DataFrame) -> pd.DataFrame:
    """格式化输出数值：数量为整数时去小数，金额保留两位小数。"""
    result = df.copy()

    qty = result["数量"]
    result["数量"] = qty.apply(
        lambda value: int(value) if float(value).is_integer() else round(float(value), 4)
    )

    result["金额"] = result["金额"].apply(lambda value: round(float(value), 2))
    return result


def aggregate_excel(input_path: Union[str, Path]) -> AggregateResult:
    """
    读取 Excel 并按「销售渠道 + 货品名称」汇总数量与金额。

    :param input_path: 输入 Excel 路径
    :return: 汇总结果
    """
    path = Path(input_path)
    if not path.exists():
        raise AggregateError(f"文件不存在：{path}")

    if path.suffix.lower() not in {".xlsx", ".xls"}:
        raise AggregateError("仅支持 .xlsx 或 .xls 格式的 Excel 文件")

    try:
        df = pd.read_excel(path, sheet_name=0)
    except Exception as exc:
        raise AggregateError(f"读取 Excel 失败：{exc}") from exc

    if df.empty:
        raise AggregateError("Excel 中没有数据行")

    _validate_columns(df)

    work_df = df[REQUIRED_COLUMNS].copy()
    raw_row_count = len(work_df)

    # 过滤货品名称为空的行
    work_df["货品名称"] = work_df["货品名称"].astype(str).str.strip()
    work_df["销售渠道"] = work_df["销售渠道"].astype(str).str.strip()
    work_df = work_df[work_df["货品名称"] != ""]

    if work_df.empty:
        raise AggregateError("过滤空货品名称后没有可汇总的数据")

    for column in SUM_COLUMNS:
        _convert_numeric_column(work_df, column)

    grouped = (
        work_df.groupby(GROUP_BY_COLUMNS, as_index=False)[SUM_COLUMNS]
        .sum()
        .loc[:, OUTPUT_COLUMNS]
    )

    grouped = grouped.sort_values(
        by=GROUP_BY_COLUMNS,
        kind="stable",
    ).reset_index(drop=True)

    grouped = _format_output_values(grouped)

    return AggregateResult(
        data=grouped,
        raw_row_count=raw_row_count,
        result_row_count=len(grouped),
    )


def export_excel(result: AggregateResult, output_path: Union[str, Path]) -> None:
    """导出汇总结果到 Excel。"""
    path = Path(output_path)
    if path.suffix.lower() != ".xlsx":
        path = path.with_suffix(".xlsx")

    # Excel 列宽（字符数），便于阅读长货品名称
    excel_column_widths = {
        "销售渠道": 44,
        "货品名称": 90,
        "数量": 14,
        "金额": 16,
    }

    try:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            result.data.to_excel(writer, index=False, sheet_name="汇总结果")
            worksheet = writer.sheets["汇总结果"]

            # 表头样式：灰色底、深绿色加粗居中文字
            header_fill = PatternFill(fill_type="solid", fgColor="D9D9D9")
            header_font = Font(name="Calibri", size=10, bold=True, color="006400")
            header_alignment = Alignment(horizontal="center", vertical="center")

            for index, column in enumerate(OUTPUT_COLUMNS, start=1):
                cell = worksheet.cell(row=1, column=index)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_alignment

                column_letter = cell.column_letter
                worksheet.column_dimensions[column_letter].width = excel_column_widths.get(
                    column, 16
                )

            worksheet.row_dimensions[1].height = 18
            worksheet.auto_filter.ref = worksheet.dimensions
            worksheet.freeze_panes = "A2"
    except Exception as exc:
        raise AggregateError(f"导出 Excel 失败：{exc}") from exc
