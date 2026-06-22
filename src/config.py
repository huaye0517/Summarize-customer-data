# -*- coding: utf-8 -*-
"""Excel 汇总工具模板配置。"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List


@dataclass(frozen=True)
class TemplateConfig:
    """单种 Excel 模板的汇总配置。"""

    id: str
    label: str
    required_columns: List[str]
    group_by: List[str]
    sum_columns: List[str]
    output_columns: List[str]
    keep_columns: List[str] = field(default_factory=list)
    name_filter_column: str = ""
    dimension_label: str = ""
    excel_column_widths: Dict[str, int] = field(default_factory=dict)
    preview_column_widths: Dict[str, int] = field(default_factory=dict)


INTERNAL_TEMPLATE = TemplateConfig(
    id="internal",
    label="内部销售单",
    required_columns=["销售渠道", "货品名称", "数量", "金额"],
    group_by=["销售渠道", "货品名称"],
    sum_columns=["数量", "金额"],
    output_columns=["销售渠道", "货品名称", "数量", "金额"],
    name_filter_column="货品名称",
    dimension_label="销售渠道数",
    excel_column_widths={
        "销售渠道": 44,
        "货品名称": 90,
        "数量": 14,
        "金额": 16,
    },
    preview_column_widths={
        "销售渠道": 460,
        "货品名称": 820,
        "数量": 110,
        "金额": 130,
    },
)

EXTERNAL_TEMPLATE = TemplateConfig(
    id="external",
    label="外部销售单",
    required_columns=["客户", "产品名称", "数量", "金额", "销售"],
    group_by=["客户", "产品名称"],
    sum_columns=["数量", "金额"],
    keep_columns=["销售"],
    output_columns=["客户", "产品名称", "数量", "金额", "销售"],
    name_filter_column="产品名称",
    dimension_label="客户数",
    excel_column_widths={
        "客户": 28,
        "产品名称": 75,
        "数量": 14,
        "金额": 16,
        "销售": 14,
    },
    preview_column_widths={
        "客户": 220,
        "产品名称": 680,
        "数量": 110,
        "金额": 130,
        "销售": 120,
    },
)

TEMPLATES: List[TemplateConfig] = [INTERNAL_TEMPLATE, EXTERNAL_TEMPLATE]

# 界面预览最大行数
PREVIEW_MAX_ROWS = 200
