# -*- coding: utf-8 -*-
"""Excel 汇总工具列名配置。"""

# 输入文件必须包含的列
REQUIRED_COLUMNS = ["销售渠道", "货品名称", "数量", "金额"]

# 分组依据
GROUP_BY_COLUMNS = ["销售渠道", "货品名称"]

# 需要求和的列
SUM_COLUMNS = ["数量", "金额"]

# 输出列顺序
OUTPUT_COLUMNS = ["销售渠道", "货品名称", "数量", "金额"]

# 界面预览最大行数
PREVIEW_MAX_ROWS = 200
