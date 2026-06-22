# -*- coding: utf-8 -*-
"""销售单 Excel 汇总桌面工具。"""

from __future__ import annotations

import sys
import threading
from pathlib import Path

# 确保从项目根目录运行 `python src/main.py` 时也能正确导入
SRC_DIR = Path(__file__).resolve().parent
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

from aggregator import AggregateError, AggregateResult, aggregate_excel, export_excel
from config import INTERNAL_TEMPLATE, PREVIEW_MAX_ROWS, TEMPLATES

# 界面主题色（与导出 Excel 表头深绿色保持一致）
COLOR_PRIMARY = "#006400"
COLOR_PRIMARY_HOVER = "#004D00"
COLOR_PRIMARY_LIGHT = "#E8F5E9"
COLOR_BG = "#F3F6F4"
COLOR_CARD = "#FFFFFF"
COLOR_BORDER = "#DDE3DF"
COLOR_TEXT = "#1F2933"
COLOR_TEXT_MUTED = "#6B7280"
COLOR_HEADER_BG = "#D9D9D9"
FONT_FAMILY = "Microsoft YaHei UI"

TEMPLATE_MAP = {template.id: template for template in TEMPLATES}


class ExcelAggregatorApp:
    """Excel 汇总桌面应用。"""

    def __init__(self, root: ctk.CTk) -> None:
        self.root = root
        self.root.title("客户产品 Excel 汇总工具")
        self.root.geometry("1400x720")
        self.root.minsize(1200, 580)
        self.root.configure(fg_color=COLOR_BG)

        self.input_path: Path | None = None
        self.result: AggregateResult | None = None
        self._processing = False
        self.preview_columns: list[str] = list(INTERNAL_TEMPLATE.output_columns)
        self.tree: ttk.Treeview | None = None
        self.table_wrap: ctk.CTkFrame | None = None

        self._setup_styles()
        self._build_ui()

    def _setup_styles(self) -> None:
        """配置表格样式，与导出表头风格一致。"""
        style = ttk.Style()
        style.theme_use("clam")

        style.configure(
            "Result.Treeview",
            background=COLOR_CARD,
            fieldbackground=COLOR_CARD,
            foreground=COLOR_TEXT,
            rowheight=30,
            font=(FONT_FAMILY, 10),
            borderwidth=0,
        )
        style.configure(
            "Result.Treeview.Heading",
            background=COLOR_HEADER_BG,
            foreground=COLOR_PRIMARY,
            font=(FONT_FAMILY, 10, "bold"),
            relief="flat",
            padding=(8, 6),
        )
        style.map(
            "Result.Treeview",
            background=[("selected", COLOR_PRIMARY_LIGHT)],
            foreground=[("selected", COLOR_PRIMARY)],
        )
        style.map(
            "Result.Treeview.Heading",
            background=[("active", "#CFCFCF")],
        )

    def _build_ui(self) -> None:
        """构建界面。"""
        self._build_header()

        content = ctk.CTkFrame(self.root, fg_color="transparent")
        content.pack(fill=tk.BOTH, expand=True, padx=20, pady=(16, 20))

        self._build_file_card(content)
        self._build_status_card(content)
        self._build_preview_card(content)

    def _build_header(self) -> None:
        """顶部标题栏。"""
        header = ctk.CTkFrame(
            self.root,
            fg_color=COLOR_PRIMARY,
            corner_radius=0,
            height=78,
        )
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        title_wrap = ctk.CTkFrame(header, fg_color="transparent")
        title_wrap.pack(side=tk.LEFT, padx=24, pady=14)

        ctk.CTkLabel(
            title_wrap,
            text="客户产品 Excel 汇总工具",
            font=(FONT_FAMILY, 24, "bold"),
            text_color="#FFFFFF",
        ).pack(anchor=tk.W)

        ctk.CTkLabel(
            title_wrap,
            text="支持内部/外部销售单 · 自动识别格式 · 一键汇总导出",
            font=(FONT_FAMILY, 12),
            text_color="#B7DFB7",
        ).pack(anchor=tk.W, pady=(4, 0))

    def _create_card(self, parent: ctk.CTkFrame, title: str) -> ctk.CTkFrame:
        """创建统一样式的卡片容器。"""
        card = ctk.CTkFrame(
            parent,
            fg_color=COLOR_CARD,
            corner_radius=14,
            border_width=1,
            border_color=COLOR_BORDER,
        )
        card.pack(fill=tk.X, pady=(0, 14))

        ctk.CTkLabel(
            card,
            text=title,
            font=(FONT_FAMILY, 14, "bold"),
            text_color=COLOR_PRIMARY,
        ).pack(anchor=tk.W, padx=18, pady=(14, 8))

        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill=tk.BOTH, expand=True, padx=18, pady=(0, 16))
        return body

    def _build_file_card(self, parent: ctk.CTkFrame) -> None:
        """文件选择与操作区。"""
        body = self._create_card(parent, "文件操作")

        path_row = ctk.CTkFrame(body, fg_color="transparent")
        path_row.pack(fill=tk.X)
        path_row.columnconfigure(0, weight=1)

        self.file_var = tk.StringVar(value="尚未选择文件")
        self.file_entry = ctk.CTkEntry(
            path_row,
            textvariable=self.file_var,
            height=40,
            corner_radius=10,
            border_color=COLOR_BORDER,
            fg_color="#FAFBFA",
            text_color=COLOR_TEXT,
            font=(FONT_FAMILY, 11),
            state="readonly",
        )
        self.file_entry.grid(row=0, column=0, sticky="ew", padx=(0, 12))

        button_row = ctk.CTkFrame(path_row, fg_color="transparent")
        button_row.grid(row=0, column=1, sticky="e")

        ctk.CTkButton(
            button_row,
            text="选择 Excel",
            width=110,
            height=40,
            corner_radius=10,
            fg_color="#FFFFFF",
            hover_color="#EEF2EE",
            border_width=1,
            border_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            font=(FONT_FAMILY, 12, "bold"),
            command=self._choose_file,
        ).pack(side=tk.LEFT, padx=(0, 8))

        self.aggregate_button = ctk.CTkButton(
            button_row,
            text="开始汇总",
            width=110,
            height=40,
            corner_radius=10,
            fg_color=COLOR_PRIMARY,
            hover_color=COLOR_PRIMARY_HOVER,
            font=(FONT_FAMILY, 12, "bold"),
            command=self._start_aggregate,
            state="disabled",
        )
        self.aggregate_button.pack(side=tk.LEFT, padx=(0, 8))

        self.export_button = ctk.CTkButton(
            button_row,
            text="导出结果",
            width=110,
            height=40,
            corner_radius=10,
            fg_color="#2E7D32",
            hover_color="#1B5E20",
            font=(FONT_FAMILY, 12, "bold"),
            command=self._export_result,
            state="disabled",
        )
        self.export_button.pack(side=tk.LEFT)

    def _build_status_card(self, parent: ctk.CTkFrame) -> None:
        """状态信息区。"""
        body = self._create_card(parent, "汇总状态")

        status_row = ctk.CTkFrame(body, fg_color="transparent")
        status_row.pack(fill=tk.X)

        self.status_dot = ctk.CTkLabel(
            status_row,
            text="●",
            font=(FONT_FAMILY, 14),
            text_color=COLOR_TEXT_MUTED,
            width=20,
        )
        self.status_dot.pack(side=tk.LEFT, padx=(0, 8))

        self.status_var = tk.StringVar(value="请选择销售单 Excel 文件后开始汇总")
        ctk.CTkLabel(
            status_row,
            textvariable=self.status_var,
            font=(FONT_FAMILY, 12),
            text_color=COLOR_TEXT,
            anchor=tk.W,
            justify=tk.LEFT,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.progress_frame = ctk.CTkFrame(body, fg_color="transparent")
        self.progress_frame.pack(fill=tk.X, pady=(10, 0))

        progress_top = ctk.CTkFrame(self.progress_frame, fg_color="transparent")
        progress_top.pack(fill=tk.X)

        self.progress_label_var = tk.StringVar(value="0%")
        ctk.CTkLabel(
            progress_top,
            textvariable=self.progress_label_var,
            font=(FONT_FAMILY, 11, "bold"),
            text_color=COLOR_PRIMARY,
            width=52,
        ).pack(side=tk.LEFT)

        self.progress_bar = ctk.CTkProgressBar(
            progress_top,
            height=16,
            corner_radius=8,
            progress_color=COLOR_PRIMARY,
            fg_color="#E5EAE6",
        )
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))
        self.progress_bar.set(0)

        self.progress_detail_var = tk.StringVar(value="等待开始汇总")
        ctk.CTkLabel(
            self.progress_frame,
            textvariable=self.progress_detail_var,
            font=(FONT_FAMILY, 10),
            text_color=COLOR_TEXT_MUTED,
            anchor=tk.W,
        ).pack(fill=tk.X, pady=(6, 0))

        stats_row = ctk.CTkFrame(body, fg_color="transparent")
        stats_row.pack(fill=tk.X, pady=(12, 0))

        self.raw_count_label = self._create_stat_badge(stats_row, "原始行数", "--")
        self.result_count_label = self._create_stat_badge(stats_row, "汇总行数", "--")
        self.dimension_title_var = tk.StringVar(value="销售渠道数")
        self.dimension_count_label = self._create_stat_badge(
            stats_row,
            self.dimension_title_var.get(),
            "--",
            title_var=self.dimension_title_var,
        )

    def _create_stat_badge(
        self,
        parent: ctk.CTkFrame,
        title: str,
        value: str,
        title_var: tk.StringVar | None = None,
    ) -> ctk.CTkLabel:
        """创建统计徽章。"""
        badge = ctk.CTkFrame(
            parent,
            fg_color=COLOR_PRIMARY_LIGHT,
            corner_radius=10,
            border_width=1,
            border_color="#C8E6C9",
        )
        badge.pack(side=tk.LEFT, padx=(0, 10))

        ctk.CTkLabel(
            badge,
            textvariable=title_var if title_var is not None else None,
            text=None if title_var is not None else title,
            font=(FONT_FAMILY, 10),
            text_color=COLOR_TEXT_MUTED,
        ).pack(padx=14, pady=(8, 0))

        value_label = ctk.CTkLabel(
            badge,
            text=value,
            font=(FONT_FAMILY, 18, "bold"),
            text_color=COLOR_PRIMARY,
        )
        value_label.pack(padx=14, pady=(0, 8))
        return value_label

    def _build_preview_card(self, parent: ctk.CTkFrame) -> None:
        """结果预览区。"""
        card = ctk.CTkFrame(
            parent,
            fg_color=COLOR_CARD,
            corner_radius=14,
            border_width=1,
            border_color=COLOR_BORDER,
        )
        card.pack(fill=tk.BOTH, expand=True)

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill=tk.X, padx=18, pady=(14, 8))

        ctk.CTkLabel(
            header,
            text="结果预览",
            font=(FONT_FAMILY, 14, "bold"),
            text_color=COLOR_PRIMARY,
        ).pack(side=tk.LEFT)

        ctk.CTkLabel(
            header,
            text=f"最多显示前 {PREVIEW_MAX_ROWS} 行",
            font=(FONT_FAMILY, 11),
            text_color=COLOR_TEXT_MUTED,
        ).pack(side=tk.RIGHT)

        self.table_wrap = ctk.CTkFrame(
            card,
            fg_color="#FAFBFA",
            corner_radius=10,
            border_width=1,
            border_color=COLOR_BORDER,
        )
        self.table_wrap.pack(fill=tk.BOTH, expand=True, padx=18, pady=(0, 16))
        self.table_wrap.rowconfigure(0, weight=1)
        self.table_wrap.columnconfigure(0, weight=1)

        self._rebuild_preview_table(self.preview_columns, INTERNAL_TEMPLATE.id)

    def _rebuild_preview_table(self, columns: list[str], template_id: str) -> None:
        """根据汇总结果动态重建预览表格列。"""
        if self.table_wrap is None:
            return

        for widget in self.table_wrap.winfo_children():
            widget.destroy()

        self.preview_columns = list(columns)
        template = TEMPLATE_MAP.get(template_id, INTERNAL_TEMPLATE)
        column_widths = template.preview_column_widths
        stretch_columns = set(template.group_by)

        self.tree = ttk.Treeview(
            self.table_wrap,
            columns=columns,
            show="headings",
            style="Result.Treeview",
        )

        center_columns = {"数量", "金额"}
        for column in columns:
            self.tree.heading(column, text=column)
            anchor = tk.CENTER if column in center_columns else tk.W
            self.tree.column(
                column,
                width=column_widths.get(column, 140),
                minwidth=column_widths.get(column, 120),
                anchor=anchor,
                stretch=column in stretch_columns,
            )

        self.tree.tag_configure("odd", background="#FAFBFA")
        self.tree.tag_configure("even", background="#FFFFFF")

        y_scroll = ttk.Scrollbar(
            self.table_wrap,
            orient=tk.VERTICAL,
            command=self.tree.yview,
        )
        x_scroll = ttk.Scrollbar(
            self.table_wrap,
            orient=tk.HORIZONTAL,
            command=self.tree.xview,
        )
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

    def _set_status(self, message: str, tone: str = "normal") -> None:
        """更新状态文字与指示点颜色。"""
        colors = {
            "normal": COLOR_TEXT_MUTED,
            "info": "#1976D2",
            "success": COLOR_PRIMARY,
            "error": "#C62828",
            "working": "#F57C00",
        }
        self.status_var.set(message)
        self.status_dot.configure(text_color=colors.get(tone, COLOR_TEXT_MUTED))

    def _reset_progress(self) -> None:
        """重置进度条。"""
        self.progress_bar.set(0)
        self.progress_label_var.set("0%")
        self.progress_detail_var.set("等待开始汇总")

    def _report_progress(self, percent: float, message: str) -> None:
        """线程安全地更新汇总进度。"""
        safe_percent = max(0.0, min(float(percent), 1.0))

        def update() -> None:
            self.progress_bar.set(safe_percent)
            self.progress_label_var.set(f"{int(safe_percent * 100)}%")
            self.progress_detail_var.set(message)
            self._set_status(message, "working")
            self.root.update_idletasks()

        self.root.after(0, update)

    def _finish_progress(self, success: bool = True) -> None:
        """汇总结束后的进度条状态。"""
        if success:
            self.progress_bar.set(1)
            self.progress_label_var.set("100%")
            self.progress_detail_var.set("汇总完成")
        else:
            self._reset_progress()

    def _reset_stats(self) -> None:
        """重置统计徽章。"""
        self.raw_count_label.configure(text="--")
        self.result_count_label.configure(text="--")
        self.dimension_count_label.configure(text="--")
        self.dimension_title_var.set("销售渠道数")

    def _update_stats(self, result: AggregateResult) -> None:
        """更新统计徽章。"""
        template = TEMPLATE_MAP.get(result.template_id, INTERNAL_TEMPLATE)
        dimension_column = template.group_by[0]
        dimension_count = result.data[dimension_column].nunique()

        self.raw_count_label.configure(text=str(result.raw_row_count))
        self.result_count_label.configure(text=str(result.result_row_count))
        self.dimension_title_var.set(result.dimension_label)
        self.dimension_count_label.configure(text=str(dimension_count))

    def _choose_file(self) -> None:
        """选择输入 Excel 文件。"""
        file_path = filedialog.askopenfilename(
            title="选择销售单 Excel",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xls"),
                ("所有文件", "*.*"),
            ],
        )
        if not file_path:
            return

        self.input_path = Path(file_path)
        self.file_var.set(str(self.input_path))
        self.result = None
        self._rebuild_preview_table(
            list(INTERNAL_TEMPLATE.output_columns),
            INTERNAL_TEMPLATE.id,
        )
        self._reset_stats()
        self.export_button.configure(state="disabled")
        self.aggregate_button.configure(state="normal")
        self._set_status("已选择文件，点击「开始汇总」处理数据", "info")

    def _start_aggregate(self) -> None:
        """在后台线程执行汇总，避免界面卡顿。"""
        if self._processing or not self.input_path:
            return

        self._processing = True
        self.aggregate_button.configure(state="disabled")
        self.export_button.configure(state="disabled")
        self._reset_progress()
        self._report_progress(0.01, "正在启动汇总任务...")
        self.root.update_idletasks()

        thread = threading.Thread(target=self._run_aggregate, daemon=True)
        thread.start()

    def _run_aggregate(self) -> None:
        """后台汇总任务。"""
        assert self.input_path is not None

        try:
            result = aggregate_excel(
                self.input_path,
                progress_callback=self._report_progress,
            )
        except AggregateError as exc:
            self.root.after(0, lambda: self._on_aggregate_failed(str(exc)))
            return
        except Exception as exc:
            self.root.after(0, lambda: self._on_aggregate_failed(f"汇总失败：{exc}"))
            return

        self.root.after(0, lambda: self._on_aggregate_success(result))

    def _on_aggregate_success(self, result: AggregateResult) -> None:
        """汇总成功后的界面更新。"""
        self.result = result
        self._processing = False
        self.aggregate_button.configure(state="normal")
        self.export_button.configure(state="normal")
        self._update_stats(result)
        self._rebuild_preview_table(
            list(result.data.columns),
            result.template_id,
        )
        self._fill_preview(result.data.head(PREVIEW_MAX_ROWS))
        self._finish_progress(success=True)

        preview_tip = ""
        if len(result.data) > PREVIEW_MAX_ROWS:
            preview_tip = f"，预览前 {PREVIEW_MAX_ROWS} 行"

        self._set_status(
            f"已识别：{result.template_label} | 汇总完成："
            f"原始 {result.raw_row_count} 行 -> 汇总后 {result.result_row_count} 行"
            f"{preview_tip}",
            "success",
        )

    def _on_aggregate_failed(self, message: str) -> None:
        """汇总失败后的界面更新。"""
        self._processing = False
        self.result = None
        self.aggregate_button.configure(
            state="normal" if self.input_path else "disabled"
        )
        self.export_button.configure(state="disabled")
        self._rebuild_preview_table(
            list(INTERNAL_TEMPLATE.output_columns),
            INTERNAL_TEMPLATE.id,
        )
        self._reset_stats()
        self._finish_progress(success=False)
        self._set_status(message, "error")
        messagebox.showerror("汇总失败", message)

    def _fill_preview(self, df: pd.DataFrame) -> None:
        """填充预览表格。"""
        if self.tree is None:
            return

        for index, (_, row) in enumerate(df.iterrows()):
            values = [row[column] for column in self.preview_columns]
            tag = "even" if index % 2 == 0 else "odd"
            self.tree.insert("", tk.END, values=values, tags=(tag,))

    def _export_result(self) -> None:
        """导出汇总结果。"""
        if not self.result:
            messagebox.showwarning("提示", "请先完成汇总")
            return

        default_name = "汇总结果.xlsx"
        if self.input_path:
            default_name = f"{self.input_path.stem}_汇总结果.xlsx"

        output_path = filedialog.asksaveasfilename(
            title="保存汇总结果",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not output_path:
            return

        try:
            export_excel(self.result, output_path)
        except AggregateError as exc:
            messagebox.showerror("导出失败", str(exc))
            return

        self._set_status(
            f"导出成功：{output_path}（共 {self.result.result_row_count} 行）",
            "success",
        )
        messagebox.showinfo("导出成功", f"已保存到：\n{output_path}")


def main() -> None:
    """程序入口。"""
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("green")

    root = ctk.CTk()
    ExcelAggregatorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
