#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通用Excel估算文件生成脚本
从metadata.json生成Excel估算文件（使用xlsxwriter确保边框正确）

schema 校验复用共享模块 metadata_schema（与 archive_estimation.py 同源，
消除此前两份复制粘贴且已漂移的校验逻辑）。
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# 共享模块：schema 校验的单一真相源
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from metadata_schema import ValidationError, validate_metadata_schema  # noqa: E402
from logging_setup import setup_logging  # noqa: E402

setup_logging()
log = logging.getLogger(__name__)


# ── 布局/样式常量（替代散落的 magic literal）─────────────────────────
SHEET_NAME = "工时估算"

# 列布局: A=#, B=系统, C=模块, D=功能, E=工时, F=描述
COL_SEQ = 0
COL_SYS = 1
COL_MOD = 2
COL_NAME = 3
COL_EFFORT = 4
COL_DESC = 5
NUM_COLS = 6

HEADERS = ["#", "系统", "模块", "功能", "工时估算(人天)", "描述"]

# 列宽
COLUMN_WIDTHS = {
    COL_SEQ: 5,
    COL_SYS: 15,
    COL_MOD: 20,
    COL_NAME: 30,
    COL_EFFORT: 15,
    COL_DESC: 40,
}

# 配色
COLOR_HEADER_BG = "#4472C4"
COLOR_WHITE = "#FFFFFF"
COLOR_BLACK = "#000000"


def round_half(value: float) -> float:
    """将数值四舍五入到0.5的整数倍（最小估算单位0.5人天）"""
    return round(value * 2) / 2


def load_metadata(metadata_path: str) -> dict:
    """加载metadata.json文件"""
    with open(metadata_path, "r", encoding="utf-8") as f:
        return json.load(f)


def _build_formats(wb: xlsxwriter.Workbook) -> dict:
    """构造所有单元格样式，返回格式字典。"""
    return {
        "header": wb.add_format({
            "bold": True, "font_color": COLOR_WHITE, "bg_color": COLOR_HEADER_BG,
            "border": 1, "border_color": COLOR_BLACK,
            "align": "center", "valign": "vcenter", "text_wrap": True,
        }),
        "cell": wb.add_format({
            "border": 1, "border_color": COLOR_BLACK,
            "valign": "vcenter", "text_wrap": True,
        }),
        "cell_center": wb.add_format({
            "border": 1, "border_color": COLOR_BLACK,
            "align": "center", "valign": "vcenter",
        }),
        "merge": wb.add_format({
            "border": 1, "border_color": COLOR_BLACK,
            "align": "left", "valign": "vcenter", "text_wrap": True,
        }),
        "merge_center": wb.add_format({
            "border": 1, "border_color": COLOR_BLACK,
            "align": "center", "valign": "vcenter",
        }),
        "total": wb.add_format({"bold": True, "valign": "vcenter"}),
        "total_center": wb.add_format({
            "bold": True, "align": "center", "valign": "vcenter",
        }),
    }


def _collect_rows(metadata: dict) -> list[tuple]:
    """
    从 metadata 收集所有数据行（功能行 + 其他工时行）。
    每行: (seq, system, module, name, effort, desc)
    """
    features = metadata.get("features", [])
    summary = metadata.get("summary", {})
    other_fixed = summary.get("other_fixed", {})

    rows: list[tuple[int, str, str, str, float, str]] = []
    seq_num = 0
    prev_system_module = None

    for feature in features:
        system = feature.get("system", "")
        module = feature.get("module", "")
        key = (system, module)
        if key != prev_system_module:
            seq_num += 1
            prev_system_module = key
        rows.append((
            seq_num, system, module,
            feature.get("name", ""),
            round_half(feature.get("effort_days", 0)),
            feature.get("description", ""),
        ))

    # UI/UX 设计工时（兼容 ui_ux_design / ui_design 两种字段名）
    ui_effort_total = other_fixed.get("ui_ux_design", other_fixed.get("ui_design", 0))
    ui_desc = other_fixed.get("ui_design_desc", "")
    if not ui_desc and ui_effort_total > 0:
        ui_parts = []
        for f in features:
            ue = f.get("ui_effort", 0)
            if ue > 0:
                ui_parts.append(f"{f.get('name', '')}({ue}人天)")
        ui_desc = "、".join(ui_parts) if ui_parts else "界面设计、交互设计"

    other_items = [
        ("UI/UX设计", ui_effort_total, ui_desc),
        ("系统测试", summary.get("testing_effort", 0), "集成测试、系统测试"),
        ("需求沟通", other_fixed.get("communication", 0), "沟通、需求文档"),
        ("环境搭建", other_fixed.get("environment", 0), "开发/测试/生产环境"),
        ("验收发布", other_fixed.get("acceptance", 0), "项目验收和上线发布"),
    ]
    for label, effort, desc in other_items:
        if effort > 0:
            seq_num += 1
            rows.append((seq_num, "其他", label, "", round_half(effort), desc))

    return rows


def _apply_merges(ws, rows: list[tuple], fmts: dict) -> None:
    """对数据行应用合并单元格：先按 (system,module) 合并 # 和模块列，再按 system 合并系统列。"""
    total = len(rows)
    if total == 0:
        return

    # 第一轮：按 (system, module) 分组，合并 # 列与 模块 列
    i = 0
    while i < total:
        sys_val = rows[i][1]
        mod_val = rows[i][2]
        group_end = i
        while group_end + 1 < total and rows[group_end + 1][1] == sys_val \
                and rows[group_end + 1][2] == mod_val:
            group_end += 1
        if group_end > i:
            r_start, r_end = i + 1, group_end + 1  # +1: header 占第 0 行
            ws.merge_range(r_start, COL_SEQ, r_end, COL_SEQ,
                           rows[i][0], fmts["merge_center"])
            ws.merge_range(r_start, COL_MOD, r_end, COL_MOD,
                           mod_val, fmts["merge"])
        i = group_end + 1

    # 第二轮：按 system 合并系统列（跨模块）
    i = 0
    while i < total:
        sys_val = rows[i][1]
        sys_end = i
        while sys_end + 1 < total and rows[sys_end + 1][1] == sys_val:
            sys_end += 1
        if sys_end > i:
            ws.merge_range(i + 1, COL_SYS, sys_end + 1, COL_SYS,
                           sys_val, fmts["merge"])
        i = sys_end + 1


def _write_totals(ws, rows: list[tuple], fmts: dict) -> None:
    """写总计行：用 SUM 公式（同时提供缓存值，便于不支持公式重算的读取器）。"""
    total_row = len(rows) + 1  # 0-based，header 在第 0 行
    for c in range(NUM_COLS):
        ws.write(total_row, c, "", fmts["total_center"])

    total_effort = sum(r[4] for r in rows)
    effort_col_letter = xl_col_to_name(COL_EFFORT)
    first_data_row = 2          # Excel 1-based: header=1, 首数据行=2
    last_data_row = len(rows) + 1
    ws.write_formula(
        total_row, COL_EFFORT,
        f"=SUM({effort_col_letter}{first_data_row}:{effort_col_letter}{last_data_row})",
        fmts["total"], total_effort,
    )
    ws.write(total_row, COL_DESC, "人天", fmts["total"])


def generate_excel(metadata_path: str, output_path: str | None = None) -> str:
    """
    从metadata.json生成Excel估算文件。

    Args:
        metadata_path: metadata.json文件路径
        output_path:   输出Excel文件路径（可选；缺省时输出到 metadata 同目录）

    Raises:
        ValidationError: metadata schema 校验失败时抛出（由调用方决定如何处理）
    """
    metadata = load_metadata(metadata_path)

    # === Schema 校验（复用共享模块）===
    log.info("Validating metadata schema...")
    validate_metadata_schema(metadata)
    log.info("   [OK] Schema validation passed")

    if output_path is None:
        project_name = metadata.get("name", "项目")
        metadata_dir = os.path.dirname(metadata_path)
        output_path = os.path.join(metadata_dir, f"{project_name}估算.xlsx")

    wb = xlsxwriter.Workbook(output_path)
    ws = wb.add_worksheet(SHEET_NAME)
    fmts = _build_formats(wb)

    # 表头
    for c, h in enumerate(HEADERS):
        ws.write(0, c, h, fmts["header"])

    # 数据行
    rows = _collect_rows(metadata)
    for i, (seq, system, module, name, effort, desc) in enumerate(rows):
        r = i + 1  # header 在第 0 行
        ws.write(r, COL_SEQ, seq, fmts["merge_center"])
        ws.write(r, COL_SYS, system, fmts["cell"])
        ws.write(r, COL_MOD, module, fmts["cell"])
        ws.write(r, COL_NAME, name, fmts["cell"])
        ws.write(r, COL_EFFORT, effort, fmts["cell_center"])
        ws.write(r, COL_DESC, desc, fmts["cell"])

    # 合并单元格 + 总计
    _apply_merges(ws, rows, fmts)
    _write_totals(ws, rows, fmts)

    # 列宽 + 冻结首行
    for col, width in COLUMN_WIDTHS.items():
        ws.set_column(col, col, width)
    ws.freeze_panes(1, 0)

    wb.close()
    log.info("Excel file generated: %s", output_path)
    return output_path


def main() -> int:
    parser = argparse.ArgumentParser(
        description="从 metadata.json 生成 Excel 估算文件"
    )
    parser.add_argument("metadata_path", help="metadata.json 路径")
    parser.add_argument("output_path", nargs="?", default=None,
                        help="输出 Excel 路径（可选）")
    args = parser.parse_args()

    if not os.path.exists(args.metadata_path):
        log.error("File not found: %s", args.metadata_path)
        return 1

    try:
        output_file = generate_excel(args.metadata_path, args.output_path)
        print(f"\nDone! Excel file generated: {output_file}")
        return 0
    except ValidationError as e:
        print(f"\n{str(e)}\n")
        log.error("Excel generation cancelled: fix the errors above and retry")
        return 1


if __name__ == "__main__":
    sys.exit(main())
