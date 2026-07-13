#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通用Excel估算文件生成脚本 v2.0
从metadata.json生成Excel估算文件（匹配用户的新Excel格式）

新格式支持：
- sub_systems：一个功能可以有多个系统实现
- 需求类型列
- 系统列：企微、PC后台等（而不是用户端/管理后台）
- Excel公式：=ROUNDUP(SUM(...)*0.15,0)
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# 共享模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from logging_setup import setup_logging

setup_logging()
log = logging.getLogger(__name__)

# ── 布局/样式常量 ────────────────────────────────────────────────
SHEET_NAME = "工时估算"

# 新列布局: A=#, B=模块, C=功能, D=需求类型, E=描述, F=系统, G=工时(开发), H=备注
COL_SEQ = 0
COL_MOD = 1
COL_NAME = 2
COL_TYPE = 3
COL_DESC = 4
COL_SYS = 5
COL_EFFORT = 6
COL_NOTE = 7
NUM_COLS = 8

HEADERS = ["#", "模块", "功能", "需求类型", "描述", "系统", "工时(人天)", "备注"]

# 列宽
COLUMN_WIDTHS = {
    COL_SEQ: 5,
    COL_MOD: 15,
    COL_NAME: 30,
    COL_TYPE: 10,
    COL_DESC: 40,
    COL_SYS: 10,
    COL_EFFORT: 12,
    COL_NOTE: 30,
}

# 配色
COLOR_HEADER_BG = "#4472C4"
COLOR_WHITE = "#FFFFFF"
COLOR_BLACK = "#000000"


def round_half(value: float) -> float:
    """将数值四舍五入到0.5的整数倍"""
    return round(value * 2) / 2


def load_metadata(metadata_path: str) -> dict:
    """加载metadata.json文件"""
    with open(metadata_path, "r", encoding="utf-8") as f:
        return json.load(f)


def _build_formats(wb: xlsxwriter.Workbook) -> dict:
    """构造所有单元格样式"""
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
        "merge_center": wb.add_format({
            "border": 1, "border_color": COLOR_BLACK,
            "align": "center", "valign": "vcenter",
        }),
        "total": wb.add_format({"bold": True, "valign": "vcenter"}),
        "total_center": wb.add_format({
            "bold": True, "align": "center", "valign": "vcenter",
        }),
    }


def _collect_rows(metadata: dict, interactive: bool = False) -> list[tuple]:
    """
    从 metadata 收集所有数据行。

    新格式：每个feature可以有多个sub_systems（多系统实现）
    返回: (seq, module, name, type, desc, system, effort, note)
    """
    features = metadata.get("features", [])

    rows: list[tuple[int, str, str, str, str, str, float, str]] = []
    seq_num = 0

    for feature in features:
        module = feature.get("module", "")
        name = feature.get("name", "")
        req_type = feature.get("requirement_type", "")

        # 获取sub_systems（新格式）或单个系统（旧格式兼容）
        sub_systems = feature.get("sub_systems", [])
        if not sub_systems:
            # 旧格式兼容：单个系统
            sub_systems = [{
                "system": feature.get("system", ""),
                "effort_days": feature.get("effort_days", 0),
                "description": feature.get("description", ""),
                "reference": feature.get("reference", ""),
            }]

        # 同一功能的所有sub_system共享同一个seq_num
        seq_num += 1

        # 为每个sub_system创建一行
        first_row = True
        for sub in sub_systems:
            if not isinstance(sub, dict):
                continue

            # 第一行显示完整信息，续行只显示系统+工时+备注
            if first_row:
                rows.append((
                    seq_num,
                    module,
                    name,
                    req_type,
                    sub.get("description", ""),
                    sub.get("system", ""),
                    round_half(sub.get("effort_days", 0)),
                    sub.get("reference", ""),
                ))
                first_row = False
            else:
                rows.append((
                    seq_num,
                    "",  # 模块列空
                    "",  # 功能列空
                    "",  # 需求类型列空
                    "",  # 描述列空
                    sub.get("system", ""),
                    round_half(sub.get("effort_days", 0)),
                    sub.get("reference", ""),
                ))

    # 添加其他工时行
    rows = _add_other_rows(rows, interactive)
    return rows


def _add_other_rows(rows: list, interactive: bool) -> list:
    """添加其他工时行：功能测试、验收发布、需求沟通、UI设计"""
    import builtins
    input_func = builtins.input

    if interactive:
        print("\n=== 其他工时配置 ===")

    seq_num = len(rows)

    # 功能测试（公式行）
    seq_num += 1
    rows.append((
        seq_num,
        "功能测试",
        "",
        "",
        "集成测试，系统测试",
        "",
        None,  # None表示使用公式
        "",
    ))

    # 验收发布（默认值改为5）
    if interactive:
        print("验收发布工时（人天）[默认5，直接回车使用默认值]:")
        acceptance_input = input_func().strip()
        acceptance = float(acceptance_input) if acceptance_input else 5.0
    else:
        acceptance = 5.0

    if acceptance > 0:
        seq_num += 1
        rows.append((
            seq_num,
            "验收发布",
            "",
            "",
            "验收技术支持，系统发布",
            "",
            round_half(acceptance),
            "",
        ))

    # 需求沟通（默认值改为8）
    if interactive:
        print("需求沟通工时（人天）[默认8，直接回车使用默认值]:")
        communication_input = input_func().strip()
        communication = float(communication_input) if communication_input else 8.0
    else:
        communication = 8.0

    if communication > 0:
        seq_num += 1
        rows.append((
            seq_num,
            "需求沟通",
            "",
            "",
            "沟通，需求文档",
            "",
            round_half(communication),
            "",
        ))

    # UI设计（可选）
    if interactive:
        print("UI/UX设计工时（人天）[不需要请直接回车]:")
        ui_input = input_func().strip()
        if ui_input:
            ui_effort = float(ui_input)
            if ui_effort > 0:
                seq_num += 1
                rows.append((
                    seq_num,
                    "UI/UX设计",
                    "",
                    "",
                    "界面设计，交互设计",
                    "",
                    round_half(ui_effort),
                    "",
                ))

    return rows


def _apply_merges(ws, rows: list[tuple], fmts: dict) -> None:
    """
    对数据行应用合并单元格。

    合并规则：
    - 同一功能的所有行（相同seq_num）合并：模块列、功能列、需求类型列、描述列
    """
    total = len(rows)
    if total == 0:
        return

    # 按seq_num分组（直接使用rows数组中的值，不从单元格读取）
    seq_groups = {}
    for i, row in enumerate(rows):
        seq_num = row[0]  # 第1列是seq_num
        # 只跳过真正的空行，但包含有seq_num但其他列为空的行
        if seq_num is None or seq_num == "":
            continue
        if seq_num not in seq_groups:
            seq_groups[seq_num] = []
        seq_groups[seq_num].append(i)

    # 对每个seq组进行合并
    for seq_val, row_indices in seq_groups.items():
        if len(row_indices) <= 1:
            continue  # 只有一行，不需要合并

        first_row_idx = row_indices[0]
        last_row_idx = row_indices[-1]

        # 获取第一行的值（用于合并）
        mod_val = rows[first_row_idx][1]
        name_val = rows[first_row_idx][2]
        type_val = rows[first_row_idx][3]
        desc_val = rows[first_row_idx][4]

        # Excel行号（1-based）
        r_start = first_row_idx + 1
        r_end = last_row_idx + 1

        # 合并 # 列（A列）
        ws.merge_range(r_start, COL_SEQ, r_end, COL_SEQ,
                       seq_val, fmts["merge_center"])

        # 合并 模块 列（B列）
        ws.merge_range(r_start, COL_MOD, r_end, COL_MOD,
                       mod_val, fmts["merge_center"])

        # 合并 功能 列（C列）
        ws.merge_range(r_start, COL_NAME, r_end, COL_NAME,
                       name_val, fmts["merge_center"])

        # 合并 需求类型 列（D列）
        ws.merge_range(r_start, COL_TYPE, r_end, COL_TYPE,
                       type_val, fmts["merge_center"])

        # 合并 描述 列（E列）
        ws.merge_range(r_start, COL_DESC, r_end, COL_DESC,
                       desc_val, fmts["merge_center"])


def _write_formulas(ws, rows: list[tuple], fmts: dict) -> None:
    """
    写入Excel公式：
    - 功能测试行：=ROUNDUP(SUM(开发工时)*0.15,0)
    - 合计行：=ROUNDUP(SUM(所有工时),0)
    """
    # 找到功能测试行（effort=None的那行）
    test_row_idx = None
    for i, row in enumerate(rows):
        if row[1] == "功能测试" and row[6] is None:
            test_row_idx = i
            break

    if test_row_idx is None:
        log.warning("Function test row not found, skipping formula")
        return

    # 写入功能测试公式
    test_row_num = test_row_idx + 1  # +1: header 在第 0 行
    effort_col_letter = xl_col_to_name(COL_EFFORT)

    # 计算开发工时范围（功能测试行之前的所有行）
    first_data_row = 2  # Excel 1-based: header=1, 首数据行=2
    last_dev_row = test_row_num  # 功能测试行

    # 写入公式: =ROUNDUP(SUM(开发工时)*0.15,0)
    test_formula = f"=ROUNDUP(SUM({effort_col_letter}{first_data_row}:{effort_col_letter}{last_dev_row})*0.15,0)"
    ws.write_formula(
        test_row_num, COL_EFFORT,
        test_formula,
        fmts["total_center"],
    )

    # 合计行
    total_row = len(rows) + 1  # 0-based，header 在第 0 行

    # 写入"合计"文本到各列
    for c in range(NUM_COLS):
        if c == COL_NOTE:  # 备注列写"合计"
            ws.write(total_row, c, "合计", fmts["total_center"])
        else:
            ws.write(total_row, c, "", fmts["total_center"])

    # 写入合计公式: =ROUNDUP(SUM(所有工时),0)
    total_formula = f"=ROUNDUP(SUM({effort_col_letter}{first_data_row}:{effort_col_letter}{total_row}),0)"
    ws.write_formula(
        total_row, COL_EFFORT,
        total_formula,
        fmts["total"],
    )


def generate_excel(metadata_path: str, output_path: str | None = None, interactive: bool = False) -> str:
    """
    从metadata.json生成Excel估算文件。

    Args:
        metadata_path: metadata.json文件路径
        output_path:   输出Excel文件路径（可选）
        interactive:   是否交互式询问UI设计、验收发布、需求沟通工时
    """
    metadata = load_metadata(metadata_path)

    if output_path is None:
        project_name = metadata.get("name", "项目")
        metadata_dir = os.path.dirname(metadata_path)
        output_path = os.path.join(metadata_dir, f"{project_name}估算.xlsx")

    log.info("Generating Excel: %s", output_path)

    wb = xlsxwriter.Workbook(output_path)
    ws = wb.add_worksheet(SHEET_NAME)
    fmts = _build_formats(wb)

    # 表头
    for c, h in enumerate(HEADERS):
        ws.write(0, c, h, fmts["header"])

    # 数据行
    rows = _collect_rows(metadata, interactive=interactive)
    for i, (seq, module, name, req_type, desc, system, effort, note) in enumerate(rows):
        r = i + 1  # header 在第 0 行

        ws.write(r, COL_SEQ, seq if seq else "", fmts["merge_center"])
        ws.write(r, COL_MOD, module, fmts["cell"])
        ws.write(r, COL_NAME, name, fmts["cell"])
        ws.write(r, COL_TYPE, req_type, fmts["cell"])
        ws.write(r, COL_DESC, desc, fmts["cell"])
        ws.write(r, COL_SYS, system, fmts["cell_center"])

        # 功能测试行 effort=None：写公式（稍后处理）
        if effort is not None:
            ws.write(r, COL_EFFORT, effort, fmts["cell_center"])

        ws.write(r, COL_NOTE, note, fmts["cell"])

    # 合并单元格 + 写入公式
    _apply_merges(ws, rows, fmts)
    _write_formulas(ws, rows, fmts)

    # 列宽 + 冻结首行
    for col, width in COLUMN_WIDTHS.items():
        ws.set_column(col, col, width)
    ws.freeze_panes(1, 0)

    wb.close()
    log.info("Excel file generated: %s", output_path)
    return output_path


def main() -> int:
    parser = argparse.ArgumentParser(
        description="从 metadata.json 生成 Excel 估算文件 v2.0（支持sub_systems）"
    )
    parser.add_argument("metadata_path", help="metadata.json 路径")
    parser.add_argument("output_path", nargs="?", default=None,
                        help="输出 Excel 路径（可选）")
    parser.add_argument("-i", "--interactive", action="store_true",
                        help="交互式询问UI设计、验收发布、需求沟通工时")
    args = parser.parse_args()

    if not os.path.exists(args.metadata_path):
        log.error("File not found: %s", args.metadata_path)
        return 1

    try:
        output_file = generate_excel(args.metadata_path, args.output_path, interactive=args.interactive)
        print(f"\nDone! Excel file generated: {output_file}")
        return 0
    except Exception as e:
        print(f"\nError: {str(e)}\n")
        log.exception("Excel generation failed")
        return 1


if __name__ == "__main__":
    sys.exit(main())
