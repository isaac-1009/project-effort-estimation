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
    """将数值四舍五入到0.5的整数倍（最小估算单位0.5人天）"""
    return round(value * 2) / 2


def load_metadata(metadata_path: str) -> dict:
    """加载metadata.json文件"""
    with open(metadata_path, "r", encoding="utf-8") as f:
        return json.load(f)


def _build_formats(wb: xlsxwriter.Workbook) -> dict:
    """构造所有单元格样式，返回格式字典。

    对齐规则（按用户要求）：
    - A/B/C/D/E/F/H 列：左对齐 + 上对齐
    - G 列（工时）：右对齐 + 上对齐
    - 表头：居中 + 垂直居中（保持可读性）
    """
    return {
        "header": wb.add_format({
            "bold": True, "font_color": COLOR_WHITE, "bg_color": COLOR_HEADER_BG,
            "border": 1, "border_color": COLOR_BLACK,
            "align": "center", "valign": "vcenter", "text_wrap": True,
        }),
        # 普通单元格：左对齐 + 上对齐（A/B/C/D/E/F/H 列）
        "cell": wb.add_format({
            "border": 1, "border_color": COLOR_BLACK,
            "align": "left", "valign": "top", "text_wrap": True,
        }),
        # 工时单元格：右对齐 + 上对齐（G 列）
        "cell_effort": wb.add_format({
            "border": 1, "border_color": COLOR_BLACK,
            "align": "right", "valign": "top",
        }),
        # 系统单元格：左对齐 + 上对齐（F 列）
        "cell_system": wb.add_format({
            "border": 1, "border_color": COLOR_BLACK,
            "align": "left", "valign": "top",
        }),
        # 合并单元格：左对齐 + 上对齐
        "merge": wb.add_format({
            "border": 1, "border_color": COLOR_BLACK,
            "align": "left", "valign": "top", "text_wrap": True,
        }),
        # 合并单元格（序号/其他行 C-E）：居中 + 上对齐
        "merge_center": wb.add_format({
            "border": 1, "border_color": COLOR_BLACK,
            "align": "center", "valign": "top", "text_wrap": True,
        }),
        # 合计行工时：右对齐 + 上对齐
        "total": wb.add_format({
            "bold": True, "align": "right", "valign": "top",
        }),
        # 合计行其他单元格：左对齐 + 上对齐
        "total_left": wb.add_format({
            "bold": True, "align": "left", "valign": "top",
        }),
    }


def _collect_rows(metadata: dict, interactive: bool = False) -> list[tuple]:
    """
    从 metadata 收集所有数据行（功能行 + 其他工时行）。
    每行: (seq, module, name, type, desc, system, effort, note)

    新格式支持：
    - 一个功能可能有多行（如企微+PC后台）
    - 不再按系统分组，而是按功能分组
    - 增加需求类型列和备注列
    - 交互式询问UI设计、验收发布、需求沟通工时
    - 支持 sub_systems 结构（新格式）和直接字段（旧格式）
    """
    features = metadata.get("features", [])
    summary = metadata.get("summary", {})

    rows: list[tuple[int, str, str, str, str, str, float, str]] = []
    seq_num = 0
    prev_module_name = None

    # 收集功能行
    for feature in features:
        # 检查是否使用新结构（sub_systems）
        if "sub_systems" in feature:
            # 新结构：从 sub_systems 数组展开多行
            module = feature.get("module", "")
            name = feature.get("name", "")
            req_type = feature.get("requirement_type", "")
            key = (module, name)
            if key != prev_module_name:
                seq_num += 1
                prev_module_name = key

            # 遍历 sub_systems，每个系统一行
            # 描述优先用 sub_system 自己的，没有则回退到 feature 级描述
            feature_desc = feature.get("description", "")
            for sub in feature.get("sub_systems", []):
                rows.append((
                    seq_num,
                    module,
                    name,
                    req_type,
                    sub.get("description") or feature_desc,
                    sub.get("system", ""),
                    round_half(sub.get("effort_days", 0)),
                    sub.get("remark", ""),
                ))
        else:
            # 旧结构：直接从 feature 读取字段
            module = feature.get("module", "")
            name = feature.get("name", "")
            key = (module, name)
            if key != prev_module_name:
                seq_num += 1
                prev_module_name = key

            # 空行占位（一个功能可能有多行）
            rows.append((
                seq_num, module, name,
                feature.get("requirement_type", ""),
                feature.get("description", ""),
                feature.get("system", ""),
                round_half(feature.get("effort_days", 0)),
                feature.get("remark", ""),
            ))

    # 交互式询问其他工时
    other_rows = _collect_other_rows(metadata, rows, interactive, seq_num)
    return other_rows


def _collect_other_rows(metadata: dict, rows: list, interactive: bool, seq_num: int) -> list:
    """
    收集其他工时行（功能测试、验收发布、需求沟通、UI设计）。
    支持交互式询问。
    """
    if interactive:
        return _interactive_collect_other_rows(metadata, rows, seq_num)

    # 非交互模式：从metadata读取
    summary = metadata.get("summary", {})
    other_fixed = summary.get("other_fixed", {})

    # 功能测试（公式行，effort=None）
    seq_num += 1
    rows.append((
        seq_num, "功能测试", "-", "", "集成测试，系统测试", "-",
        None,  # None表示使用公式
        "",
    ))

    # 验收发布
    acceptance = other_fixed.get("acceptance", 1)
    if acceptance > 0:
        seq_num += 1
        rows.append((
            seq_num, "验收发布", "-", "", "验收技术支持，系统发布", "-",
            round_half(acceptance), "",
        ))

    # 需求沟通
    communication = other_fixed.get("communication", 2)
    if communication > 0:
        seq_num += 1
        rows.append((
            seq_num, "需求沟通", "-", "", "沟通，需求文档", "-",
            round_half(communication), "",
        ))

    # UI设计
    ui_effort = other_fixed.get("ui_ux_design", other_fixed.get("ui_design", 0))
    if ui_effort > 0:
        seq_num += 1
        rows.append((
            seq_num, "UI/UX设计", "-", "", "界面设计，交互设计", "-",
            round_half(ui_effort), "",
        ))

    return rows


def _interactive_collect_other_rows(metadata: dict, rows: list, seq_num: int) -> list:
    """
    交互式收集其他工时行。
    询问用户：UI设计、验收发布、需求沟通工时。
    """
    import builtins
    input_func = builtins.input

    print("\n=== 其他工时配置 ===")

    # 功能测试（公式行）
    seq_num += 1
    rows.append((
        seq_num, "功能测试", "-", "", "集成测试，系统测试", "-",
        None,  # None表示使用公式
        "",
    ))

    # 验收发布
    print("验收发布工时（人天）[默认1，直接回车使用默认值]:")
    acceptance_input = input_func().strip()
    acceptance = float(acceptance_input) if acceptance_input else 1.0
    if acceptance > 0:
        seq_num += 1
        rows.append((
            seq_num, "验收发布", "-", "", "验收技术支持，系统发布", "-",
            round_half(acceptance), "",
        ))

    # 需求沟通
    print("需求沟通工时（人天）[默认2，直接回车使用默认值]:")
    communication_input = input_func().strip()
    communication = float(communication_input) if communication_input else 2.0
    if communication > 0:
        seq_num += 1
        rows.append((
            seq_num, "需求沟通", "-", "", "沟通，需求文档", "-",
            round_half(communication), "",
        ))

    # UI设计（可选）
    print("UI/UX设计工时（人天）[不需要请直接回车]:")
    ui_input = input_func().strip()
    if ui_input:
        ui_effort = float(ui_input)
        if ui_effort > 0:
            seq_num += 1
            rows.append((
                seq_num, "UI/UX设计", "-", "", "界面设计，交互设计", "-",
                round_half(ui_effort), "",
            ))

    return rows


def _merge_consecutive(ws, rows: list[tuple], col: int, fmt_key: str,
                       fmts: dict, is_other: bool = False) -> None:
    """
    对指定列按连续相同值合并（仅在该列值非空时）。
    用于 B(模块)、D(需求类型) 列的连续合并。
    """
    total = len(rows)
    if total == 0:
        return
    i = 0
    while i < total:
        val = rows[i][col]
        if val in (None, ""):
            i += 1
            continue
        end = i
        while end + 1 < total and rows[end + 1][col] == val:
            end += 1
        if end > i:
            r_start, r_end = i + 1, end + 1  # +1: header 占第 0 行
            ws.merge_range(r_start, col, r_end, col, val, fmts[fmt_key])
        i = end + 1


def _apply_merges(ws, rows: list[tuple], fmts: dict) -> None:
    """
    对数据行应用合并单元格。

    合并规则：
    1. 功能区行（非"其他"模块）：
       - A(#)/C(功能) 列：按 (module, name) 分组合并
       - E(描述) 列：**不合并**，每行独立显示（企微行/PC行各有自己的描述）
       - B(模块) 列：按连续相同值合并（同模块跨多功能）
       - D(需求类型) 列：按连续相同值合并
    2. 其他工时行（功能测试/验收发布/需求沟通）：合并 C-E 列
    """
    OTHER_MODULES = {"功能测试", "验收发布", "需求沟通", "UI/UX设计", "环境搭建"}
    total = len(rows)
    if total == 0:
        return

    # 第一轮：功能区按 (module, name) 分组合并 A(#) 和 C(功能) 列
    # 注意：E(描述) 不合并，每行独立显示各自的描述
    i = 0
    while i < total:
        mod_val = rows[i][1]
        if mod_val in OTHER_MODULES:
            i += 1
            continue
        name_val = rows[i][2]
        group_end = i
        while group_end + 1 < total and rows[group_end + 1][1] == mod_val \
                and rows[group_end + 1][2] == name_val \
                and rows[group_end + 1][1] not in OTHER_MODULES:
            group_end += 1
        if group_end > i:
            r_start, r_end = i + 1, group_end + 1
            ws.merge_range(r_start, COL_SEQ, r_end, COL_SEQ,
                           rows[i][0], fmts["merge_center"])
            ws.merge_range(r_start, COL_NAME, r_end, COL_NAME,
                           name_val, fmts["merge"])
        i = group_end + 1

    # 第二轮：B(模块) 和 D(需求类型) 列按连续相同值合并
    _merge_consecutive(ws, rows, COL_MOD, "merge_center", fmts)
    _merge_consecutive(ws, rows, COL_TYPE, "merge_center", fmts)

    # 第三轮：其他工时行合并 C-E 列
    for i, row in enumerate(rows):
        if row[1] in OTHER_MODULES:
            r = i + 1
            ws.merge_range(r, COL_NAME, r, COL_DESC,
                           row[2], fmts["merge_center"])


def _write_totals(ws, rows: list[tuple], fmts: dict) -> None:
    """
    写总计行：测试工时公式 + 合计公式。
    新格式：
      - 功能测试行：公式 =ROUNDUP(SUM(开发工时)*0.15,0)
      - 合计行：公式 =ROUNDUP(SUM(所有工时),0)，备注列显示"合计"
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

    # 计算开发工时范围（从首数据行到功能测试行的前一行，不含测试行本身）
    first_data_row = 2  # Excel 1-based: header=1, 首数据行=2
    last_dev_row = test_row_num - 1  # 功能测试行的前一行（开发工时最后一行）

    # 写入公式: =ROUNDUP(SUM(开发工时)*0.15,0)
    test_formula = f"=ROUNDUP(SUM({effort_col_letter}{first_data_row}:{effort_col_letter}{last_dev_row})*0.15,0)"
    ws.write_formula(
        test_row_num, COL_EFFORT,
        test_formula,
        fmts["cell_effort"],
    )

    # 合计行
    total_row = len(rows) + 1  # 0-based，header 在第 0 行
    for c in range(NUM_COLS):
        ws.write(total_row, c, "", fmts["total_left"])

    # 写入合计公式: =ROUNDUP(SUM(所有数据行工时),0)
    # 求和范围：从首数据行到合计行的前一行（包含功能测试公式行 + 验收/需求等）
    total_formula = f"=ROUNDUP(SUM({effort_col_letter}{first_data_row}:{effort_col_letter}{total_row}),0)"
    ws.write_formula(
        total_row, COL_EFFORT,
        total_formula,
        fmts["total"],
    )
    ws.write(total_row, COL_SYS, "合计", fmts["total_left"])


def generate_excel(metadata_path: str, output_path: str | None = None, interactive: bool = False) -> str:
    """
    从metadata.json生成Excel估算文件。

    Args:
        metadata_path: metadata.json文件路径
        output_path:   输出Excel文件路径（可选；缺省时输出到 metadata 同目录）
        interactive:   是否交互式询问UI设计、验收发布、需求沟通工时

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
    rows = _collect_rows(metadata, interactive=interactive)
    for i, (seq, module, name, req_type, desc, system, effort, note) in enumerate(rows):
        r = i + 1  # header 在第 0 行
        # A/B/C/D/E/F/H 列：左对齐 + 上对齐
        ws.write(r, COL_SEQ, seq, fmts["merge_center"])
        ws.write(r, COL_MOD, module, fmts["cell"])
        ws.write(r, COL_NAME, name, fmts["cell"])
        ws.write(r, COL_TYPE, req_type, fmts["cell"])
        ws.write(r, COL_DESC, desc, fmts["cell"])
        ws.write(r, COL_SYS, system, fmts["cell_system"])
        # G 列（工时）：右对齐 + 上对齐
        if effort is not None:
            ws.write(r, COL_EFFORT, effort, fmts["cell_effort"])
        else:
            # 功能测试行，稍后写入公式
            pass
        ws.write(r, COL_NOTE, note, fmts["cell"])

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
    except ValidationError as e:
        print(f"\n{str(e)}\n")
        log.error("Excel generation cancelled: fix the errors above and retry")
        return 1


if __name__ == "__main__":
    sys.exit(main())
