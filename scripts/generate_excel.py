#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通用Excel估算文件生成脚本
从metadata.json生成Excel估算文件（使用xlsxwriter确保边框正确）

包含 metadata schema 校验，确保生成前数据完整性
"""

import json
import os
import sys
from typing import Any, Dict

import xlsxwriter


class ValidationError(Exception):
    """元数据验证错误"""
    pass


def validate_metadata_schema(metadata: Dict[str, Any]) -> None:
    """
    验证 metadata.json 的 schema 完整性

    Args:
        metadata: 加载的 metadata 字典

    Raises:
        ValidationError: 当验证失败时
    """
    errors = []

    # === 必填字段检查 ===
    required_fields = [
        "project_id", "name", "created_date", "total_effort",
        "description", "systems", "tags", "tech_stack", "features", "summary"
    ]

    for field in required_fields:
        if field not in metadata:
            errors.append(f"[ERROR] Missing required field: {field}")

    # === 字段类型检查 ===
    if "project_id" in metadata and not isinstance(metadata["project_id"], str):
        errors.append("[ERROR] project_id must be a string")
    if "name" in metadata and not isinstance(metadata["name"], str):
        errors.append("[ERROR] name must be a string")
    if "total_effort" in metadata and not isinstance(metadata["total_effort"], (int, float)):
        errors.append("[ERROR] total_effort must be a number")

    # === 数值范围检查 ===
    if "total_effort" in metadata:
        effort = metadata["total_effort"]
        if effort <= 0:
            errors.append(f"[ERROR] total_effort must be > 0, current value: {effort}")
        if effort > 1000:
            errors.append(f"[WARNING] total_effort is too large ({effort} person-days), please confirm")

    # === systems 检查 ===
    if "systems" in metadata:
        if not isinstance(metadata["systems"], list):
            errors.append("[ERROR] systems must be an array")
        elif len(metadata["systems"]) == 0:
            errors.append("[ERROR] systems cannot be empty")

    # === tech_stack 检查 ===
    if "tech_stack" in metadata:
        required_tech_keys = ["frontend", "backend", "database"]
        for key in required_tech_keys:
            if key not in metadata["tech_stack"]:
                errors.append(f"[ERROR] tech_stack.{key} is missing")

    # === features 检查 ===
    if "features" in metadata:
        if not isinstance(metadata["features"], list):
            errors.append("[ERROR] features must be an array")
        elif len(metadata["features"]) == 0:
            errors.append("[ERROR] features cannot be empty, at least one feature is required")
        else:
            for i, feature in enumerate(metadata["features"]):
                feature_prefix = f"features[{i}]"
                required_feature_fields = ["system", "module", "name", "effort_days", "description", "complexity"]
                for field in required_feature_fields:
                    if field not in feature:
                        errors.append(f"[ERROR] {feature_prefix} missing field: {field}")

                if "effort_days" in feature:
                    effort = feature["effort_days"]
                    if not isinstance(effort, (int, float)):
                        errors.append(f"[ERROR] {feature_prefix}.effort_days must be a number")
                    elif effort < 0:
                        errors.append(f"[ERROR] {feature_prefix}.effort_days cannot be negative")

                if "complexity" in feature:
                    valid_complexity = ["低", "中", "高"]
                    if feature["complexity"] not in valid_complexity:
                        errors.append(f"[ERROR] {feature_prefix}.complexity must be one of {valid_complexity}")

    # === summary 检查 ===
    if "summary" in metadata:
        required_summary_keys = [
            "user_development", "admin_development",
            "total_development_effort", "testing_effort", "total_effort"
        ]
        for key in required_summary_keys:
            if key not in metadata["summary"]:
                errors.append(f"[ERROR] summary.{key} is missing")

        # 动态计算所有 *_development 字段总和并验证
        if "total_development_effort" in metadata["summary"]:
            calc_dev_total = 0
            for key, value in metadata["summary"].items():
                if key.endswith("_development") and key != "total_development_effort":
                    calc_dev_total += value
            stated_total = metadata["summary"]["total_development_effort"]
            if abs(calc_dev_total - stated_total) > 0.5:
                errors.append(
                    f"[ERROR] summary.total_development_effort ({stated_total}) does not match "
                    f"sum of system development effort ({calc_dev_total})"
                )

    # === 抛出错误 ===
    if errors:
        error_message = "[ERROR] Metadata Schema Validation Failed:\n\n" + "\n".join(errors)
        raise ValidationError(error_message)


def round_half(value: float) -> float:
    """将数值四舍五入到0.5的整数倍（最小估算单位0.5人天）"""
    return round(value * 2) / 2


def load_metadata(metadata_path: str) -> dict:
    """加载metadata.json文件"""
    with open(metadata_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def col_letter(index: int) -> str:
    """将列索引(0-based)转为Excel列字母，0=A"""
    result = ''
    while True:
        result = chr(ord('A') + index % 26) + result
        index = index // 26 - 1
        if index < 0:
            break
    return result


def generate_excel(metadata_path: str, output_path: str = None) -> str:
    """
    从metadata.json生成Excel估算文件

    Args:
        metadata_path: metadata.json文件路径
        output_path: 输出Excel文件路径（可选）
    """
    metadata = load_metadata(metadata_path)

    # === Schema 校验 ===
    print("[INFO] Validating metadata schema...")
    try:
        validate_metadata_schema(metadata)
        print("   [OK] Schema validation passed")
    except ValidationError as e:
        print(f"\n{str(e)}\n")
        print("[ERROR] Excel generation cancelled: please fix the errors above and retry")
        sys.exit(1)

    if output_path is None:
        project_name = metadata.get('name', '项目')
        # 生成在 metadata.json 同目录中
        metadata_dir = os.path.dirname(metadata_path)
        output_path = os.path.join(metadata_dir, f"{project_name}估算.xlsx")

    wb = xlsxwriter.Workbook(output_path)
    ws = wb.add_worksheet("工时估算")

    # ── 样式定义 ──
    fmt_header = wb.add_format({
        'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#4472C4',
        'border': 1, 'border_color': '#000000',
        'align': 'center', 'valign': 'vcenter',
        'text_wrap': True,
    })
    fmt_cell = wb.add_format({
        'border': 1, 'border_color': '#000000',
        'valign': 'vcenter', 'text_wrap': True,
    })
    fmt_cell_center = wb.add_format({
        'border': 1, 'border_color': '#000000',
        'align': 'center', 'valign': 'vcenter',
    })
    fmt_merge = wb.add_format({
        'border': 1, 'border_color': '#000000',
        'align': 'left', 'valign': 'vcenter', 'text_wrap': True,
    })
    fmt_merge_center = wb.add_format({
        'border': 1, 'border_color': '#000000',
        'align': 'center', 'valign': 'vcenter',
    })
    fmt_total = wb.add_format({
        'bold': True,
        'valign': 'vcenter',
    })
    fmt_total_center = wb.add_format({
        'bold': True,
        'align': 'center', 'valign': 'vcenter',
    })

    # 列布局: A=#, B=系统, C=模块, D=功能, E=工时, F=描述
    COL_SEQ = 0       # A: #
    COL_SYS = 1       # B: 系统
    COL_MOD = 2       # C: 模块
    COL_NAME = 3      # D: 功能
    COL_EFFORT = 4    # E: 工时
    COL_DESC = 5      # F: 描述
    NUM_COLS = 6

    # ── 写入表头 ──
    headers = ["#", "系统", "模块", "功能", "工时估算(人天)", "描述"]
    for c, h in enumerate(headers):
        ws.write(0, c, h, fmt_header)

    # ── 收集所有数据行 ──
    features = metadata.get('features', [])
    summary = metadata.get('summary', {})
    other_fixed = summary.get('other_fixed', {})

    rows: list[tuple[int, str, str, str, float, str]] = []
    seq_num = 0
    prev_system_module = None

    for feature in features:
        system = feature.get('system', '')
        module = feature.get('module', '')
        key = (system, module)
        if key != prev_system_module:
            seq_num += 1
            prev_system_module = key
        rows.append((
            seq_num,
            system,
            module,
            feature.get('name', ''),
            round_half(feature.get('effort_days', 0)),
            feature.get('description', ''),
        ))

    # 其他工时
    ui_effort_total = other_fixed.get('ui_ux_design', other_fixed.get('ui_design', 0))
    ui_desc = other_fixed.get('ui_design_desc', '')
    if not ui_desc and ui_effort_total > 0:
        # 兼容：从 features 中收集有 ui_effort 的用户端功能
        ui_parts = []
        for f in features:
            ue = f.get('ui_effort', 0)
            if ue > 0:
                ui_parts.append(f"{f.get('name', '')}({ue}人天)")
        ui_desc = "、".join(ui_parts) if ui_parts else "界面设计、交互设计"

    other_items = [
        ("UI/UX设计", ui_effort_total, ui_desc),
        ("系统测试", summary.get('testing_effort', 0), "集成测试、系统测试"),
        ("需求沟通", other_fixed.get('communication', 0), "沟通、需求文档"),
        ("环境搭建", other_fixed.get('environment', 0), "开发/测试/生产环境"),
        ("验收发布", other_fixed.get('acceptance', 0), "项目验收和上线发布"),
    ]
    for label, effort, desc in other_items:
        if effort > 0:
            seq_num += 1
            rows.append((seq_num, "其他", label, "", round_half(effort), desc))

    # ── 写入数据行 ──
    for i, (seq, system, module, name, effort, desc) in enumerate(rows):
        r = i + 1  # Excel row (0-based header is row 0)
        ws.write(r, COL_SEQ, seq, fmt_merge_center)
        ws.write(r, COL_SYS, system, fmt_cell)
        ws.write(r, COL_MOD, module, fmt_cell)
        ws.write(r, COL_NAME, name, fmt_cell)
        ws.write(r, COL_EFFORT, effort, fmt_cell_center)
        ws.write(r, COL_DESC, desc, fmt_cell)

    total_data_rows = len(rows)

    # ── 合并单元格 ──
    # 第一轮：按 (system, module) 分组，合并 # 和 模块 列（系统列留给第二轮）
    i = 0
    while i < total_data_rows:
        sys_val = rows[i][1]
        mod_val = rows[i][2]
        group_start = i
        group_end = i

        while group_end + 1 < total_data_rows:
            if rows[group_end + 1][1] == sys_val and rows[group_end + 1][2] == mod_val:
                group_end += 1
            else:
                break

        if group_end > group_start:
            r_start = group_start + 1  # 0-based row
            r_end = group_end + 1
            # 合并 # 列
            ws.merge_range(
                r_start, COL_SEQ, r_end, COL_SEQ,
                rows[group_start][0], fmt_merge_center)
            # 合并 模块 列
            ws.merge_range(
                r_start, COL_MOD, r_end, COL_MOD,
                mod_val, fmt_merge)

        i = group_end + 1

    # 第二轮：按系统列合并（同一系统的所有行，跨模块）
    i = 0
    while i < total_data_rows:
        sys_val = rows[i][1]
        sys_start = i
        sys_end = i
        while sys_end + 1 < total_data_rows and rows[sys_end + 1][1] == sys_val:
            sys_end += 1
        if sys_end > sys_start:
            ws.merge_range(
                sys_start + 1, COL_SYS, sys_end + 1, COL_SYS,
                sys_val, fmt_merge)
        i = sys_end + 1

    # ── 总计行 ──
    total_row = total_data_rows + 1  # 0-based
    for c in range(NUM_COLS):
        ws.write(total_row, c, '', fmt_total_center)

    # 计算总工时
    total_effort = sum(r[4] for r in rows)  # rows中的第5列是工时
    ws.write(total_row, COL_EFFORT, total_effort, fmt_total)
    # 同时写入公式（可选，用于Excel自动更新）
    ws.write_formula(total_row, COL_EFFORT,
                     f'=SUM({col_letter(COL_EFFORT)}2:{col_letter(COL_EFFORT)}{total_data_rows + 1})',
                     fmt_total, total_effort)
    ws.write(total_row, COL_DESC, "人天", fmt_total)

    # ── 列宽 ──
    ws.set_column('A:A', 5)     # #
    ws.set_column('B:B', 15)    # 系统
    ws.set_column('C:C', 20)    # 模块
    ws.set_column('D:D', 30)    # 功能
    ws.set_column('E:E', 15)    # 工时
    ws.set_column('F:F', 40)    # 描述

    # ── 冻结首行 ──
    ws.freeze_panes(1, 0)

    wb.close()
    print(f"Excel file generated: {output_path}")
    return output_path


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python generate_excel.py <metadata.json路径> [输出Excel路径]")
        print("示例: python generate_excel.py data/current/metadata.json")
        sys.exit(1)

    metadata_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None

    if not os.path.exists(metadata_path):
        print(f"Error: File not found {metadata_path}")
        sys.exit(1)

    output_file = generate_excel(metadata_path, output_path)
    print(f"\nDone! Excel file generated: {output_file}")
