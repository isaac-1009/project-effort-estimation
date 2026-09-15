#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Medtronic 客户报价单（报价单模板）填写脚本

读取估算 Excel（基准版/报价版/任意配平版均可），把功能与固定项填入
templates/Medtronic_template.xlsx 的「报价单」工作表：
- 功能行写入「2. Digital Technology」区；系统名映射：企微/对外H5→用户端、PC后台→管理后台、其他→其他
- 多端功能（同一功能多行，如 用户与角色权限 企微+PC后台）合并为一行、数量求和，
  系统名取最后一行（与人工填写习惯一致）
- 固定项按 系统测试 → 需求沟通 → 验收发布 顺序追加；系统测试按 开发合计(含环境)×15% ROUNDUP 重算
- 数量列(D)填人天数；单价列(C)留空由商务手动填写；F=C*D、区小计、总计公式按实际行数重写
- I 列统一填「低于费率卡」；模板数据行（39-50）不够时在其后插行并复制样式，
  插入点以下的合并单元格与公式一并修正

用法：
    python scripts/fill_medtronic_quote.py <估算.xlsx> --name "RGI 荣誉勋章项目" [-o 输出.xlsx]
"""

from __future__ import annotations

import argparse
import logging
import math
import os
import sys
from copy import copy
from typing import Any

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from logging_setup import setup_logging  # noqa: E402
from path_safety import sanitize_filename  # noqa: E402
from read_excel import read_excel  # noqa: E402
from generate_quote import excel_roundup  # noqa: E402

setup_logging()
log = logging.getLogger(__name__)

TEMPLATE_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "templates", "Medtronic_template.xlsx",
)
SHEET_NAME = "报价单"

# 模板「2. Digital Technology」数据区：39-50 行共 12 个槽位，39 行为样式母本
DATA_START = 39
TEMPLATE_SLOTS = 12
INSERT_AT = 51          # 模板数据区之后的插行位置（51 行为模板空行）
COL_MIN, COL_MAX = 1, 9  # A..I

# 系统名 → 报价单"内容"列前缀
SYS_CN = {"企微": "用户端", "对外H5": "用户端", "PC后台": "管理后台", "其他": "其他"}
# 人工填写时的系统名覆盖（按功能名），如雷达图在客户报价中划入管理后台
SYSTEM_OVERRIDES: dict[str, str] = {"六边形雷达图": "管理后台"}
# 固定项行文案（与模板样例一致；数量由脚本按估算数据填）
FIXED_ROW_TEXT = {
    "系统测试": "其他\t\n系统测试\t\t\n集成测试、系统测试",
    "需求沟通": "其他\t\n需求沟通\t\t\n沟通、需求文档",
    "验收发布": "其他\t\n验收发布\t\t\n项目验收和上线发布",
}
RATE_CARD_TEXT = "低于费率卡"
UNIT_TEXT = "人/天"
# read_excel 行中固定项的识别特征（功能名为 "-"，模块为固定项名）
FIXED_MODULES = ("需求沟通", "验收发布", "UI/UX设计")


def collect_quote_rows(xlsx_path: str) -> list[tuple[float, str]]:
    """
    从估算 Excel 提取报价行，返回 [(数量人天, 内容文案), ...]，顺序与模板填写一致。

    功能行按功能名分组（多端合并、数量求和，保持出现顺序），固定项按
    系统测试 → 需求沟通 → 验收发布 追加；系统测试数量按 开发合计(含环境)×15% ROUNDUP。
    """
    result = read_excel(xlsx_path, print_report=False)
    rows = result["rows"]

    features: dict[str, dict[str, Any]] = {}   # name -> {days, desc, system}
    fixed: dict[str, float] = {}
    feature_days_total = 0.0
    for r in rows:
        effort = r["effort"]
        if not isinstance(effort, (int, float)):
            continue
        name = str(r["name"]).strip()
        if name == "-":
            if r["module"] in FIXED_MODULES:
                fixed[r["module"]] = fixed.get(r["module"], 0.0) + float(effort)
            continue
        item = features.setdefault(name, {"days": 0.0, "desc": "", "system": ""})
        item["days"] += float(effort)
        if not item["desc"] and r["desc"]:
            item["desc"] = str(r["desc"])
        if r["system"] and r["system"] != "-":
            item["system"] = r["system"]     # 多端合并取最后一行系统名
        feature_days_total += float(effort)

    quote_rows: list[tuple[float, str]] = []
    for name, item in features.items():
        sys_cn = SYSTEM_OVERRIDES.get(name) or SYS_CN.get(item["system"], item["system"] or "其他")
        quote_rows.append((item["days"], f"{sys_cn}\t\n{item['desc']}"))

    testing = excel_roundup(feature_days_total * 0.15)
    quote_rows.append((float(testing), FIXED_ROW_TEXT["系统测试"]))
    for module in ("需求沟通", "验收发布"):
        if module in fixed:
            quote_rows.append((fixed[module], FIXED_ROW_TEXT[module]))
        else:
            log.warning("估算文件中缺少 %s 行，报价单对应行数量为 0", module)
            quote_rows.append((0.0, FIXED_ROW_TEXT[module]))
    return quote_rows


def fill_template(quote_rows: list[tuple[float, str]], project_name: str, output_path: str) -> str:
    """把报价行填入 Medtronic 模板并另存，返回输出路径。"""
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb[SHEET_NAME]

    n = len(quote_rows)
    data_end = DATA_START + n - 1
    spare_end = data_end + 1        # 数据区后保留一个备用空行（与人工填写版一致）
    extra = max(0, n - TEMPLATE_SLOTS)
    if extra:
        # 先摘除插入点以下的合并单元格（insert_rows 不保证移动它们），插完后按新位置恢复
        shifted = [rng for rng in list(ws.merged_cells.ranges) if rng.min_row >= INSERT_AT]
        for rng in shifted:
            ws.unmerge_cells(str(rng))
        ws.insert_rows(INSERT_AT, extra + 1)   # +1：保留数据区与「3. Others」之间的空行结构
        donor = DATA_START                     # 样式母本行（含备用行一并复制样式）
        for r in range(INSERT_AT, spare_end + 1):
            for c in range(COL_MIN, COL_MAX + 1):
                ws.cell(row=r, column=c)._style = copy(ws.cell(row=donor, column=c)._style)
            if ws.row_dimensions.get(donor) is not None and ws.row_dimensions[donor].height:
                ws.row_dimensions[r].height = ws.row_dimensions[donor].height
            ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)  # G:H 合并
        for rng in shifted:
            ws.merge_cells(start_row=rng.min_row + extra + 1, start_column=rng.min_col,
                           end_row=rng.max_row + extra + 1, end_column=rng.max_col)
    elif n < TEMPLATE_SLOTS:
        for r in range(data_end + 1, DATA_START + TEMPLATE_SLOTS):  # 清空多余样例槽位
            for c in (1, 5, 6, 7, 9):
                ws.cell(row=r, column=c).value = None

    # 模板末位备用槽（50 行）样式与其他数据行不一致，写入内容后会暴露；
    # 数据区所有模板槽位统一按母本行（39 行）样式归一（合并单元格本就已存在）
    for r in range(DATA_START + 1, min(n, TEMPLATE_SLOTS) + DATA_START):
        for c in range(COL_MIN, COL_MAX + 1):
            ws.cell(row=r, column=c)._style = copy(ws.cell(row=DATA_START, column=c)._style)

    for i, (days, text) in enumerate(quote_rows):
        r = DATA_START + i
        ws.cell(row=r, column=1, value=i + 1)                       # A 编号
        ws.cell(row=r, column=4, value=days)                        # D 数量（人天）
        ws.cell(row=r, column=5, value=UNIT_TEXT)                   # E 单位
        ws.cell(row=r, column=6, value=f"=C{r}*D{r}")               # F 总金额
        ws.cell(row=r, column=7, value=text)                        # G 内容
        ws.cell(row=r, column=9, value=RATE_CARD_TEXT)              # I 费率卡

    section3 = 41 + n   # 「3. Others其他」标题行：52 + (n - 12) + 1
    ws["C13"] = project_name
    ws["G37"] = f"=SUM(F{DATA_START}:F{spare_end})"   # 含备用空行，与人工填写版一致
    ws[f"G{section3}"] = f"=SUM(F{section3 + 2}:F{section3 + 3})"
    # 第 3 区末位槽的 C*D 公式随行移动但引用不会平移，按新行号重写
    ws.cell(row=section3 + 3, column=6, value=f"=C{section3 + 3}*D{section3 + 3}")
    ws["E20"] = f"=G29+G37+G{section3}"

    directory = os.path.dirname(os.path.abspath(output_path))
    os.makedirs(directory, exist_ok=True)
    wb.save(output_path)
    log.info("报价单已生成: %s（%d 行）", output_path, n)
    return output_path


def main() -> int:
    parser = argparse.ArgumentParser(description="按 Medtronic 模板生成客户报价单")
    parser.add_argument("estimate_xlsx", help="估算 Excel（基准版/报价版/配平版）")
    parser.add_argument("--name", required=True, help="项目/活动名称（写入 C13）")
    parser.add_argument("-o", "--output", default=None,
                        help="输出 xlsx 路径（默认与输入同目录，<名称>_客户报价单.xlsx）")
    args = parser.parse_args()

    if not os.path.exists(args.estimate_xlsx):
        log.error("File not found: %s", args.estimate_xlsx)
        return 1
    output = args.output
    if output is None:
        stem = sanitize_filename(os.path.splitext(os.path.basename(args.estimate_xlsx))[0])
        output = os.path.join(os.path.dirname(os.path.abspath(args.estimate_xlsx)),
                              f"{stem}_客户报价单.xlsx")
    try:
        fill_template(collect_quote_rows(args.estimate_xlsx), args.name, output)
    except Exception:
        log.exception("报价单生成失败")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
