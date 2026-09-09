#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
报价版估算生成器

给定目标总工时，把基准估算的开发工时按比例上浮（0.5 人天粒度、最大余数法配平），
使生成的报价版 Excel 合计恰好等于目标值。

规则：
- 只上浮开发工时行（功能行）；每行不低于基准值
- 系统测试仍按公式 ROUNDUP(开发×15%) 自动计提（随开发上浮而增大）
- 需求沟通 / 验收发布 / UI设计 等固定项维持基准值不动
- 目标值必须 ≥ 基准总工时，且为 0.5 人天的整数倍

用法：
    python scripts/generate_quote.py <metadata.json> --total 249.5 [-o 输出.xlsx]
"""

from __future__ import annotations

import argparse
import copy
import logging
import math
import os
import sys
from typing import Any

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from metadata_schema import ValidationError  # noqa: E402
from logging_setup import setup_logging  # noqa: E402
import generate_excel as ge  # noqa: E402

setup_logging()
log = logging.getLogger(__name__)

# 工时最小粒度（人天）
GRID = 0.5
# 系统测试计提比例（与 generate_excel 公式一致）
TESTING_RATE = 0.15
# 其他固定工时在 summary.other_fixed 中的键（与 _collect_other_rows 渲染口径一致）
FIXED_KEYS = ("acceptance", "communication", "ui_ux_design", "ui_design")

# _collect_rows 中"其他工时行"的模块名（这些行不参与上浮）
OTHER_MODULES = {"功能测试", "验收发布", "需求沟通", "UI/UX设计", "环境搭建"}


def excel_roundup(x: float) -> int:
    """Excel 的 ROUNDUP（正数场景等价于向上取整）。"""
    return math.ceil(round(x, 9))


def rendered_fixed_sum(summary: dict[str, Any]) -> float:
    """求会渲染进 Excel 的固定工时合计（验收/沟通/UI），与环境搭建等未渲染项区分。"""
    other_fixed = summary.get("other_fixed", {})
    total = 0.0
    for key in FIXED_KEYS:
        value = other_fixed.get(key)
        if value is None:
            continue
        value = float(value)
        if value > 0:
            total += ge.round_half(value)
    return total


def testing_enabled(summary: dict[str, Any]) -> bool:
    """是否包含功能测试公式行（other_fixed.testing=false 时不包含）。"""
    return summary.get("other_fixed", {}).get("testing", True) is not False


def base_totals(metadata: dict[str, Any]) -> dict[str, Any]:
    """拆解基准 metadata：开发行基准值列表 + 各分项合计。"""
    rows = ge._collect_rows(metadata, interactive=False)
    dev_rows = [r for r in rows if r[1] not in OTHER_MODULES]
    base_dev_values = [float(r[ge.COL_EFFORT]) for r in dev_rows]
    base_dev = ge.round_half(sum(base_dev_values))
    fixed = rendered_fixed_sum(metadata.get("summary", {}))
    has_testing = testing_enabled(metadata.get("summary", {}))
    testing = excel_roundup(base_dev * TESTING_RATE) if has_testing else 0
    return {
        "dev_values": base_dev_values,
        "dev_total": base_dev,
        "fixed": fixed,
        "has_testing": has_testing,
        "testing": testing,
        "total": base_dev + testing + fixed,
    }


def solve_dev_target(
    base_dev: float, fixed: float, has_testing: bool, target_total: float
) -> float:
    """
    求满足 dev + ROUNDUP(dev×15%) + fixed == target_total 的开发合计。

    dev 与 ROUNDUP 的联动存在不可达的 0.5 空隙；不可达时抛出并列出最接近的可行值。
    """
    need = round(target_total - fixed, 2)
    if not has_testing:
        if abs(need * 2 - round(need * 2)) > 1e-9:
            raise ValidationError(f"目标总工时 {target_total} 不是 {GRID} 的整数倍")
        return need

    def actual_total(dev: float) -> float:
        return dev + excel_roundup(dev * TESTING_RATE) + fixed

    # 在 [base_dev-1, need] 范围内按 0.5 步长扫描可行解
    lo = max(GRID, math.floor((base_dev - 1) / GRID) * GRID)
    hi = math.ceil(need / GRID) * GRID
    feasible: list[float] = []
    steps = int(round((hi - lo) / GRID))
    for k in range(steps + 1):
        dev = round(lo + k * GRID, 2)
        if abs(actual_total(dev) - target_total) < 1e-9 and dev >= base_dev:
            feasible.append(dev)

    if feasible:
        return min(feasible)

    # 不可达：给出最接近的可行总价，便于人工调整目标
    lo_total = actual_total(lo)
    scan = [round(lo + k * GRID, 2) for k in range(int(round((hi + 4 - lo) / GRID)) + 1)]
    totals = sorted({actual_total(d) for d in scan})
    nearest = min(totals, key=lambda t: (abs(t - target_total), t))
    raise ValidationError(
        f"目标总工时 {target_total} 不可达（测试计提 ROUNDUP 的取整空隙）；"
        f"最接近的可行报价总价为 {nearest} 人天"
    )


def distribute_uplift(base_values: list[float], dev_target: float) -> list[float]:
    """
    把开发合计从 base_sum 上浮到 dev_target：
    每行按比例放大 → 向下取到 0.5 网格（且不低于基准值）→ 最大余数法补足差额。
    """
    base_sum = ge.round_half(sum(base_values))
    if dev_target < base_sum:
        raise ValidationError(
            f"开发目标 {dev_target} 低于基准 {base_sum}，报价版只支持上浮"
        )
    scale = dev_target / base_sum
    raws = [v * scale for v in base_values]
    floored = [max(math.floor(r / GRID) * GRID, base) for r, base in zip(raws, base_values)]

    diff_units = int(round((dev_target - sum(floored)) / GRID))
    if diff_units < 0:
        raise ValidationError("配平差额为负，请检查目标值")

    # 最大余数法：按（raw - floor）降序补 0.5，余数相同则基准值大的优先
    order = sorted(
        range(len(raws)),
        key=lambda i: (-(raws[i] - floored[i]), -base_values[i], i),
    )
    adjusted = list(floored)
    for k in range(diff_units):
        adjusted[order[k % len(order)]] += GRID
    return [round(v, 2) for v in adjusted]


def build_quote_metadata(
    metadata: dict[str, Any],
    dev_values: list[float],
    dev_target: float,
    target_total: float,
    scale: float,
) -> dict[str, Any]:
    """把上浮后的工时写回 metadata 副本，并更新 summary 与 notes。"""
    quote = copy.deepcopy(metadata)

    # 按 _collect_rows 的展平顺序回写（sub_systems 逐行 / 旧结构单行）
    idx = 0
    for feature in quote.get("features", []):
        if "sub_systems" in feature:
            for sub in feature["sub_systems"]:
                sub["effort_days"] = dev_values[idx]
                idx += 1
        else:
            feature["effort_days"] = dev_values[idx]
            idx += 1
    if idx != len(dev_values):
        raise ValidationError(
            f"工时行数不匹配：metadata 展开 {idx} 行，分配到 {len(dev_values)} 行"
        )

    # 重算 summary（企微→用户端 / PC后台→管理后台 / 其余→其他）
    user = admin = other = 0.0
    for feature in quote.get("features", []):
        if "sub_systems" in feature:
            entries = feature["sub_systems"]
        else:
            entries = [feature]
        for entry in entries:
            system = entry.get("system", "")
            effort = float(entry.get("effort_days", 0))
            if system == "企微":
                user += effort
            elif system == "PC后台":
                admin += effort
            else:
                other += effort

    testing = excel_roundup(dev_target * TESTING_RATE) if testing_enabled(quote["summary"]) else 0
    summary = quote.setdefault("summary", {})
    summary["user_development"] = ge.round_half(user)
    summary["admin_development"] = ge.round_half(admin)
    if other:
        summary["other_development"] = ge.round_half(other)
    summary["total_development_effort"] = dev_target
    summary["testing_effort"] = testing
    summary["total_effort"] = target_total
    quote["total_effort"] = target_total

    note = (
        f"报价版说明：以基准估算上浮生成，目标总工时 {target_total} 人天；"
        f"开发工时按比例约 {scale:.2f} 上浮（0.5 人天粒度配平），"
        f"系统测试按开发×15% 公式计提 {testing} 人天，"
        f"需求沟通/验收发布维持标准工时不变。"
    )
    quote.setdefault("notes", []).append(note)
    return quote


def generate_quote(metadata_path: str, target_total: float, output_path: str) -> dict[str, Any]:
    """主流程：读基准 → 求解 → 分配 → 生成报价版 Excel，返回逐行对照表数据。"""
    if abs(target_total / GRID - round(target_total / GRID)) > 1e-9:
        raise ValidationError(f"目标总工时 {target_total} 不是 {GRID} 的整数倍")

    metadata = ge.load_metadata(metadata_path)
    base = base_totals(metadata)

    if target_total < base["total"]:
        raise ValidationError(
            f"目标总工时 {target_total} 低于基准 {base['total']}，报价版只支持上浮"
        )
    if abs(target_total - base["total"]) < 1e-9:
        log.warning("目标等于基准总工时，报价版将与基准一致")

    dev_target = solve_dev_target(
        base["dev_total"], base["fixed"], base["has_testing"], target_total
    )
    dev_values = distribute_uplift(base["dev_values"], dev_target)
    scale = dev_target / base["dev_total"]

    quote_metadata = build_quote_metadata(metadata, dev_values, dev_target, target_total, scale)
    ge.generate_excel(None, output_path=output_path, metadata_dict=quote_metadata)

    return {
        "base_total": base["total"],
        "target_total": target_total,
        "base_dev": base["dev_total"],
        "dev_target": dev_target,
        "testing": excel_roundup(dev_target * TESTING_RATE) if base["has_testing"] else 0,
        "fixed": base["fixed"],
        "scale": scale,
        "rows": list(zip(base["dev_values"], dev_values)),
        "output": output_path,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="生成报价版估算 Excel（合计恰好等于给定总工时）")
    parser.add_argument("metadata_path", help="基准 metadata.json 路径")
    parser.add_argument("--total", type=float, required=True, dest="target_total",
                        help="报价版总工时（人天，0.5 粒度，须不低于基准）")
    parser.add_argument("-o", "--output", default=None, help="输出 xlsx 路径（默认 data/current/{项目名}估算_报价版.xlsx）")
    args = parser.parse_args()

    if not os.path.exists(args.metadata_path):
        log.error("File not found: %s", args.metadata_path)
        return 1

    try:
        output = args.output
        if output is None:
            from archive_estimation import resolve_base_path
            metadata = ge.load_metadata(args.metadata_path)
            name = metadata.get("name", "项目")
            output = os.path.join(
                resolve_base_path(), "data", "current", f"{name}估算_报价版.xlsx"
            )

        result = generate_quote(args.metadata_path, args.target_total, output)
    except ValidationError as e:
        print(f"\n{str(e)}\n")
        log.error("报价版生成取消：请修正上述问题后重试")
        return 1

    log.info("基准总工时 %s → 报价版 %s 人天", result["base_total"], result["target_total"])
    log.info(
        "开发 %s → %s（约 ×%.3f），测试公式计提 %s，固定项 %s 不变",
        result["base_dev"], result["dev_target"], result["scale"],
        result["testing"], result["fixed"],
    )
    log.info("报价版已生成: %s", result["output"])
    return 0


if __name__ == "__main__":
    sys.exit(main())
