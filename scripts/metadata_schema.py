#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
metadata.json schema 校验的单一真相源。

此前 generate_excel.py 与 archive_estimation.py 各自复制了一份校验逻辑，
两份副本已悄然漂移（archive 版多检查了 0 工时与 total_effort 一致性）。
本模块取 archive 版（规则更全）作为唯一实现，两个脚本都从此 import。

CLAUDE.md 中所有 metadata.json 字段约定在此强制。
"""

from __future__ import annotations

from typing import Any

# 复杂度系数与 CLAUDE.md 的约定保持一致
VALID_COMPLEXITY = ["低", "中", "高"]

# 数值校验阈值（命名常量，消除散落的 magic number）
_TOTAL_EFFORT_MAX = 1000            # 单项目总工时上限（人天），超出给警告
_SUMMARY_TOLERANCE = 0.5            # summary 求和一致性容差（人天）


class ValidationError(Exception):
    """元数据验证错误"""
    pass


def validate_metadata_schema(metadata: dict[str, Any]) -> None:
    """
    验证 metadata.json 的 schema 完整性。

    Args:
        metadata: 加载的 metadata 字典

    Raises:
        ValidationError: 当存在 [ERROR] 级别的问题时抛出。
                         [WARNING] 也并入消息但同样抛出，便于人工复核。
    """
    errors: list[str] = []

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

    # === 数值范围检查（仅在类型正确时进行，避免 TypeError）===
    if "total_effort" in metadata and isinstance(metadata["total_effort"], (int, float)):
        effort = metadata["total_effort"]
        if effort <= 0:
            errors.append(f"[ERROR] total_effort must be > 0, current value: {effort}")
        if effort > _TOTAL_EFFORT_MAX:
            errors.append(
                f"[WARNING] total_effort is too large ({effort} person-days), please confirm"
            )

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
                _validate_feature(feature, i, errors)

    # === summary 检查 ===
    if "summary" in metadata:
        _validate_summary(metadata, errors)

    # === phases 检查（可选字段：阶段细分明细）===
    if "phases" in metadata:
        _validate_phases(metadata, errors)

    # === 抛出错误 ===
    if errors:
        error_message = "[ERROR] Metadata Schema Validation Failed:\n\n" + "\n".join(errors)
        raise ValidationError(error_message)


def _validate_feature(feature: dict[str, Any], index: int, errors: list[str]) -> None:
    """
    校验单个 feature 条目。

    支持两种结构：
    1. 旧结构：feature 直接包含 system, effort_days, description 等字段
    2. 新结构：feature 包含 sub_systems 数组，每个子项包含 system, effort_days, description
       （新结构对应新 Excel 8 列布局：一个功能跨企微+PC 多行，C/E 列按功能合并）

    可选字段（两种结构均支持）：
    - requirement_type: 需求类型（如 "功能优化"/"新功能"），输出到 D 列
    - remark: 备注，输出到 H 列（sub_systems 子项也可有 remark）
    """
    prefix = f"features[{index}]"

    # 可选字段类型检查（两种结构通用）
    if "requirement_type" in feature and not isinstance(feature["requirement_type"], str):
        errors.append(f"[ERROR] {prefix}.requirement_type must be a string")
    if "remark" in feature and not isinstance(feature["remark"], str):
        errors.append(f"[ERROR] {prefix}.remark must be a string")

    # 检查是否使用新结构（sub_systems）
    if "sub_systems" in feature:
        # 新结构校验
        if not isinstance(feature["sub_systems"], list):
            errors.append(f"[ERROR] {prefix}.sub_systems must be an array")
        elif len(feature["sub_systems"]) == 0:
            errors.append(f"[ERROR] {prefix}.sub_systems cannot be empty")
        else:
            for j, sub in enumerate(feature["sub_systems"]):
                sub_prefix = f"{prefix}.sub_systems[{j}]"
                required_sub_fields = ["system", "effort_days"]
                for field in required_sub_fields:
                    if field not in sub:
                        errors.append(f"[ERROR] {sub_prefix} missing field: {field}")

                # 可选字段类型检查
                if "remark" in sub and not isinstance(sub["remark"], str):
                    errors.append(f"[ERROR] {sub_prefix}.remark must be a string")
                if "effort_days" in sub:
                    effort = sub["effort_days"]
                    if not isinstance(effort, (int, float)):
                        errors.append(f"[ERROR] {sub_prefix}.effort_days must be a number")
                    elif effort < 0:
                        errors.append(f"[ERROR] {sub_prefix}.effort_days cannot be negative")

        # 新结构的必填字段
        new_required_fields = ["module", "name"]
        for field in new_required_fields:
            if field not in feature:
                errors.append(f"[ERROR] {prefix} missing field: {field}")
    else:
        # 旧结构校验
        required_feature_fields = ["system", "module", "name", "effort_days", "description", "complexity"]
        for field in required_feature_fields:
            if field not in feature:
                errors.append(f"[ERROR] {prefix} missing field: {field}")

        # 工时检查
        if "effort_days" in feature:
            effort = feature["effort_days"]
            if not isinstance(effort, (int, float)):
                errors.append(f"[ERROR] {prefix}.effort_days must be a number")
            elif effort < 0:
                errors.append(f"[ERROR] {prefix}.effort_days cannot be negative")
            elif effort == 0:
                # 工时为 0 只允许在"需求未定"的情况下
                desc = feature.get("description", "")
                if "需求未定" not in desc and "undefined" not in desc.lower():
                    errors.append(
                        f"[WARNING] {prefix}.effort_days=0 but description doesn't mention 'undefined'"
                    )

        # 复杂度检查
        if "complexity" in feature and feature["complexity"] not in VALID_COMPLEXITY:
            errors.append(
                f"[ERROR] {prefix}.complexity must be one of {VALID_COMPLEXITY}"
            )


def _validate_summary(metadata: dict[str, Any], errors: list[str]) -> None:
    """校验 summary 字段及其数值一致性。"""
    summary = metadata["summary"]
    required_summary_keys = [
        "user_development", "admin_development",
        "total_development_effort", "testing_effort", "total_effort"
    ]
    for key in required_summary_keys:
        if key not in summary:
            errors.append(f"[ERROR] summary.{key} is missing")

    # 数值一致性检查 - 动态计算所有 _development 字段（仅当值都是数字时）
    if "total_development_effort" in summary:
        dev_values = [
            value for key, value in summary.items()
            if key.endswith("_development") and key != "total_development_effort"
        ]
        stated_total = summary["total_development_effort"]
        if all(isinstance(v, (int, float)) for v in dev_values) \
                and isinstance(stated_total, (int, float)):
            calc_dev_total = sum(dev_values)
            if abs(calc_dev_total - stated_total) > _SUMMARY_TOLERANCE:
                errors.append(
                    f"[ERROR] summary.total_development_effort ({stated_total}) does not match "
                    f"sum of system development effort ({calc_dev_total})"
                )

    # 总工时一致性检查：summary.total_effort 必须与顶层 total_effort 一致
    # （仅当两者都是数字时才比较，避免类型错误的输入引发 TypeError）
    if "total_effort" in summary and isinstance(summary["total_effort"], (int, float)) \
            and "total_effort" in metadata and isinstance(metadata["total_effort"], (int, float)):
        stated_grand = summary["total_effort"]
        if abs(metadata["total_effort"] - stated_grand) > _SUMMARY_TOLERANCE:
            errors.append(
                f"[ERROR] Top-level total_effort ({metadata['total_effort']}) does not match "
                f"summary.total_effort ({stated_grand})"
            )


def _validate_phases(metadata: dict[str, Any], errors: list[str]) -> None:
    """
    校验顶层 phases 字段（可选）：阶段细分明细。

    结构约定：
        phases: [
            {
                "name": "设计",              # 阶段名（必填）
                "total_effort": 30.0,        # 阶段总工时（必填，= subitems 之和）
                "subitems": [                # 必填，至少 1 项
                    {"name": "需求沟通", "effort_days": 2.0, "description": "..."}
                ]
            }
        ]

    校验规则：
    - phases 必须为非空数组
    - 每个阶段含 name / total_effort / subitems
    - 阶段 total_effort = 其 subitems 的 effort_days 之和（容差 0.5）
    - subitem 的 effort_days >= 0
    """
    phases = metadata["phases"]
    if not isinstance(phases, list):
        errors.append("[ERROR] phases must be an array")
        return
    if len(phases) == 0:
        errors.append("[ERROR] phases cannot be empty (if present)")
        return

    for i, phase in enumerate(phases):
        prefix = f"phases[{i}]"
        if not isinstance(phase, dict):
            errors.append(f"[ERROR] {prefix} must be an object")
            continue

        # 必填字段
        for field in ("name", "total_effort", "subitems"):
            if field not in phase:
                errors.append(f"[ERROR] {prefix} missing field: {field}")

        if "name" in phase and not isinstance(phase["name"], str):
            errors.append(f"[ERROR] {prefix}.name must be a string")

        if "total_effort" in phase:
            te = phase["total_effort"]
            if not isinstance(te, (int, float)):
                errors.append(f"[ERROR] {prefix}.total_effort must be a number")
            elif te < 0:
                errors.append(f"[ERROR] {prefix}.total_effort cannot be negative")

        if "subitems" in phase:
            subs = phase["subitems"]
            if not isinstance(subs, list) or len(subs) == 0:
                errors.append(f"[ERROR] {prefix}.subitems must be a non-empty array")
                continue
            sub_sum = 0.0
            sub_sum_valid = True
            for j, sub in enumerate(subs):
                sp = f"{prefix}.subitems[{j}]"
                if not isinstance(sub, dict):
                    errors.append(f"[ERROR] {sp} must be an object")
                    sub_sum_valid = False
                    continue
                for field in ("name", "effort_days"):
                    if field not in sub:
                        errors.append(f"[ERROR] {sp} missing field: {field}")
                if "effort_days" in sub:
                    e = sub["effort_days"]
                    if not isinstance(e, (int, float)):
                        errors.append(f"[ERROR] {sp}.effort_days must be a number")
                        sub_sum_valid = False
                    elif e < 0:
                        errors.append(f"[ERROR] {sp}.effort_days cannot be negative")
                    else:
                        sub_sum += e
            # 子项之和一致性
            if sub_sum_valid and "total_effort" in phase \
                    and isinstance(phase["total_effort"], (int, float)):
                if abs(sub_sum - phase["total_effort"]) > _SUMMARY_TOLERANCE:
                    errors.append(
                        f"[ERROR] {prefix}.total_effort ({phase['total_effort']}) does not match "
                        f"sum of subitems ({sub_sum})"
                    )
