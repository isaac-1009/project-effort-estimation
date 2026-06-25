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

    # === 数值范围检查 ===
    if "total_effort" in metadata:
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

    # === 抛出错误 ===
    if errors:
        error_message = "[ERROR] Metadata Schema Validation Failed:\n\n" + "\n".join(errors)
        raise ValidationError(error_message)


def _validate_feature(feature: dict[str, Any], index: int, errors: list[str]) -> None:
    """校验单个 feature 条目。"""
    prefix = f"features[{index}]"
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

    # 数值一致性检查 - 动态计算所有 _development 字段
    if "total_development_effort" in summary:
        calc_dev_total = 0
        for key, value in summary.items():
            if key.endswith("_development") and key != "total_development_effort":
                calc_dev_total += value
        stated_total = summary["total_development_effort"]
        if abs(calc_dev_total - stated_total) > _SUMMARY_TOLERANCE:
            errors.append(
                f"[ERROR] summary.total_development_effort ({stated_total}) does not match "
                f"sum of system development effort ({calc_dev_total})"
            )

    # 总工时一致性检查：summary.total_effort 必须与顶层 total_effort 一致
    if "total_effort" in summary and "total_effort" in metadata:
        stated_grand = summary["total_effort"]
        if abs(metadata["total_effort"] - stated_grand) > _SUMMARY_TOLERANCE:
            errors.append(
                f"[ERROR] Top-level total_effort ({metadata['total_effort']}) does not match "
                f"summary.total_effort ({stated_grand})"
            )
