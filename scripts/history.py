#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
历史功能模式库的自学习写入逻辑。

data/history.json 是估算系统的"越用越准"基础：
- projects[]        — 每个已归档项目的 feature 列表（带 keywords/patterns）
- feature_patterns  — 24 种标准功能模式，examples[] 记录命中该模式的历史 feature id

本模块在归档时被调用，把新项目的 feature 追加进 projects，并把每个 feature
匹配到的标准模式的 examples 补上对应 feature_id。关键词匹配完全由
feature_patterns[].keywords 数据驱动，不硬编码任何中文业务词。
"""

from __future__ import annotations

import json
import os
from typing import Any

from metadata_schema import ValidationError, validate_metadata_schema


def load_history(history_path: str) -> dict[str, Any]:
    """加载 history.json，缺失时抛出明确的错误。"""
    try:
        with open(history_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        raise FileNotFoundError(
            f"history.json not found at {history_path}"
        )
    except json.JSONDecodeError as e:
        raise ValidationError(f"history.json is not valid JSON: {e}")


def save_history(history_path: str, history: dict[str, Any]) -> None:
    """写回 history.json（UTF-8、不转义中文、缩进 2）。

    本函数只允许写 history.json 本身，防止调用方传入其它路径误写文件。
    """
    if os.path.basename(os.path.normpath(history_path)) != "history.json":
        raise ValidationError(
            f"save_history refuses to write files other than history.json: {history_path!r}"
        )
    with open(history_path, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)
        f.write("\n")


def _feature_text(feature: dict[str, Any]) -> str:
    """把 feature 的 name + description 拼成一个待匹配文本块。"""
    return f"{feature.get('name', '')} {feature.get('description', '')}"


def feature_effort_summary(feature: dict[str, Any]) -> tuple[float, list[dict[str, Any]]]:
    """
    取一个 feature 的工时合计与多端明细。

    单端功能：返回 (effort_days, [])。
    多端功能（sub_systems，如企微+PC后台）：顶层无工时字段，返回
    (各端工时合计, 各端明细)，避免 history.json 把多端功能记成 0 工时。
    sub_systems 存在时优先生效（它是比顶层 system/effort_days 更完整的形态）。
    """
    sub = feature.get("sub_systems") or []
    if sub:
        total = sum(s.get("effort_days", 0) for s in sub)
        detail = [
            {"system": s.get("system", ""), "effort_days": s.get("effort_days", 0)}
            for s in sub
        ]
        return total, detail
    return feature.get("effort_days", 0), []


def match_patterns(
    feature: dict[str, Any],
    feature_patterns: dict[str, Any],
) -> list[str]:
    """
    判断一个 feature 命中了哪些标准功能模式。

    数据驱动：对每个 pattern，若其 keywords 中的任一词出现在 feature 的
    name/description 文本中，则视为命中。返回命中的 pattern 名称列表。
    """
    text = _feature_text(feature)
    hits: list[str] = []
    for pattern_name, pattern in feature_patterns.items():
        for kw in pattern.get("keywords", []):
            if kw and kw in text:
                hits.append(pattern_name)
                break  # 一个关键词命中即可，避免重复
    return hits


def _build_project_record(
    project_id: str,
    metadata: dict[str, Any],
    archived_date: str,
    feature_patterns: dict[str, Any],
) -> dict[str, Any]:
    """
    从 metadata 构造一条 history.projects 记录。
    为每个 feature 补充 keywords（来自命中模式的并集）与 patterns（命中模式名）。
    """
    features_record = []
    for i, feat in enumerate(metadata.get("features", [])):
        hits = match_patterns(feat, feature_patterns)
        # keywords = feature 自身名 + 命中模式的关键词并集（去重保序）
        keywords: list[str] = []
        seen = set()
        for kw in [feat.get("name", "")] + [
            k for p in hits for k in feature_patterns[p].get("keywords", [])
        ]:
            if kw and kw not in seen:
                keywords.append(kw)
                seen.add(kw)

        # feature_id: 优先用 metadata 里的 id，否则用 {project_id}_{序号}
        fid = feat.get("id") or f"{project_id}_{i + 1:03d}"

        effort, sub_detail = feature_effort_summary(feat)
        rec: dict[str, Any] = {
            "id": fid,
            "system": feat.get("system", ""),
            "module": feat.get("module", ""),
            "name": feat.get("name", ""),
            "effort_days": effort,
            "description": feat.get("description", ""),
            "complexity": feat.get("complexity", ""),
            "keywords": keywords,
            "patterns": hits,
        }
        if sub_detail:
            rec["sub_systems"] = sub_detail
        features_record.append(rec)

    return {
        "project_id": project_id,
        "name": metadata.get("name", ""),
        "created_date": archived_date,
        "tech_stack": metadata.get("tech_stack", {}),
        "features": features_record,
    }


def update_history(
    history_path: str,
    project_id: str,
    metadata: dict[str, Any],
    archived_date: str,
) -> dict[str, Any]:
    """
    把一个已归档项目并入 history.json 的自学习库。

    步骤：
      1. 校验 metadata schema（与归档脚本同一套规则）
      2. 若 project_id 已存在于 history.projects，则替换；否则追加
      3. 把每个 feature 的 id 回填到命中的 feature_patterns[].examples（去重）
      4. 更新 last_updated / total_features，写回磁盘

    Args:
        history_path: data/history.json 路径
        project_id:   项目唯一 id（应与 index.json 一致，格式 name_YYYYMMDD）
        metadata:     已校验过的 metadata dict
        archived_date:归档日期 YYYY-MM-DD

    Returns:
        更新后的 history dict（也已写盘）
    """
    validate_metadata_schema(metadata)

    history = load_history(history_path)
    feature_patterns = history.setdefault("feature_patterns", {})

    # 1. 构造项目记录（会顺手算出每个 feature 命中的 patterns）
    record = _build_project_record(
        project_id, metadata, archived_date, feature_patterns
    )

    # 2. 替换或追加项目
    projects = history.setdefault("projects", [])
    projects = [p for p in projects if p.get("project_id") != project_id]
    projects.append(record)
    history["projects"] = projects

    # 3. 回填 feature_patterns[].examples（去重，保序）
    new_examples_by_pattern: dict[str, list[str]] = {}
    for feat in record["features"]:
        for pattern_name in feat["patterns"]:
            new_examples_by_pattern.setdefault(pattern_name, []).append(
                feat["id"]
            )

    for pattern_name, feat_ids in new_examples_by_pattern.items():
        pattern = feature_patterns.get(pattern_name)
        if pattern is None:
            continue
        examples = pattern.setdefault("examples", [])
        for fid in feat_ids:
            if fid not in examples:
                examples.append(fid)

    # 4. 汇总元信息
    history["total_features"] = sum(
        len(p.get("features", [])) for p in history["projects"]
    )
    history["last_updated"] = archived_date

    save_history(history_path, history)
    return history
