#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
归档当前估算到历史估算文件夹

当current文件夹中的估算确认无误后，使用此脚本将其移动到历史估算文件夹中

包含 metadata schema 校验，确保数据完整性
"""

import argparse
import json
import logging
import os
import shutil
import sys
from datetime import datetime
from typing import Any

# 共享模块：schema 校验的单一真相源（消除与 generate_excel.py 的复制粘贴）
from metadata_schema import ValidationError, validate_metadata_schema
# 共享模块：history.json 自学习写入逻辑
import history
# 统一日志配置（UTF-8 + [LEVEL] 前缀）
from logging_setup import setup_logging

setup_logging()
log = logging.getLogger(__name__)


# 可选的保留 ID 校验格式：name_YYYYMMDD
_ID_DATE_TAIL_LEN = len("_20260101")  # 9


def _normalize_project_id(metadata: dict, archived_date_compact: str) -> str:
    """
    规范化 project_id：统一用归档日期，消除双时间源漂移。

    规则：
      - 若 metadata.project_id 末尾是 _YYYYMMDD 且该日期 == 归档日期，则原样使用
      - 否则用 {name}_{归档日期} 生成（与 folder 名同源）
    这样 project_id 与 index.json.folder 始终一致。
    """
    raw_id = metadata.get("project_id", "")
    name = metadata.get("name", "Unnamed Project")
    if (
        len(raw_id) > _ID_DATE_TAIL_LEN
        and raw_id[-_ID_DATE_TAIL_LEN:].startswith("_")
        and raw_id[-8:].isdigit()
        and raw_id[-8:] == archived_date_compact
    ):
        return raw_id
    return f"{name}_{archived_date_compact}"


def load_metadata(metadata_path):
    """加载metadata.json文件"""
    with open(metadata_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def update_index(index_path, project_id, metadata, feature_patterns):
    """
    更新 estimations/index.json。

    Args:
        index_path:       data/estimations/index.json 路径
        project_id:       已规范化的项目 id（与 folder 同源日期）
        metadata:         已校验的 metadata dict
        feature_patterns: history.json 的 feature_patterns，用于数据驱动的关键词匹配
    """
    # 读取现有索引
    with open(index_path, 'r', encoding='utf-8') as f:
        index = json.load(f)

    # 统一时间源：folder 与 id 共用同一个归档日期
    archived_date = datetime.now()
    archived_date_compact = archived_date.strftime('%Y%m%d')
    archived_date_iso = archived_date.strftime('%Y-%m-%d')
    folder_name = f"{metadata['name']}_{archived_date_compact}"

    # 添加新项目
    project_info = {
        "id": project_id,
        "name": metadata["name"],
        "date": archived_date_iso,
        "folder": folder_name,
        "total_effort": metadata["total_effort"],
        "systems": metadata.get("systems", []),
        "tech_stack": metadata.get("tech_stack", {}),
        "features_count": len(metadata.get("features", [])),
        "status": "completed",
        "tags": metadata.get("tags", [])
    }

    # 检查是否已存在
    existing_ids = [p.get("id") for p in index.get("projects", [])]
    if project_id in existing_ids:
        log.warning("Project ID %s already exists, will update record", project_id)
        # 删除旧记录
        index["projects"] = [p for p in index["projects"] if p.get("id") != project_id]

    index.setdefault("projects", []).append(project_info)
    index["total_projects"] = len(index["projects"])
    index["last_updated"] = archived_date_iso

    # 更新功能索引（数据驱动：复用 history.match_patterns 得到命中的模式名，
    # 把模式名 + feature 自身名作为 keyword 写入 feature_index，不再硬编码中文词）
    feature_index = index.setdefault("feature_index", {})
    for feature in metadata.get("features", []):
        keywords = [feature.get("name", "")]
        keywords.extend(history.match_patterns(feature, feature_patterns))

        for keyword in keywords:
            if not keyword:
                continue
            if keyword not in feature_index:
                feature_index[keyword] = []
            if project_id not in feature_index[keyword]:
                feature_index[keyword].append(project_id)

    # 保存索引
    with open(index_path, 'w', encoding='utf-8') as f:
        json.dump(index, f, ensure_ascii=False, indent=2)
        f.write("\n")

    return folder_name


def resolve_base_path():
    """自动推断项目根目录

    优先级：
    1. 环境变量 PROJECT_EFFORT_ESTIMATION_HOME
    2. 脚本文件所在目录的上一级（scripts/ 的父目录）
    """
    env_home = os.environ.get("PROJECT_EFFORT_ESTIMATION_HOME")
    if env_home and os.path.isdir(env_home):
        return os.path.abspath(env_home)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidate = os.path.dirname(script_dir)  # scripts/ 的父目录
    # 校验候选目录确实是项目根目录（包含 data/ 文件夹）
    if os.path.isdir(os.path.join(candidate, "data")):
        return candidate

    raise FileNotFoundError(
        "[ERROR] Cannot locate project root. "
        "Set PROJECT_EFFORT_ESTIMATION_HOME or run from the project directory."
    )


def archive_current_estimation(force=False):
    """归档当前估算"""
    base_path = resolve_base_path()
    current_path = os.path.join(base_path, "data", "current")
    estimations_path = os.path.join(base_path, "data", "estimations")
    history_path = os.path.join(base_path, "data", "history.json")

    # 检查current文件夹是否存在
    if not os.path.exists(current_path):
        log.error("current folder does not exist")
        return False

    # 检查metadata.json是否存在
    metadata_path = os.path.join(current_path, "metadata.json")
    if not os.path.exists(metadata_path):
        log.error("metadata.json not found")
        return False

    # 加载元数据
    log.info("Loading project information...")
    metadata = load_metadata(metadata_path)

    # === Schema 校验 ===
    log.info("Validating metadata schema...")
    try:
        validate_metadata_schema(metadata)
        log.info("   [OK] Schema validation passed")
    except ValidationError as e:
        print(f"\n{str(e)}\n")
        log.error("Archive cancelled: please fix the errors above and retry")
        return False

    # === 规范化 project_id：统一用归档日期，消除双时间源漂移 ===
    archived_date_compact = datetime.now().strftime('%Y%m%d')
    archived_date_iso = datetime.now().strftime('%Y-%m-%d')
    project_id = _normalize_project_id(metadata, archived_date_compact)
    project_name = metadata.get("name", "Unnamed Project")

    log.info("   Project Name: %s", project_name)
    log.info("   Project ID: %s", project_id)
    log.info("   Total Effort: %s person-days", metadata.get('total_effort', 0))
    log.info("   Feature Count: %d", len(metadata.get('features', [])))

    # 确认
    if not force:
        confirm = input("\nConfirm to archive this project? (yes/no): ")
        if confirm.lower() not in ['yes', 'y']:
            log.info("Archive cancelled")
            return False

    # 加载 history 的 feature_patterns，供数据驱动的关键词匹配使用
    history_data = history.load_history(history_path)
    feature_patterns = history_data.get("feature_patterns", {})

    # 更新索引（project_id 与 folder 同源日期，关键词数据驱动）
    log.info("Updating project index...")
    index_path = os.path.join(estimations_path, "index.json")
    folder_name = update_index(index_path, project_id, metadata, feature_patterns)

    # 更新历史功能模式库（自学习：把新 feature 并入 history.json）
    log.info("Updating history feature-pattern library (self-learning)...")
    history.update_history(history_path, project_id, metadata, archived_date_iso)
    log.info("   [OK] history.json updated")

    # 移动文件
    log.info("Moving files to history folder...")
    target_path = os.path.join(estimations_path, folder_name)

    # 安全处理：若目标已存在，先备份为 .bak.{时间戳}，避免静默丢数据
    if os.path.exists(target_path):
        backup_name = f"{folder_name}.bak.{datetime.now().strftime('%Y%m%d%H%M%S')}"
        backup_path = os.path.join(estimations_path, backup_name)
        log.warning("Target exists, backing up old folder to: %s", backup_name)
        shutil.move(target_path, backup_path)

    # 移动整个current文件夹
    shutil.move(current_path, target_path)

    # 重新创建current文件夹
    os.makedirs(current_path)

    # 创建README
    readme_path = os.path.join(current_path, "README.md")
    with open(readme_path, 'w', encoding='utf-8') as f:
        f.write("# Current Estimation\n\n")
        f.write("Current estimation project files are stored here.\n\n")
        f.write("When estimation is confirmed, use `python scripts/archive_estimation.py` to move it to history.\n")

    log.info("[SUCCESS] Archive completed!")
    log.info("   Project saved to: %s", target_path)
    log.info("   Index updated:    %s", index_path)
    log.info("   History updated:  %s", history_path)
    log.info("   current folder cleared, ready for new estimation")

    return True


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description="归档当前估算到历史估算文件夹"
    )
    parser.add_argument(
        '-f', '--force', action='store_true',
        help="跳过交互确认"
    )
    args = parser.parse_args()
    try:
        ok = archive_current_estimation(force=args.force)
        sys.exit(0 if ok else 1)
    except Exception:
        log.exception("Archive failed")
        sys.exit(1)
