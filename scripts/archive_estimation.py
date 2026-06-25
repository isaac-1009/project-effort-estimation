#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
归档当前估算到历史估算文件夹

当current文件夹中的估算确认无误后，使用此脚本将其移动到历史估算文件夹中

包含 metadata schema 校验，确保数据完整性
"""

import json
import os
import shutil
from datetime import datetime
import sys
import io
from typing import Any, Dict, List

# 设置标准输出编码为 UTF-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


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
            # 每个功能的必填字段
            for i, feature in enumerate(metadata["features"]):
                feature_prefix = f"features[{i}]"
                required_feature_fields = ["system", "module", "name", "effort_days", "description", "complexity"]
                for field in required_feature_fields:
                    if field not in feature:
                        errors.append(f"[ERROR] {feature_prefix} missing field: {field}")

                # 工时检查
                if "effort_days" in feature:
                    effort = feature["effort_days"]
                    if not isinstance(effort, (int, float)):
                        errors.append(f"[ERROR] {feature_prefix}.effort_days must be a number")
                    elif effort < 0:
                        errors.append(f"[ERROR] {feature_prefix}.effort_days cannot be negative")
                    elif effort == 0:
                        # 工时为0只允许在"需求未定"的情况下
                        desc = feature.get("description", "")
                        if "需求未定" not in desc and "undefined" not in desc.lower():
                            errors.append(f"[WARNING] {feature_prefix}.effort_days=0 but description doesn't mention 'undefined'")

                # 复杂度检查
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

        # 数值一致性检查 - 动态计算所有 _development 字段
        if "total_development_effort" in metadata["summary"] and "features" in metadata:
            # 计算所有以 _development 结尾的字段总和
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

        # 总工时一致性检查
        if "total_effort" in metadata["summary"]:
            stated_grand = metadata["summary"]["total_effort"]
            if "total_effort" in metadata and abs(metadata["total_effort"] - stated_grand) > 0.5:
                errors.append(
                    f"[ERROR] Top-level total_effort ({metadata['total_effort']}) does not match "
                    f"summary.total_effort ({stated_grand})"
                )

    # === 抛出错误 ===
    if errors:
        error_message = "[ERROR] Metadata Schema Validation Failed:\n\n" + "\n".join(errors)
        raise ValidationError(error_message)


def load_metadata(metadata_path):
    """加载metadata.json文件"""
    with open(metadata_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def update_index(index_path, project_id, metadata):
    """更新estimations/index.json"""
    # 读取现有索引
    with open(index_path, 'r', encoding='utf-8') as f:
        index = json.load(f)

    # 生成文件夹名
    folder_name = f"{metadata['name']}_{datetime.now().strftime('%Y%m%d')}"

    # 添加新项目
    project_info = {
        "id": project_id,
        "name": metadata["name"],
        "date": datetime.now().strftime('%Y-%m-%d'),
        "folder": folder_name,
        "total_effort": metadata["total_effort"],
        "systems": metadata.get("systems", []),
        "tech_stack": metadata.get("tech_stack", {}),
        "features_count": len(metadata.get("features", [])),
        "status": "completed",
        "tags": metadata.get("tags", [])
    }

    # 检查是否已存在
    existing_ids = [p["id"] for p in index["projects"]]
    if project_id in existing_ids:
        print(f"[WARNING] Project ID {project_id} already exists, will update record")
        # 删除旧记录
        index["projects"] = [p for p in index["projects"] if p["id"] != project_id]

    index["projects"].append(project_info)
    index["total_projects"] = len(index["projects"])
    index["last_updated"] = datetime.now().strftime('%Y-%m-%d')

    # 更新功能索引
    for feature in metadata.get("features", []):
        keywords = []

        # 从功能名提取关键词
        keywords.append(feature.get("name", ""))

        # 从描述提取关键词
        description = feature.get("description", "")
        if "登录" in description or "认证" in feature.get("name", ""):
            keywords.extend(["登录", "认证"])
        if "上传" in feature.get("name", ""):
            keywords.extend(["上传", "图片上传"])
        if "AI" in feature.get("name", "") or "AI" in description:
            keywords.append("AI")

        # 添加到索引
        for keyword in keywords:
            if not keyword:
                continue
            if keyword not in index["feature_index"]:
                index["feature_index"][keyword] = []
            if project_id not in index["feature_index"][keyword]:
                index["feature_index"][keyword].append(project_id)

    # 保存索引
    with open(index_path, 'w', encoding='utf-8') as f:
        json.dump(index, f, ensure_ascii=False, indent=2)

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

    # 检查current文件夹是否存在
    if not os.path.exists(current_path):
        print("[ERROR] current folder does not exist")
        return False

    # 检查metadata.json是否存在
    metadata_path = os.path.join(current_path, "metadata.json")
    if not os.path.exists(metadata_path):
        print("[ERROR] metadata.json not found")
        return False

    # 加载元数据
    print("[INFO] Loading project information...")
    metadata = load_metadata(metadata_path)

    # === Schema 校验 ===
    print("[INFO] Validating metadata schema...")
    try:
        validate_metadata_schema(metadata)
        print("   [OK] Schema validation passed")
    except ValidationError as e:
        print(f"\n{str(e)}\n")
        print("[ERROR] Archive cancelled: please fix the errors above and retry")
        return False

    project_id = metadata.get("project_id", f"project_{datetime.now().strftime('%Y%m%d')}")
    project_name = metadata.get("name", "Unnamed Project")

    print(f"   Project Name: {project_name}")
    print(f"   Project ID: {project_id}")
    print(f"   Total Effort: {metadata.get('total_effort', 0)} person-days")
    print(f"   Feature Count: {len(metadata.get('features', []))}")

    # 确认
    if not force:
        confirm = input(f"\nConfirm to archive this project? (yes/no): ")
        if confirm.lower() not in ['yes', 'y']:
            print("[INFO] Archive cancelled")
            return False

    # 更新索引
    print("\n[INFO] Updating project index...")
    index_path = os.path.join(estimations_path, "index.json")
    folder_name = update_index(index_path, project_id, metadata)

    # 移动文件
    print(f"[INFO] Moving files to history folder...")
    target_path = os.path.join(estimations_path, folder_name)

    # 如果目标文件夹已存在，先删除
    if os.path.exists(target_path):
        shutil.rmtree(target_path)

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

    print(f"\n[SUCCESS] Archive completed!")
    print(f"   Project saved to: {target_path}")
    print(f"   Index updated: {index_path}")
    print(f"\n   current folder cleared, ready for new estimation")

    return True


if __name__ == '__main__':
    try:
        # 支持命令行参数 --force 跳过交互确认
        force = '--force' in sys.argv or '-f' in sys.argv
        archive_current_estimation(force=force)
    except Exception as e:
        print(f"\n[ERROR] {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
