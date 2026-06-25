#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
一次性数据修复脚本：把现有已归档项目补录进 history.json，
并修正 index.json 中 project_id 年份漂移的问题。

背景：
  - history.json 此前只含 1 个 2025 年种子项目，归档脚本从不写它，
    导致 README 宣称的"越用越准"从未生效。
  - index.json 中部分 project_id 末尾日期（2025）与 date 字段（2026）不一致，
    污染了 feature_index 的键。

用法：
    python scripts/migrate_history.py --dry-run     # 预览将发生的变更，不写盘
    python scripts/migrate_history.py --apply       # 实际写盘

本脚本可安全重复执行（幂等）：update_history 内部按 project_id 去重替换。
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sys
from datetime import datetime

# 让 scripts/ 内的相对 import 在直接 `python scripts/migrate_history.py` 时也能工作
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import history  # noqa: E402
from metadata_schema import ValidationError  # noqa: E402
from logging_setup import setup_logging  # noqa: E402

setup_logging()
log = logging.getLogger(__name__)


# 匹配 project_id 末尾的 _YYYYMMDD
_ID_DATE_RE = re.compile(r"_(\d{8})$")


def resolve_base_path() -> str:
    """与 archive_estimation.resolve_base_path 同构的最小实现。"""
    env_home = os.environ.get("PROJECT_EFFORT_ESTIMATION_HOME")
    if env_home and os.path.isdir(env_home):
        return os.path.abspath(env_home)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidate = os.path.dirname(script_dir)
    if os.path.isdir(os.path.join(candidate, "data")):
        return candidate
    raise FileNotFoundError(
        "[ERROR] Cannot locate project root. "
        "Set PROJECT_EFFORT_ESTIMATION_HOME or run from the project directory."
    )


def _derive_correct_id(project: dict) -> str | None:
    """
    根据 project.date 推导"正确"的 project_id。
    若当前 id 末尾日期 != date 派生日期，则返回修正后的 id；否则返回 None。
    """
    current_id = project.get("id", "")
    name = project.get("name", "")
    date_str = project.get("date", "")  # 形如 2026-03-29
    if not date_str:
        return None
    try:
        compact = datetime.strptime(date_str, "%Y-%m-%d").strftime("%Y%m%d")
    except ValueError:
        return None
    correct_id = f"{name}_{compact}"

    m = _ID_DATE_RE.search(current_id)
    if m:
        # 末尾是日期：若与正确日期不同则需要修正
        return correct_id if m.group(1) != compact else None
    # 末尾不是日期格式：统一规整为 name_YYYYMMDD
    return correct_id if current_id != correct_id else None


def fix_index_ids(index_path: str, dry_run: bool) -> int:
    """
    修正 index.json 中 project_id 的年份漂移，并同步更新 feature_index 的键。

    Returns: 实际修正的 project 数量。
    """
    with open(index_path, "r", encoding="utf-8") as f:
        index = json.load(f)

    projects = index.get("projects", [])
    feature_index = index.get("feature_index", {})

    rename_map: dict[str, str] = {}  # old_id -> new_id
    fixed = 0
    for p in projects:
        new_id = _derive_correct_id(p)
        if new_id and new_id != p.get("id"):
            rename_map[p["id"]] = new_id
            p["id"] = new_id
            fixed += 1

    if rename_map:
        # 同步 feature_index：把引用了旧 id 的键值替换为新 id
        for keyword, proj_ids in feature_index.items():
            feature_index[keyword] = [
                rename_map.get(pid, pid) for pid in proj_ids
            ]
        index["feature_index"] = feature_index
        index["last_updated"] = datetime.now().strftime("%Y-%m-%d")

    log.info("[index.json] %d project_id(s) need fixing:", fixed)
    for old, new in rename_map.items():
        log.info("    %s", old)
        log.info("     -> %s", new)

    if fixed and not dry_run:
        with open(index_path, "w", encoding="utf-8") as f:
            json.dump(index, f, ensure_ascii=False, indent=2)
            f.write("\n")
        log.info("[index.json] written (%d fix(es) applied)", fixed)
    elif fixed and dry_run:
        log.info("[index.json] (dry-run, not written)")

    return fixed


def backfill_history(
    estimations_path: str,
    history_path: str,
    dry_run: bool,
) -> int:
    """
    遍历每个已归档项目的 metadata.json，调用 update_history 补录进 history.json。

    对 schema 校验失败的项目跳过并报告，不中断整体迁移。
    Returns: 成功补录的项目数。
    """
    # 先修正过的 index.json，拿到权威的 (id, folder, date) 列表
    with open(os.path.join(estimations_path, "index.json"), "r", encoding="utf-8") as f:
        index = json.load(f)

    # 建立 folder -> (id, date) 映射，以 index 为准（确保用修正后的 id）
    folder_to_meta = {
        p["folder"]: (p["id"], p.get("date", ""))
        for p in index.get("projects", [])
        if "folder" in p
    }

    success = 0
    skipped = 0
    for folder in sorted(os.listdir(estimations_path)):
        folder_path = os.path.join(estimations_path, folder)
        if not os.path.isdir(folder_path):
            continue
        meta_file = os.path.join(folder_path, "metadata.json")
        if not os.path.exists(meta_file):
            continue

        with open(meta_file, "r", encoding="utf-8") as f:
            metadata = json.load(f)

        # 用 index 的权威 id（已修正年份），而非 metadata 里手填的旧 id
        project_id, archived_date = folder_to_meta.get(
            folder,
            (metadata.get("project_id", folder), metadata.get("created_date", "")),
        )

        try:
            # 无论 dry-run 还是 apply，都先做 schema 校验 + 模式匹配预览
            from metadata_schema import validate_metadata_schema
            validate_metadata_schema(metadata)
            hist = history.load_history(history_path)
            fp = hist.get("feature_patterns", {})
            hits_summary = {
                feat.get("name", ""): history.match_patterns(feat, fp)
                for feat in metadata.get("features", [])
            }
            successful_hits = sum(1 for v in hits_summary.values() if v)

            if dry_run:
                log.info("[history] WOULD backfill: %s (%d features)",
                         project_id, len(metadata.get("features", [])))
                log.info("           date=%s", archived_date or "(missing)")
                log.info("           features matching a pattern: %d/%d",
                         successful_hits, len(hits_summary))
            else:
                history.update_history(history_path, project_id, metadata, archived_date)
                log.info("[history] backfilled: %s (%d features, %d matched)",
                         project_id, len(metadata.get("features", [])), successful_hits)
            success += 1
        except ValidationError as e:
            skipped += 1
            log.warning("[history] SKIP %s: schema validation failed", folder)
            log.warning("          %s", str(e).splitlines()[0])

    mode = "dry-run" if dry_run else "written"
    log.info("[history.json] %d project(s) processed (%s), %d skipped",
             success, mode, skipped)
    return success


def main() -> int:
    parser = argparse.ArgumentParser(
        description="把现有已归档项目补录进 history.json，并修正 index.json 的 project_id 漂移"
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--dry-run", action="store_true", help="只预览，不写盘")
    group.add_argument("--apply", action="store_true", help="实际写盘")
    args = parser.parse_args()

    base_path = resolve_base_path()
    estimations_path = os.path.join(base_path, "data", "estimations")
    history_path = os.path.join(base_path, "data", "history.json")
    index_path = os.path.join(estimations_path, "index.json")

    print("=" * 60)
    print("STEP 1: Fix project_id year drift in index.json")
    print("=" * 60)
    fix_index_ids(index_path, dry_run=args.dry_run)

    print()
    print("=" * 60)
    print("STEP 2: Backfill archived projects into history.json")
    print("=" * 60)
    backfill_history(estimations_path, history_path, dry_run=args.dry_run)

    print()
    print("=" * 60)
    if args.dry_run:
        print("DRY RUN complete. Re-run with --apply to write changes.")
    else:
        print("Migration applied. Review with git diff.")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    sys.exit(main())
