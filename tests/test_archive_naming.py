# -*- coding: utf-8 -*-
"""归档命名规则测试：folder 与 project_id 日期前置（YYYYMMDD_name）。"""

import json
import sys
import os
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "scripts"))

from archive_estimation import _normalize_project_id, update_index


def _metadata(name="测试项目"):
    return {
        "name": name,
        "total_effort": 10,
        "systems": ["企微"],
        "tech_stack": {},
        "tags": [],
        "features": [{"name": "对话标签两态管理", "effort_days": 2}],
    }


class TestNormalizeProjectId:
    ARCHIVE_DATE = "20260916"

    def test_fallback_uses_date_first(self):
        metadata = {"name": "渠道随行客服响应优化", "project_id": "渠道随行客服响应优化_20260915"}
        assert (
            _normalize_project_id(metadata, self.ARCHIVE_DATE)
            == f"{self.ARCHIVE_DATE}_渠道随行客服响应优化"
        )

    def test_keep_new_prefix_when_same_date(self):
        raw = f"{self.ARCHIVE_DATE}_某项目"
        assert _normalize_project_id({"name": "某项目", "project_id": raw}, self.ARCHIVE_DATE) == raw

    def test_keep_old_suffix_when_same_date(self):
        raw = f"某项目_{self.ARCHIVE_DATE}"
        assert _normalize_project_id({"name": "某项目", "project_id": raw}, self.ARCHIVE_DATE) == raw

    def test_old_suffix_different_date_regenerated(self):
        metadata = {"name": "某项目", "project_id": "某项目_20260101"}
        assert (
            _normalize_project_id(metadata, self.ARCHIVE_DATE)
            == f"{self.ARCHIVE_DATE}_某项目"
        )

    def test_missing_id_regenerated(self):
        assert _normalize_project_id({"name": "某项目"}, self.ARCHIVE_DATE) == f"{self.ARCHIVE_DATE}_某项目"


class TestUpdateIndexFolderName:
    def test_folder_date_first(self, tmp_path):
        index_path = tmp_path / "index.json"
        index_path.write_text(json.dumps({"projects": [], "feature_index": {}}), encoding="utf-8")
        folder = update_index(
            str(index_path), "20260916_测试项目", _metadata(), {}
        )
        assert folder == "20260916_测试项目"
        record = json.loads(index_path.read_text(encoding="utf-8"))["projects"][0]
        assert record["folder"] == "20260916_测试项目"
        assert record["id"] == "20260916_测试项目"

    def test_folder_name_sanitized(self, tmp_path):
        index_path = tmp_path / "index.json"
        index_path.write_text(json.dumps({"projects": [], "feature_index": {}}), encoding="utf-8")
        folder = update_index(
            str(index_path), "20260916_测试项目", _metadata(name='A/B:C*D'), {}
        )
        assert "/" not in folder and ":" not in folder and folder.startswith("20260916_")
