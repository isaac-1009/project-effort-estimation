"""generate_excel.round_half 与 history.match_patterns 的单元测试。"""

import pytest

from generate_excel import round_half
from history import match_patterns, _build_project_record


# ── round_half ─────────────────────────────────────────────────────
class TestRoundHalf:
    @pytest.mark.parametrize("value, expected", [
        (0.0, 0.0),
        (0.3, 0.5),       # 0.3 -> 0.6 -> round -> 1 -> 0.5
        (0.5, 0.5),
        (0.7, 0.5),       # 0.7 -> 1.4 -> round -> 1 -> 0.5
        (1.0, 1.0),
        (1.2, 1.0),       # 1.2 -> 2.4 -> 2 -> 1.0
        (1.3, 1.5),       # 1.3 -> 2.6 -> 3 -> 1.5
        (2.5, 2.5),
        (3.0, 3.0),
        (0.25, 0.0),      # banker's rounding edge: round(0.5)=0 in py3
        (1.25, 1.0),      # round(2.5)=2 (banker's)
    ])
    def test_rounding(self, value, expected):
        assert round_half(value) == expected


# ── match_patterns ─────────────────────────────────────────────────
class TestMatchPatterns:
    def test_match_by_name_keyword(self, feature_patterns):
        feature = {"name": "用户登录", "description": ""}
        hits = match_patterns(feature, feature_patterns)
        assert "认证授权" in hits  # "登录" 命中

    def test_match_by_description_keyword(self, feature_patterns):
        feature = {"name": "文章管理", "description": "支持富文本编辑"}
        hits = match_patterns(feature, feature_patterns)
        assert "富文本编辑" in hits  # "富文本" 命中

    def test_multiple_pattern_hits(self, feature_patterns):
        feature = {"name": "内容发布", "description": "富文本编辑后发布"}
        hits = match_patterns(feature, feature_patterns)
        assert "内容管理" in hits     # "发布" 命中
        assert "富文本编辑" in hits   # "富文本" 命中

    def test_no_match_returns_empty(self, feature_patterns):
        feature = {"name": "天气查询", "description": "查询天气预报"}
        assert match_patterns(feature, feature_patterns) == []

    def test_empty_feature_text(self, feature_patterns):
        feature = {"name": "", "description": ""}
        assert match_patterns(feature, feature_patterns) == []

    def test_no_duplicate_pattern(self, feature_patterns):
        """同一 pattern 的多个关键词命中只应返回该 pattern 一次。"""
        feature = {"name": "登录认证权限", "description": ""}
        hits = match_patterns(feature, feature_patterns)
        assert hits.count("认证授权") == 1


# ── _build_project_record ──────────────────────────────────────────
class TestBuildProjectRecord:
    def test_record_structure(self, feature_patterns):
        metadata = {
            "name": "测试项目",
            "tech_stack": {"frontend": "Vue"},
            "features": [
                {"name": "用户登录", "description": "OAuth 认证"},
                {"name": "无关功能", "description": "完全无关的内容"},
            ],
        }
        record = _build_project_record(
            "测试项目_20260101", metadata, "2026-01-01", feature_patterns
        )
        assert record["project_id"] == "测试项目_20260101"
        assert record["name"] == "测试项目"
        assert record["created_date"] == "2026-01-01"
        assert len(record["features"]) == 2

        # 命中模式的 feature 应记录 patterns
        login_feat = record["features"][0]
        assert "认证授权" in login_feat["patterns"]
        assert "认证授权" in login_feat["keywords"] or "登录" in login_feat["keywords"]

        # 未命中的 feature patterns 为空
        assert record["features"][1]["patterns"] == []

    def test_feature_id_fallback(self, feature_patterns):
        """feature 没有 id 字段时，用 {project_id}_{序号} 兜底。"""
        metadata = {
            "name": "X",
            "tech_stack": {},
            "features": [{"name": "f", "description": ""}],
        }
        record = _build_project_record(
            "X_20260101", metadata, "2026-01-01", feature_patterns
        )
        assert record["features"][0]["id"] == "X_20260101_001"

    def test_sub_systems_effort_summed(self, feature_patterns):
        """多端功能（sub_systems）：effort_days 记各端合计，明细随行保留。"""
        metadata = {
            "name": "X",
            "tech_stack": {},
            "features": [
                {
                    "name": "EAG",
                    "description": "以物易物",
                    "sub_systems": [
                        {"system": "企微", "effort_days": 26, "remark": "r1"},
                        {"system": "PC后台", "effort_days": 30, "remark": "r2"},
                    ],
                },
            ],
        }
        record = _build_project_record(
            "X_20260101", metadata, "2026-01-01", feature_patterns
        )
        feat = record["features"][0]
        assert feat["effort_days"] == 56
        assert feat["system"] == ""
        assert feat["sub_systems"] == [
            {"system": "企微", "effort_days": 26},
            {"system": "PC后台", "effort_days": 30},
        ]

    def test_single_system_feature_has_no_sub_systems_key(self, feature_patterns):
        """单端功能：effort 原样透传，不产生 sub_systems 键。"""
        metadata = {
            "name": "X",
            "tech_stack": {},
            "features": [
                {"name": "天气预警", "description": "", "system": "企微", "effort_days": 5},
            ],
        }
        record = _build_project_record(
            "X_20260101", metadata, "2026-01-01", feature_patterns
        )
        feat = record["features"][0]
        assert feat["effort_days"] == 5
        assert "sub_systems" not in feat

    def test_empty_sub_systems_falls_back_to_top_level(self, feature_patterns):
        """sub_systems 为空列表时按单端功能处理。"""
        metadata = {
            "name": "X",
            "tech_stack": {},
            "features": [
                {"name": "f", "description": "", "system": "企微",
                 "effort_days": 2, "sub_systems": []},
            ],
        }
        record = _build_project_record(
            "X_20260101", metadata, "2026-01-01", feature_patterns
        )
        feat = record["features"][0]
        assert feat["effort_days"] == 2
        assert "sub_systems" not in feat
