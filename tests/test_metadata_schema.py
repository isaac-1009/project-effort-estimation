"""metadata_schema.validate_metadata_schema 的单元测试。"""

import pytest

from metadata_schema import (
    VALID_COMPLEXITY,
    ValidationError,
    validate_metadata_schema,
)


# ── 合法用例 ────────────────────────────────────────────────────────
class TestValid:
    def test_valid_metadata_passes(self, valid_metadata):
        """完全合法的 metadata 不应抛异常。"""
        validate_metadata_schema(valid_metadata)  # 不 raise 即通过

    def test_extra_fields_are_tolerated(self, valid_metadata):
        """额外的非 schema 字段不应导致失败。"""
        valid_metadata["custom_field"] = "anything"
        validate_metadata_schema(valid_metadata)


# ── 必填字段缺失 ────────────────────────────────────────────────────
class TestMissingRequired:
    @pytest.mark.parametrize("missing", [
        "project_id", "name", "created_date", "total_effort",
        "description", "systems", "tags", "tech_stack",
        "features", "summary",
    ])
    def test_missing_top_level_field_fails(self, valid_metadata_factory, missing):
        metadata = valid_metadata_factory(**{missing: ...})
        with pytest.raises(ValidationError, match="Missing required field"):
            validate_metadata_schema(metadata)


# ── 类型/数值校验 ──────────────────────────────────────────────────
class TestTypeAndRange:
    def test_total_effort_must_be_number(self, valid_metadata):
        valid_metadata["total_effort"] = "很多"
        with pytest.raises(ValidationError, match="total_effort must be a number"):
            validate_metadata_schema(valid_metadata)

    def test_non_number_total_effort_does_not_crash(self, valid_metadata):
        """字符串类型的 total_effort 不应引发 TypeError，而应优雅报错。"""
        valid_metadata["total_effort"] = "很多"
        valid_metadata["summary"]["total_effort"] = "也是字符串"
        # 应抛 ValidationError，而非 TypeError
        with pytest.raises(ValidationError):
            validate_metadata_schema(valid_metadata)

    def test_total_effort_must_be_positive(self, valid_metadata):
        valid_metadata["total_effort"] = 0
        with pytest.raises(ValidationError, match="total_effort must be > 0"):
            validate_metadata_schema(valid_metadata)

    def test_systems_must_be_nonempty_list(self, valid_metadata):
        valid_metadata["systems"] = []
        with pytest.raises(ValidationError, match="systems cannot be empty"):
            validate_metadata_schema(valid_metadata)

    def test_tech_stack_missing_key(self, valid_metadata):
        del valid_metadata["tech_stack"]["backend"]
        with pytest.raises(ValidationError, match="tech_stack.backend is missing"):
            validate_metadata_schema(valid_metadata)


# ── features 校验 ──────────────────────────────────────────────────
class TestFeatures:
    def test_empty_features_fails(self, valid_metadata):
        valid_metadata["features"] = []
        with pytest.raises(ValidationError, match="features cannot be empty"):
            validate_metadata_schema(valid_metadata)

    def test_feature_missing_field(self, valid_metadata):
        del valid_metadata["features"][0]["effort_days"]
        with pytest.raises(ValidationError, match="missing field: effort_days"):
            validate_metadata_schema(valid_metadata)

    def test_negative_effort_fails(self, valid_metadata):
        valid_metadata["features"][0]["effort_days"] = -1
        with pytest.raises(ValidationError, match="cannot be negative"):
            validate_metadata_schema(valid_metadata)

    def test_invalid_complexity_fails(self, valid_metadata):
        valid_metadata["features"][0]["complexity"] = "极高"
        with pytest.raises(ValidationError, match="complexity must be one of"):
            validate_metadata_schema(valid_metadata)

    def test_valid_complexity_values(self, valid_metadata):
        """三种合法复杂度都应通过。"""
        for i, c in enumerate(VALID_COMPLEXITY):
            valid_metadata["features"][0]["complexity"] = c
            validate_metadata_schema(valid_metadata)


# ── summary 一致性 ─────────────────────────────────────────────────
class TestSummaryConsistency:
    def test_dev_total_mismatch_fails(self, valid_metadata):
        """total_development_effort 与各 _development 之和不一致应失败。"""
        valid_metadata["summary"]["total_development_effort"] = 999
        with pytest.raises(ValidationError, match="does not match"):
            validate_metadata_schema(valid_metadata)

    def test_grand_total_mismatch_fails(self, valid_metadata):
        """顶层 total_effort 与 summary.total_effort 不一致应失败。"""
        # 让顶层与 summary 不一致（顶层改了，summary 保持原值）
        valid_metadata["total_effort"] = 999
        with pytest.raises(ValidationError, match="Top-level total_effort"):
            validate_metadata_schema(valid_metadata)

    def test_dev_total_within_tolerance_passes(self, valid_metadata):
        """0.5 人天容差内的偏差应通过。"""
        valid_metadata["summary"]["total_development_effort"] = 8.3
        validate_metadata_schema(valid_metadata)  # |8.3 - 8| = 0.3 <= 0.5
