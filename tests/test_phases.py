"""phases 阶段细分明细的单元测试。

覆盖两个维度：
1. metadata_schema._validate_phases —— 阶段结构校验
2. generate_excel._collect_rows / _collect_phase_rows —— 阶段明细行渲染

设计要点：
- 阶段汇总行 effort=None（不参与 SUM，避免重复计入）
- 阶段明细行 effort=数值（参与 SUM）
- 子项之和 == 阶段 total_effort（容差 0.5）
- 无 phases 字段时回退到旧 other_fixed 逻辑（向后兼容）
"""

import pytest

from generate_excel import _collect_rows
from metadata_schema import validate_metadata_schema, ValidationError


# ── _validate_phases ───────────────────────────────────────────────
class TestValidatePhases:
    def test_valid_phases_pass(self, valid_metadata_factory):
        """合法的 phases 结构应通过校验。"""
        m = valid_metadata_factory(
            total_effort=15,
            phases=[
                {
                    "name": "设计",
                    "total_effort": 4.0,
                    "subitems": [
                        {"name": "需求沟通", "effort_days": 1.0, "description": "x"},
                        {"name": "详细设计", "effort_days": 3.0, "description": "y"},
                    ],
                }
            ],
        )
        m["summary"]["total_effort"] = 15
        # 不应抛异常
        validate_metadata_schema(m)

    def test_phases_must_be_array(self, valid_metadata_factory):
        m = valid_metadata_factory(phases="not an array")
        with pytest.raises(ValidationError, match="phases must be an array"):
            validate_metadata_schema(m)

    def test_phases_cannot_be_empty(self, valid_metadata_factory):
        m = valid_metadata_factory(phases=[])
        with pytest.raises(ValidationError, match="phases cannot be empty"):
            validate_metadata_schema(m)

    def test_phase_missing_required_field(self, valid_metadata_factory):
        m = valid_metadata_factory(
            phases=[{"name": "设计", "subitems": []}]  # 缺 total_effort
        )
        with pytest.raises(ValidationError, match="missing field: total_effort"):
            validate_metadata_schema(m)

    def test_phase_subitems_empty_rejected(self, valid_metadata_factory):
        m = valid_metadata_factory(
            phases=[{"name": "设计", "total_effort": 0, "subitems": []}]
        )
        with pytest.raises(ValidationError, match="subitems must be a non-empty array"):
            validate_metadata_schema(m)

    def test_phase_total_mismatch_subitems_sum(self, valid_metadata_factory):
        """阶段 total_effort 与子项之和不符应报错。"""
        m = valid_metadata_factory(
            phases=[
                {
                    "name": "设计",
                    "total_effort": 10.0,  # 但子项之和=3
                    "subitems": [
                        {"name": "a", "effort_days": 1.0},
                        {"name": "b", "effort_days": 2.0},
                    ],
                }
            ]
        )
        with pytest.raises(ValidationError, match="does not match sum of subitems"):
            validate_metadata_schema(m)

    def test_phase_negative_effort_rejected(self, valid_metadata_factory):
        m = valid_metadata_factory(
            phases=[
                {
                    "name": "设计",
                    "total_effort": -1.0,
                    "subitems": [{"name": "a", "effort_days": -1.0}],
                }
            ]
        )
        with pytest.raises(ValidationError):
            validate_metadata_schema(m)

    def test_no_phases_field_still_valid(self, valid_metadata_factory):
        """不提供 phases 字段（旧格式）应继续通过校验。"""
        m = valid_metadata_factory()
        # 确认没有 phases 键
        assert "phases" not in m
        validate_metadata_schema(m)  # 不抛
