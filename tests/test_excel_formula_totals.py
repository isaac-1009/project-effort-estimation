"""_write_totals 公式行范围的回归测试。

曾出现 off-by-one：功能测试公式 SUM 范围漏掉最后一个功能行
（开发 153 被按 148 计提 15%，Excel 合计比 metadata 少 6 人天，
且被 ROUNDUP 掩盖）。本文件锁定公式串必须覆盖全部数据行。
"""

import json

import openpyxl
import pytest

from generate_excel import generate_excel


@pytest.fixture
def formula_metadata_path(tmp_path):
    """两个单行功能（3+5=8 开发），无验收/沟通，自动追加功能测试行。"""
    metadata = {
        "project_id": "公式测试_20260902",
        "name": "公式测试",
        "created_date": "2026-09-02",
        "total_effort": 10,
        "description": "测试公式范围",
        "systems": ["企微"],
        "tags": [],
        "tech_stack": {"frontend": "H5", "backend": "Java", "database": "MySQL"},
        "features": [
            {
                "module": "功能开发",
                "name": "功能A",
                "requirement_type": "新功能",
                "description": "A",
                "complexity": "低",
                "system": "企微",
                "effort_days": 3,
            },
            {
                "module": "功能开发",
                "name": "功能B",
                "requirement_type": "新功能",
                "description": "B",
                "complexity": "低",
                "system": "企微",
                "effort_days": 5,
            },
        ],
        "summary": {
            "user_development": 8,
            "admin_development": 0,
            "total_development_effort": 8,
            "testing_effort": 2,
            "other_fixed": {"acceptance": 0, "communication": 0, "ui_ux_design": 0},
            "total_effort": 10,
        },
    }
    path = tmp_path / "metadata.json"
    path.write_text(json.dumps(metadata, ensure_ascii=False), encoding="utf-8")
    return path


class TestFormulaRanges:
    def test_testing_formula_covers_all_dev_rows(self, tmp_path, formula_metadata_path):
        out = generate_excel(str(formula_metadata_path), str(tmp_path / "out.xlsx"))
        ws = openpyxl.load_workbook(out)["工时估算"]
        formulas = {}
        for r in range(2, ws.max_row + 1):
            v = ws.cell(row=r, column=7).value
            if isinstance(v, str) and v.startswith("="):
                formulas[r] = v
        # 数据展示行 2-3 为功能，行 4 为功能测试行：公式必须覆盖 G2:G3
        assert formulas[4] == "=ROUNDUP(SUM(G2:G3)*0.15,0)"
        # 合计行（展示行 5）覆盖 G2:G4，直接 SUM 不做 ROUNDUP
        assert formulas[5] == "=SUM(G2:G4)"

    def test_many_rows_testing_formula_includes_last(self, tmp_path):
        """多行场景（曾出 bug 的形态）：最后一行功能必须计入测试公式。"""
        features = [
            {
                "module": "功能开发",
                "name": f"功能{i}",
                "requirement_type": "新功能",
                "description": "d",
                "complexity": "低",
                "system": "企微",
                "effort_days": 1,
            }
            for i in range(10)
        ]
        metadata = {
            "project_id": "公式测试2_20260902",
            "name": "公式测试2",
            "created_date": "2026-09-02",
            "total_effort": 12,
            "description": "测试公式范围",
            "systems": ["企微"],
            "tags": [],
            "tech_stack": {"frontend": "H5", "backend": "Java", "database": "MySQL"},
            "features": features,
            "summary": {
                "user_development": 10,
                "admin_development": 0,
                "total_development_effort": 10,
                "testing_effort": 2,
                "other_fixed": {"acceptance": 0, "communication": 0, "ui_ux_design": 0},
                "total_effort": 12,
            },
        }
        path = tmp_path / "metadata.json"
        path.write_text(json.dumps(metadata, ensure_ascii=False), encoding="utf-8")
        out = generate_excel(str(path), str(tmp_path / "out2.xlsx"))
        ws = openpyxl.load_workbook(out)["工时估算"]
        # 10 个功能行：展示行 2-11，功能测试行在展示行 12
        v = ws.cell(row=12, column=7).value
        assert v == "=ROUNDUP(SUM(G2:G11)*0.15,0)"
