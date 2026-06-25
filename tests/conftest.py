"""pytest 公共 fixture：提供合法与非法的 sample metadata。"""

import copy

import pytest


def _valid_metadata() -> dict:
    """一份完全合法的 metadata，用作各种变体测试的基底。"""
    return {
        "project_id": "测试项目_20260101",
        "name": "测试项目",
        "created_date": "2026-01-01",
        "total_effort": 23,
        "description": "用于单元测试的示例项目",
        "systems": ["用户端", "管理后台"],
        "tags": ["测试"],
        "tech_stack": {
            "frontend": "Vue",
            "backend": "Java",
            "database": "MySQL",
        },
        "features": [
            {
                "system": "用户端",
                "module": "登录",
                "name": "用户登录",
                "effort_days": 3,
                "description": "用户名密码登录，含 OAuth 认证",
                "complexity": "中",
            },
            {
                "system": "管理后台",
                "module": "内容管理",
                "name": "文章发布",
                "effort_days": 5,
                "description": "富文本编辑与发布",
                "complexity": "中",
            },
        ],
        "summary": {
            "user_development": 3,
            "admin_development": 5,
            "total_development_effort": 8,
            "testing_effort": 2,
            "total_effort": 23,
            "other_fixed": {
                "ui_ux_design": 8,
                "communication": 2,
                "environment": 2,
                "acceptance": 1,
            },
        },
    }


@pytest.fixture
def valid_metadata() -> dict:
    """一份合法的 metadata（每个测试拿到独立副本，可安全修改）。"""
    return copy.deepcopy(_valid_metadata())


@pytest.fixture
def valid_metadata_factory():
    """返回一个工厂，调用得到可修改的合法 metadata 副本。"""
    def _make(**overrides):
        data = copy.deepcopy(_valid_metadata())
        for k, v in overrides.items():
            if v is ...:  # 用 Ellipsis 表示"删除该字段"
                data.pop(k, None)
            else:
                data[k] = v
        return data
    return _make


@pytest.fixture
def feature_patterns() -> dict:
    """一份精简的 feature_patterns，用于 history 匹配测试。"""
    return {
        "认证授权": {
            "base_effort": 4,
            "complexity": "中",
            "keywords": ["登录", "认证", "权限"],
            "examples": [],
            "effort_range": [3, 8],
        },
        "内容管理": {
            "base_effort": 10,
            "complexity": "中",
            "keywords": ["内容管理", "发布", "文章"],
            "examples": [],
            "effort_range": [8, 12],
        },
        "富文本编辑": {
            "base_effort": 5,
            "complexity": "中",
            "keywords": ["富文本", "编辑器"],
            "examples": [],
            "effort_range": [4, 8],
        },
    }
