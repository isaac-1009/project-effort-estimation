"""path_safety 模块的单元测试：文件名清洗 + 子路径越界校验。"""

import os

import pytest

from path_safety import ensure_within, sanitize_filename


class TestSanitizeFilename:
    @pytest.mark.parametrize("raw,expected", [
        # Windows 非法字符与路径分隔符
        ("a/b\\c:d*e?f\"g<h>i|j", "a_b_c_d_e_f_g_h_i_j"),
        # 分隔符替换后，首尾的点也被清掉（防止 "../" 片段保留前导点）
        ("../etc/passwd", "_etc_passwd"),
        ("..\\..\\win", "_.._win"),
        # 控制字符
        ("bad\x01name", "bad_name"),
        # 首尾空格与点被去掉
        ("  name. ", "name"),
        # 空与纯点 -> unnamed
        ("", "unnamed"),
        ("   ", "unnamed"),
        ("..", "unnamed"),
        (".", "unnamed"),
    ])
    def test_replaces_or_rejects(self, raw, expected):
        assert sanitize_filename(raw) == expected

    def test_keeps_chinese_and_fullwidth_punctuation(self):
        # 中文与全角括号是合法文件名字符，不能被误伤（归档项目名大量使用）
        name = "废品与样品管理系统（WSMS）"
        assert sanitize_filename(name) == name

    def test_custom_replacement(self):
        assert sanitize_filename("a/b", "-") == "a-b"

    def test_result_never_contains_separator(self):
        for raw in ["a/b", "..\\x", "nul:", "x*y"]:
            assert "/" not in sanitize_filename(raw)
            assert "\\" not in sanitize_filename(raw)


class TestEnsureWithin:
    def test_returns_absolute_path_when_inside(self, tmp_path):
        base = str(tmp_path)
        target = ensure_within(base, os.path.join(base, "sub", "file.json"))
        assert os.path.isabs(target)
        assert target == os.path.realpath(os.path.join(base, "sub", "file.json"))

    def test_allows_nonexistent_target_under_base(self, tmp_path):
        base = str(tmp_path / "not_created_yet")
        target = ensure_within(base, os.path.join(base, "proj_20260902"))
        assert target == os.path.realpath(os.path.join(base, "proj_20260902"))

    def test_rejects_parent_escape(self, tmp_path):
        with pytest.raises(ValueError, match="escapes base"):
            ensure_within(str(tmp_path), os.path.join(str(tmp_path), "..", "evil"))

    def test_rejects_absolute_target_outside_base(self, tmp_path):
        with pytest.raises(ValueError, match="escapes base"):
            ensure_within(str(tmp_path), str(tmp_path.parent / "other"))

    def test_rejects_dotdot_only_fragment(self, tmp_path):
        with pytest.raises(ValueError, match="escapes base"):
            ensure_within(str(tmp_path), os.path.join(str(tmp_path), ".."))
