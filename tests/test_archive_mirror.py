"""archive_estimation._copy_to_mirror 的单元测试。

覆盖镜像复制（best-effort，把整个归档文件夹同步到外部目录）的关键分支：
  - 镜像目录不存在 -> 自动创建并复制成功
  - 正常复制 -> 目标内容与源一致（含子文件）
  - 目标已存在 -> 旧目录备份为 .bak.{时间戳}，新内容写入
  - 源不存在 / 复制异常 -> 返回 None，不抛异常（保护主归档）

环境变量读取本身（archive_current_estimation 内的胶水代码）不在此测试。
"""

import os

import pytest

from archive_estimation import _copy_to_mirror


def _make_src(tmp_path, name="src_proj"):
    """在 tmp_path 下造一个带子文件的源归档文件夹，返回其路径。"""
    src = tmp_path / name
    src.mkdir()
    (src / "metadata.json").write_text("{}", encoding="utf-8")
    (src / "report.xlsx").write_text("fake-xlsx", encoding="utf-8")
    sub = src / "attachments"
    sub.mkdir()
    (sub / "a.txt").write_text("hello", encoding="utf-8")
    return src


class TestCopyToMirror:
    def test_creates_missing_mirror_dir_and_copies(self, tmp_path):
        src = _make_src(tmp_path)
        # 镜像目录尚不存在
        mirror_dir = str(tmp_path / "WeDrive" / "项目估算")

        target = _copy_to_mirror(mirror_dir, str(src), "src_proj")

        assert target is not None
        assert os.path.isdir(mirror_dir), "镜像根目录应被自动创建"
        assert os.path.isfile(os.path.join(target, "metadata.json"))
        assert os.path.isfile(os.path.join(target, "report.xlsx"))
        assert os.path.isfile(os.path.join(target, "attachments", "a.txt"))

    def test_copy_preserves_content(self, tmp_path):
        src = _make_src(tmp_path)
        mirror_dir = str(tmp_path / "mirror")

        target = _copy_to_mirror(mirror_dir, str(src), "src_proj")

        assert target is not None
        # 内容一致
        assert (src / "report.xlsx").read_text(encoding="utf-8") \
            == open(os.path.join(target, "report.xlsx"), encoding="utf-8").read()

    def test_existing_target_is_backed_up(self, tmp_path):
        src = _make_src(tmp_path)
        mirror_dir = tmp_path / "mirror"
        mirror_dir.mkdir()
        # 预先存在的同名目录（代表上一次归档）
        old_target = mirror_dir / "src_proj"
        old_target.mkdir()
        (old_target / "STALE.txt").write_text("old", encoding="utf-8")

        target = _copy_to_mirror(str(mirror_dir), str(src), "src_proj")

        assert target is not None
        # 新内容写入
        assert os.path.isfile(os.path.join(target, "metadata.json"))
        # 旧目录被备份为 .bak.{时间戳}
        backups = [d for d in os.listdir(mirror_dir)
                   if d.startswith("src_proj.bak.")]
        assert len(backups) == 1, f"应有一个 .bak 备份，实际：{os.listdir(mirror_dir)}"
        assert os.path.isfile(os.path.join(mirror_dir, backups[0], "STALE.txt"))

    def test_returns_none_when_source_missing(self, tmp_path):
        mirror_dir = str(tmp_path / "mirror")
        missing_src = str(tmp_path / "does_not_exist")

        target = _copy_to_mirror(mirror_dir, missing_src, "src_proj")

        assert target is None, "源不存在时应返回 None 而非抛异常"

    def test_returns_none_on_copy_error(self, tmp_path, monkeypatch):
        src = _make_src(tmp_path)
        mirror_dir = str(tmp_path / "mirror")

        # 让 shutil.copytree 抛错，验证异常被吞掉
        import archive_estimation as arch
        def _boom(*args, **kwargs):
            raise OSError("simulated WeDrive failure")
        monkeypatch.setattr(arch.shutil, "copytree", _boom)

        target = _copy_to_mirror(mirror_dir, str(src), "src_proj")

        assert target is None, "复制异常时应返回 None，保护主归档"
