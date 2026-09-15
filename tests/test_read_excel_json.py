"""read_excel.py --json 程序化输出模式的测试。

约定：读取估算 Excel 一律走 read_excel.py（程序化读取用 --json），
因此结果必须可直接 json.dumps，CLI --json 必须输出可解析 JSON。
"""

import json
import subprocess
import sys
from pathlib import Path

from generate_excel import generate_excel
from read_excel import read_excel

SCRIPTS_DIR = Path(__file__).resolve().parent.parent / "scripts"


def _make_excel(tmp_path: Path, metadata: dict) -> str:
    out = tmp_path / "json模式测试.xlsx"
    generate_excel(None, output_path=str(out), metadata_dict=metadata)
    return str(out)


class TestJsonSerializable:
    def test_result_is_json_serializable(self, tmp_path, valid_metadata):
        xlsx = _make_excel(tmp_path, valid_metadata)
        result = read_excel(xlsx, print_report=False)
        # 不抛 TypeError 即可序列化；关键字段齐全
        payload = json.loads(json.dumps(result, ensure_ascii=False))
        assert set(payload) == {"rows", "systems", "total_effort"}
        assert payload["rows"] and payload["total_effort"] > 0


class TestCliJsonFlag:
    def test_cli_json_outputs_parseable_json(self, tmp_path, valid_metadata):
        xlsx = _make_excel(tmp_path, valid_metadata)
        proc = subprocess.run(
            [sys.executable, str(SCRIPTS_DIR / "read_excel.py"), xlsx, "--json"],
            capture_output=True, text=True, encoding="utf-8",
        )
        assert proc.returncode == 0
        payload = json.loads(proc.stdout)
        assert payload["rows"]
        # 静态合计（不含公式行）：开发 3+5 + UI 8 + 验收 1 + 沟通 2
        # （测试 2 人天为公式行被跳过；环境 2 仅记 other_fixed 不出独立行）
        assert payload["total_effort"] == 3 + 5 + 8 + 1 + 2

    def test_cli_default_still_human_report(self, tmp_path, valid_metadata):
        xlsx = _make_excel(tmp_path, valid_metadata)
        proc = subprocess.run(
            [sys.executable, str(SCRIPTS_DIR / "read_excel.py"), xlsx],
            capture_output=True, text=True, encoding="utf-8",
        )
        assert proc.returncode == 0
        assert "总计工时" in proc.stdout
