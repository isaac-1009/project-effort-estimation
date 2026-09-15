"""fill_medtronic_quote.collect_quote_rows 的单元测试。

验证：功能按名分组保持出现顺序、固定项按 系统测试→需求沟通→验收发布 追加、
系统测试数量按 开发合计×15% ROUNDUP 重算、内容文案带系统名前缀。
"""

from generate_excel import generate_excel
from fill_medtronic_quote import FIXED_ROW_TEXT, collect_quote_rows


def _make_estimate(tmp_path, metadata):
    out = tmp_path / "估算.xlsx"
    generate_excel(None, output_path=str(out), metadata_dict=metadata)
    return str(out)


def test_collect_quote_rows_order_and_days(tmp_path, valid_metadata):
    xlsx = _make_estimate(tmp_path, valid_metadata)
    rows = collect_quote_rows(xlsx)

    # 2 个功能行 + 系统测试(ROUNDUP(8*15%)=2) + 沟通2 + 验收1
    assert [days for days, _ in rows] == [3, 5, 2, 2, 1]
    assert [text for _, text in rows][0] == "用户端\t\n用户名密码登录，含 OAuth 认证"
    assert rows[2] == (2, FIXED_ROW_TEXT["系统测试"])
    assert rows[3] == (2, FIXED_ROW_TEXT["需求沟通"])
    assert rows[4] == (1, FIXED_ROW_TEXT["验收发布"])


def test_multi_system_feature_grouped(tmp_path, valid_metadata_factory):
    meta = valid_metadata_factory()
    # 用户登录改为多端功能：用户端 3 + 管理后台 2 → 合并一行 5 人天
    meta["features"][0] = {
        "module": "功能开发", "name": "登录与权限", "requirement_type": "新功能",
        "description": "登录与权限配置",
        "sub_systems": [
            {"system": "企微", "effort_days": 3, "remark": "OAuth"},
            {"system": "PC后台", "effort_days": 2, "remark": "RBAC"},
        ],
    }
    meta["features"][1]["name"] = "文章发布"
    meta["features"][1]["system"] = "PC后台"
    xlsx = _make_estimate(tmp_path, meta)
    rows = collect_quote_rows(xlsx)

    days = [d for d, _ in rows]
    assert days[0] == 5                                  # 多端合并求和
    assert rows[0][1].startswith("管理后台\t\n")          # 系统名取最后一行
    assert rows[0][1].endswith("登录与权限配置")
    assert days[-3] == 2                                 # 测试 ROUNDUP(10×15%)=2（开发=5+5）
