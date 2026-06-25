#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
工时估算Excel读取脚本

用法：
    python scripts/read_excel.py <excel文件路径>
    python scripts/read_excel.py data/current/estimate.xlsx
"""

import sys
import io
import openpyxl

# 设置标准输出编码为 UTF-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


def read_excel(file_path):
    """
    读取工时估算Excel文件

    Args:
        file_path: Excel文件路径
    """
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb['工时估算']

        print("=" * 80)
        print(f"工时估算文件: {file_path}")
        print("=" * 80)

        # 读取所有数据行
        rows = []
        seen_modules = {}   # 记录已见过的模块及其工时，用于检测重复行

        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[4] is not None:  # 工时列有值
                # 跳过总计行
                if isinstance(row[4], str) and '=' in str(row[4]):
                    continue

                # 检测重复行：功能名为空或None且模块+工时组合已见过
                name = row[3]
                module = row[2] if row[2] else ""
                effort = row[4]

                # name为None或空字符串时检查是否重复
                if (not name or name is None or str(name).strip() == "") and module:
                    if isinstance(effort, (int, float)):
                        module_effort_key = (module, effort)
                        if module_effort_key in seen_modules:
                            continue
                        seen_modules[module_effort_key] = True

                rows.append(row)

        if not rows:
            print("[WARNING] 未找到数据行")
            return

        # 按系统分组统计（保留原始顺序）
        systems = {}
        total_effort = 0
        current_system = None

        for row in rows:
            seq = row[0]
            system = row[1] if row[1] else current_system
            module = row[2] if row[2] else ""
            name = row[3] if row[3] else ""
            effort = row[4]
            desc = row[5] if len(row) > 5 else ""

            if system:
                current_system = system

            if system not in systems:
                systems[system] = []

            systems[system].append({
                'seq': seq,
                'module': module,
                'name': name,
                'effort': effort,
                'desc': desc
            })

            if isinstance(effort, (int, float)):
                total_effort += effort

        # 输出详细信息
        for system in list(systems.keys()):
            if system == "其他":
                continue
            if system:
                print(f"\n【{system}】")
            system_total = 0
            for item in systems[system]:
                effort = item['effort']
                if isinstance(effort, (int, float)):
                    system_total += effort
                    module_str = f" - {item['module']}" if item['module'] else ""
                    print(f"  {item['seq']}. {item['name']}{module_str} | {effort}人天")
                    if item['desc']:
                        desc = item['desc']
                        if len(desc) > 60:
                            desc = desc[:60] + "..."
                        print(f"     {desc}")
            if system_total > 0:
                print(f"  小计: {system_total}人天")

        # 其他类最后输出
        if "其他" in systems:
            print(f"\n【其他】")
            system_total = 0
            for item in systems["其他"]:
                effort = item['effort']
                if isinstance(effort, (int, float)):
                    system_total += effort
                    module_str = f" - {item['module']}" if item['module'] else ""
                    print(f"  {item['seq']}. {item['name']}{module_str} | {effort}人天")
                    if item['desc']:
                        desc = item['desc']
                        if len(desc) > 60:
                            desc = desc[:60] + "..."
                        print(f"     {desc}")
            if system_total > 0:
                print(f"  小计: {system_total}人天")

        print("\n" + "=" * 80)
        print(f"总计工时: {total_effort}人天")
        print(f"功能数量: {len(rows)}项")
        print("=" * 80)

        wb.close()

    except FileNotFoundError:
        print(f"[ERROR] 文件不存在: {file_path}")
        sys.exit(1)
    except Exception as e:
        print(f"[ERROR] 读取文件失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python scripts/read_excel.py <excel文件路径>")
        print("示例: python scripts/read_excel.py data/current/estimate.xlsx")
        sys.exit(1)

    file_path = sys.argv[1]
    read_excel(file_path)
