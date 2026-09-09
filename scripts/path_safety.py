#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""路径安全助手：文件名清洗 + 子路径包含校验（供各脚本复用）。

归档路径中的动态片段（如 metadata.name 生成的文件夹名）必须经过
sanitize_filename 清洗；由「受信基目录 + 动态片段」拼出的目标路径
用 ensure_within 做越界校验，防止形如 ../ 的片段逃出基目录。
"""

import os
import re

# Windows 文件名非法字符 + 路径分隔符 + 控制字符。
# 注意：只替换 ASCII 非法字符，中文与全角标点（如 （））不受影响。
_ILLEGAL_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def sanitize_filename(name: str, replacement: str = "_") -> str:
    """把任意字符串清洗为安全的文件名/文件夹名片段。

    - 非法字符与路径分隔符替换为 replacement
    - 去掉首尾的空格与点（Windows 文件名不以点结尾）
    - 清洗后为空或只剩 '..' 时返回 'unnamed'，防止相对路径逃逸
    """
    cleaned = _ILLEGAL_CHARS.sub(replacement, str(name))
    cleaned = cleaned.strip(" .")
    if cleaned in ("", ".", ".."):
        return "unnamed"
    return cleaned


def ensure_within(base: str, target: str) -> str:
    """校验 target 落在 base 目录内，返回 target 的绝对路径；越界抛 ValueError。

    用于「基目录可信、拼接片段可能被污染」的场景：即使片段含 '..'
    或绝对路径，也在写盘前拦截。
    """
    base_abs = os.path.realpath(base)
    target_abs = os.path.realpath(target)
    try:
        inside = os.path.commonpath([base_abs, target_abs]) == base_abs
    except ValueError:
        # 不同盘符等无法计算公共路径的情形，一律视为越界
        inside = False
    if not inside:
        raise ValueError(
            f"path escapes base directory: {target!r} is not under {base!r}"
        )
    return target_abs
