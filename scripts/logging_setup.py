#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
统一日志配置。

此前各脚本用 print("[INFO] ...") / print("[ERROR] ...") 手工拼前缀，
输出无法分级或重定向。本模块提供统一的 logging 配置，保留原有的
[LEVEL] 前缀风格，但改用标准库 logging。

用法（在脚本入口处）：
    from logging_setup import setup_logging
    setup_logging()          # 之后直接用 logging.info / logging.error
"""

from __future__ import annotations

import io
import logging
import sys

# 自定义格式：保留项目原有的 [LEVEL] 前缀风格
_FORMAT = "[%(levelname)s] %(message)s"

# levelname 映射：让 logging 的 INFO/WARNING/ERROR 显示为项目惯用的简写
_LEVEL_DISPLAY = {
    "DEBUG": "DEBUG",
    "INFO": "INFO",
    "WARNING": "WARNING",
    "ERROR": "ERROR",
    "CRITICAL": "ERROR",
}


class _PrefixFormatter(logging.Formatter):
    """用 _LEVEL_DISPLAY 映射 levelname，保持 [INFO]/[ERROR] 等简写风格。"""

    def format(self, record: logging.LogRecord) -> str:
        record.levelname = _LEVEL_DISPLAY.get(record.levelname, record.levelname)
        return super().format(record)


def setup_logging(level: int = logging.INFO) -> None:
    """
    配置根 logger：UTF-8 输出 + [LEVEL] 前缀格式。

    在 Windows 上 stdout 默认编码可能是 cp936/gbk，会导致中文崩溃，
    因此显式用 UTF-8 包裹 stdout（与原脚本行为一致）。
    """
    # 仅当 stdout 是真实缓冲区时才包裹，避免对已替换/管道场景重复包裹
    if hasattr(sys.stdout, "buffer"):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(_PrefixFormatter(_FORMAT))

    root = logging.getLogger()
    # 清除可能存在的旧 handler，避免重复输出（脚本被 import 时）
    for h in list(root.handlers):
        root.removeHandler(h)
    root.addHandler(handler)
    root.setLevel(level)
