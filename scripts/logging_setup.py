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


# 幂等标记：本模块已配置过日志则置 True，避免多脚本被同一进程 import 时
# 重复包裹 stdout（重复包裹会让旧 wrapper 的缓冲区被提前关闭，
# 进而触发 pytest capture 的 "I/O operation on closed file"）。
_LOGGING_CONFIGURED = False


def setup_logging(level: int = logging.INFO) -> None:
    """
    配置根 logger：UTF-8 输出 + [LEVEL] 前缀格式。

    在 Windows 上 stdout 默认编码可能是 cp936/gbk，会导致中文崩溃，
    因此显式用 UTF-8 包裹 stdout（与原脚本行为一致）。

    幂等：同一进程内多次调用（含被多个脚本 import）只会在首次包裹 stdout，
    后续调用仅刷新 handler。

    pytest 兼容：pytest 会用 capture 对象替换 sys.stdout 并在用例间 seek/读它，
    包裹它会触发 "I/O operation on closed file"。检测到 pytest 在跑时跳过包裹，
    把日志交给 pytest 的 capture 统一管理。
    """
    global _LOGGING_CONFIGURED

    # 检测是否在 pytest 下运行。
    # 注意 PYTEST_CURRENT_TEST 只在用例执行期间设置，模块导入（collection）时还没有，
    # 因此用 sys.modules 判断 pytest 是否已加载——它在 collection 阶段就为真。
    under_pytest = "pytest" in sys.modules

    # 仅在真实 stdout 且尚未包裹时包裹一次。pytest 下不包裹（交给 capture）。
    if (
        not _LOGGING_CONFIGURED
        and not under_pytest
        and hasattr(sys.stdout, "buffer")
    ):
        wrapped = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
        setattr(wrapped, "_UTF8_WRAPPED", True)  # 标记，便于外部诊断
        sys.stdout = wrapped
        _LOGGING_CONFIGURED = True

    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(_PrefixFormatter(_FORMAT))

    root = logging.getLogger()
    # 清除可能存在的旧 handler，避免重复输出（脚本被 import 时）
    for h in list(root.handlers):
        root.removeHandler(h)
    root.addHandler(handler)
    root.setLevel(level)
