"""
logger.py
─────────
统一日志配置：
- 控制台输出（带颜色，使用 rich）
- 文件输出（run.log）
- 错误专用日志（error_log.txt，纯文本，便于非技术用户查阅）
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path
from typing import Optional

# 尝试使用 rich 美化控制台输出
try:
    from rich.logging import RichHandler
    _HAS_RICH = True
except ImportError:
    _HAS_RICH = False

_logger: Optional[logging.Logger] = None
_error_file_path: Optional[Path] = None


def setup_logging(
    log_file: str | Path,
    error_log: str | Path,
    level: int = logging.INFO,
) -> logging.Logger:
    """初始化日志系统，返回主 logger。"""
    global _logger, _error_file_path

    log_file = Path(log_file)
    error_log = Path(error_log)
    log_file.parent.mkdir(parents=True, exist_ok=True)
    error_log.parent.mkdir(parents=True, exist_ok=True)
    _error_file_path = error_log

    logger = logging.getLogger("hotel_case_writer")
    logger.setLevel(level)
    logger.handlers.clear()

    # ── 控制台 Handler ─────────────────────────────────────────────────────────
    if _HAS_RICH:
        console_handler = RichHandler(
            rich_tracebacks=True,
            show_path=False,
            markup=True,
        )
        console_handler.setFormatter(logging.Formatter("%(message)s", datefmt="[%X]"))
    else:
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
        )
    console_handler.setLevel(level)
    logger.addHandler(console_handler)

    # ── 文件 Handler（完整日志）────────────────────────────────────────────────
    file_handler = logging.FileHandler(log_file, encoding="utf-8", mode="a")
    file_handler.setFormatter(
        logging.Formatter(
            "%(asctime)s [%(levelname)-8s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    )
    file_handler.setLevel(logging.DEBUG)
    logger.addHandler(file_handler)

    _logger = logger
    return logger


def get_logger() -> logging.Logger:
    """获取全局 logger（若未初始化则返回基础 logger）。"""
    if _logger is not None:
        return _logger
    # 降级：未初始化时返回基础配置
    fallback = logging.getLogger("hotel_case_writer")
    if not fallback.handlers:
        fallback.setLevel(logging.INFO)
        fallback.addHandler(logging.StreamHandler(sys.stdout))
    return fallback


def log_error_to_file(
    source_file: str,
    row_index: int,
    error_msg: str,
    content_preview: str = "",
):
    """
    将错误信息追加写入 error_log.txt（便于非技术用户查阅）。
    格式简洁，每条独立成段。
    """
    if _error_file_path is None:
        return
    try:
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        preview = (content_preview[:100] + "...") if len(content_preview) > 100 else content_preview
        entry = (
            f"[{timestamp}]\n"
            f"  文件: {source_file}\n"
            f"  行号: {row_index}\n"
            f"  错误: {error_msg}\n"
            f"  内容预览: {preview}\n"
            f"{'─' * 60}\n"
        )
        with open(_error_file_path, "a", encoding="utf-8") as f:
            f.write(entry)
    except Exception:
        pass  # 日志写入失败不应影响主流程
