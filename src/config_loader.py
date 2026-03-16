"""
config_loader.py
────────────────
加载并合并配置：
  1. config.yaml（项目配置）
  2. .env（敏感信息/API Key）
返回统一的配置字典。
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import yaml
from dotenv import load_dotenv


def load_config(config_path: str | Path = "config.yaml") -> dict[str, Any]:
    """
    加载配置文件并将环境变量合并进来。
    工作目录优先查找，其次查找脚本所在目录。
    """
    config_path = Path(config_path)

    # 搜索顺序：当前目录 → 脚本目录
    if not config_path.is_absolute():
        candidates = [
            Path.cwd() / config_path,
            Path(__file__).parent.parent / config_path,
        ]
        for c in candidates:
            if c.exists():
                config_path = c
                break

    cfg: dict[str, Any] = {}
    if config_path.exists():
        with open(config_path, encoding="utf-8") as f:
            cfg = yaml.safe_load(f) or {}
    else:
        print(f"[警告] 未找到配置文件 {config_path}，使用默认配置。")

    # 加载 .env（搜索顺序：当前目录 → 脚本目录）
    _load_dotenv_auto()

    return cfg


def _load_dotenv_auto():
    """自动查找并加载 .env 文件。"""
    candidates = [
        Path.cwd() / ".env",
        Path(__file__).parent.parent / ".env",
    ]
    for c in candidates:
        if c.exists():
            load_dotenv(c, override=False)
            return
    # 即使没有 .env 文件也不报错
    load_dotenv(override=False)


def get_paths(cfg: dict) -> dict[str, Path]:
    """从配置中提取并规范化路径配置。"""
    paths_cfg = cfg.get("paths", {})
    base = Path.cwd()

    def resolve(key: str, default: str) -> Path:
        raw = paths_cfg.get(key, default)
        p = Path(raw)
        return p if p.is_absolute() else base / p

    return {
        "input_dir": resolve("input_dir", "./input"),
        "output_dir": resolve("output_dir", "./output"),
        "log_file": resolve("log_file", "./logs/run.log"),
        "error_log": resolve("error_log", "./logs/error_log.txt"),
        "progress_db": resolve("progress_db", "./logs/progress.db"),
        "style_ref_file": Path(paths_cfg.get("style_ref_file", "") or ""),
    }
