"""
file_writer.py
──────────────
负责将生成的案例内容保存为 .md 文件：
- 自动提取标题、清洗文件名
- 处理文件名冲突（追加序号）
- 可选添加元数据注释头
- 按来源文件名创建子目录（可配置）
"""

from __future__ import annotations

import re
import time
from pathlib import Path
from typing import Optional

from .excel_reader import EventRecord
from .prompt_manager import extract_title, sanitize_filename


def save_case(
    content: str,
    record: EventRecord,
    output_dir: str | Path,
    output_subdir_by_file: bool = True,
    add_metadata_header: bool = True,
    encoding: str = "utf-8",
    filename_max_length: int = 60,
    conflict_suffix: bool = True,
) -> Path:
    """
    将生成的 Markdown 内容保存为文件。

    返回：保存的文件路径。
    """
    output_dir = Path(output_dir)

    # ── 确定输出目录 ──────────────────────────────────────────────────────────
    if output_subdir_by_file:
        # 用来源 xlsx 文件名（不含扩展名）作为子目录名
        subdir_name = sanitize_filename(
            Path(record.source_file).stem, max_length=40
        )
        target_dir = output_dir / subdir_name
    else:
        target_dir = output_dir
    target_dir.mkdir(parents=True, exist_ok=True)

    # ── 提取标题 ──────────────────────────────────────────────────────────────
    title = extract_title(content)
    if title:
        file_stem = sanitize_filename(title, max_length=filename_max_length)
    else:
        # 无标题时用来源信息兜底
        file_stem = sanitize_filename(
            f"{Path(record.source_file).stem}_row{record.row_index}",
            max_length=filename_max_length,
        )

    # ── 处理文件名冲突 ────────────────────────────────────────────────────────
    file_path = target_dir / f"{file_stem}.md"
    if conflict_suffix and file_path.exists():
        counter = 2
        while (target_dir / f"{file_stem}_{counter}.md").exists():
            counter += 1
        file_path = target_dir / f"{file_stem}_{counter}.md"

    # ── 构建最终内容 ──────────────────────────────────────────────────────────
    final_content = content
    if add_metadata_header:
        metadata = _build_metadata_header(record)
        final_content = metadata + "\n\n" + content

    # ── 写入文件 ──────────────────────────────────────────────────────────────
    file_path.write_text(final_content, encoding=encoding)
    return file_path


def _build_metadata_header(record: EventRecord) -> str:
    """生成 Markdown 注释形式的元数据头，不影响渲染效果。"""
    from datetime import datetime
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [
        "<!--",
        f"  自动生成时间: {now}",
        f"  来源文件: {record.source_file}",
        f"  工作表: {record.sheet_name}",
        f"  原始行号: {record.row_index}",
        "-->",
    ]
    return "\n".join(lines)
