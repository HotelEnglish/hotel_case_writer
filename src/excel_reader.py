"""
excel_reader.py
───────────────
读取 .xlsx 文件，提取 Resolution Notes 列，返回结构化记录列表。
支持多 Sheet 扫描，自动跳过空值/无效值。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import pandas as pd


# ── 无效内容的固定值集合（全部转小写比较）──────────────────────────────────
_DEFAULT_INVALID_VALUES = {"无", "n/a", "nil", "none", "-", "/", "na", ""}


@dataclass
class EventRecord:
    """单条 Logbook 事件记录"""
    source_file: str          # 来源文件名（不含路径）
    sheet_name: str           # 来源 Sheet 名
    row_index: int            # 原始行号（1-based，含表头）
    content: str              # Resolution Notes 原文
    extra_fields: dict = field(default_factory=dict)   # 保留其他列供 Prompt 参考
    desensitized_content: str = ""   # 脱敏后内容（处理后填充）


class ExcelReader:
    """
    遍历单个 .xlsx 文件，提取所有有效的事件记录。

    参数
    ────
    column_name     : Resolution Notes 列名（支持精确匹配或模糊匹配）
    skip_sheets     : 跳过的 sheet 名称列表
    min_length      : 最短有效内容字符数
    invalid_values  : 需要跳过的固定值集合（小写比较）
    extra_columns   : 额外保留的列（用于 Prompt 上下文，如 Description、Member 等）
    """

    def __init__(
        self,
        column_name: str = "Resolution Notes",
        skip_sheets: list[str] | None = None,
        min_length: int = 20,
        invalid_values: set[str] | None = None,
        extra_columns: list[str] | None = None,
    ):
        self.column_name = column_name
        self.skip_sheets = set(skip_sheets or [])
        self.min_length = min_length
        self.invalid_values = invalid_values or _DEFAULT_INVALID_VALUES
        self.extra_columns = extra_columns or ["Description", "Case Contact Name", "Member", "Location"]

    # ── 公开方法 ──────────────────────────────────────────────────────────────

    def read(self, xlsx_path: str | Path) -> list[EventRecord]:
        """
        读取一个 .xlsx 文件，返回有效事件记录列表。
        """
        path = Path(xlsx_path)
        if not path.exists():
            raise FileNotFoundError(f"文件不存在: {path}")
        if path.suffix.lower() != ".xlsx":
            raise ValueError(f"仅支持 .xlsx 格式，收到: {path.suffix}")

        records: list[EventRecord] = []
        xl = pd.ExcelFile(str(path), engine="openpyxl")

        for sheet_name in xl.sheet_names:
            if sheet_name in self.skip_sheets:
                continue
            try:
                sheet_records = self._parse_sheet(xl, sheet_name, path.name)
                records.extend(sheet_records)
            except Exception as exc:
                # 单个 Sheet 出错不影响其他 Sheet
                from .logger import get_logger
                get_logger().warning(f"[{path.name}] Sheet '{sheet_name}' 解析失败: {exc}")

        return records

    def read_all(self, input_dir: str | Path) -> dict[str, list[EventRecord]]:
        """
        遍历目录下所有 .xlsx 文件，返回 {文件名: [EventRecord, ...]} 字典。
        """
        input_dir = Path(input_dir)
        if not input_dir.is_dir():
            raise NotADirectoryError(f"输入目录不存在: {input_dir}")

        result: dict[str, list[EventRecord]] = {}
        xlsx_files = sorted(input_dir.glob("*.xlsx"))

        if not xlsx_files:
            raise FileNotFoundError(f"目录 '{input_dir}' 中没有找到 .xlsx 文件")

        for xlsx_file in xlsx_files:
            try:
                records = self.read(xlsx_file)
                if records:
                    result[xlsx_file.name] = records
            except Exception as exc:
                from .logger import get_logger
                get_logger().error(f"读取文件失败 [{xlsx_file.name}]: {exc}")

        return result

    # ── 内部方法 ──────────────────────────────────────────────────────────────

    def _parse_sheet(
        self, xl: pd.ExcelFile, sheet_name: str, source_file: str
    ) -> list[EventRecord]:
        """解析单个 Sheet，返回有效记录列表。"""
        df = xl.parse(sheet_name, header=None)

        if df.empty:
            return []

        # 自动检测表头行：找第一行包含目标列名的行
        header_row = self._detect_header_row(df)
        if header_row is None:
            return []   # 该 Sheet 没有目标列，跳过

        # 重新读取，以检测到的行作为表头
        df = xl.parse(sheet_name, header=header_row)

        # 查找目标列（支持模糊匹配）
        notes_col = self._find_column(df.columns.tolist(), self.column_name)
        if notes_col is None:
            return []

        records: list[EventRecord] = []
        for idx, row in df.iterrows():
            raw_value = row.get(notes_col)
            content = self._clean_content(raw_value)

            if not self._is_valid(content):
                continue

            # 提取额外字段（用于增强 Prompt 上下文）
            extra = {}
            for col in self.extra_columns:
                matched = self._find_column(df.columns.tolist(), col)
                if matched:
                    val = row.get(matched)
                    if pd.notna(val) and str(val).strip():
                        extra[col] = str(val).strip()

            # row_index: header_row(0-based) + idx(0-based后行) + 2 = Excel行号(1-based含表头)
            excel_row = header_row + int(idx) + 2

            records.append(
                EventRecord(
                    source_file=source_file,
                    sheet_name=sheet_name,
                    row_index=excel_row,
                    content=content,
                    extra_fields=extra,
                )
            )

        return records

    def _detect_header_row(self, df: pd.DataFrame) -> Optional[int]:
        """
        自动检测表头行：遍历前20行，找到包含目标列名的行。
        支持模糊匹配（忽略大小写和空格）。
        """
        target_lower = self.column_name.lower().replace(" ", "")
        for i, row in df.head(20).iterrows():
            for cell in row:
                if pd.notna(cell) and target_lower in str(cell).lower().replace(" ", ""):
                    return int(i)
        return None

    def _find_column(self, columns: list, target: str) -> Optional[str]:
        """模糊查找列名（忽略大小写和空格）。"""
        target_lower = target.lower().replace(" ", "")
        for col in columns:
            if target_lower in str(col).lower().replace(" ", ""):
                return col
        return None

    @staticmethod
    def _clean_content(raw: object) -> str:
        """清理单元格内容：转字符串、去首尾空白、统一换行符。"""
        if raw is None or (isinstance(raw, float) and pd.isna(raw)):
            return ""
        text = str(raw).strip()
        # 统一换行符
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        # 去除多余空行
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text

    def _is_valid(self, content: str) -> bool:
        """判断内容是否有效（非空、非无效值、达到最小长度）。"""
        if not content:
            return False
        if content.lower() in self.invalid_values:
            return False
        if len(content) < self.min_length:
            return False
        return True
