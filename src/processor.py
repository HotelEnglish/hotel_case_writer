"""
processor.py
────────────
核心处理流程：将 EventRecord 列表通过 LLM 改写为案例 .md 文件。
整合：脱敏 → Prompt构建 → LLM调用 → 字数验证 → 文件保存 → 进度记录
"""

from __future__ import annotations

import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Union

from .excel_reader import EventRecord
from .desensitizer import Desensitizer, DesensitizeConfig
from .llm_client import LLMClient, GenerateResult
from .model_pool import ModelPool
from .prompt_manager import PromptManager, count_chinese_chars, has_guide_questions, append_guide_questions_to_content
from .file_writer import save_case
from .progress_tracker import ProgressTracker
from .logger import get_logger, log_error_to_file


# ── 暂停控制器 ────────────────────────────────────────────────────────────────

class PauseController:
    """
    线程安全的暂停/恢复控制器。
    Processor 在每条记录处理前检查是否需要暂停。
    Streamlit UI 通过 pause() / resume() 控制状态。
    """

    def __init__(self):
        self._paused = False
        self._stopped = False
        self._event = threading.Event()
        self._event.set()   # 初始状态：未暂停（event set = 可继续）
        self._lock = threading.Lock()

    @property
    def is_paused(self) -> bool:
        return self._paused

    @property
    def is_stopped(self) -> bool:
        return self._stopped

    def pause(self):
        """暂停处理（下一条记录开始前生效）。"""
        with self._lock:
            self._paused = True
            self._event.clear()

    def resume(self):
        """恢复处理。"""
        with self._lock:
            self._paused = False
            self._event.set()

    def stop(self):
        """停止处理（不可恢复，用于彻底终止）。"""
        with self._lock:
            self._stopped = True
            self._paused = False
            self._event.set()   # 唤醒等待中的线程，让其检测到 stopped 后退出

    def wait_if_paused(self):
        """
        阻塞直到恢复或停止。
        在 Processor 主循环每条记录前调用。
        """
        self._event.wait()   # 暂停时 event 被 clear，此处阻塞



@dataclass
class ProcessConfig:
    output_dir: Path = Path("./output")
    output_subdir_by_file: bool = True
    add_metadata_header: bool = True
    encoding: str = "utf-8"
    filename_max_length: int = 60
    conflict_suffix: bool = True
    # 字数控制
    target_word_count: int = 2000
    min_word_count: int = 1200
    max_word_count: int = 2200
    retry_if_short: bool = True
    max_word_count_retries: int = 2
    # 引导问题检测
    ensure_guide_questions: bool = True   # 是否检测并补写引导问题
    # 脱敏
    desensitize_config: DesensitizeConfig = field(default_factory=DesensitizeConfig)


@dataclass
class ProcessStats:
    total: int = 0
    done: int = 0
    skipped: int = 0
    failed: int = 0
    paused_count: int = 0          # 累计暂停次数
    total_input_tokens: int = 0
    total_output_tokens: int = 0
    total_cost_usd: float = 0.0
    start_time: float = field(default_factory=time.time)

    @property
    def elapsed(self) -> float:
        return time.time() - self.start_time

    @property
    def success_rate(self) -> float:
        processed = self.done + self.failed
        return self.done / processed * 100 if processed > 0 else 0.0

    def summary(self, usd_to_cny: float = 7.2) -> str:
        cost_cny = self.total_cost_usd * usd_to_cny
        return (
            f"\n{'=' * 55}\n"
            f"  处理完成\n"
            f"  总计: {self.total} 条  |  成功: {self.done}  |  "
            f"跳过: {self.skipped}  |  失败: {self.failed}\n"
            f"  成功率: {self.success_rate:.1f}%\n"
            f"  Token消耗: 输入 {self.total_input_tokens:,}  输出 {self.total_output_tokens:,}\n"
            f"  预估费用: ${self.total_cost_usd:.4f} (≈ ¥{cost_cny:.2f})\n"
            f"  耗时: {self.elapsed:.1f}s\n"
            f"{'=' * 55}"
        )


class Processor:
    """
    批量处理器：协调所有模块，完成从 EventRecord → .md 文件的全流程。
    支持单模型（LLMClient）或多模型轮换池（ModelPool）。
    支持暂停/恢复（通过 PauseController）。
    """

    def __init__(
        self,
        llm_client: Union[LLMClient, ModelPool],
        prompt_manager: PromptManager,
        progress_tracker: ProgressTracker,
        process_config: ProcessConfig,
        pause_controller: Optional[PauseController] = None,
    ):
        self.llm = llm_client
        self.prompts = prompt_manager
        self.tracker = progress_tracker
        self.cfg = process_config
        self.pause_ctrl = pause_controller or PauseController()
        self.desensitizer = Desensitizer(process_config.desensitize_config)
        self.logger = get_logger()
        self._stats = ProcessStats()
        self._system_prompt = prompt_manager.build_system_prompt()

    # ── 公开方法 ──────────────────────────────────────────────────────────────

    def process_records(
        self,
        records: list[EventRecord],
        progress_callback=None,   # 可选回调，供 Streamlit 更新进度条
    ) -> ProcessStats:
        """
        处理 EventRecord 列表。
        progress_callback(current, total, message) -> None
        """
        self._stats = ProcessStats(total=len(records))

        for idx, record in enumerate(records):

            # ── 暂停检测 ──────────────────────────────────────────────────────
            if self.pause_ctrl.is_paused:
                self._stats.paused_count += 1
                if progress_callback:
                    progress_callback(idx + 1, self._stats.total, "⏸ 已暂停，等待恢复...")
                self.pause_ctrl.wait_if_paused()

            # ── 停止检测 ──────────────────────────────────────────────────────
            if self.pause_ctrl.is_stopped:
                self.logger.info("处理已被用户终止")
                break

            # 断点续传：已处理的跳过
            if self.tracker.is_done(record.source_file, record.sheet_name, record.row_index):
                self._stats.skipped += 1
                self.logger.info(
                    f"[跳过] {record.source_file} 第{record.row_index}行（已完成）"
                )
                if progress_callback:
                    progress_callback(idx + 1, self._stats.total, "跳过（已完成）")
                continue

            msg = f"[{idx+1}/{self._stats.total}] {record.source_file} 第{record.row_index}行"
            self.logger.info(f"处理: {msg}")
            if progress_callback:
                progress_callback(idx + 1, self._stats.total, msg)

            try:
                output_path = self._process_one(record)
                self._stats.done += 1
                self.tracker.mark_done(
                    record.source_file, record.sheet_name, record.row_index,
                    output_file=str(output_path)
                )
                self.logger.info(f"  ✓ 已保存: {output_path.name}")

            except Exception as e:
                self._stats.failed += 1
                error_msg = str(e)
                self.tracker.mark_failed(
                    record.source_file, record.sheet_name, record.row_index,
                    error_msg=error_msg
                )
                self.logger.error(f"  ✗ 失败: {error_msg}")
                log_error_to_file(
                    record.source_file, record.row_index,
                    error_msg, record.content[:200]
                )

        self.logger.info(self._stats.summary())
        return self._stats

    def estimate_cost(self, records: list[EventRecord]) -> dict:
        """
        运行前预估总费用（基于 Token 估算）。
        """
        # 取第一个可用客户端用于 token 计数
        ref_client = self.llm._entries[0].client if isinstance(self.llm, ModelPool) else self.llm
        system_tokens = ref_client.count_tokens(self._system_prompt)
        total_input_tokens = 0
        for record in records:
            user_msg = self.prompts.build_user_message(record)
            total_input_tokens += system_tokens + ref_client.count_tokens(user_msg)

        # 假设每条输出约 2000 汉字 ≈ 1500 tokens
        estimated_output_per_record = 1500
        total_output_tokens = len(records) * estimated_output_per_record
        total_input_tokens_all = total_input_tokens

        cost_usd = ref_client.estimate_cost(total_input_tokens_all, total_output_tokens)

        return {
            "record_count": len(records),
            "estimated_input_tokens": total_input_tokens_all,
            "estimated_output_tokens": total_output_tokens,
            "estimated_total_tokens": total_input_tokens_all + total_output_tokens,
            "estimated_cost_usd": cost_usd,
            "estimated_cost_cny": cost_usd * 7.2,
        }

    # ── 内部方法 ──────────────────────────────────────────────────────────────

    def _generate(self, system_prompt: str, user_message: str) -> GenerateResult:
        """统一调用入口：兼容单客户端和模型池。"""
        return self.llm.generate(system_prompt, user_message)

    def _process_one(self, record: EventRecord) -> Path:
        """处理单条记录，返回保存的文件路径。"""
        # 1. 脱敏 Resolution Notes 正文
        record.desensitized_content = self.desensitizer.desensitize(record.content)

        # 1b. 脱敏 extra_fields 中的高风险字段
        #     Case Contact Name 直接含客人全名，必须专项处理
        _NAME_FIELDS = {"Case Contact Name", "Member"}
        _TEXT_FIELDS = {"Description"}
        if self.desensitizer.config.enabled:
            for key in list(record.extra_fields.keys()):
                val = record.extra_fields[key]
                if key in _NAME_FIELDS:
                    record.extra_fields[key] = self.desensitizer.desensitize_name_field(val)
                elif key in _TEXT_FIELDS:
                    record.extra_fields[key] = self.desensitizer.desensitize(val)

        # 2. 构建 User Message
        user_message = self.prompts.build_user_message(record)

        # 3. 首次 LLM 调用
        result = self._generate(self._system_prompt, user_message)
        if not result.success:
            raise RuntimeError(f"LLM 调用失败: {result.error}")

        content = result.content
        self._stats.total_input_tokens += result.input_tokens
        self._stats.total_output_tokens += result.output_tokens
        self._stats.total_cost_usd += result.estimated_cost_usd

        # 4. 字数检查与重试
        if self.cfg.retry_if_short:
            content = self._ensure_word_count(content, record)

        # 5. 引导问题检测与补写
        if self.cfg.ensure_guide_questions:
            content = self._ensure_guide_questions(content)

        # 6. 保存文件
        output_path = save_case(
            content=content,
            record=record,
            output_dir=self.cfg.output_dir,
            output_subdir_by_file=self.cfg.output_subdir_by_file,
            add_metadata_header=self.cfg.add_metadata_header,
            encoding=self.cfg.encoding,
            filename_max_length=self.cfg.filename_max_length,
            conflict_suffix=self.cfg.conflict_suffix,
        )
        return output_path

    def _ensure_word_count(self, content: str, record: EventRecord) -> str:
        """字数不足时自动追加扩写指令重试。"""
        for attempt in range(self.cfg.max_word_count_retries):
            char_count = count_chinese_chars(content)
            if char_count >= self.cfg.min_word_count:
                break

            self.logger.warning(
                f"  字数不足（{char_count}字 < {self.cfg.min_word_count}字），"
                f"触发扩写重试 (第{attempt+1}次)..."
            )

            retry_message = self.prompts.build_retry_message(content, char_count)
            result = self._generate(self._system_prompt, retry_message)
            if result.success and result.content:
                content = result.content
                self._stats.total_input_tokens += result.input_tokens
                self._stats.total_output_tokens += result.output_tokens
                self._stats.total_cost_usd += result.estimated_cost_usd
            else:
                self.logger.warning(f"  扩写重试失败: {result.error}")
                break

        final_count = count_chinese_chars(content)
        self.logger.debug(f"  最终字数: {final_count} 汉字")
        return content

    def _ensure_guide_questions(self, content: str) -> str:
        """检测是否包含引导问题，缺失时自动触发补写。"""
        if has_guide_questions(content):
            return content

        self.logger.warning("  未检测到'引导问题'章节，触发补写...")
        supplement_msg = self.prompts.build_missing_guide_questions_message(content)
        result = self._generate(self._system_prompt, supplement_msg)

        if result.success and result.content:
            self._stats.total_input_tokens += result.input_tokens
            self._stats.total_output_tokens += result.output_tokens
            self._stats.total_cost_usd += result.estimated_cost_usd

            # 如果模型返回了完整重写版本（含引导问题），直接用新版
            if has_guide_questions(result.content) and count_chinese_chars(result.content) > 500:
                self.logger.info("  引导问题补写成功（完整重写版）")
                return result.content
            # 否则将补写内容追加到原文末尾
            merged = append_guide_questions_to_content(content, result.content)
            if has_guide_questions(merged):
                self.logger.info("  引导问题补写成功（追加模式）")
                return merged

        self.logger.warning("  引导问题补写失败，保留原始内容")
        return content

