"""
model_pool.py
─────────────
多模型池：支持轮换（Round-Robin）+ 自动降级的 LLM 调用策略。

策略说明：
- 正常情况下按顺序轮换使用所有已配置模型，分散负载、避免单一模型限速
- 某模型连续失败超过阈值后自动标记为「冷却中」，暂时跳过
- 冷却时间过后自动恢复（避免永久失效）
- 若所有模型均处于冷却状态，仍会使用失败次数最少的那个（兜底）
"""

from __future__ import annotations

import time
import threading
from dataclasses import dataclass, field
from typing import Optional

from .llm_client import LLMClient, LLMConfig, GenerateResult
from .logger import get_logger


@dataclass
class ModelEntry:
    """模型池中的单个模型条目。"""
    config: LLMConfig
    label: str = ""                         # 显示名称，如 "DeepSeek-Chat"
    consecutive_failures: int = 0           # 连续失败次数
    cooldown_until: float = 0.0             # 冷却截止时间戳（0 = 未冷却）
    total_calls: int = 0
    total_failures: int = 0
    total_successes: int = 0
    _client: Optional[LLMClient] = field(default=None, repr=False)

    @property
    def is_cooling(self) -> bool:
        return time.monotonic() < self.cooldown_until

    @property
    def client(self) -> LLMClient:
        if self._client is None:
            self._client = LLMClient(self.config)
        return self._client

    def reset_client(self):
        """重建客户端（配置变更后调用）。"""
        self._client = None


class ModelPool:
    """
    多模型轮换 + 自动降级池。

    使用方式：
        pool = ModelPool([entry1, entry2, ...])
        result = pool.generate(system_prompt, user_message)

    generate() 内部自动选择下一个可用模型，失败后切换下一个继续重试。
    """

    def __init__(
        self,
        entries: list[ModelEntry],
        max_consecutive_failures: int = 3,      # 连续失败N次后进入冷却
        cooldown_seconds: float = 300.0,         # 冷却时长（5分钟）
        max_pool_retries: int | None = None,     # 最多尝试多少个不同模型（None=全部）
    ):
        if not entries:
            raise ValueError("ModelPool 至少需要1个模型配置")
        self._entries = list(entries)
        self._max_failures = max_consecutive_failures
        self._cooldown = cooldown_seconds
        self._max_pool_retries = max_pool_retries or len(entries)
        self._index = 0          # 轮换指针
        self._lock = threading.Lock()
        self._logger = get_logger()

    # ── 公开方法 ──────────────────────────────────────────────────────────────

    @property
    def size(self) -> int:
        return len(self._entries)

    @property
    def current_label(self) -> str:
        """返回当前轮换到的模型名称（仅供展示）。"""
        with self._lock:
            return self._entries[self._index % len(self._entries)].label

    def generate(
        self,
        system_prompt: str,
        user_message: str,
        stream: bool = False,
    ) -> GenerateResult:
        """
        轮换调用模型，失败时自动切换下一个。
        所有模型均失败时返回最后一次失败的 GenerateResult。
        """
        tried = 0
        last_result: Optional[GenerateResult] = None
        start_index = self._get_next_index()

        for _ in range(min(len(self._entries), self._max_pool_retries)):
            entry = self._pick_entry(start_index + tried)
            tried += 1

            self._logger.debug(
                f"  [ModelPool] 使用模型: {entry.label} "
                f"(冷却: {entry.is_cooling}, 连续失败: {entry.consecutive_failures})"
            )

            result = entry.client.generate(system_prompt, user_message, stream)

            with self._lock:
                entry.total_calls += 1
                if result.success:
                    entry.total_successes += 1
                    entry.consecutive_failures = 0  # 成功则清零
                    self._logger.info(f"  [ModelPool] ✓ {entry.label} 调用成功")
                    return result
                else:
                    entry.total_failures += 1
                    entry.consecutive_failures += 1
                    if entry.consecutive_failures >= self._max_failures:
                        entry.cooldown_until = time.monotonic() + self._cooldown
                        self._logger.warning(
                            f"  [ModelPool] {entry.label} 连续失败 "
                            f"{entry.consecutive_failures} 次，进入 "
                            f"{self._cooldown:.0f}s 冷却"
                        )
                    else:
                        self._logger.warning(
                            f"  [ModelPool] {entry.label} 调用失败（{result.error[:80]}），"
                            f"切换下一个模型..."
                        )
                    last_result = result

        # 所有模型均失败
        self._logger.error(
            f"  [ModelPool] 所有 {len(self._entries)} 个模型均调用失败！"
        )
        return last_result or GenerateResult(
            content="", success=False, error="所有模型均不可用"
        )

    def get_status(self) -> list[dict]:
        """返回所有模型的当前状态（供 UI 展示）。"""
        now = time.monotonic()
        status = []
        for i, e in enumerate(self._entries):
            cooldown_remaining = max(0.0, e.cooldown_until - now)
            status.append({
                "index": i,
                "label": e.label,
                "provider": e.config.provider,
                "model": e.config.model,
                "is_cooling": e.is_cooling,
                "cooldown_remaining_s": round(cooldown_remaining),
                "consecutive_failures": e.consecutive_failures,
                "total_calls": e.total_calls,
                "total_successes": e.total_successes,
                "total_failures": e.total_failures,
                "is_current": (i == self._index % len(self._entries)),
            })
        return status

    def reset_cooldowns(self):
        """手动清除所有模型的冷却状态（UI 按钮调用）。"""
        with self._lock:
            for e in self._entries:
                e.cooldown_until = 0.0
                e.consecutive_failures = 0
        self._logger.info("[ModelPool] 已清除所有模型的冷却状态")

    # ── 内部方法 ──────────────────────────────────────────────────────────────

    def _get_next_index(self) -> int:
        """推进轮换指针，返回本次起始 index。"""
        with self._lock:
            idx = self._index
            self._index = (self._index + 1) % len(self._entries)
        return idx

    def _pick_entry(self, idx: int) -> ModelEntry:
        """
        从 idx 开始找第一个不在冷却中的模型；
        若全部冷却则返回失败次数最少的（兜底）。
        """
        n = len(self._entries)
        for offset in range(n):
            entry = self._entries[(idx + offset) % n]
            if not entry.is_cooling:
                return entry
        # 全部冷却，兜底选失败最少的
        return min(self._entries, key=lambda e: e.total_failures)


# ── 工厂函数 ──────────────────────────────────────────────────────────────────

def build_model_pool_from_configs(
    configs: list[dict],
    max_consecutive_failures: int = 3,
    cooldown_seconds: float = 300.0,
) -> ModelPool:
    """
    从配置字典列表构建 ModelPool。

    每个 config 字典格式：
    {
        "label":       "DeepSeek-Chat",      # 显示名称（可选）
        "provider":    "deepseek",
        "api_key":     "sk-...",
        "base_url":    "https://api.deepseek.com/v1",
        "model":       "deepseek-chat",
        "temperature": 0.75,                 # 可选
        "max_tokens":  4096,                 # 可选
    }
    """
    entries = []
    for cfg in configs:
        llm_cfg = LLMConfig(
            provider=cfg.get("provider", "custom"),
            api_key=cfg.get("api_key", ""),
            base_url=cfg.get("base_url", ""),
            model=cfg.get("model", ""),
            temperature=float(cfg.get("temperature", 0.75)),
            max_tokens=int(cfg.get("max_tokens", 4096)),
            input_price_per_1k=float(cfg.get("input_price_per_1k", 0.001)),
            output_price_per_1k=float(cfg.get("output_price_per_1k", 0.003)),
        )
        label = cfg.get("label") or f"{cfg.get('provider', 'custom')}:{cfg.get('model', '')}"
        entries.append(ModelEntry(config=llm_cfg, label=label))

    return ModelPool(
        entries,
        max_consecutive_failures=max_consecutive_failures,
        cooldown_seconds=cooldown_seconds,
    )
