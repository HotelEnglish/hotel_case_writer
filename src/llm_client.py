"""
llm_client.py
─────────────
统一的 LLM 调用层，支持：
  - OpenAI 官方 API
  - Ollama 本地推理（兼容 OpenAI 接口）
  - DeepSeek / 智谱 / 千问(通义千问) / 其他 OpenAI 兼容接口
  - Azure OpenAI

核心功能：
  - 指数退避自动重试（含 429 限速处理）
  - 速率限制（令牌桶）
  - Token 计数与成本估算
  - 流式输出支持（可选）
"""

from __future__ import annotations

import os
import time
import threading
from dataclasses import dataclass, field
from typing import Optional

import tiktoken
from openai import OpenAI, AzureOpenAI, RateLimitError, APITimeoutError, APIConnectionError


# ── 数据类 ────────────────────────────────────────────────────────────────────

@dataclass
class LLMConfig:
    provider: str = "ollama"              # openai | ollama | azure | deepseek | zhipu | qwen | custom
    api_key: str = "ollama"               # Ollama 不需要真实 key
    base_url: str = "http://localhost:11434/v1"
    model: str = "qwen3:4b"
    temperature: float = 0.75
    max_tokens: int = 4096
    request_timeout: int = 120
    max_retries: int = 3
    retry_base_delay: float = 5.0         # 指数退避基础等待时间（秒）
    requests_per_minute: int = 20         # 速率限制
    # Azure 专用
    azure_endpoint: str = ""
    azure_api_version: str = "2024-02-15-preview"
    azure_deployment: str = ""
    # 成本相关
    input_price_per_1k: float = 0.001     # 美元/千Token
    output_price_per_1k: float = 0.003


@dataclass
class GenerateResult:
    content: str
    input_tokens: int = 0
    output_tokens: int = 0
    total_tokens: int = 0
    estimated_cost_usd: float = 0.0
    retries: int = 0
    success: bool = True
    error: str = ""


# ── 速率限制令牌桶 ────────────────────────────────────────────────────────────

class _RateLimiter:
    """简单的滑动窗口速率限制器（线程安全）。"""

    def __init__(self, requests_per_minute: int):
        self.rpm = max(1, requests_per_minute)
        self._interval = 60.0 / self.rpm
        self._lock = threading.Lock()
        self._last_request_time = 0.0

    def acquire(self):
        with self._lock:
            now = time.monotonic()
            wait = self._interval - (now - self._last_request_time)
            if wait > 0:
                time.sleep(wait)
            self._last_request_time = time.monotonic()


# ── 主客户端 ─────────────────────────────────────────────────────────────────

class LLMClient:
    """统一 LLM 调用客户端。"""

    def __init__(self, config: LLMConfig):
        self.config = config
        self._rate_limiter = _RateLimiter(config.requests_per_minute)
        self._client = self._build_client()
        self._tokenizer = self._build_tokenizer()

    # ── 公开方法 ──────────────────────────────────────────────────────────────

    def generate(
        self,
        system_prompt: str,
        user_message: str,
        stream: bool = False,
    ) -> GenerateResult:
        """
        发送请求，自动处理重试和速率限制。
        返回 GenerateResult。
        """
        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message},
        ]

        last_error = ""
        for attempt in range(self.config.max_retries + 1):
            try:
                self._rate_limiter.acquire()
                result = self._do_request(messages, stream)
                result.retries = attempt
                return result

            except RateLimitError as e:
                last_error = f"RateLimit: {e}"
                wait = self.config.retry_base_delay * (2 ** attempt) + 10
                _log_warn(f"遇到限速(429)，等待 {wait:.0f}s 后重试 (第{attempt+1}次)...")
                time.sleep(wait)

            except (APITimeoutError, APIConnectionError) as e:
                last_error = f"连接超时/失败: {e}"
                wait = self.config.retry_base_delay * (2 ** attempt)
                _log_warn(f"请求超时/连接失败，等待 {wait:.0f}s 后重试 (第{attempt+1}次)...")
                time.sleep(wait)

            except Exception as e:
                last_error = str(e)
                if attempt < self.config.max_retries:
                    wait = self.config.retry_base_delay * (2 ** attempt)
                    _log_warn(f"请求失败: {e}，{wait:.0f}s 后重试 (第{attempt+1}次)...")
                    time.sleep(wait)
                else:
                    break

        return GenerateResult(
            content="",
            success=False,
            error=last_error,
            retries=self.config.max_retries,
        )

    def count_tokens(self, text: str) -> int:
        """估算文本的 Token 数量。"""
        if self._tokenizer:
            try:
                return len(self._tokenizer.encode(text))
            except Exception:
                pass
        # 降级估算：中文约1.5字/token，英文约4字符/token
        chinese = sum(1 for c in text if "\u4e00" <= c <= "\u9fff")
        english_chars = len(text) - chinese
        return int(chinese / 1.5 + english_chars / 4)

    def estimate_cost(self, input_tokens: int, output_tokens: int) -> float:
        """估算单次请求费用（美元）。"""
        return (
            input_tokens / 1000 * self.config.input_price_per_1k
            + output_tokens / 1000 * self.config.output_price_per_1k
        )

    # ── 内部方法 ──────────────────────────────────────────────────────────────

    def _do_request(self, messages: list[dict], stream: bool) -> GenerateResult:
        """执行实际的 API 调用。"""
        kwargs = dict(
            model=self.config.model if self.config.provider != "azure" else self.config.azure_deployment,
            messages=messages,
            temperature=self.config.temperature,
            max_tokens=self.config.max_tokens,
            stream=stream,
        )

        if stream:
            return self._handle_stream(kwargs)
        else:
            response = self._client.chat.completions.create(**kwargs)
            content = response.choices[0].message.content or ""
            usage = response.usage

            input_tokens = usage.prompt_tokens if usage else self.count_tokens(str(messages))
            output_tokens = usage.completion_tokens if usage else self.count_tokens(content)
            total = (usage.total_tokens if usage else input_tokens + output_tokens)

            return GenerateResult(
                content=content,
                input_tokens=input_tokens,
                output_tokens=output_tokens,
                total_tokens=total,
                estimated_cost_usd=self.estimate_cost(input_tokens, output_tokens),
                success=True,
            )

    def _handle_stream(self, kwargs: dict) -> GenerateResult:
        """处理流式输出，拼接完整内容。"""
        chunks = []
        stream = self._client.chat.completions.create(**kwargs)
        for chunk in stream:
            delta = chunk.choices[0].delta
            if delta and delta.content:
                chunks.append(delta.content)
                print(delta.content, end="", flush=True)
        print()  # 换行
        content = "".join(chunks)
        input_tokens = self.count_tokens(str(kwargs["messages"]))
        output_tokens = self.count_tokens(content)
        return GenerateResult(
            content=content,
            input_tokens=input_tokens,
            output_tokens=output_tokens,
            total_tokens=input_tokens + output_tokens,
            estimated_cost_usd=self.estimate_cost(input_tokens, output_tokens),
            success=True,
        )

    def _build_client(self) -> OpenAI:
        """根据 provider 配置创建 OpenAI 兼容客户端。"""
        cfg = self.config
        if cfg.provider == "azure":
            return AzureOpenAI(
                api_key=cfg.azure_endpoint and cfg.api_key,
                azure_endpoint=cfg.azure_endpoint,
                api_version=cfg.azure_api_version,
                timeout=cfg.request_timeout,
            )
        # Ollama / OpenAI / DeepSeek / 智谱 / 千问 / Custom 均兼容 OpenAI SDK
        return OpenAI(
            api_key=cfg.api_key,
            base_url=cfg.base_url,
            timeout=cfg.request_timeout,
        )

    def _build_tokenizer(self):
        """尝试加载 tiktoken tokenizer，失败则降级估算。"""
        try:
            # gpt-4o 系列
            return tiktoken.get_encoding("o200k_base")
        except Exception:
            try:
                return tiktoken.get_encoding("cl100k_base")
            except Exception:
                return None


# ── 工厂函数：从配置文件和环境变量构建 LLMClient ─────────────────────────────

def build_client_from_config(cfg_dict: dict) -> LLMClient:
    """
    从 config.yaml 的 llm 节点 + 环境变量构建 LLMClient。
    环境变量优先级高于配置文件。
    """
    provider = os.getenv("LLM_PROVIDER", "ollama").lower()

    if provider == "azure":
        llm_cfg = LLMConfig(
            provider="azure",
            api_key=os.getenv("AZURE_OPENAI_API_KEY", ""),
            azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT", ""),
            azure_deployment=os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-4o-mini"),
            azure_api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2024-02-15-preview"),
            model=os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-4o-mini"),
        )
    elif provider == "ollama":
        llm_cfg = LLMConfig(
            provider="ollama",
            api_key="ollama",
            base_url=os.getenv("OLLAMA_BASE_URL", "http://localhost:11434/v1"),
            model=os.getenv("OLLAMA_MODEL", "qwen3:4b"),
        )
    elif provider == "qwen":
        # 通义千问 DashScope OpenAI 兼容接口
        llm_cfg = LLMConfig(
            provider="qwen",
            api_key=os.getenv("DASHSCOPE_API_KEY", os.getenv("OPENAI_API_KEY", "")),
            base_url=os.getenv("DASHSCOPE_BASE_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1"),
            model=os.getenv("QWEN_MODEL", "qwen-plus"),
        )
    else:
        # openai / deepseek / zhipu / custom
        llm_cfg = LLMConfig(
            provider=provider,
            api_key=os.getenv("OPENAI_API_KEY", ""),
            base_url=os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1"),
            model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
        )

    # 应用 config.yaml 中的通用参数
    llm_cfg.temperature = float(cfg_dict.get("temperature", llm_cfg.temperature))
    llm_cfg.max_tokens = int(cfg_dict.get("max_tokens", llm_cfg.max_tokens))
    llm_cfg.request_timeout = int(cfg_dict.get("request_timeout", llm_cfg.request_timeout))
    llm_cfg.max_retries = int(cfg_dict.get("max_retries", llm_cfg.max_retries))
    llm_cfg.retry_base_delay = float(cfg_dict.get("retry_base_delay", llm_cfg.retry_base_delay))
    llm_cfg.requests_per_minute = int(cfg_dict.get("requests_per_minute", llm_cfg.requests_per_minute))

    return LLMClient(llm_cfg)


def _log_warn(msg: str):
    """临时日志输出，正式运行时由 logger 接管。"""
    print(f"[WARN] {msg}")
