"""
app_ui.py
─────────
Streamlit 图形界面 - 酒店案例批量改写工具

启动方式：
  python main.py --ui
  或直接：
  streamlit run app_ui.py
"""

import json
import os
import sys
import time
import threading
from pathlib import Path

import streamlit as st
import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))

from src.config_loader import load_config, get_paths
from src.logger import setup_logging, get_logger
from src.excel_reader import ExcelReader
from src.desensitizer import DesensitizeConfig
from src.llm_client import build_client_from_config, LLMConfig, LLMClient
from src.model_pool import ModelPool, ModelEntry, build_model_pool_from_configs
from src.prompt_manager import PromptManager
from src.progress_tracker import ProgressTracker
from src.processor import Processor, ProcessConfig, PauseController
from image_restorer import process_folder as extract_images, RestoreResult


# ── 页面配置 ──────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="酒店案例批量改写工具",
    page_icon="🏨",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 配置持久化路径 ────────────────────────────────────────────────────────────
_CONFIG_STORE_PATH = Path("./logs/ui_config.json")

# 每个服务商的默认配置（base_url、model、api_key占位）
_PROVIDER_DEFAULTS: dict[str, dict] = {
    "ollama": {
        "base_url": "http://localhost:11434/v1",
        "model": "gemma3:1b",
        "api_key": "ollama",
        "show_key": False,
    },
    "openai": {
        "base_url": "https://api.openai.com/v1",
        "model": "gpt-5-mini",
        "api_key": "",
        "show_key": True,
    },
    "deepseek": {
        "base_url": "https://api.deepseek.com/v1",
        "model": "deepseek-chat",
        "api_key": "",
        "show_key": True,
    },
    "zhipu": {
        "base_url": "https://open.bigmodel.cn/api/paas/v4",
        "model": "GLM-4.7-Flash",
        "api_key": "",
        "show_key": True,
    },
    "qwen": {
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "model": "glm-4.7",
        "api_key": "",
        "show_key": True,
    },
    "azure": {
        "base_url": "",
        "model": "gpt-5-mini",
        "api_key": "",
        "show_key": True,
    },
    "custom": {
        "base_url": "",
        "model": "",
        "api_key": "",
        "show_key": True,
    },
}

_PROVIDER_LABELS = {
    "ollama":  "Ollama（本地）",
    "openai":  "OpenAI",
    "deepseek":"DeepSeek",
    "zhipu":   "智谱AI（GLM）",
    "qwen":    "通义千问（Qwen）",
    "azure":   "Azure OpenAI",
    "custom":  "自定义（Custom）",
}

_QWEN_MODELS = [
    "qwen3-vl-flash-2026-01-22", "glm-4.7", "tongyi-xiaomi-analysis-flash",  "tongyi-xiaomi-analysis-pro","MiniMax-M2.1","qwen3-max-2026-01-23", "kimi-k2.5","qwen3.5-flash",
    "qwen3.5-397b-a17b", "glm-5", "qwen3.5-plus-2026-02-15","qwen3.5-plus",  "MiniMax-M2.5",
]
_OPENAI_MODELS = ["gpt-5-mini", "gpt-5.4", "gpt-5.3-codex", "gpt-5.2"]
_DEEPSEEK_MODELS = ["deepseek-chat", "deepseek-reasoner"]
_ZHIPU_MODELS = ["GLM-4.7-Flash", "GLM-4.6V-Flash", "GLM-4-Flash-250414"]


# ── 配置持久化工具函数 ─────────────────────────────────────────────────────────

def _load_stored_configs() -> dict:
    """从 JSON 文件加载已保存的服务商配置。"""
    if _CONFIG_STORE_PATH.exists():
        try:
            return json.loads(_CONFIG_STORE_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def _save_pool_models(pool_models: list[dict]):
    """将轮换池模型配置持久化到 JSON 文件。"""
    _CONFIG_STORE_PATH.parent.mkdir(parents=True, exist_ok=True)
    stored = _load_stored_configs()
    stored["_pool_models"] = pool_models
    _CONFIG_STORE_PATH.write_text(
        json.dumps(stored, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _save_provider_config(provider: str, api_key: str, base_url: str, model: str):
    """将当前服务商的配置持久化到 JSON 文件。"""
    _CONFIG_STORE_PATH.parent.mkdir(parents=True, exist_ok=True)
    stored = _load_stored_configs()
    stored[provider] = {"api_key": api_key, "base_url": base_url, "model": model}
    _CONFIG_STORE_PATH.write_text(
        json.dumps(stored, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _get_provider_config(provider: str) -> dict:
    """获取服务商配置：优先读已保存，其次读环境变量，最后用默认值。"""
    stored = _load_stored_configs().get(provider, {})
    defaults = _PROVIDER_DEFAULTS.get(provider, {})

    # 从环境变量读取（作为初始值参考，不强制覆盖已保存的配置）
    env_api_key = _env_api_key_for(provider)
    env_base_url = _env_base_url_for(provider)
    env_model    = _env_model_for(provider)

    return {
        "api_key":  stored.get("api_key")  or env_api_key  or defaults.get("api_key", ""),
        "base_url": stored.get("base_url") or env_base_url or defaults.get("base_url", ""),
        "model":    stored.get("model")    or env_model    or defaults.get("model", ""),
    }


def _env_api_key_for(provider: str) -> str:
    if provider == "ollama":  return "ollama"
    if provider == "azure":   return os.getenv("AZURE_OPENAI_API_KEY", "")
    if provider == "qwen":    return os.getenv("DASHSCOPE_API_KEY", os.getenv("OPENAI_API_KEY", ""))
    return os.getenv("OPENAI_API_KEY", "")


def _env_base_url_for(provider: str) -> str:
    if provider == "ollama":  return os.getenv("OLLAMA_BASE_URL", "http://localhost:11434/v1")
    if provider == "azure":   return os.getenv("AZURE_OPENAI_ENDPOINT", "")
    if provider == "qwen":    return os.getenv("DASHSCOPE_BASE_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1")
    return os.getenv("OPENAI_BASE_URL", _PROVIDER_DEFAULTS.get(provider, {}).get("base_url", ""))


def _env_model_for(provider: str) -> str:
    if provider == "ollama":  return os.getenv("OLLAMA_MODEL", "qwen3:4b")
    if provider == "azure":   return os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-4o-mini")
    if provider == "qwen":    return os.getenv("QWEN_MODEL", "qwen-plus")
    return os.getenv("OPENAI_MODEL", _PROVIDER_DEFAULTS.get(provider, {}).get("model", ""))


# ── 工具函数 ──────────────────────────────────────────────────────────────────

@st.cache_data(ttl=300)
def load_cfg():
    return load_config("config.yaml")


# ── 侧边栏：配置区 ────────────────────────────────────────────────────────────

def render_sidebar():
    st.sidebar.title("⚙️ 配置")

    with st.sidebar.expander("📡 LLM 服务配置", expanded=True):

        provider_keys = list(_PROVIDER_LABELS.keys())
        provider_display = list(_PROVIDER_LABELS.values())

        # 从 session_state 恢复上次选择的服务商
        saved_idx = provider_keys.index(
            st.session_state.get("last_provider", "ollama")
        ) if st.session_state.get("last_provider", "ollama") in provider_keys else 0

        provider_label = st.selectbox(
            "服务商",
            provider_display,
            index=saved_idx,
            key="provider_selectbox",
        )
        provider = provider_keys[provider_display.index(provider_label)]

        # 服务商切换时自动加载已保存的配置
        if st.session_state.get("last_provider") != provider:
            st.session_state["last_provider"] = provider
            # 触发 Streamlit 重渲染，使下方 widget 值更新
            st.rerun()

        # 读取当前服务商的已保存配置
        saved_cfg = _get_provider_config(provider)

        if provider == "ollama":
            base_url = st.text_input(
                "Ollama 地址", value=saved_cfg["base_url"], key=f"{provider}_url"
            )
            model = st.text_input(
                "模型名称", value=saved_cfg["model"], key=f"{provider}_model",
                help="示例: qwen3:4b, llama3:8b, gemma3:4b"
            )
            api_key = "ollama"

        elif provider == "azure":
            base_url = st.text_input(
                "Azure Endpoint", value=saved_cfg["base_url"], key=f"{provider}_url"
            )
            model = st.text_input(
                "Deployment 名称", value=saved_cfg["model"], key=f"{provider}_model"
            )
            api_key = st.text_input(
                "Azure API Key", value=saved_cfg["api_key"],
                type="password", key=f"{provider}_key"
            )

        elif provider == "qwen":
            base_url = st.text_input(
                "DashScope Base URL", value=saved_cfg["base_url"], key=f"{provider}_url"
            )
            # 千问提供常用模型下拉选择
            qwen_model_options = _QWEN_MODELS
            current_model = saved_cfg["model"]
            if current_model not in qwen_model_options:
                qwen_model_options = [current_model] + qwen_model_options
            model = st.selectbox(
                "模型", qwen_model_options,
                index=qwen_model_options.index(current_model),
                key=f"{provider}_model"
            )
            api_key = st.text_input(
                "DashScope API Key", value=saved_cfg["api_key"],
                type="password", key=f"{provider}_key",
                help="在阿里云百炼控制台获取：https://bailian.console.aliyun.com/"
            )

        elif provider == "deepseek":
            base_url = st.text_input(
                "API Base URL", value=saved_cfg["base_url"], key=f"{provider}_url"
            )
            model = st.selectbox(
                "模型", _DEEPSEEK_MODELS,
                index=_DEEPSEEK_MODELS.index(saved_cfg["model"]) if saved_cfg["model"] in _DEEPSEEK_MODELS else 0,
                key=f"{provider}_model"
            )
            api_key = st.text_input(
                "API Key", value=saved_cfg["api_key"],
                type="password", key=f"{provider}_key"
            )

        elif provider == "zhipu":
            base_url = st.text_input(
                "API Base URL", value=saved_cfg["base_url"], key=f"{provider}_url"
            )
            model = st.selectbox(
                "模型", _ZHIPU_MODELS,
                index=_ZHIPU_MODELS.index(saved_cfg["model"]) if saved_cfg["model"] in _ZHIPU_MODELS else 0,
                key=f"{provider}_model"
            )
            api_key = st.text_input(
                "API Key", value=saved_cfg["api_key"],
                type="password", key=f"{provider}_key"
            )

        elif provider == "openai":
            base_url = st.text_input(
                "API Base URL", value=saved_cfg["base_url"], key=f"{provider}_url"
            )
            model_options = _OPENAI_MODELS
            current_model = saved_cfg["model"]
            if current_model not in model_options:
                model_options = [current_model] + model_options
            model = st.selectbox(
                "模型", model_options,
                index=model_options.index(current_model),
                key=f"{provider}_model"
            )
            api_key = st.text_input(
                "API Key", value=saved_cfg["api_key"],
                type="password", key=f"{provider}_key"
            )

        else:  # custom
            base_url = st.text_input(
                "API Base URL", value=saved_cfg["base_url"], key=f"{provider}_url"
            )
            model = st.text_input(
                "模型名称", value=saved_cfg["model"], key=f"{provider}_model"
            )
            api_key = st.text_input(
                "API Key", value=saved_cfg["api_key"],
                type="password", key=f"{provider}_key"
            )

        temperature = st.slider("Temperature", 0.0, 1.0, 0.75, 0.05)
        max_tokens = st.number_input("Max Tokens", min_value=1000, max_value=8192, value=4096, step=256)

        # 保存配置按钮
        col_save, col_hint = st.columns([1, 2])
        with col_save:
            if st.button("💾", width='stretch'):
                _save_provider_config(provider, api_key, base_url, model)
                st.success("已保存！")
        with col_hint:
            st.caption("保存后切换服务商再切回，配置自动填入")

    with st.sidebar.expander("🔀 多模型轮换池（可选）", expanded=False):
        st.caption(
            "配置后，系统将**轮换使用**所有已启用的模型，有效避免单个模型被限速或余额耗尽。"
            "某个模型连续失败后自动冷却，稳定后自动恢复。"
        )
        pool_enabled = st.checkbox("启用多模型轮换", value=False, key="pool_enabled")

        pool_models_raw = st.session_state.get("pool_models", [])
        if not pool_models_raw:
            pool_models_raw = _load_stored_configs().get("_pool_models", [])
            if pool_models_raw:
                st.session_state["pool_models"] = pool_models_raw

        # 显示已添加的模型列表
        pool_models = list(st.session_state.get("pool_models", []))

        if pool_models:
            st.write(f"**已配置 {len(pool_models)} 个模型：**")
            to_remove = []
            for i, m in enumerate(pool_models):
                col_info, col_del = st.columns([4, 1])
                with col_info:
                    st.caption(f"**{i+1}.** {m.get('label', m.get('model',''))} ({m.get('provider','')})")
                with col_del:
                    if st.button("✕", key=f"del_pool_{i}", help="移除此模型"):
                        to_remove.append(i)
            if to_remove:
                for idx in sorted(to_remove, reverse=True):
                    pool_models.pop(idx)
                st.session_state["pool_models"] = pool_models
                _save_pool_models(pool_models)
                st.rerun()

        st.write("**添加模型到轮换池：**")
        new_provider_label = st.selectbox(
            "服务商", list(_PROVIDER_LABELS.values()),
            key="pool_new_provider"
        )
        new_provider = list(_PROVIDER_LABELS.keys())[
            list(_PROVIDER_LABELS.values()).index(new_provider_label)
        ]
        new_saved = _get_provider_config(new_provider)
        pool_new_url = st.text_input("Base URL", value=new_saved["base_url"], key="pool_new_url")
        pool_new_model = st.text_input("模型名称", value=new_saved["model"], key="pool_new_model")
        pool_new_key = st.text_input("API Key", value=new_saved["api_key"], type="password", key="pool_new_key")
        pool_new_label = st.text_input(
            "显示名称（可选）",
            value=f"{new_provider_label}:{new_saved['model']}",
            key="pool_new_label"
        )

        if st.button("➕ 添加到轮换池", key="add_pool_model"):
            if pool_new_model.strip() and pool_new_url.strip():
                new_entry = {
                    "label": pool_new_label.strip() or f"{new_provider}:{pool_new_model}",
                    "provider": new_provider,
                    "api_key": pool_new_key.strip(),
                    "base_url": pool_new_url.strip(),
                    "model": pool_new_model.strip(),
                }
                pool_models.append(new_entry)
                st.session_state["pool_models"] = pool_models
                _save_pool_models(pool_models)
                st.success(f"已添加：{new_entry['label']}")
                st.rerun()
            else:
                st.warning("请填写模型名称和 Base URL")

        if pool_models:
            cooldown_min = st.slider("冷却时间（分钟）", 1, 30, 5, 1, key="pool_cooldown")
            max_failures = st.slider("触发冷却的连续失败次数", 1, 10, 3, 1, key="pool_max_fail")
        else:
            cooldown_min = 5
            max_failures = 3

    with st.sidebar.expander("📂 路径配置", expanded=True):
        input_dir = st.text_input("输入目录（.xlsx 文件所在）", value="./input")
        output_dir = st.text_input("输出目录", value="./output")
        style_ref = st.text_input("范文路径（可选 .md 文件）", value="")

    with st.sidebar.expander("🔒 脱敏与字数", expanded=False):
        desensitize = st.checkbox("开启脱敏", value=True)
        dedup_names = st.checkbox("替换中文姓名", value=True)
        dedup_phone = st.checkbox("替换手机号", value=True)
        dedup_id = st.checkbox("替换身份证号", value=True)
        min_words = st.number_input("触发扩写重试的最低字数", 500, 2000, 1200, 100)
        max_words = st.number_input("最多汉字数", 1500, 4000, 2200, 100)

    with st.sidebar.expander("💰 成本估算", expanded=False):
        input_price = st.number_input("输入价格（$/千Token）", 0.0, 1.0, 0.001, 0.0001, format="%.4f")
        output_price = st.number_input("输出价格（$/千Token）", 0.0, 1.0, 0.003, 0.0001, format="%.4f")
        usd_to_cny = st.number_input("美元/人民币汇率", 5.0, 10.0, 7.2, 0.1)

    return {
        "provider": provider,
        "api_key": api_key,
        "base_url": base_url,
        "model": model,
        "temperature": temperature,
        "max_tokens": max_tokens,
        "input_dir": input_dir,
        "output_dir": output_dir,
        "style_ref": style_ref,
        "desensitize": desensitize,
        "dedup_names": dedup_names,
        "dedup_phone": dedup_phone,
        "dedup_id": dedup_id,
        "min_words": min_words,
        "max_words": max_words,
        "input_price": input_price,
        "output_price": output_price,
        "usd_to_cny": usd_to_cny,
        # 多模型池
        "pool_enabled": pool_enabled,
        "pool_models": pool_models,
        "pool_cooldown_seconds": cooldown_min * 60,
        "pool_max_failures": max_failures,
    }


# ── 主界面 ────────────────────────────────────────────────────────────────────

def main():
    st.title("🏨 酒店案例批量改写工具")
    st.caption("将酒店 Logbook 中的 Resolution Notes 自动扩写为专业培训案例")

    sidebar_cfg = render_sidebar()

    tabs = st.tabs(["📋 数据预览", "🚀 批量处理", "🖼️ 图片提取", "📊 进度统计", "📝 使用说明"])

    # ── Tab 1: 数据预览 ───────────────────────────────────────────────────────
    with tabs[0]:
        st.subheader("读取 & 预览 Excel 数据")
        col1, col2 = st.columns([3, 1])
        with col1:
            input_dir = st.text_input("输入目录", value=sidebar_cfg["input_dir"], key="preview_input")
        with col2:
            st.write("")
            st.write("")
            scan_btn = st.button("🔍 扫描文件", width='stretch')

        if scan_btn:
            input_path = Path(input_dir)
            if not input_path.exists():
                st.error(f"目录不存在: {input_path}")
            else:
                xlsx_files = list(input_path.glob("*.xlsx"))
                if not xlsx_files:
                    st.warning(f"目录 '{input_path}' 中没有找到 .xlsx 文件")
                else:
                    st.success(f"找到 {len(xlsx_files)} 个 .xlsx 文件")
                    reader = ExcelReader()
                    all_records = []
                    for f in xlsx_files:
                        try:
                            records = reader.read(f)
                            all_records.extend(records)
                        except Exception as e:
                            st.warning(f"读取 {f.name} 失败: {e}")

                    if all_records:
                        preview_data = [
                            {
                                "来源文件": r.source_file,
                                "Sheet": r.sheet_name,
                                "行号": r.row_index,
                                "内容预览": r.content[:80] + "..." if len(r.content) > 80 else r.content,
                                "字数": len(r.content),
                            }
                            for r in all_records
                        ]
                        st.dataframe(pd.DataFrame(preview_data), width='stretch', height=400)
                        st.info(f"共 {len(all_records)} 条有效记录待处理")
                        st.session_state["all_records"] = all_records
                    else:
                        st.warning("未找到有效的 Resolution Notes 记录")

    # ── Tab 2: 批量处理 ───────────────────────────────────────────────────────
    with tabs[1]:
        st.subheader("批量改写处理")

        col_a, col_b, col_c = st.columns(3)
        with col_a:
            dry_run = st.checkbox("🧮 仅估算成本（不调用 LLM）", value=False)
        with col_b:
            reset_progress = st.checkbox("🔄 重置断点续传记录", value=False)
        with col_c:
            single_file = st.text_input("只处理单个文件（留空则处理全部）", value="")

        # 初始化暂停控制器
        if "pause_ctrl" not in st.session_state:
            st.session_state["pause_ctrl"] = PauseController()

        pause_ctrl: PauseController = st.session_state["pause_ctrl"]
        is_running = st.session_state.get("processing_running", False)

        # 开始/暂停/恢复/停止 按钮区
        btn_cols = st.columns(4)
        with btn_cols[0]:
            start_clicked = st.button(
                "▶ 开始处理", type="primary", width='stretch',
                disabled=is_running
            )
        with btn_cols[1]:
            if is_running and not pause_ctrl.is_paused:
                if st.button("⏸ 暂停", width='stretch'):
                    pause_ctrl.pause()
                    st.rerun()
            elif is_running and pause_ctrl.is_paused:
                if st.button("▶ 恢复", type="primary", width='stretch'):
                    pause_ctrl.resume()
                    st.rerun()
            else:
                st.button("⏸ 暂停", width='stretch', disabled=True)
        with btn_cols[2]:
            if st.button("⏹ 停止", width='stretch', disabled=not is_running):
                pause_ctrl.stop()
                st.session_state["processing_running"] = False
                st.info("已发送停止信号，当前记录处理完成后终止。")
        with btn_cols[3]:
            # 模型池状态快速查看
            if sidebar_cfg.get("pool_enabled") and sidebar_cfg.get("pool_models"):
                pool_status = st.session_state.get("pool_status", [])
                if pool_status:
                    cooling_count = sum(1 for s in pool_status if s["is_cooling"])
                    if cooling_count > 0:
                        st.warning(f"🔀 {cooling_count}个模型冷却中")
                    else:
                        st.success(f"🔀 {len(pool_status)}个模型就绪")

        # 当前状态提示
        if is_running:
            if pause_ctrl.is_paused:
                st.warning("⏸ **任务已暂停** — 点击「恢复」继续，或「停止」终止任务")
            else:
                st.info("⏳ **处理中...** — 点击「暂停」可随时暂停，进度不会丢失")

        if start_clicked:
            # 重置暂停控制器
            st.session_state["pause_ctrl"] = PauseController()
            st.session_state["processing_running"] = True
            _run_processing(sidebar_cfg, dry_run, reset_progress, single_file)
            st.session_state["processing_running"] = False

    # ── Tab 3: 图片提取 ───────────────────────────────────────────────────────
    with tabs[2]:
        st.subheader("从 Excel 提取图片并修复变形")
        st.caption("支持从单个文件或整个文件夹批量提取图片，自动检测并修复被挤压变形的图像。")

        col1, col2 = st.columns(2)
        with col1:
            img_input = st.text_input(
                "输入路径（Excel 文件或文件夹）",
                value=sidebar_cfg["input_dir"],
                key="img_input",
            )
        with col2:
            img_output = st.text_input(
                "图片输出目录",
                value="./extracted_images",
                key="img_output",
            )

        col_a, col_b, col_c = st.columns(3)
        with col_a:
            fix_distortion = st.checkbox("自动修复变形图片", value=True)
        with col_b:
            threshold_pct = st.slider("变形检测阈值", 1, 20, 5, 1, help="宽高比差异超过此百分比则判定为变形")
        with col_c:
            img_recursive = st.checkbox("递归子文件夹", value=True)

        if st.button("🖼️ 开始提取", type="primary", width='stretch', key="img_btn"):
            _run_image_extraction(img_input, img_output, fix_distortion, threshold_pct / 100, img_recursive)

    # ── Tab 4: 进度统计 ───────────────────────────────────────────────────────
    with tabs[3]:
        st.subheader("处理进度统计")
        if st.button("🔄 刷新统计"):
            _show_stats(sidebar_cfg)
        _show_stats(sidebar_cfg)

    # ── Tab 5: 使用说明 ───────────────────────────────────────────────────────
    with tabs[4]:
        st.markdown(_USAGE_DOC)


def _run_processing(cfg: dict, dry_run: bool, reset_progress: bool, single_file: str):
    """在 Streamlit 中执行处理流程。"""
    progress_placeholder = st.empty()
    log_placeholder = st.empty()
    status_text = st.empty()

    # ── 构建 LLM 客户端 / 模型池 ──────────────────────────────────────────────
    pool_enabled = cfg.get("pool_enabled", False)
    pool_model_cfgs = cfg.get("pool_models", [])

    if pool_enabled and pool_model_cfgs:
        # 将主模型也加入轮换池（排第一位）
        primary_cfg = {
            "label": f"{cfg['provider']}:{cfg['model']}（主）",
            "provider": cfg["provider"],
            "api_key": cfg["api_key"],
            "base_url": cfg["base_url"],
            "model": cfg["model"],
            "temperature": cfg["temperature"],
            "max_tokens": cfg["max_tokens"],
            "input_price_per_1k": cfg["input_price"],
            "output_price_per_1k": cfg["output_price"],
        }
        # 为备用模型补充 temperature/max_tokens
        all_cfgs = [primary_cfg] + [
            {**m, "temperature": cfg["temperature"], "max_tokens": cfg["max_tokens"],
             "input_price_per_1k": cfg["input_price"], "output_price_per_1k": cfg["output_price"]}
            for m in pool_model_cfgs
        ]
        llm_or_pool = build_model_pool_from_configs(
            all_cfgs,
            max_consecutive_failures=cfg.get("pool_max_failures", 3),
            cooldown_seconds=cfg.get("pool_cooldown_seconds", 300),
        )
        st.info(
            f"🔀 **多模型轮换已启用** — 共 {len(all_cfgs)} 个模型："
            f" {', '.join(c['label'] for c in all_cfgs)}"
        )
    else:
        llm_cfg = LLMConfig(
            provider=cfg["provider"],
            api_key=cfg["api_key"],
            base_url=cfg["base_url"],
            model=cfg["model"],
            temperature=cfg["temperature"],
            max_tokens=cfg["max_tokens"],
            input_price_per_1k=cfg["input_price"],
            output_price_per_1k=cfg["output_price"],
        )
        llm_or_pool = LLMClient(llm_cfg)

    style_ref = cfg["style_ref"] if cfg["style_ref"].strip() else None
    prompt_manager = PromptManager(style_ref_file=style_ref)

    db_path = Path("./logs/progress.db")
    tracker = ProgressTracker(db_path)

    if reset_progress:
        tracker.reset()
        st.info("已重置断点续传记录")

    proc_cfg = ProcessConfig(
        output_dir=Path(cfg["output_dir"]),
        min_word_count=cfg["min_words"],
        max_word_count=cfg["max_words"],
        desensitize_config=DesensitizeConfig(
            enabled=cfg["desensitize"],
            replace_chinese_names=cfg["dedup_names"],
            replace_phone=cfg["dedup_phone"],
            replace_id_card=cfg["dedup_id"],
        ),
    )

    # 取出暂停控制器
    pause_ctrl: PauseController = st.session_state.get("pause_ctrl", PauseController())

    processor = Processor(llm_or_pool, prompt_manager, tracker, proc_cfg, pause_ctrl)

    # 读取数据
    reader = ExcelReader()
    all_records = []
    try:
        if single_file.strip():
            all_records = reader.read(single_file.strip())
        else:
            file_map = reader.read_all(cfg["input_dir"])
            for recs in file_map.values():
                all_records.extend(recs)
    except Exception as e:
        st.error(f"读取 Excel 失败: {e}")
        return

    if not all_records:
        st.warning("没有找到有效记录")
        return

    # 成本预估
    estimate = processor.estimate_cost(all_records)
    with st.expander("💰 成本预估", expanded=True):
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("记录数", estimate["record_count"])
        c2.metric("预估Token", f"{estimate['estimated_total_tokens']:,}")
        c3.metric("预估费用(USD)", f"${estimate['estimated_cost_usd']:.4f}")
        c4.metric("预估费用(CNY)", f"¥{estimate['estimated_cost_cny']:.2f}")

    if dry_run:
        st.success("--仅估算模式-- 未调用 LLM，处理完毕。")
        return

    # 进度条
    progress_bar = st.progress(0)
    status_text.text("准备处理...")
    logs = []

    # 模型池状态展示区（若启用）
    pool_status_placeholder = st.empty()

    def progress_callback(current, total, message):
        pct = current / total if total > 0 else 0
        progress_bar.progress(pct)
        status_text.text(f"[{current}/{total}] {message}")
        logs.append(message)
        if len(logs) > 20:
            logs.pop(0)
        log_placeholder.code("\n".join(logs[-10:]))

        # 更新模型池状态（若为 ModelPool）
        if isinstance(llm_or_pool, ModelPool):
            pool_status = llm_or_pool.get_status()
            st.session_state["pool_status"] = pool_status
            _render_pool_status(pool_status_placeholder, pool_status)

    stats = processor.process_records(all_records, progress_callback=progress_callback)

    if pause_ctrl.is_stopped:
        st.warning(f"⏹ 任务已终止。已完成 {stats.done} 条，失败 {stats.failed} 条。")
    else:
        st.success("处理完成！")
        st.balloons()

    st.text(stats.summary(usd_to_cny=cfg["usd_to_cny"]))


def _render_pool_status(placeholder, pool_status: list[dict]):
    """在占位符中渲染模型池状态表格。"""
    if not pool_status:
        return
    rows = []
    for s in pool_status:
        cooling_str = f"❄️ 冷却中({s['cooldown_remaining_s']}s)" if s["is_cooling"] else "✅ 就绪"
        current_str = "👉 当前" if s["is_current"] else ""
        rows.append({
            "模型": s["label"],
            "状态": cooling_str,
            "连续失败": s["consecutive_failures"],
            "成功/总调用": f"{s['total_successes']}/{s['total_calls']}",
            "": current_str,
        })
    with placeholder.container():
        st.caption("🔀 模型池实时状态")
        st.dataframe(pd.DataFrame(rows), width='stretch', hide_index=True)


def _run_image_extraction(
    input_path_str: str,
    output_dir_str: str,
    fix_distortion: bool,
    threshold: float,
    recursive: bool,
):
    """在 Streamlit 中执行图片提取流程。"""
    input_path = Path(input_path_str.strip())
    output_dir = Path(output_dir_str.strip())

    if not input_path.exists():
        st.error(f"路径不存在: {input_path}")
        return

    with st.spinner("正在提取图片，请稍候..."):
        try:
            results = extract_images(
                input_path=input_path,
                output_dir=output_dir,
                fix_distortion=fix_distortion,
                distortion_threshold=threshold,
                recursive=recursive,
            )
        except Exception as e:
            st.error(f"提取过程中发生错误: {e}")
            return

    if not results:
        st.warning("未找到任何 Excel 文件或图片。")
        return

    total_images = sum(r.total_images for r in results)
    total_extracted = sum(r.extracted for r in results)
    total_fixed = sum(r.fixed for r in results)
    all_errors = [e for r in results for e in r.errors]

    st.success(f"提取完成！共处理 {len(results)} 个文件")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("发现图片", total_images)
    c2.metric("成功提取", total_extracted)
    c3.metric("修复变形", total_fixed)
    c4.metric("输出目录", str(output_dir))

    if all_errors:
        with st.expander(f"⚠️ 错误信息（{len(all_errors)} 条）"):
            for err in all_errors:
                st.text(err)

    with st.expander("📋 各文件详情"):
        detail_data = [
            {
                "文件": Path(r.excel_file).name,
                "图片总数": r.total_images,
                "提取成功": r.extracted,
                "修复变形": r.fixed,
                "错误数": len(r.errors),
            }
            for r in results
        ]
        if detail_data:
            st.dataframe(pd.DataFrame(detail_data), width='stretch')


def _show_stats(cfg: dict):
    db_path = Path("./logs/progress.db")
    if not db_path.exists():
        st.info("尚无处理记录。请先运行一次批量处理。")
        return
    tracker = ProgressTracker(db_path)
    stats = tracker.get_stats()
    if not stats:
        st.info("暂无数据")
        return

    total = sum(stats.values())
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("总计", total)
    c2.metric("✅ 完成", stats.get("done", 0))
    c3.metric("⏭️ 跳过", stats.get("skipped", 0))
    c4.metric("❌ 失败", stats.get("failed", 0))
    c5.metric("⏳ 待处理", stats.get("pending", 0))

    failed_records = tracker.get_failed_records()
    if failed_records:
        st.subheader("❌ 失败记录详情")
        fail_data = [
            {"文件": r.source_file, "行号": r.row_index, "错误": r.error_msg[:100]}
            for r in failed_records
        ]
        st.dataframe(pd.DataFrame(fail_data), width='stretch')


# ── 使用说明文档 ──────────────────────────────────────────────────────────────

_USAGE_DOC = """
## 快速上手

### 1. 安装依赖
```bash
pip install -r requirements.txt
```

### 2. 配置 API
复制 `.env.example` 为 `.env`，填入 API Key：
```bash
cp .env.example .env
```
如果使用 Ollama，确保本地服务已启动：
```bash
ollama run qwen3:4b
```

### 3. 准备数据
将 `.xlsx` 文件放入 `input/` 目录。
Excel 文件需包含 `Resolution Notes` 列。

### 4. 运行
**图形界面（推荐）：**
```bash
python main.py --ui
```

**命令行：**
```bash
python main.py                        # 处理所有文件
python main.py --dry-run              # 仅估算成本
python main.py --file ./input/a.xlsx  # 处理单文件
python main.py --reset                # 重置进度
```

### 5. 查看结果
生成的 `.md` 文件保存在 `output/` 目录，按来源文件名分子文件夹。

---

## 常见问题

**Q: 遇到 429 限速怎么办？**
A: 工具会自动等待并重试，可在 `config.yaml` 中调低 `requests_per_minute`。

**Q: 如何使用范文参考？**
A: 在侧边栏填入范文 `.md` 文件路径，工具会将其注入 Prompt 供 AI 参考语言风格。

**Q: 中断后如何续跑？**
A: 直接再次运行即可，已处理的记录会自动跳过（断点续传）。

**Q: 如何关闭脱敏？**
A: 在侧边栏取消勾选"开启脱敏"，或在 `config.yaml` 中设置 `desensitization.enabled: false`。
"""


if __name__ == "__main__":
    main()
