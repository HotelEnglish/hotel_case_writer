#!/usr/bin/env python3
"""
main.py
───────
酒店案例批量改写工具 - 命令行入口

用法示例：
  # 基本运行（处理 input/ 目录下所有 xlsx）
  python main.py

  # 指定输入/输出目录
  python main.py --input ./my_excel --output ./my_cases

  # 只处理单个文件
  python main.py --file ./data/logbook_jan.xlsx

  # 运行前预估成本（不实际调用 LLM）
  python main.py --dry-run

  # 重置指定文件的断点续传进度（重新处理）
  python main.py --reset --file ./data/logbook_jan.xlsx

  # 使用范文参考
  python main.py --style-ref ./某温泉酒店Mini吧的陷阱.md

  # 启动 Streamlit 图形界面
  python main.py --ui
"""

import argparse
import sys
from pathlib import Path

# ── 确保 src 模块可以被导入 ───────────────────────────────────────────────────
sys.path.insert(0, str(Path(__file__).parent))

from src.config_loader import load_config, get_paths
from src.logger import setup_logging, get_logger
from src.excel_reader import ExcelReader
from src.desensitizer import Desensitizer, DesensitizeConfig
from src.llm_client import build_client_from_config
from src.prompt_manager import PromptManager
from src.progress_tracker import ProgressTracker
from src.processor import Processor, ProcessConfig


def parse_args():
    parser = argparse.ArgumentParser(
        description="酒店Logbook事件批量改写为标准化案例工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("--config", default="config.yaml", help="配置文件路径（默认: config.yaml）")
    parser.add_argument("--input", help="输入目录，覆盖 config.yaml 中的设置")
    parser.add_argument("--output", help="输出目录，覆盖 config.yaml 中的设置")
    parser.add_argument("--file", help="只处理单个 .xlsx 文件")
    parser.add_argument("--style-ref", help="范文文件路径（.md），用于风格参考")
    parser.add_argument("--dry-run", action="store_true", help="只预估成本，不实际调用 LLM")
    parser.add_argument("--reset", action="store_true", help="重置断点续传进度（与 --file 配合可只重置单个文件）")
    parser.add_argument("--stats", action="store_true", help="查看当前处理进度统计后退出")
    parser.add_argument("--ui", action="store_true", help="启动 Streamlit 图形界面")
    parser.add_argument("--verbose", action="store_true", help="显示详细调试信息")
    return parser.parse_args()


def main():
    args = parse_args()

    # ── 启动 UI 模式 ──────────────────────────────────────────────────────────
    if args.ui:
        import subprocess
        ui_path = Path(__file__).parent / "app_ui.py"
        subprocess.run([sys.executable, "-m", "streamlit", "run", str(ui_path)], check=True)
        return

    # ── 加载配置 ──────────────────────────────────────────────────────────────
    cfg = load_config(args.config)
    paths = get_paths(cfg)

    # 命令行参数覆盖配置文件
    if args.input:
        paths["input_dir"] = Path(args.input)
    if args.output:
        paths["output_dir"] = Path(args.output)
    if args.style_ref:
        paths["style_ref_file"] = Path(args.style_ref)

    # ── 初始化日志 ────────────────────────────────────────────────────────────
    import logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    setup_logging(paths["log_file"], paths["error_log"], level=log_level)
    logger = get_logger()

    logger.info("=" * 55)
    logger.info("  酒店案例批量改写工具 启动")
    logger.info("=" * 55)

    # ── 初始化进度追踪器 ──────────────────────────────────────────────────────
    tracker = ProgressTracker(paths["progress_db"])

    # ── 查看进度统计 ──────────────────────────────────────────────────────────
    if args.stats:
        stats = tracker.get_stats()
        print("\n当前进度统计：")
        for status, count in stats.items():
            print(f"  {status:10s}: {count}")
        print(f"  {'合计':10s}: {sum(stats.values())}")
        return

    # ── 重置进度 ──────────────────────────────────────────────────────────────
    if args.reset:
        target = args.file if args.file else None
        tracker.reset(source_file=Path(target).name if target else None)
        reset_target = f"文件 [{Path(target).name}]" if target else "全部进度"
        logger.info(f"已重置 {reset_target} 的断点续传记录。")
        if not args.dry_run and not args.input and not args.file:
            return

    # ── 读取 Excel 数据 ───────────────────────────────────────────────────────
    excel_cfg = cfg.get("excel", {})
    reader = ExcelReader(
        column_name=excel_cfg.get("resolution_notes_column", "Resolution Notes"),
        skip_sheets=excel_cfg.get("skip_sheets", []),
        min_length=excel_cfg.get("min_content_length", 20),
        invalid_values=set(v.lower() for v in excel_cfg.get("invalid_values", [])) or None,
    )

    all_records = []
    if args.file:
        file_path = Path(args.file)
        logger.info(f"读取单个文件: {file_path}")
        try:
            records = reader.read(file_path)
            all_records.extend(records)
            logger.info(f"  读取到 {len(records)} 条有效记录")
        except Exception as e:
            logger.error(f"读取文件失败: {e}")
            sys.exit(1)
    else:
        input_dir = paths["input_dir"]
        logger.info(f"扫描输入目录: {input_dir}")
        try:
            all_file_records = reader.read_all(input_dir)
            for fname, records in all_file_records.items():
                logger.info(f"  [{fname}] {len(records)} 条记录")
                all_records.extend(records)
        except FileNotFoundError as e:
            logger.error(str(e))
            logger.info(f"提示：请将 .xlsx 文件放入目录：{input_dir}")
            sys.exit(1)

    logger.info(f"共读取到 {len(all_records)} 条有效记录")

    if not all_records:
        logger.warning("没有找到任何有效记录，退出。")
        return

    # ── 初始化 LLM 客户端 ─────────────────────────────────────────────────────
    llm_client = build_client_from_config(cfg.get("llm", {}))
    logger.info(f"LLM: {llm_client.config.provider} / {llm_client.config.model}")

    # ── 初始化 Prompt 管理器 ──────────────────────────────────────────────────
    prompt_cfg = cfg.get("prompt", {})
    style_ref_file = str(paths["style_ref_file"]) if paths["style_ref_file"].name else None
    prompt_manager = PromptManager(
        template_file=prompt_cfg.get("template_file"),
        style_ref_file=style_ref_file,
        style_ref_max_chars=int(prompt_cfg.get("style_ref_max_chars", 3000)),
    )

    # ── 预估成本 ──────────────────────────────────────────────────────────────
    wc_cfg = cfg.get("word_count", {})
    de_cfg = cfg.get("desensitization", {})
    proc_cfg = ProcessConfig(
        output_dir=paths["output_dir"],
        output_subdir_by_file=cfg.get("excel", {}).get("output_subdir_by_file", True),
        add_metadata_header=cfg.get("output", {}).get("add_metadata_header", True),
        encoding=cfg.get("output", {}).get("encoding", "utf-8"),
        filename_max_length=int(cfg.get("output", {}).get("filename_max_length", 60)),
        conflict_suffix=cfg.get("output", {}).get("conflict_suffix", True),
        target_word_count=int(wc_cfg.get("target", 2000)),
        min_word_count=int(wc_cfg.get("min", 1800)),
        max_word_count=int(wc_cfg.get("max", 2200)),
        retry_if_short=bool(wc_cfg.get("retry_if_short", True)),
        max_word_count_retries=int(wc_cfg.get("max_word_count_retries", 2)),
        desensitize_config=DesensitizeConfig(
            enabled=bool(de_cfg.get("enabled", True)),
            replace_chinese_names=bool(de_cfg.get("replace_chinese_names", True)),
            replace_phone=bool(de_cfg.get("replace_phone", True)),
            replace_id_card=bool(de_cfg.get("replace_id_card", True)),
            replace_room_number=bool(de_cfg.get("replace_room_number", False)),
        ),
    )

    processor = Processor(llm_client, prompt_manager, tracker, proc_cfg)

    cost_cfg = cfg.get("cost", {})
    if cost_cfg.get("show_estimate_before_run", True) or args.dry_run:
        estimate = processor.estimate_cost(all_records)
        usd_to_cny = float(cost_cfg.get("usd_to_cny_rate", 7.2))
        print("\n── 预估成本 ────────────────────────────────────────")
        print(f"  待处理记录数: {estimate['record_count']}")
        print(f"  预估输入 Token: {estimate['estimated_input_tokens']:,}")
        print(f"  预估输出 Token: {estimate['estimated_output_tokens']:,}")
        print(f"  预估总 Token:  {estimate['estimated_total_tokens']:,}")
        print(f"  预估费用:       ${estimate['estimated_cost_usd']:.4f}"
              f" (≈ ¥{estimate['estimated_cost_cny']:.2f})")
        print("────────────────────────────────────────────────────\n")

        if args.dry_run:
            logger.info("--dry-run 模式，不实际调用 LLM，退出。")
            return

        # 确认继续
        try:
            confirm = input("确认开始处理？[y/N] ").strip().lower()
            if confirm not in ("y", "yes"):
                logger.info("用户取消，退出。")
                return
        except (EOFError, KeyboardInterrupt):
            # 非交互环境（如批量脚本调用）直接继续
            pass

    # ── 开始批量处理 ──────────────────────────────────────────────────────────
    logger.info("开始批量处理...")
    processor.process_records(all_records)


if __name__ == "__main__":
    main()
