"""
image_restorer.py
─────────────────────────────────────────────────────────────
从 Excel (.xlsx) 文件中提取嵌入图片，并自动检测和修复被挤压变形的图片。

工作原理：
  xlsx 文件本质是 ZIP 压缩包，图片存放在内部的 xl/media/ 目录中。
  同时 xl/drawings/drawing*.xml 中记录了每张图片对应的单元格位置和尺寸（EMU单位）。
  通过对比图片原始宽高比（像素）与其在 Excel 中显示的宽高比（EMU），
  判断是否存在变形，若变形则用 Pillow 恢复到原始宽高比后保存。

用法（命令行）：
  # 处理单个文件
  python image_restorer.py input.xlsx

  # 处理整个文件夹（递归）
  python image_restorer.py ./excel_files/

  # 指定输出目录
  python image_restorer.py ./excel_files/ --output ./extracted_images/

  # 仅提取，不修复变形（原样保存）
  python image_restorer.py input.xlsx --no-fix

  # 设置变形检测阈值（默认5%）
  python image_restorer.py input.xlsx --threshold 0.08

依赖：
  pip install Pillow openpyxl lxml
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

try:
    from PIL import Image
except ImportError:
    print("[错误] 缺少依赖: pip install Pillow")
    sys.exit(1)

# ── 常量 ─────────────────────────────────────────────────────────────────────

# EMU (English Metric Units): 1 inch = 914400 EMU, 1 cm = 360000 EMU
EMU_PER_PIXEL = 9144  # 假设 96 DPI: 914400 / 96 = 9525, 用 9144 作为保守估算

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xlam"}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif", ".webp", ".emf", ".wmf"}

# XML 命名空间
NS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p":   "http://schemas.openxmlformats.org/drawingml/2006/picture",
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


# ── 数据类 ────────────────────────────────────────────────────────────────────

@dataclass
class ImageInfo:
    """单张图片的元数据。"""
    zip_path: str           # 在 ZIP 包内的路径，如 xl/media/image1.png
    filename: str           # 原始文件名
    display_cx: int = 0     # 在 Excel 中显示的宽度（EMU）
    display_cy: int = 0     # 在 Excel 中显示的高度（EMU）
    sheet_name: str = ""    # 所在工作表名
    anchor_info: str = ""   # 锚点信息（用于日志）


@dataclass
class RestoreResult:
    """单个 Excel 文件的处理结果。"""
    excel_file: str
    total_images: int = 0
    extracted: int = 0
    fixed: int = 0
    skipped: int = 0        # 不支持的格式（如 EMF/WMF）
    errors: list[str] = field(default_factory=list)


# ── 核心解析逻辑 ───────────────────────────────────────────────────────────────

def _parse_drawings(zf: zipfile.ZipFile, sheet_name: str, drawing_path: str) -> dict[str, ImageInfo]:
    """
    解析 drawing*.xml，提取每个图片的 rId → 显示尺寸（EMU）映射。
    返回 {rId: ImageInfo}
    """
    result: dict[str, ImageInfo] = {}
    try:
        with zf.open(drawing_path) as f:
            tree = ET.parse(f)
        root = tree.getroot()
    except Exception as e:
        logger.debug(f"    解析 {drawing_path} 失败: {e}")
        return result

    # 兼容 twoCellAnchor 和 oneCellAnchor 两种锚点类型
    for anchor_tag in ("xdr:twoCellAnchor", "xdr:oneCellAnchor", "xdr:absoluteAnchor"):
        for anchor in root.findall(anchor_tag, NS):
            # 找到 sp:pic 下的 blipFill → blip r:embed
            pic = anchor.find(".//p:pic", NS)
            if pic is None:
                # 尝试通用路径
                pic = anchor.find(".//{http://schemas.openxmlformats.org/drawingml/2006/picture}pic")
            if pic is None:
                continue

            blip = pic.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip")
            if blip is None:
                continue

            r_embed = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
            if not r_embed:
                continue

            # 尝试读取 ext (extent) 以获取显示尺寸
            cx, cy = 0, 0
            ext = anchor.find("xdr:ext", NS)
            if ext is not None:
                try:
                    cx = int(ext.get("cx", 0))
                    cy = int(ext.get("cy", 0))
                except (ValueError, TypeError):
                    pass

            # 如果是 twoCellAnchor，通过 from/to 坐标差估算（不依赖 ext）
            # 这里优先使用 ext，因为它更准确

            anchor_info = f"sheet={sheet_name}, cx={cx}, cy={cy}"

            result[r_embed] = ImageInfo(
                zip_path="",  # 稍后通过 rels 填充
                filename="",
                display_cx=cx,
                display_cy=cy,
                sheet_name=sheet_name,
                anchor_info=anchor_info,
            )

    return result


def _parse_rels(zf: zipfile.ZipFile, rels_path: str) -> dict[str, str]:
    """
    解析 .rels 文件，返回 {rId: target_path}。
    target_path 是相对于 xl/drawings/ 的路径，如 ../media/image1.png。
    """
    result: dict[str, str] = {}
    try:
        with zf.open(rels_path) as f:
            tree = ET.parse(f)
        for rel in tree.getroot():
            rid = rel.get("Id", "")
            target = rel.get("Target", "")
            rtype = rel.get("Type", "")
            if "image" in rtype.lower() and rid and target:
                result[rid] = target
    except Exception as e:
        logger.debug(f"    解析 rels 失败 ({rels_path}): {e}")
    return result


def _collect_images(zf: zipfile.ZipFile) -> list[ImageInfo]:
    """
    遍历整个 xlsx ZIP 包，收集所有图片信息（含显示尺寸）。
    """
    all_names = zf.namelist()
    images: list[ImageInfo] = []
    processed_media: set[str] = set()

    # 找所有 drawing*.xml 文件
    drawing_paths = [n for n in all_names if re.match(r"xl/drawings/drawing\d+\.xml$", n)]

    for drawing_path in drawing_paths:
        # 找对应的 .rels 文件
        drawing_dir = drawing_path.rsplit("/", 1)[0]
        drawing_file = drawing_path.rsplit("/", 1)[1]
        rels_path = f"{drawing_dir}/_rels/{drawing_file}.rels"

        # 尝试从 workbook.xml.rels 找到 sheet 名称（简化处理）
        sheet_name = drawing_path  # 默认用路径作为标识

        # 解析 drawing 得到 rId → 尺寸
        rid_to_info = _parse_drawings(zf, sheet_name, drawing_path)

        if not rid_to_info:
            continue

        # 解析 rels 得到 rId → 媒体路径
        rid_to_target: dict[str, str] = {}
        if rels_path in all_names:
            rid_to_target = _parse_rels(zf, rels_path)
        else:
            logger.debug(f"    未找到 rels 文件: {rels_path}")

        for rid, info in rid_to_info.items():
            if rid not in rid_to_target:
                continue
            target = rid_to_target[rid]
            # target 类似 "../media/image1.png"，转换为绝对路径
            media_path = _resolve_path(drawing_dir, target)
            if media_path not in all_names:
                # 尝试不区分大小写匹配
                media_path_lower = media_path.lower()
                matched = next((n for n in all_names if n.lower() == media_path_lower), None)
                if matched:
                    media_path = matched
                else:
                    logger.debug(f"    媒体文件不存在: {media_path}")
                    continue

            if media_path in processed_media:
                # 同一图片被多个 drawing 引用，更新尺寸信息（取最后一次）
                for existing in images:
                    if existing.zip_path == media_path:
                        if info.display_cx > 0:
                            existing.display_cx = info.display_cx
                            existing.display_cy = info.display_cy
                        break
                continue

            processed_media.add(media_path)
            info.zip_path = media_path
            info.filename = Path(media_path).name
            images.append(info)

    # 兜底：收集 xl/media/ 下所有未被 drawing 引用的图片
    for name in all_names:
        if name.startswith("xl/media/") and name not in processed_media:
            suffix = Path(name).suffix.lower()
            if suffix in IMAGE_EXTENSIONS:
                images.append(ImageInfo(
                    zip_path=name,
                    filename=Path(name).name,
                    display_cx=0,
                    display_cy=0,
                    sheet_name="unknown",
                    anchor_info="未关联drawing（直接从media提取）",
                ))

    return images


def _resolve_path(base_dir: str, relative: str) -> str:
    """将相对路径（含 ../）解析为 ZIP 内的绝对路径。"""
    parts = base_dir.split("/") + relative.split("/")
    resolved = []
    for part in parts:
        if part == "..":
            if resolved:
                resolved.pop()
        elif part and part != ".":
            resolved.append(part)
    return "/".join(resolved)


# ── 变形检测与修复 ─────────────────────────────────────────────────────────────

def _is_distorted(img_width: int, img_height: int, display_cx: int, display_cy: int,
                  threshold: float = 0.05) -> tuple[bool, float, float]:
    """
    判断图片是否被挤压变形。

    返回 (is_distorted, img_ratio, display_ratio)
    threshold: 宽高比差异容忍度（默认5%）
    """
    if display_cx <= 0 or display_cy <= 0 or img_width <= 0 or img_height <= 0:
        return False, 0.0, 0.0

    img_ratio = img_width / img_height
    display_ratio = display_cx / display_cy
    diff = abs(img_ratio - display_ratio) / max(img_ratio, display_ratio)
    return diff > threshold, img_ratio, display_ratio


def _restore_image(img: Image.Image, img_ratio: float, display_ratio: float) -> Image.Image:
    """
    恢复被挤压的图片到其原始宽高比。

    策略：以图片的**较大边**为基准，按原始宽高比计算另一边。
    这样不会损失任何内容，只是重新调整比例。
    """
    orig_w, orig_h = img.size

    if display_ratio > img_ratio:
        # 显示时被横向拉伸（宽 > 原始比例）→ 恢复宽度
        # 以高度为基准，按原始宽高比重新计算宽度
        new_h = orig_h
        new_w = round(orig_h * img_ratio)
    else:
        # 显示时被纵向拉伸（高 > 原始比例）→ 恢复高度
        new_w = orig_w
        new_h = round(orig_w / img_ratio)

    if new_w <= 0 or new_h <= 0:
        return img

    return img.resize((new_w, new_h), Image.LANCZOS)


# ── 主流程 ────────────────────────────────────────────────────────────────────

def process_excel_file(
    excel_path: Path,
    output_dir: Path,
    fix_distortion: bool = True,
    distortion_threshold: float = 0.05,
) -> RestoreResult:
    """
    处理单个 Excel 文件：提取图片，检测并修复变形。
    """
    result = RestoreResult(excel_file=str(excel_path))

    if not zipfile.is_zipfile(excel_path):
        result.errors.append(f"{excel_path} 不是有效的 ZIP/XLSX 文件")
        return result

    # 输出子目录：以 Excel 文件名命名
    file_output_dir = output_dir / excel_path.stem
    file_output_dir.mkdir(parents=True, exist_ok=True)

    try:
        with zipfile.ZipFile(excel_path, "r") as zf:
            images = _collect_images(zf)
            result.total_images = len(images)

            if not images:
                logger.info(f"  [{excel_path.name}] 未发现嵌入图片")
                return result

            logger.info(f"  [{excel_path.name}] 发现 {len(images)} 张图片")

            for img_info in images:
                _process_single_image(
                    zf=zf,
                    img_info=img_info,
                    output_dir=file_output_dir,
                    fix_distortion=fix_distortion,
                    distortion_threshold=distortion_threshold,
                    result=result,
                )

    except Exception as e:
        error_msg = f"处理 {excel_path.name} 时出错: {e}"
        result.errors.append(error_msg)
        logger.error(f"  {error_msg}")

    return result


def _process_single_image(
    zf: zipfile.ZipFile,
    img_info: ImageInfo,
    output_dir: Path,
    fix_distortion: bool,
    distortion_threshold: float,
    result: RestoreResult,
):
    """处理单张图片：提取、检测变形、修复并保存。"""
    suffix = Path(img_info.filename).suffix.lower()

    # EMF/WMF 是矢量格式，Pillow 不支持，直接原样提取
    if suffix in (".emf", ".wmf"):
        try:
            raw = zf.read(img_info.zip_path)
            out_path = _unique_path(output_dir, img_info.filename)
            out_path.write_bytes(raw)
            result.extracted += 1
            result.skipped += 1
            logger.info(f"    [原样] {img_info.filename} (矢量格式，跳过变形检测)")
        except Exception as e:
            result.errors.append(f"{img_info.filename}: {e}")
        return

    try:
        raw = zf.read(img_info.zip_path)
    except Exception as e:
        result.errors.append(f"读取 {img_info.filename} 失败: {e}")
        return

    # 用 Pillow 打开
    try:
        from io import BytesIO
        img = Image.open(BytesIO(raw))
        img.load()  # 确保完全加载
    except Exception as e:
        # Pillow 无法识别，原样保存
        out_path = _unique_path(output_dir, img_info.filename)
        out_path.write_bytes(raw)
        result.extracted += 1
        result.skipped += 1
        logger.warning(f"    [原样] {img_info.filename} (Pillow无法解析: {e})")
        return

    orig_w, orig_h = img.size
    was_fixed = False

    if fix_distortion and img_info.display_cx > 0 and img_info.display_cy > 0:
        distorted, img_ratio, display_ratio = _is_distorted(
            orig_w, orig_h,
            img_info.display_cx, img_info.display_cy,
            threshold=distortion_threshold,
        )
        if distorted:
            img = _restore_image(img, img_ratio, display_ratio)
            was_fixed = True
            logger.info(
                f"    [修复] {img_info.filename} "
                f"原始比例={img_ratio:.3f} 显示比例={display_ratio:.3f} "
                f"{orig_w}x{orig_h} → {img.size[0]}x{img.size[1]}"
            )
        else:
            logger.info(
                f"    [正常] {img_info.filename} "
                f"比例={img_ratio:.3f}（无变形）"
            )
    else:
        logger.info(f"    [提取] {img_info.filename} ({orig_w}x{orig_h})")

    # 保存
    out_filename = img_info.filename
    if was_fixed:
        stem = Path(img_info.filename).stem
        ext = Path(img_info.filename).suffix
        out_filename = f"{stem}_restored{ext}"

    out_path = _unique_path(output_dir, out_filename)

    try:
        # 保持原格式保存；JPEG 特殊处理 EXIF 和 RGB 模式
        save_fmt = img.format or _guess_format(suffix)
        if save_fmt in ("JPEG", "JPG"):
            save_fmt = "JPEG"
            if img.mode in ("RGBA", "P", "LA"):
                img = img.convert("RGB")
            img.save(out_path, format=save_fmt, quality=95)
        else:
            img.save(out_path, format=save_fmt)

        result.extracted += 1
        if was_fixed:
            result.fixed += 1

    except Exception as e:
        # 格式不兼容时降级保存为 PNG
        try:
            png_path = out_path.with_suffix(".png")
            png_path = _unique_path(output_dir, png_path.name)
            if img.mode not in ("RGB", "RGBA", "L", "P"):
                img = img.convert("RGBA")
            img.save(png_path, format="PNG")
            result.extracted += 1
            if was_fixed:
                result.fixed += 1
            logger.warning(f"    [降级] {img_info.filename} → PNG 格式 ({e})")
        except Exception as e2:
            result.errors.append(f"保存 {img_info.filename} 失败: {e2}")


def _unique_path(directory: Path, filename: str) -> Path:
    """如果文件已存在，自动追加序号后缀。"""
    target = directory / filename
    if not target.exists():
        return target
    stem = Path(filename).stem
    suffix = Path(filename).suffix
    counter = 2
    while (directory / f"{stem}_{counter}{suffix}").exists():
        counter += 1
    return directory / f"{stem}_{counter}{suffix}"


def _guess_format(suffix: str) -> str:
    """根据扩展名猜测 Pillow 保存格式。"""
    mapping = {
        ".jpg": "JPEG", ".jpeg": "JPEG",
        ".png": "PNG", ".gif": "GIF",
        ".bmp": "BMP", ".tiff": "TIFF", ".tif": "TIFF",
        ".webp": "WEBP",
    }
    return mapping.get(suffix.lower(), "PNG")


# ── 批量处理 ───────────────────────────────────────────────────────────────────

def process_folder(
    input_path: Path,
    output_dir: Path,
    fix_distortion: bool = True,
    distortion_threshold: float = 0.05,
    recursive: bool = True,
) -> list[RestoreResult]:
    """递归处理文件夹下的所有 Excel 文件。"""
    if input_path.is_file():
        excel_files = [input_path]
    else:
        pattern = "**/*.xlsx" if recursive else "*.xlsx"
        excel_files = list(input_path.glob(pattern))
        for ext in (".xlsm", ".xlam"):
            p = f"**/*{ext}" if recursive else f"*{ext}"
            excel_files.extend(input_path.glob(p))

    if not excel_files:
        logger.warning(f"在 {input_path} 下未找到 Excel 文件")
        return []

    logger.info(f"找到 {len(excel_files)} 个 Excel 文件，开始处理...")
    results = []

    for idx, excel_file in enumerate(excel_files, 1):
        logger.info(f"[{idx}/{len(excel_files)}] 处理: {excel_file.name}")
        result = process_excel_file(
            excel_path=excel_file,
            output_dir=output_dir,
            fix_distortion=fix_distortion,
            distortion_threshold=distortion_threshold,
        )
        results.append(result)

    return results


def print_summary(results: list[RestoreResult]):
    """打印汇总报告。"""
    total_images = sum(r.total_images for r in results)
    total_extracted = sum(r.extracted for r in results)
    total_fixed = sum(r.fixed for r in results)
    total_skipped = sum(r.skipped for r in results)
    all_errors = [e for r in results for e in r.errors]

    print("\n" + "=" * 60)
    print("  图片提取与修复完成")
    print("=" * 60)
    print(f"  处理文件数 : {len(results)}")
    print(f"  发现图片总数: {total_images}")
    print(f"  成功提取   : {total_extracted}")
    print(f"  修复变形   : {total_fixed}")
    print(f"  跳过(矢量) : {total_skipped}")

    if all_errors:
        print(f"\n  错误 ({len(all_errors)} 条):")
        for err in all_errors[:10]:
            print(f"    - {err}")
        if len(all_errors) > 10:
            print(f"    ... 还有 {len(all_errors) - 10} 条错误（查看日志）")
    print("=" * 60)


# ── CLI 入口 ──────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="从 Excel 文件中提取图片并修复被挤压变形的图像",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python image_restorer.py input.xlsx
  python image_restorer.py ./excel_files/ --output ./images/
  python image_restorer.py input.xlsx --no-fix
  python image_restorer.py input.xlsx --threshold 0.08
        """,
    )
    parser.add_argument("input", help="Excel 文件路径或包含 Excel 文件的文件夹路径")
    parser.add_argument(
        "--output", "-o",
        default="./extracted_images",
        help="图片输出目录（默认: ./extracted_images）",
    )
    parser.add_argument(
        "--no-fix",
        action="store_true",
        help="只提取图片，不修复变形",
    )
    parser.add_argument(
        "--threshold", "-t",
        type=float,
        default=0.05,
        help="变形检测阈值，宽高比差异超过此值视为变形（默认: 0.05，即5%%）",
    )
    parser.add_argument(
        "--no-recursive",
        action="store_true",
        help="不递归处理子文件夹（仅处理指定目录的直接子文件）",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="显示详细调试日志",
    )

    args = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    input_path = Path(args.input).resolve()
    output_dir = Path(args.output).resolve()

    if not input_path.exists():
        print(f"[错误] 路径不存在: {input_path}")
        sys.exit(1)

    output_dir.mkdir(parents=True, exist_ok=True)
    logger.info(f"输出目录: {output_dir}")
    logger.info(f"变形修复: {'关闭' if args.no_fix else f'开启（阈值={args.threshold:.0%}）'}")

    results = process_folder(
        input_path=input_path,
        output_dir=output_dir,
        fix_distortion=not args.no_fix,
        distortion_threshold=args.threshold,
        recursive=not args.no_recursive,
    )

    print_summary(results)


if __name__ == "__main__":
    main()
