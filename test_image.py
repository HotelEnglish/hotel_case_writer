import sys
sys.path.insert(0, 'd:/miro/dev/anliku/hotel_case_writer')
from image_restorer import process_folder
from pathlib import Path

results = process_folder(
    input_path=Path('d:/miro/dev/anliku/JW GSM Log 2025.01.01.xlsx'),
    output_dir=Path('d:/miro/dev/anliku/hotel_case_writer/output/images_test'),
)

for r in results:
    print(f"文件: {r.excel_file}")
    print(f"  图片总数: {r.total_images}")
    print(f"  提取成功: {r.extracted}")
    print(f"  修复变形: {r.fixed}")
    print(f"  跳过(矢量): {r.skipped}")
    if r.errors:
        print(f"  错误: {r.errors}")
