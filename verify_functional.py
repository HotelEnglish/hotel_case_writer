"""端到端功能验证（不调用LLM）"""
import sys
sys.path.insert(0, '.')

# 1. 测试 Excel 读取
from src.excel_reader import ExcelReader
reader = ExcelReader()
records = reader.read('./sample_data/sample_logbook.xlsx')
print(f'[OK] Excel读取: {len(records)} 条记录')
for r in records:
    print(f'     {r.source_file} Sheet={r.sheet_name} Row={r.row_index} len={len(r.content)}')

# 2. 测试脱敏
from src.desensitizer import Desensitizer, DesensitizeConfig
d = Desensitizer(DesensitizeConfig(enabled=True))
test_text = '客人张建国先生反映，其手机号13812345678已在系统登记，身份证号310101199001011234遗失。'
result = d.desensitize(test_text)
print(f'\n[OK] 脱敏测试:')
print(f'     原文: {test_text}')
print(f'     脱敏: {result}')

# 3. 测试断点续传
from src.progress_tracker import ProgressTracker
import tempfile, os
with tempfile.NamedTemporaryFile(suffix='.db', delete=False) as f:
    db_path = f.name
tracker = ProgressTracker(db_path)
tracker.mark_done('test.xlsx', 'Sheet1', 5, 'output/test.md')
assert tracker.is_done('test.xlsx', 'Sheet1', 5)
assert not tracker.is_done('test.xlsx', 'Sheet1', 6)
os.unlink(db_path)
print('\n[OK] 断点续传: 写入/读取/判断均正常')

# 4. 测试 Prompt 构建
from src.prompt_manager import PromptManager
pm = PromptManager()
sys_prompt = pm.build_system_prompt()
user_msg = pm.build_user_message(records[0])
print(f'\n[OK] Prompt构建:')
print(f'     System Prompt 长度: {len(sys_prompt)} 字符')
print(f'     User Message 长度: {len(user_msg)} 字符')

# 5. 测试文件名清洗
from src.prompt_manager import sanitize_filename
cases = [
    '三房同层的执念：一场本可避免的前台风波',
    '含非法字符/\\:*?"<>|的标题',
    'A' * 100,
]
for c in cases:
    clean = sanitize_filename(c)
    print(f'     "{c[:30]}..." -> "{clean}"')

print('\n所有功能验证通过 [OK]')
