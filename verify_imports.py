"""验证所有模块可以正常导入"""
import sys
sys.path.insert(0, '.')

errors = []

modules = [
    ('src.config_loader', 'load_config'),
    ('src.logger', 'setup_logging'),
    ('src.excel_reader', 'ExcelReader'),
    ('src.desensitizer', 'Desensitizer'),
    ('src.llm_client', 'LLMClient'),
    ('src.prompt_manager', 'PromptManager'),
    ('src.progress_tracker', 'ProgressTracker'),
    ('src.file_writer', 'save_case'),
    ('src.processor', 'Processor'),
]

for mod_name, cls_name in modules:
    try:
        mod = __import__(mod_name, fromlist=[cls_name])
        obj = getattr(mod, cls_name)
        print(f'  OK  {mod_name}.{cls_name}')
    except Exception as e:
        print(f'  FAIL {mod_name}: {e}')
        errors.append((mod_name, str(e)))

print()
if errors:
    print(f'有 {len(errors)} 个模块导入失败')
else:
    print('所有模块导入成功 [OK]')
