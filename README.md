# 🏨 酒店案例批量改写工具

将酒店 Logbook（`.xlsx`）中的 `Resolution Notes` 事件记录，通过大语言模型自动扩写为 **2000 字左右的专业培训案例**（`.md` 格式）。

---

## 功能特性

| 功能 | 说明 |
|------|------|
| 📂 批量处理 | 遍历整个目录下的所有 `.xlsx` 文件 |
| 🤖 多模型支持 | Ollama 本地 / OpenAI / DeepSeek / 智谱 / Azure |
| 🔄 断点续传 | 中断后重启自动跳过已处理记录 |
| 🔒 数据脱敏 | 自动替换客人姓名、手机号、身份证号 |
| 📝 风格参考 | 可注入范文，让 AI 模仿特定写作风格 |
| 💰 成本估算 | 运行前显示预估 Token 消耗和费用 |
| 🛡️ 自动重试 | 遇到限速(429)或字数不足时自动重试 |
| 💬 引导问题 | 每篇案例末尾自动生成2个开放式讨论问题 |
| 🖼️ 图片提取 | 从 Excel 提取嵌入图片，自动修复被挤压变形的图像 |
| 📊 图形界面 | Streamlit 网页界面，无需命令行操作 |
| 📋 错误日志 | 所有失败记录写入 `error_log.txt` |

---

## 快速开始

### 第一步：安装依赖

> 推荐使用 Python 3.10+，建议在虚拟环境中安装。

```bash
# 创建虚拟环境（可选但推荐）
python -m venv venv
venv\Scripts\activate        # Windows
source venv/bin/activate     # macOS/Linux

# 安装依赖
pip install -r requirements.txt
```

### 第二步：配置 API

复制配置模板并填写你的 API 信息：

```bash
cp .env.example .env
```

用文本编辑器打开 `.env`，根据你使用的服务商填写：

**方案 A：本地 Ollama（免费，需要本地 GPU）**
```env
LLM_PROVIDER=ollama
OLLAMA_BASE_URL=http://localhost:11434
OLLAMA_MODEL=qwen3:4b
```
确保 Ollama 已启动并拉取模型：
```bash
ollama run qwen3:4b
```

**方案 B：DeepSeek API（性价比高）**
```env
LLM_PROVIDER=deepseek
OPENAI_API_KEY=sk-你的key
OPENAI_BASE_URL=https://api.deepseek.com/v1
OPENAI_MODEL=deepseek-chat
```

**方案 C：OpenAI**
```env
LLM_PROVIDER=openai
OPENAI_API_KEY=sk-你的key
OPENAI_MODEL=gpt-4o-mini
```

### 第三步：准备数据

将你的 `.xlsx` 文件放入 `input/` 目录：
```
hotel_case_writer/
├── input/
│   ├── GSM_Log_2025_01.xlsx
│   ├── GSM_Log_2025_02.xlsx
│   └── ...
```

**Excel 文件格式要求：**
- 需包含 `Resolution Notes` 列（工具会自动在前20行内查找表头）
- 支持多个 Sheet，空 Sheet 自动跳过
- 其他列（如 `Description`、`Member`、`Location`）会作为背景信息注入 Prompt

### 第四步：运行

**图形界面（推荐，无需命令行基础）：**
```bash
python main.py --ui
```
浏览器会自动打开 `http://localhost:8501`

**命令行方式：**
```bash
python main.py                              # 处理 input/ 目录所有文件
python main.py --dry-run                   # 仅预估成本
python main.py --file ./input/log.xlsx     # 处理单个文件
python main.py --style-ref ./范文.md       # 使用范文作为风格参考
python main.py --reset                     # 重置所有断点续传记录
python main.py --stats                     # 查看进度统计
```

---

## 项目结构

```
hotel_case_writer/
├── main.py                  # 命令行主入口
├── app_ui.py                # Streamlit 图形界面
├── image_restorer.py        # 图片提取与变形修复工具（独立脚本）
├── config.yaml              # 主配置文件
├── .env.example             # 环境变量示例（复制为 .env 并填写）
├── requirements.txt         # Python 依赖列表
├── generate_sample_data.py  # 生成测试数据
│
├── src/                     # 核心模块
│   ├── __init__.py
│   ├── config_loader.py     # 配置加载（yaml + .env）
│   ├── excel_reader.py      # Excel 读取与解析
│   ├── desensitizer.py      # 数据脱敏
│   ├── llm_client.py        # LLM 调用层（统一接口）
│   ├── prompt_manager.py    # Prompt 模板管理
│   ├── processor.py         # 核心处理流程
│   ├── progress_tracker.py  # 断点续传（SQLite）
│   ├── file_writer.py       # 文件保存与命名
│   └── logger.py            # 日志系统
│
├── prompts/
│   └── system_prompt.md     # System Prompt 模板（可自定义）
│
├── input/                   # 放置待处理 xlsx 文件（需手动创建或在配置中指定）
├── output/                  # 生成的 md 文件保存位置
├── sample_data/
│   └── sample_logbook.xlsx  # 5 条测试数据
└── logs/
    ├── run.log              # 运行日志
    ├── error_log.txt        # 错误日志（人类可读格式）
    └── progress.db          # 断点续传数据库（SQLite）
```

---

## 配置详解

所有配置均在 `config.yaml` 中，常用项说明：

```yaml
excel:
  resolution_notes_column: "Resolution Notes"  # 目标列名（支持模糊匹配）
  min_content_length: 20                        # 低于此字数视为无效，跳过

word_count:
  min: 1800    # 生成字数低于此值时触发重试
  max: 2200    # 超出此值仍然保留（不截断）

desensitization:
  enabled: true              # 全局开关
  replace_room_number: false # 房间号默认保留（设为 true 则脱敏）

llm:
  requests_per_minute: 20   # 降低此值可避免 429 错误
  concurrent_workers: 1     # 并发数（云端 API 建议 1-3）

paths:
  style_ref_file: "./某范文.md"   # 范文路径（留空则不使用）
```

---

## 自定义 Prompt

编辑 `prompts/system_prompt.md` 即可修改 AI 的写作指令。

文件中 `{style_reference_section}` 是占位符，当指定范文时会自动替换为范文内容，无需手动修改。

---

## 输出结构

生成的案例文件按来源 Excel 文件名分子文件夹存放：

```
output/
├── GSM_Log_2025_01/
│   ├── 三房同层的执念：一场本可避免的前台风波.md
│   ├── 走出酒店的那一刻：大东海遗失案的警示录.md
│   └── ...
└── GSM_Log_2025_02/
    └── ...
```

每个 `.md` 文件开头包含元数据注释（HTML 注释格式，不影响 Markdown 渲染）：
```html
<!--
  自动生成时间: 2025-01-08 14:32:11
  来源文件: GSM_Log_2025_01.xlsx
  工作表: Sheet1
  原始行号: 5
-->
```

---

## 常见问题

**Q：提示 `Resolution Notes 列未找到` 怎么办？**
> A：检查 Excel 列名是否完全匹配，或在 `config.yaml` 中修改 `resolution_notes_column` 的值。工具支持模糊匹配（大小写不敏感）。

**Q：Ollama 连接失败？**
> A：确认 Ollama 服务已启动（`ollama serve`），并且模型已下载（`ollama pull qwen3:4b`）。

**Q：生成速度太慢？**
> A：本地 Ollama 速度取决于硬件。使用云端 API（DeepSeek/OpenAI）速度更快。可适当提高 `concurrent_workers` 开启并发处理。

**Q：如何重新处理已完成的文件？**
> A：运行 `python main.py --reset` 清空所有进度，或 `python main.py --reset --file xxx.xlsx` 只重置指定文件。

**Q：生成内容质量不满意怎么办？**
> A：提供一篇高质量的范文（`--style-ref` 参数），或修改 `prompts/system_prompt.md` 中的写作要求。

---

## 测试运行

使用内置测试数据验证工具是否正常运行：

```bash
# 1. 生成测试数据
python generate_sample_data.py

# 2. 复制到 input 目录
cp sample_data/sample_logbook.xlsx input/

# 3. 先预估成本（无需真实 API Key）
python main.py --dry-run --input ./input --output ./output

# 4. 实际运行（需要配置好 .env）
python main.py --input ./input --output ./output
```

---

## 图片提取与变形修复

部分 Excel 文件中嵌有图片，有时因单元格尺寸拉伸导致图片显示变形。`image_restorer.py` 可以：

1. **提取**：将 Excel 内嵌图片按文件名分子文件夹保存为独立图片文件
2. **检测变形**：对比图片原始宽高比与 Excel 中显示的宽高比
3. **修复变形**：将被挤压的图片恢复为原始宽高比（文件名加 `_restored` 后缀）

**命令行用法：**

```bash
# 提取并修复单个文件中的图片
python image_restorer.py "GSM Log 2025.01.xlsx"

# 批量处理整个文件夹
python image_restorer.py ./input/ --output ./extracted_images/

# 只提取，不修复变形
python image_restorer.py ./input/ --no-fix

# 调整变形检测灵敏度（默认5%，值越小越灵敏）
python image_restorer.py ./input/ --threshold 0.08
```

**图形界面：** 在 Streamlit UI 中切换到「🖼️ 图片提取」标签页操作。

**输出结构：**
```
extracted_images/
├── GSM Log 2025.01/
│   ├── image1.png           # 正常图片原样保存
│   ├── image2_restored.png  # 变形图片修复后保存
│   └── image3.jpg
└── GSM Log 2025.02/
    └── image1.png
```

> **注意：** EMF/WMF 矢量格式图片会原样提取，不作修复处理（Pillow 不支持矢量格式）。

---

## 依赖说明

| 包名 | 用途 |
|------|------|
| `pandas` + `openpyxl` | Excel 读取 |
| `openai` | LLM API 调用（兼容 Ollama/DeepSeek 等） |
| `python-dotenv` | 加载 `.env` 配置 |
| `pyyaml` | 加载 `config.yaml` |
| `tiktoken` | Token 数量估算 |
| `tqdm` + `rich` | 进度显示美化 |
| `streamlit` | 图形界面 |
| `Pillow` | 图片处理与变形修复 |
| `xlsxwriter` | 生成测试数据 |
