"""
prompt_manager.py
─────────────────
Prompt 模板管理模块：
- 加载外部 system_prompt.md 模板文件
- 支持注入范文（风格参考）
- 构建 user_message（将事件记录 + 上下文字段组合）
- 字数验证与补充扩写指令
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

from .excel_reader import EventRecord


# ── 默认 System Prompt（内置，当外部文件不可用时使用）────────────────────────

DEFAULT_SYSTEM_PROMPT = """# 角色设定
你是一位拥有20年从业经验的五星级酒店服务培训专家，同时也是一位擅长将真实事件转化为深度教学案例的专业作者。你对高端酒店的服务标准、客诉处理流程、跨部门协作机制以及会员关怀体系有极为深厚的理解。

# 任务说明
你将收到一段来自酒店Logbook的简短事件记录（通常几十到两百字）。你的任务是将其改写为一篇**专业、深度、叙事生动**的酒店服务案例文章，供培训使用。

## 改写原则
1. **事实忠实**：核心事件（投诉原因、处理方式、最终结果）必须与原始记录保持一致，严禁编造不存在的结果或颠倒事实。
2. **合理扩写**：在事实框架内，可以合理补充：
   - 客人的神态、情绪、语气和心理活动
   - 现场环境描写（时间、氛围、季节感）
   - 员工的操作细节和内心决策过程
   - 符合逻辑的对话片段（用引号标注，标注为"场景还原"）
   - 深层原因分析（硬件、制度、培训、文化等维度）
3. **严禁篡改**：不得改变投诉的根本原因、不得美化或丑化任何一方、不得捏造原文未提及的赔偿金额或法律结果。

# 输出结构（必须严格遵守，7个章节缺一不可）

你的输出必须包含且仅包含以下7个章节，按顺序排列，不得遗漏任何一个：

# [案例标题]
（标题单独一行，以 # 开头）

## 案例背景
（时间、地点、人物背景、入住渠道、客人期待——约200-300字）

## 事件经过
（事件的完整发展脉络，包含关键转折点、冲突升级过程、涉及部门的联动情况——约500-600字）

## 处理难点
（分析当时处置的真正困境：为什么不好处理？涉及哪些两难选择？——约200-300字）

## 解决方案
（具体行动步骤、使用的话术策略、资源调配过程、跨部门协调细节——约300-400字）

## 结果与反馈
（客人的最终反应、事件对品牌/团队的影响、后续跟进情况——约150-200字）

## 案例启示
（从管理、员工、制度三个维度提出3-5条可操作的改进建议，附具体SOP优化方向——约300-400字）

## 引导问题
⚠️ 此章节为必填项，不得省略！
请针对本案例的核心矛盾，提出恰好2个开放式问题，聚焦一线操作层面的具体困境，第二个问题聚焦制度设计、沟通策略或客户关系管理层面，供培训学员讨论。
问题要求：聚焦实操困境或决策两难，无标准答案，能引发辩论。


# 案例标题要求
- 标题要有文学感和吸引力，不得直白照抄原始记录内容
- 好标题示例："三房同层的执念：一场本可避免的前台风波"、"当维修工单遇上摔跤的孩子"
- 避免标题示例："客人不满意房间安排事件"（过于平淡）

# 字数要求
**全文总字数必须在1800-2200汉字之间**（不含标题和章节标题，含引导问题正文）。这是硬性要求，请在生成前自行规划每个章节的篇幅分配，确保达到字数要求。

# 语言风格
- 文笔专业而不刻板，叙述流畅，逻辑清晰
- 事件经过部分可适当使用叙事性文学笔法，增强临场感
- 案例启示部分需保持条理性，语言精炼、可操作性强
- 整体语调客观中立，既不过度美化服务方，也不夸大客人的无理诉求

{style_reference_section}"""

STYLE_REFERENCE_TEMPLATE = """
# 风格参考范文
以下是一篇高质量的参考案例（请学习其叙事节奏、段落结构和语言风格，但不要直接复制其内容）：

---
{style_text}
---
"""

RETRY_PROMPT_APPENDIX = """

⚠️ 注意：你的上一次输出字数不足（约{current_count}字），请在保持所有已有内容的基础上，大幅扩充以下方面：
- 在"事件经过"中补充更多细节：现场对话、员工心理活动、时间线描述
- 在"案例启示"中增加更具体的SOP建议和制度改进方案
- 在"案例背景"中增加对酒店环境、节假日氛围的描写
- 目标：确保全文达到2000字左右。"""


class PromptManager:
    """
    负责构建每次 LLM 调用的 System Prompt 和 User Message。
    """

    def __init__(
        self,
        template_file: Optional[str] = None,
        style_ref_file: Optional[str] = None,
        style_ref_max_chars: int = 3000,
    ):
        self.style_ref_max_chars = style_ref_max_chars
        self._system_template = self._load_template(template_file)
        self._style_text = self._load_style_ref(style_ref_file)

    # ── 公开方法 ──────────────────────────────────────────────────────────────

    def build_system_prompt(self) -> str:
        """构建完整的 System Prompt（含风格参考）。"""
        if self._style_text:
            style_section = STYLE_REFERENCE_TEMPLATE.format(style_text=self._style_text)
        else:
            style_section = ""

        return self._system_template.replace("{style_reference_section}", style_section).strip()

    def build_user_message(self, record: EventRecord) -> str:
        """
        将 EventRecord 转化为 User Message，提供足够的上下文。
        使用脱敏后的内容（若已填充）。
        """
        content = record.desensitized_content or record.content

        # 整理上下文字段
        context_parts = []
        field_labels = {
            "Description": "事件简述",
            "Case Contact Name": "客人信息",
            "Member": "会员级别",
            "Location": "房间/区域",
        }
        for key, label in field_labels.items():
            if key in record.extra_fields:
                val = record.extra_fields[key]
                # Description 本身是中英双语，截取中文部分
                if key == "Description":
                    # 只保留第一个换行前的内容（通常是中文部分）
                    val = val.split("\n")[0].strip()
                context_parts.append(f"- **{label}**：{val}")

        context_str = "\n".join(context_parts) if context_parts else "（无额外信息）"

        return f"""## 原始事件记录

**来源文件**：{record.source_file}（第{record.row_index}行）

**背景信息**：
{context_str}

**事件详细记录（Resolution Notes）**：
{content}

---

请根据以上记录，严格按照 System Prompt 中规定的8个章节结构撰写完整的酒店服务案例。
⚠️ 特别提醒："## 引导问题"是必填章节，必须包含在输出中，提出2个针对本案例的开放式讨论问题，绝对不可遗漏！"""

    def build_retry_message(self, original_content: str, current_count: int) -> str:
        """字数不足时的追加扩写指令（追加到原内容之后让模型继续扩写）。"""
        appendix = RETRY_PROMPT_APPENDIX.format(current_count=current_count)
        return f"{original_content}\n{appendix}"

    def build_missing_guide_questions_message(self, original_content: str) -> str:
        """引导问题缺失时，要求模型补写引导问题的指令。"""
        return f"{original_content}\n{MISSING_GUIDE_QUESTIONS_PROMPT}"

    def update_style_ref(self, style_ref_file: str):
        """动态更新范文路径（Streamlit UI 使用）。"""
        self._style_text = self._load_style_ref(style_ref_file)

    # ── 内部方法 ──────────────────────────────────────────────────────────────

    def _load_template(self, template_file: Optional[str]) -> str:
        """加载外部 Prompt 模板文件，失败则使用内置模板。"""
        if template_file:
            path = Path(template_file)
            if path.exists():
                try:
                    return path.read_text(encoding="utf-8")
                except Exception:
                    pass
        return DEFAULT_SYSTEM_PROMPT

    def _load_style_ref(self, style_ref_file: Optional[str]) -> str:
        """加载范文文件，截取前 N 字符以控制 Token 消耗。"""
        if not style_ref_file:
            return ""
        path = Path(style_ref_file)
        if not path.exists():
            return ""
        try:
            text = path.read_text(encoding="utf-8")
            if len(text) > self.style_ref_max_chars:
                text = text[: self.style_ref_max_chars] + "\n\n[...节选至此...]"
            return text
        except Exception:
            return ""


# ── 工具函数 ─────────────────────────────────────────────────────────────────

def count_chinese_chars(text: str) -> int:
    """统计中文字符数量（作为'汉字数'的近似值）。"""
    return sum(1 for c in text if "\u4e00" <= c <= "\u9fff")


def count_all_chars(text: str) -> int:
    """统计全文字符数（含英文，排除标点和空白）。"""
    return sum(1 for c in text if not c.isspace())


def has_guide_questions(text: str) -> bool:
    """检测文章是否包含引导问题章节。"""
    # 匹配 "## 引导问题" 或 "## 7. 引导问题" 等变体
    return bool(re.search(r"^##\s*([\d\.]*\s*)?引导问题", text, re.MULTILINE))


def append_guide_questions_to_content(base_content: str, supplement: str) -> str:
    """
    将模型补写的引导问题内容合并到原文末尾。
    如果补写内容里已含有 ## 引导问题 标题则直接追加，否则加上标题再追加。
    """
    # 提取补写内容中的引导问题部分（从 ## 引导问题 开始到末尾）
    match = re.search(r"(^##\s*[\d\.]*\s*引导问题.*)", supplement, re.MULTILINE | re.DOTALL)
    if match:
        return base_content.rstrip() + "\n\n" + match.group(1).strip()
    # 补写内容没有标题，直接追加
    return base_content.rstrip() + "\n\n## 引导问题\n\n" + supplement.strip()


def extract_title(content: str) -> Optional[str]:
    """从生成内容中提取第一个 # 标题。"""
    match = re.search(r"^#\s+(.+)$", content, re.MULTILINE)
    if match:
        return match.group(1).strip()
    return None


def sanitize_filename(title: str, max_length: int = 60) -> str:
    """
    将标题转换为合法的文件名：
    - 去除非法字符: / \\ : * ? " < > |
    - 截断过长文件名
    - 替换连续空白为单个空格
    """
    illegal = r'[/\\:*?"<>|]'
    name = re.sub(illegal, "", title)
    name = re.sub(r"\s+", " ", name).strip()
    # 按字符长度截断（中文字符宽度约为2个英文字符，但文件系统通常按字节计）
    if len(name) > max_length:
        name = name[:max_length].rstrip()
    return name or "未命名案例"
