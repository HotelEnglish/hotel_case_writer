"""
desensitizer.py
───────────────
对事件内容进行脱敏处理：
- 中文姓名替换（张先生/李女士等）
- 手机号脱敏
- 身份证号脱敏
- 银行卡号脱敏
- 可配置：是否保留房间号
支持关闭脱敏（直接返回原文）。
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Optional


@dataclass
class DesensitizeConfig:
    enabled: bool = True
    replace_chinese_names: bool = True
    replace_phone: bool = True
    replace_id_card: bool = True
    replace_bank_card: bool = True
    replace_room_number: bool = False   # 默认保留房间号（有助于案例真实感）
    custom_patterns: list[tuple[str, str]] = None   # [(pattern, replacement), ...]


class Desensitizer:
    """
    正则 + 规则驱动的轻量级脱敏器。
    不依赖 NLP 模型，在无法安装 presidio 的环境下同样可用。
    """

    # 常见中文姓氏（用于姓名识别）
    _SURNAMES = (
        "赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜"
        "戚谢邹喻柏水窦章云苏潘葛奚范彭郎鲁韦昌马苗凤花方俞任袁柳酆鲍史唐"
        "费廉岑薛雷贺倪汤滕殷罗毕郝邬安常乐于时傅皮卞齐康伍余元卜顾孟平黄"
        "和穆萧尹姚邵湛汪祁毛禹狄米贝明臧计伏成戴谈宋茅庞熊纪舒屈项祝董梁"
        "杜阮蓝闵席季麻强贾路娄危江童颜郭梅盛林刁钟徐邱骆高夏蔡田樊胡凌霍"
        "虞万支柯咸管卢莫经房裘缪干解应宗丁宣贲邓郁单杭洪包诸左石崔吉钮龚"
        "程嵇邢滑裴陆荣翁荀羊於惠甄麹家封芮羿储靳汲邴糜松井段富巫乌焦巴弓"
        "牧隗山谷车侯宓蓬全郗班仰秋仲伊宫宁仇栾暴甘钭厉戎祖武符刘景詹束龙"
        "叶幸司韶郜黎蓟薄印宿白怀蒲邰从鄂索咸籍赖卓蔺屠蒙池乔阴鬱胥能苍双"
        "闻莘党翟谭贡劳逄姬申扶堵冉宰郦雍郤璩桑桂濮牛寿通边扈燕冀郏浦尚农"
        "温别庄晏柴瞿阎充慕连茹习宦艾鱼容向古易慎戈廖庾终暨居衡步都耿满弘匡"
        "国文寇广禄阙东殴殳沃利蔚越夔隆师巩厍聂晁勾敖融冷訾辛阚那简饶空曾"
        "毋沙乜养鞠须丰巢关蒯相查后荆红游竺权逯盖益桓公仉督晋楚阳"
    )

    # ── 正则表达式 ────────────────────────────────────────────────────────────

    # 中文手机号（支持 +86 前缀）
    _RE_PHONE = re.compile(
        r"(?<!\d)(?:\+?86[-\s]?)?1[3-9]\d{9}(?!\d)"
    )

    # 身份证号（18位，末位可为X/x）
    _RE_ID_CARD = re.compile(
        r"\b\d{6}(?:19|20)\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])\d{3}[\dXx]\b"
    )

    # 银行卡号（13-19位纯数字，避免误伤一般数字）
    _RE_BANK_CARD = re.compile(
        r"(?<!\d)(?:6\d{15,18}|4\d{12}(?:\d{3})?|5[1-5]\d{14})(?!\d)"
    )

    # 房间号：3-4位数字房间号（如 788、872、1201）
    _RE_ROOM = re.compile(r"\b(?:[0-9]{1,2}楼|[0-9]{3,4}(?:号|房|室)?)\b")

    def __init__(self, config: Optional[DesensitizeConfig] = None):
        self.config = config or DesensitizeConfig()
        self._surname_set = set(self._SURNAMES)

        surname_alts = "|".join(re.escape(s) for s in sorted(self._surname_set, key=len, reverse=True))

        # 含称谓姓名：姓+名+先生/女士等（最高精度）
        self._re_name_with_title = re.compile(
            rf"(?:{surname_alts})"
            r"[^\s，。！？,.]{0,3}"
            r"(?:先生|女士|小姐|总|经理|主任|老师|工程师|医生|店长)"
        )
        # 裸姓名（无称谓），名字后接标点/空格/行尾/括号
        self._re_name_bare = re.compile(
            rf"(?:{surname_alts})"
            r"[\u4e00-\u9fff]{1,2}"
            r"(?=[，。！？、；\s\(（【「」]|$)"
        )
        # 前缀触发全名：「客人/宾客/会员 等2字前缀 + 全名」
        # 捕获组1=前缀，捕获组2=姓，捕获组3=名（1-2字，非贪婪）
        _name_prefixes = r"(?:客人|宾客|会员|住客|旅客|来宾|用户|顾客|住户)"
        self._re_name_with_prefix = re.compile(
            rf"({_name_prefixes})"
            rf"({surname_alts})"
            r"([\u4e00-\u9fff]{1,2})"           # 名字1-2字
        )
        # 全名精确匹配（用于高风险字段，名字后不接汉字）
        self._re_name_full = re.compile(
            rf"(?:{surname_alts})"
            r"([\u4e00-\u9fff]{1,3})"
            r"(?![\u4e00-\u9fff])"
        )

    # ── 公开方法 ──────────────────────────────────────────────────────────────

    def desensitize(self, text: str) -> str:
        """对文本执行脱敏，返回脱敏后的文本。"""
        if not self.config.enabled or not text:
            return text

        # 1. 手机号
        if self.config.replace_phone:
            text = self._RE_PHONE.sub(self._phone_replacement, text)

        # 2. 身份证
        if self.config.replace_id_card:
            text = self._RE_ID_CARD.sub("[身份证号码]", text)

        # 3. 银行卡
        if self.config.replace_bank_card:
            text = self._RE_BANK_CARD.sub("[银行卡号]", text)

        # 4. 房间号（可选）
        if self.config.replace_room_number:
            text = self._RE_ROOM.sub("[房间号]", text)

        # 5. 中文姓名（含称谓，精度更高）
        if self.config.replace_chinese_names:
            text = self._re_name_with_title.sub(self._name_replacement, text)
            # 前缀触发全名（"客人/会员/宾客/住客 + 全名"等高置信度场景）
            text = self._re_name_with_prefix.sub(self._name_replacement_with_prefix, text)
            # 裸姓名（无称谓，后接标点/空格/行尾）
            text = self._re_name_bare.sub(self._name_replacement_bare, text)
            # 注：正文中不使用 _re_name_full，该正则仅用于高风险姓名字段

        # 6. 自定义规则
        if self.config.custom_patterns:
            for pattern, replacement in self.config.custom_patterns:
                text = re.sub(pattern, replacement, text)

        return text

    def desensitize_name_field(self, text: str) -> str:
        """
        专门用于高风险姓名字段（如 Case Contact Name）的脱敏。
        姓名字段通常只包含姓名本身（不是完整句子），精度最高。
        策略：先用完整正文流程，再用全名正则兜底。
        """
        if not self.config.enabled or not text:
            return text
        # 第一步：完整正文流程
        text = self._re_name_with_title.sub(self._name_replacement, text)
        text = self._re_name_with_prefix.sub(self._name_replacement_with_prefix, text)
        text = self._re_name_bare.sub(self._name_replacement_bare, text)
        # 第二步：全名兜底（姓名字段极少有长句，误伤风险低）
        text = self._re_name_full.sub(self._name_replacement_bare, text)
        return text

    # ── 替换函数 ──────────────────────────────────────────────────────────────

    @staticmethod
    def _phone_replacement(m: re.Match) -> str:
        """手机号中间4位打码：138****8888"""
        digits = re.sub(r"\D", "", m.group())
        if len(digits) >= 11:
            return digits[:3] + "****" + digits[-4:]
        return "***手机号***"

    def _name_replacement(self, m: re.Match) -> str:
        """姓名替换：提取姓氏，附加[先生/女士]"""
        name = m.group()
        # 判断性别称谓
        if any(t in name for t in ("女士", "小姐")):
            return name[0] + "女士"
        return name[0] + "先生"

    def _name_replacement_bare(self, m: re.Match) -> str:
        """无称谓姓名替换：统一为 X先生"""
        name = m.group()
        return name[0] + "先生"

    def _name_replacement_with_prefix(self, m: re.Match) -> str:
        """前缀触发全名替换：保留前缀词，仅替换姓名部分。
        捕获组1=前缀, 捕获组2=姓, 捕获组3=名
        """
        prefix = m.group(1)    # "客人"、"会员" 等
        surname = m.group(2)   # 姓氏
        return prefix + surname + "先生"
