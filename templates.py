# -*- coding: utf-8 -*-
"""排版模板知识库。
借鉴 GB/T 9704-2012 党政公文、学位论文、工程设计公司商务技术文件通行规范。
"""
from dataclasses import dataclass, field
from typing import Dict


@dataclass
class Rule:
    """单类段落的排版规则。"""
    font_cn: str = "宋体"
    font_en: str = "Times New Roman"
    size: float = 12.0          # 磅
    bold: bool = False
    align: str = "justify"      # left/center/right/justify
    space_before: float = 0.0
    space_after: float = 0.0
    line_multiple: float = 1.5  # 多倍行距
    line_exact: float = 0.0     # >0 时用固定值磅
    indent: int = 0             # 首行缩进字符数
    color: str = None           # 如 "C00000"


@dataclass
class Template:
    tid: str
    name: str
    desc: str
    margin: tuple = (2.54, 2.54, 3.0, 2.5)  # 上 下 左 右 cm
    rules: Dict[str, Rule] = field(default_factory=dict)
    header_kind: str = "none"      # none/title/company/project
    footer_kind: str = "page"      # none/page/company_page
    table_uniform: bool = True     # 是否统一华通表格样式

    def get(self, ptype: str) -> Rule:
        if ptype in self.rules:
            return self.rules[ptype]
        if "body" in self.rules:
            return self.rules["body"]
        return Rule()


# ---------- 通用助手 ----------
def _body(**kw) -> Rule:
    base = dict(font_cn="宋体", size=12.0, align="justify",
                line_multiple=1.5, indent=2)
    base.update(kw)
    return Rule(**base)


TEMPLATES: Dict[str, Template] = {}


def _reg(t: Template):
    TEMPLATES[t.tid] = t


# 1. 党政公文 GB/T 9704-2012
_reg(Template(
    tid="gongwen",
    name="公文AI智能排版",
    desc="党政机关/企业通知通告，遵循 GB/T 9704 国标",
    margin=(3.7, 3.5, 2.8, 2.6),
    header_kind="none", footer_kind="page",
    rules={
        "title":    Rule(font_cn="方正小标宋简体", size=22, bold=False, align="center",
                          space_after=18, line_multiple=1.5),
        "doc_no":   Rule(font_cn="仿宋_GB2312", size=16, align="center", space_after=12),
        "recipient": Rule(font_cn="仿宋_GB2312", size=16, align="left", indent=0),
        "h1":       Rule(font_cn="黑体", size=16, align="left", indent=2,
                         space_before=6, line_exact=28),
        "h2":       Rule(font_cn="楷体_GB2312", size=16, align="left", indent=2, line_exact=28),
        "h3":       Rule(font_cn="仿宋_GB2312", size=16, bold=True, align="left", indent=2,
                         line_exact=28),
        "body":     Rule(font_cn="仿宋_GB2312", size=16, align="justify", indent=2,
                         line_exact=28),
        "signoff":  Rule(font_cn="仿宋_GB2312", size=16, align="right", line_exact=28),
        "note":     Rule(font_cn="仿宋_GB2312", size=12, align="left"),
        "table_cell": Rule(font_cn="仿宋_GB2312", size=12, align="center", line_exact=20),
    }))

# 2. 学位论文（通行格式）
_reg(Template(
    tid="thesis",
    name="学术论文排版",
    desc="学位论文/研究报告，多高校通行格式",
    margin=(2.54, 2.54, 3.0, 2.5),
    header_kind="title", footer_kind="page",
    rules={
        "title":    Rule(font_cn="黑体", size=22, align="center", space_after=12),
        "subtitle": Rule(font_cn="楷体", size=16, align="center", space_after=18),
        "h1":       Rule(font_cn="黑体", size=16, align="left", indent=2,
                         space_before=24, space_after=18),
        "h2":       Rule(font_cn="黑体", size=14, align="left",
                         space_before=12, space_after=6),
        "h3":       Rule(font_cn="黑体", size=12, align="left",
                         space_before=6, space_after=3),
        "h4":       Rule(font_cn="宋体", size=12, bold=True, align="left", indent=2),
        "body":     _body(size=12, line_multiple=1.5),
        "table_caption": Rule(font_cn="宋体", size=10.5, align="center"),
        "figure_caption": Rule(font_cn="宋体", size=10.5, align="center"),
        "table_cell": Rule(font_cn="宋体", size=10.5, align="center", line_multiple=1.0),
        "signoff":  Rule(font_cn="宋体", size=12, align="right"),
        "note":     Rule(font_cn="宋体", size=10.5, align="left"),
    }))

# 3. 工程合同
_reg(Template(
    tid="contract",
    name="工程合同排版",
    desc="各类法律协议、合同、补充协议",
    margin=(2.54, 2.54, 3.0, 2.5),
    header_kind="none", footer_kind="page",
    rules={
        "title":    Rule(font_cn="黑体", size=22, align="center", space_after=18),
        "h1":       Rule(font_cn="黑体", size=14, align="left", indent=2,
                         space_before=12, space_after=6),
        "h2":       Rule(font_cn="黑体", size=12, bold=True, align="left", indent=0,
                         space_before=6, space_after=3),
        "body":     Rule(font_cn="仿宋", size=12, align="justify", indent=2,
                         line_multiple=1.5),
        "table_cell": Rule(font_cn="仿宋", size=10.5, align="center", line_multiple=1.0),
        "signoff":  Rule(font_cn="仿宋", size=12, align="right", line_multiple=1.5),
        "note":     Rule(font_cn="宋体", size=10.5, align="left"),
    }))

# 4. 通用报告
_reg(Template(
    tid="report",
    name="通用报告排版",
    desc="工作总结/汇报/可研报告",
    margin=(2.54, 2.54, 3.0, 2.5),
    header_kind="title", footer_kind="page",
    rules={
        "title":    Rule(font_cn="黑体", size=18, align="center", space_after=12),
        "h1":       Rule(font_cn="黑体", size=16, align="left", indent=2,
                         space_before=18, space_after=12),
        "h2":       Rule(font_cn="黑体", size=14, align="left",
                         space_before=12, space_after=6),
        "h3":       Rule(font_cn="黑体", size=12, align="left",
                         space_before=6, space_after=3),
        "body":     _body(size=12, line_multiple=1.5),
        "table_caption": Rule(font_cn="宋体", size=10.5, align="center"),
        "figure_caption": Rule(font_cn="宋体", size=10.5, align="center"),
        "table_cell": Rule(font_cn="宋体", size=10.5, align="center", line_multiple=1.0),
        "signoff":  Rule(font_cn="宋体", size=12, align="right"),
        "note":     Rule(font_cn="宋体", size=10.5, align="left"),
    }))

# 5. 中网华通商务与技术文件排版格式（GS-XM-05-002 一阶段设计排版样本实测提取）
# 页边距: 上2.5 下2.0 左3.5 右2.0 cm
# 正文: 宋体/TNR 小四12pt 两端对齐 首行缩进2字符 行距固定14.4磅
# 各级标题: 黑体小四12pt 首行缩进2字符 段前6段后6 单倍行距
# 表题: 宋体小四居中; 表内: 宋体五号10.5pt居中; 表注: 宋体五号
# 封面标题: 黑体小二18pt 居中
_reg(Template(
    tid="publish",
    name="华通文件出版排版",
    desc="中网华通商务与技术文件格式（GS-XM-05 范本实测）：宋体小四/固定14.4磅/黑体小四标题",
    margin=(2.5, 2.0, 3.5, 2.0),
    header_kind="project", footer_kind="company_page",
    rules={
        "title":    Rule(font_cn="黑体", size=18, align="center", space_before=12,
                         space_after=12, line_multiple=1.0),
        "subtitle": Rule(font_cn="宋体", size=16, align="center", space_after=6,
                         line_multiple=1.0),
        "h1":       Rule(font_cn="黑体", size=12, align="left", indent=2,
                         space_before=6, space_after=6, line_multiple=1.0),
        "h2":       Rule(font_cn="黑体", size=12, align="left", indent=2,
                         space_before=6, space_after=6, line_multiple=1.0),
        "h3":       Rule(font_cn="黑体", size=12, align="left", indent=2,
                         space_before=6, space_after=6, line_multiple=1.0),
        "h4":       Rule(font_cn="宋体", size=12, bold=True, align="left", indent=2),
        "body":     Rule(font_cn="宋体", font_en="Times New Roman", size=12,
                         align="justify", indent=2, line_exact=14.4),
        "table_caption": Rule(font_cn="宋体", font_en="Times New Roman", size=12,
                              align="center", line_exact=14.4),
        "figure_caption": Rule(font_cn="宋体", font_en="Times New Roman", size=12,
                               align="center", line_exact=14.4),
        "table_cell": Rule(font_cn="宋体", font_en="Times New Roman", size=10.0,
                           align="center", line_multiple=1.0,
                           space_before=2, space_after=2),
        "note":     Rule(font_cn="宋体", font_en="Times New Roman", size=10.5,
                         align="justify", line_exact=14.4),
        "signoff":  Rule(font_cn="宋体", size=12, align="right", indent=-2, line_exact=14.4),
    }))

# 6. 招投标文书
_reg(Template(
    tid="bidding",
    name="招投标文书排版",
    desc="投标函/技术标书/商务标书",
    margin=(2.54, 2.54, 3.0, 2.5),
    header_kind="title", footer_kind="page",
    rules={
        "title":    Rule(font_cn="黑体", size=18, align="center", space_after=12),
        "h1":       Rule(font_cn="黑体", size=14, align="left", space_before=6,
                         space_after=3),
        "body":     Rule(font_cn="仿宋_GB2312", size=14, align="justify", indent=2,
                         line_multiple=1.5),
        "table_cell": Rule(font_cn="仿宋_GB2312", size=10.5, align="center", line_multiple=1.0),
        "signoff":  Rule(font_cn="仿宋_GB2312", size=14, align="right"),
        "note":     Rule(font_cn="宋体", size=10.5, align="left"),
    }))

# 7. 习题试卷
_reg(Template(
    tid="exam",
    name="习题试卷排版",
    desc="试题/试卷/作业文档",
    margin=(2.54, 2.54, 2.5, 2.5),
    header_kind="none", footer_kind="page",
    rules={
        "title":    Rule(font_cn="黑体", size=18, align="center", space_after=6),
        "h1":       Rule(font_cn="黑体", size=14, align="left", space_before=12,
                         space_after=6),
        "h2":       Rule(font_cn="黑体", size=12, bold=True, align="left", indent=0,
                         space_before=6, space_after=3),
        "body":     Rule(font_cn="宋体", size=12, align="justify", indent=0,
                         line_multiple=1.5),
        "table_cell": Rule(font_cn="宋体", size=10.5, align="center", line_multiple=1.0),
        "note":     Rule(font_cn="宋体", size=10.5, align="left"),
    }))

# 8. 通用文档
_reg(Template(
    tid="general",
    name="通用文档",
    desc="适合大部分文档的通用排版",
    margin=(2.54, 2.54, 3.0, 2.5),
    header_kind="title", footer_kind="page",
    rules={
        "title":    Rule(font_cn="黑体", size=16, align="center", space_after=10),
        "h1":       Rule(font_cn="黑体", size=14, align="left", space_before=6),
        "body":     _body(size=12, line_multiple=1.5),
        "table_cell": Rule(font_cn="宋体", size=10.5, align="center", line_multiple=1.0),
        "note":     Rule(font_cn="宋体", size=10.5, align="left"),
    }))


# 首页卡片配置（按使用频率从高到低）
HOME_CARDS = [
    ("publish",   "华通文件出版排版",  "📘"),
    ("report",    "通用报告排版",      "📊"),
    ("contract",  "工程合同排版",      "📑"),
    ("thesis",    "学术论文排版",      "🎓"),
    ("bidding",   "招投标文书排版",    "📋"),
    ("gongwen",   "公文AI智能排版",    "📄"),
    ("exam",      "习题试卷排版",      "✏️"),
    ("sample",    "导入范文排版",      "🧩"),
]

# 段落类型中文名
TYPE_CN = {
    "title": "大标题", "subtitle": "副标题", "doc_no": "发文字号",
    "h1": "一级标题", "h2": "二级标题", "h3": "三级标题", "h4": "四级标题",
    "body": "正文", "recipient": "主送机关", "signoff": "落款", "note": "注释",
    "table_caption": "表题", "figure_caption": "图题", "table_cell": "表格",
}
