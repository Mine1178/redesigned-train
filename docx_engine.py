# -*- coding: utf-8 -*-
"""docx 排版引擎：
1) 从 docx/txt 提取纯文本段落
2) 本地启发式分类段落类型（无网/无 API 也能用）
3) 按模板规则生成 .docx
4) 从范文 .docx 抽取样式指纹（导入范文排版）
"""
import os
import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn

from templates import Template, Rule

ALIGN = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
         "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}


# ---------- 文本提取 ----------
def extract_paragraphs(path):
    """从 .docx 或 .txt 提取段落文本（保留原始顺序与空行）。"""
    return extract_blocks(path)


def _open_docx_robust(path):
    """打开 docx，遇到 WPS/损坏/伪装文件时自动用 Word COM 另存为标准 docx。"""
    try:
        return Document(path)
    except Exception as e:
        msg = str(e)
        if "relationship" not in msg and "Package" not in msg and "not a zip" not in msg:
            raise
    # 兜底：用 Word COM 打开并另存为标准 docx
    import tempfile, win32com.client as win32
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    tmp = os.path.join(tempfile.gettempdir(), "_smartlayouter_converted.docx")
    if os.path.exists(tmp):
        try: os.remove(tmp)
        except OSError: pass
    try:
        doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True,
                                   ConfirmConversions=False, AddToRecentFiles=False)
        doc.SaveAs2(tmp, FileFormat=16)  # wdFormatXMLDocument
        doc.Close(False)
    finally:
        word.Quit()
    return Document(tmp)


def extract_blocks(path):
    """提取文档块序列。
    返回 list，元素两种：
      - str        : 普通段落文本
      - list[list[str]] : 表格（二维单元格）
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".txt":
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            raw = f.read()
        return [l.rstrip() for l in raw.splitlines()]

    from docx.oxml.ns import qn as _qn
    doc = _open_docx_robust(path)
    body = doc.element.body
    # 建映射：xml 元素 -> python-docx 包装对象
    p_map = {p._p: p for p in doc.paragraphs}
    t_map = {t._tbl: t for t in doc.tables}

    blocks = []
    for child in body.iterchildren():
        if child.tag == _qn("w:p"):
            p = p_map.get(child)
            if p is not None:
                blocks.append(p.text.strip())
        elif child.tag == _qn("w:tbl"):
            t = t_map.get(child)
            if t is not None:
                rows = []
                for row in t.rows:
                    rows.append([c.text.strip() for c in row.cells])
                blocks.append(rows)
    return blocks


def guess_project_name(blocks):
    """从前几段找含"工程"的行做页眉工程名。"""
    for b in blocks[:8]:
        if isinstance(b, str):
            t = b.strip()
            if 6 < len(t) < 50 and ("工程" in t or "设计" in t):
                return t
    return ""




def renumber_figures_tables(blocks):
    """扫描 blocks 里的表题/图题段落，按当前章节重编号。
    识别 h1(一、/第X章) 为章，h2(1.) 为节。
    输出新的 blocks，原地改表题/图题文字。"""
    chap = 0
    sec = 0
    out = []
    table_no = {}
    fig_no = {}
    for b in blocks:
        if isinstance(b, list):
            out.append(b)
            continue
        t = b.strip()
        # 章
        if re.match(r"^[一二三四五六七八九十]+[、．.]", t) or re.match(r"^第[一二三四五六七八九十0-9]+[章节]", t):
            chap += 1
            sec = 0
            table_no[chap] = 0
            fig_no[chap] = 0
        # 节
        elif re.match(r"^\d+\s*[\.、]\s+\S", t) and not re.match(r"^\d+\.\d+", t):
            sec += 1
        # 表题
        m = re.match(r"^(表)\s*[\d.]+", t)
        if m:
            table_no[chap] = table_no.get(chap, 0) + 1
            new_t = re.sub(r"^表\s*[\d.]+", f"表{chap}.{table_no[chap]}", t)
            out.append(new_t)
            continue
        m = re.match(r"^(图)\s*[\d.]+", t)
        if m:
            fig_no[chap] = fig_no.get(chap, 0) + 1
            new_t = re.sub(r"^图\s*[\d.]+", f"图{chap}.{fig_no[chap]}", t)
            out.append(new_t)
            continue
        out.append(b)
    return out

# ---------- 本地启发式分类 ----------
RE_H1_CN = re.compile(r"^[一二三四五六七八九十百]+[、．.]\s*\S")
RE_H1_CHAPTER = re.compile(r"^第[一二三四五六七八九十百0-9]+[章节条款部分编]\s*\S")
RE_H2_PAREN = re.compile(r"^[（(][一二三四五六七八九十]+[）)]")
RE_H2_NUM = re.compile(r"^\d+\s*[\.、]\s+\S")           # 1. 概述 / 2、概算
RE_H3_NUM = re.compile(r"^\d+\.\d+(\.\d+)?\s+\S")    # 1.1 / 1.1.1
RE_H4_PAREN = re.compile(r"^[（(]\d+[）)]")
RE_DOCNO = re.compile(r"〔\d{4}〕\s*第?\d+\s*号")
RE_DATE = re.compile(r"^[一二三四五六七八九十〇○0-9]{2,4}\s*年\s*[0-9一二三四五六七八九十]{1,2}\s*月(\s*[0-9一二三四五六七八九十]{1,2}\s*日)?$")
RE_RECIPIENT = re.compile(r"[：:]\s*$")
RE_SIGNOFF_KW = re.compile(r"(特此(通知|函告|报告|批复|公告|说明|承诺)|^[一二三四五六七八九十]{2,4}公司$|^有限公司$|^设计单位|^建设单位|^编制|^审核|^批准)")
RE_TABLE_CAPTION = re.compile(r"^(表|图)\s*\d+(\.\d+)*")
RE_FIGURE_CAPTION = re.compile(r"^图\s*\d+(\.\d+)*")
RE_NOTE = re.compile(r"^(说明|注|注释|附注)[:：]")
RE_TOC_ENTRY = re.compile(r"\t\d+\s*$")
RE_SHORT_HEADING = re.compile(r"^[\u4e00-\u9fffA-Za-z0-9·、（）()]{2,18}$")
# 目录条目：行尾是 tab + 页码（如 "一、概述\t12"）
RE_TOC_ENTRY = re.compile(r"\t\d+\s*$")
# 短标题：独立一行、无句末标点、较短
RE_SHORT_HEADING = re.compile(r"^[\u4e00-\u9fffA-Za-z0-9·、（）()]{2,18}$")


def local_classify(blocks):
    """对块序列中的段落做本地分类，返回 {块序号: ptype}。表格块跳过。"""
    result = {}
    nonempty = [i for i, b in enumerate(blocks) if isinstance(b, str) and b.strip()]
    if not nonempty:
        return result
    result[nonempty[0]] = "title"
    for k, idx in enumerate(nonempty):
        if k == 0:
            continue
        t = blocks[idx].strip()
        if RE_TOC_ENTRY.search(t) or re.match(r"^目\s*录$", t):
            result[idx] = "toc_skip"
            continue
        if RE_DOCNO.search(t):
            result[idx] = "doc_no"
        elif RE_FIGURE_CAPTION.match(t):
            result[idx] = "figure_caption"
        elif RE_TABLE_CAPTION.match(t):
            result[idx] = "table_caption"
        elif RE_NOTE.match(t):
            result[idx] = "note"
        elif RE_H1_CN.match(t) or RE_H1_CHAPTER.match(t):
            result[idx] = "h1"
        elif RE_H2_NUM.match(t) or RE_H2_PAREN.match(t):
            result[idx] = "h2"
        elif RE_H3_NUM.match(t):
            result[idx] = "h3"
        elif RE_H4_PAREN.match(t):
            result[idx] = "body"  # （1）（2）是正文中的枚举项，按正文排
        elif k >= len(nonempty) - 3 and (RE_SIGNOFF_KW.search(t) or RE_DATE.match(t)):
            result[idx] = "signoff"
        elif RE_RECIPIENT.search(t) and len(t) < 30:
            result[idx] = "recipient"
        elif RE_SHORT_HEADING.match(t) and len(t) <= 12:
            result[idx] = "h2"
        else:
            result[idx] = "body"
    return result


# ---------- docx 写出 ----------
def _set_run_font(run, rule):
    run.font.size = Pt(rule.size)
    run.font.bold = rule.bold
    run.font.name = rule.font_en
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("w:rFonts"), {})
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), rule.font_en)
    rFonts.set(qn("w:hAnsi"), rule.font_en)
    rFonts.set(qn("w:eastAsia"), rule.font_cn)
    if rule.color:
        run.font.color.rgb = RGBColor.from_string(rule.color)


def _set_first_line_chars(p, chars, size_pt):
    """首行按"字符"缩进（firstLineChars=200 表示2字符）。"""
    pPr = p._p.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = pPr.makeelement(qn("w:ind"), {})
        pPr.append(ind)
    for k in ("w:firstLineChars", "w:firstLine"):
        if ind.get(qn(k)) is not None:
            del ind.attrib[qn(k)]
    if chars > 0:
        ind.set(qn("w:firstLineChars"), str(chars * 100))
        ind.set(qn("w:firstLine"), str(int(size_pt * 2 * chars * 20)))


def _apply_paragraph(p, text, rule, ptype="body"):
    run = p.add_run(text)
    _set_run_font(run, rule)
    pf = p.paragraph_format
    pf.alignment = ALIGN.get(rule.align, WD_ALIGN_PARAGRAPH.JUSTIFY)
    pf.space_before = Pt(rule.space_before)
    pf.space_after = Pt(rule.space_after)
    pf.widow_control = True
    if ptype in ("h1", "h2", "h3", "h4"):
        pf.keep_with_next = True
    if ptype in ("table_caption", "figure_caption"):
        pf.keep_together = True
    if rule.line_exact and rule.line_exact > 0:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(rule.line_exact)
    else:
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.line_spacing = rule.line_multiple
    _set_first_line_chars(p, rule.indent, rule.size)


def _set_cell_border(cell, sz="4"):
    """给单元格四边加黑色单线。sz=4 表 0.5 磅，sz=12 表 1.5 磅。"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        tcBorders = tcPr.makeelement(qn("w:tcBorders"), {})
        tcPr.append(tcBorders)
    for edge in ("top", "left", "bottom", "right"):
        el = tcBorders.find(qn(f"w:{edge}"))
        if el is None:
            el = tcBorders.makeelement(qn(f"w:{edge}"), {})
            tcBorders.append(el)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")


def _repeat_header_row(row):
    """跨页表格：表头行在每页重复出现。"""
    trPr = row._tr.get_or_add_trPr()
    tblHeader = trPr.makeelement(qn("w:tblHeader"), {qn("w:val"): "true"})
    trPr.append(tblHeader)


def _add_styled_table(doc, rows_data, cell_rule):
    """按中网华通规范画表格：全黑单线边框、宋体五号、水平垂直居中、首行表头重复。"""
    if not rows_data:
        return
    ncols = max(len(r) for r in rows_data)
    table = doc.add_table(rows=len(rows_data), cols=ncols)
    table.autofit = True
    for ri, row_data in enumerate(rows_data):
        row = table.rows[ri]
        for ci in range(ncols):
            cell = row.cells[ci]
            text = row_data[ci] if ci < len(row_data) else ""
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(text)
            _set_run_font(run, cell_rule)
            # 纯数字右对齐，其他居中
            _text = row_data[ci] if ci < len(row_data) else ""
            import re as _re_num
            if _re_num.match(r"^[\d\.,%\s\-]+$", _text.strip()):
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            # 垂直居中
            tcPr = cell._tc.get_or_add_tcPr()
            vAlign = tcPr.makeelement(qn("w:vAlign"), {qn("w:val"): "center"})
            tcPr.append(vAlign)
            # 外框粗线，内线细线
            is_edge = (ri == 0 or ri == len(rows_data) - 1 or ci == 0 or ci == ncols - 1)
            _set_cell_border(cell, sz="12" if is_edge else "4")
        if ri == 0:
            _repeat_header_row(row)
    # 表后空一行
    doc.add_paragraph()




def _add_section_break(doc, start_type="nextPage"):
    """在文档末尾插入一个分节符（默认下一页）。"""
    new_sec = doc.add_section(WD_SECTION.NEW_PAGE)
    return new_sec


def _set_page_number_format(section, fmt="decimal", start=1):
    """设置节的页码格式：fmt=decimal(阿拉伯)/upperRoman/lowerRoman；start=起始页码。"""
    sectPr = section._sectPr
    pgNumType = sectPr.find(qn("w:pgNumType"))
    if pgNumType is None:
        pgNumType = sectPr.makeelement(qn("w:pgNumType"), {})
        sectPr.append(pgNumType)
    fmt_map = {"decimal": "decimal", "upperRoman": "upperRoman",
               "lowerRoman": "lowerRoman"}
    pgNumType.set(qn("w:fmt"), fmt_map.get(fmt, "decimal"))
    pgNumType.set(qn("w:start"), str(start))


def _set_footer_empty(section):
    """清空节的页脚。"""
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.text = ""
    for run in p.runs:
        run.text = ""


def build_document(blocks, classified, template: Template, out_path,
                   header_text="", page_num=True, with_toc=False,
                   watermark="", cover=None, renumber=False):
    if renumber:
        blocks = renumber_figures_tables(blocks)
    """把 (块序列, 类型字典) 按模板生成 docx。blocks 元素：str 段落 / list 表格。
    可选：页眉、页码、自动目录、水印、封面。"""
    doc = Document()
    sec = doc.sections[0]
    sec.page_height = Cm(29.7)
    sec.page_width = Cm(21.0)
    sec.top_margin = Cm(template.margin[0])
    sec.bottom_margin = Cm(template.margin[1])
    sec.left_margin = Cm(template.margin[2])
    sec.right_margin = Cm(template.margin[3])
    sec.header_distance = Cm(1.7)
    sec.footer_distance = Cm(1.5)

    # 默认样式字体
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    # 封面（节1：无页码）
    if cover:
        build_cover(doc, **cover)
        if page_num:
            sec1 = doc.sections[-1]
            _set_footer_empty(sec1)
        # 目录分节
        if with_toc:
            _add_section_break(doc)
    # 目录（节2：罗马数字）
    if with_toc:
        add_toc(doc)
        if page_num:
            sec2 = doc.sections[-1]
            _set_page_number_format(sec2, "lowerRoman", 1)
        # 正文分节
        if cover and page_num:
            _add_section_break(doc)
    # 正文（节3：阿拉伯数字从1开始）
    if page_num:
        sec_body = doc.sections[-1]
        _set_page_number_format(sec_body, "decimal", 1)
    # 页眉页脚
    add_header_footer(doc, header_text=header_text, page_num=page_num)
    if watermark:
        add_watermark(doc, watermark)

    cell_rule = template.get("table_cell")
    for idx, block in enumerate(blocks):
        if isinstance(block, list):
            _add_styled_table(doc, block, cell_rule)
            continue
        text = block.strip()
        if not text:
            continue
        ptype = classified.get(idx, "body")
        if ptype == "toc_skip":
            continue
        rule = template.get(ptype)
        p = doc.add_paragraph()
        _apply_paragraph(p, text, rule, ptype)
        if ptype == "h1" and idx > 0:
            p.paragraph_format.page_break_before = True
    # 让 Word 打开文档时自动更新域（目录页码）
    settings = doc.settings.element
    upd = settings.find(qn("w:updateFields"))
    if upd is None:
        upd = settings.makeelement(qn("w:updateFields"), {})
        settings.append(upd)
    upd.set(qn("w:val"), "true")

    doc.save(out_path)
    return out_path


# ---------- 范文样式抽取（导入范文排版） ----------
def extract_style_signature(path):
    """读范文 docx，按段落类型聚类得到样式指纹，返回 dict[tpl_tid-like rules]。
    简化策略：按字号/字体/对齐/缩进聚类，再用启发式命名。
    """
    doc = Document(path)
    sig = {"rules": {}}
    buckets = {}
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        f = p.runs[0].font if p.runs else None
        ea = None
        if p.runs:
            rPr = p.runs[0]._element.rPr
            rfonts = rPr.find(qn("w:rFonts")) if rPr is not None else None
            ea = rfonts.get(qn("w:eastAsia")) if rfonts is not None else None
        size = round((f.size.pt if f and f.size else 12), 1)
        # 过滤异常字号（域代码/空 run）
        if size < 9 or size > 30:
            size = 12
        cn = ea or (f.name if f else "宋体") or "宋体"
        align = p.alignment
        align_s = {0: "left", 1: "center", 2: "right", 3: "justify",
                   None: "justify"}.get(int(align) if align is not None else None, "justify")
        key = (cn, size, align_s)
        buckets.setdefault(key, []).append(t)

    # 把段落数最多的 bucket 当 body；其余按字号降序当 title/h1/h2...
    ordered = sorted(buckets.items(), key=lambda kv: -len(kv[1]))
    rules = {}
    used = set()
    if ordered:
        (cn, size, align), texts = ordered[0]
        rules["body"] = Rule(font_cn=cn, size=size, align=align, indent=2, line_multiple=1.5)
        used.add((cn, size, align))
    rest = [k for k in buckets if k not in used]
    rest.sort(key=lambda k: (-k[1], 0 if k[2] == "center" else 1))
    if rest:
        cn, size, align = rest[0]
        rules["title"] = Rule(font_cn=cn, size=size, align=align or "center",
                              space_before=12, space_after=12, line_multiple=1.5)
        used.add((cn, size, align))
    for i, k in enumerate([c for c in rest[1:] if c not in used]):
        cn, size, align = k
        rules[f"h{i+1}"] = Rule(font_cn=cn, size=size, align="left", indent=2,
                                space_before=6, space_after=6, line_multiple=1.5)
    # 页面：用范文第一节边距
    try:
        s = doc.sections[0]
        sig["margin"] = (round(s.top_margin.cm, 2), round(s.bottom_margin.cm, 2),
                         round(s.left_margin.cm, 2), round(s.right_margin.cm, 2))
    except Exception:
        sig["margin"] = (2.54, 2.54, 3.0, 2.5)
    sig["rules"] = rules
    return sig


def template_from_signature(sig, name="范文复刻模板"):
    return Template(tid="__sample__", name=name,
                    desc="从用户导入的范文自动学习的排版样式",
                    margin=sig["margin"], rules=sig["rules"])


# ============================================================
# 第一波：清理脏数据 / 页眉页脚 / 目录 / 批量
# ============================================================
def clean_blocks(blocks):
    """一键清理脏数据：合并多余空行、行首尾空格、全半角统一、连续标点。
    返回 (新blocks, 清理报告 list[str])。"""
    report = []
    out = []
    blank_run = 0
    for b in blocks:
        if isinstance(b, list):
            out.append(b)
            blank_run = 0
            continue
        t = b.strip()
        if not t:
            blank_run += 1
            if blank_run <= 1:
                out.append("")
            else:
                report.append("删除连续空行")
            continue
        # 合并行内多空格
        t2 = re.sub(r"[ \t\u3000]{2,}", " ", t)
        if t2 != t:
            report.append("合并行内多余空格")
        # 英文逗号/句号在中文语境下转全角（保守：仅当中英文混排时）
        t2 = re.sub(r"(?<=[\u4e00-\u9fff]),", "，", t2)
        t2 = re.sub(r"(?<=[\u4e00-\u9fff])\.(?=[\u4e00-\u9fff])", "。", t2)
        out.append(t2)
        blank_run = 0
    # 去重报告
    uniq = list(dict.fromkeys(report))
    return out, uniq


def add_header_footer(doc, header_text="", page_num=True,
                      company="北京中网华通设计咨询有限公司"):
    """华通范本：页眉=工程名居中宋体五号；页脚=公司名左+右对齐页码。"""
    sec = doc.sections[0]
    hdr = sec.header
    hdr.is_linked_to_previous = False
    p = hdr.paragraphs[0]
    p.text = header_text
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in p.runs:
        run.font.size = Pt(10.5)
        run.font.name = "Times New Roman"
        run._element.get_or_add_rPr().append(
            run._element.makeelement(qn("w:rFonts"), {qn("w:eastAsia"): "宋体"}))
    if page_num:
        ftr = sec.footer
        ftr.is_linked_to_previous = False
        fp = ftr.paragraphs[0]
        fp.text = ""
        run_co = fp.add_run(company)
        run_co.font.size = Pt(10.5)
        run_co.font.name = "Times New Roman"
        run_co._element.get_or_add_rPr().append(
            run_co._element.makeelement(qn("w:rFonts"), {qn("w:eastAsia"): "宋体"}))
        fp.add_run("\t")
        fldB = fp.add_run()
        fldB._r.append(fp._p.makeelement(qn("w:fldChar"), {qn("w:fldCharType"): "begin"}))
        instr = fp.add_run()
        instr._r.append(fp._p.makeelement(qn("w:instrText"), {qn("xml:space"): "preserve"}))
        instr._r[-1].text = " PAGE "
        fldE = fp.add_run()
        fldE._r.append(fp._p.makeelement(qn("w:fldChar"), {qn("w:fldCharType"): "end"}))
        for r in (fldB, instr, fldE):
            r.font.size = Pt(10.5)
        pPr = fp._p.get_or_add_pPr()
        tabs = pPr.makeelement(qn("w:tabs"), {})
        tabs.append(tabs.makeelement(qn("w:tab"), {
            qn("w:val"): "right", qn("w:leader"): "none", qn("w:pos"): "8800"}))
        pPr.append(tabs)



def add_toc(doc, title="目  录"):
    """在文档开头插入目录页（标题 + TOC 域，基于标题1/2/3）。"""
    # 页面前先加分页
    doc.add_page_break()
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_title.add_run(title)
    r.font.size = Pt(16)
    r.font.bold = False
    r._element.get_or_add_rPr().append(
        r._element.makeelement(qn("w:rFonts"), {qn("w:eastAsia"): "黑体"}))
    p_title.paragraph_format.space_before = Pt(18)
    p_title.paragraph_format.space_after = Pt(18)
    p_title.paragraph_format.line_spacing = Pt(25)

    p = doc.add_paragraph()
    run = p.add_run()
    fldBegin = run._r.makeelement(qn("w:fldChar"), {qn("w:fldCharType"): "begin"})
    instr = run._r.makeelement(qn("w:instrText"), {qn("xml:space"): "preserve"})
    instr.text = r' TOC \o "1-3" \h \z \u '
    fldSep = run._r.makeelement(qn("w:fldChar"), {qn("w:fldCharType"): "separate"})
    placeholder = run._r.makeelement(qn("w:t"), {})
    placeholder.text = "（右键此目录 → 更新域，即可生成带页码的目录）"
    fldEnd = run._r.makeelement(qn("w:fldChar"), {qn("w:fldCharType"): "end"})
    for el in (fldBegin, instr, fldSep, placeholder, fldEnd):
        run._r.append(el)
    doc.add_page_break()


# ============================================================
# 第二波：封面 / 水印 / 修订清理 / 定稿体检
# ============================================================
def build_cover(doc, project_name, book_no, company="北京中网华通设计咨询有限公司",
                date="2025年  月"):
    """按中网华通封面范本：全部黑体18pt居中。"""
    lines = [
        "中 国 联 通",
        project_name,
        "一阶段设计",
        book_no,
        company,
    ]
    for line in lines:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(line)
        r.font.size = Pt(18)
        r.font.name = "Times New Roman"
        r._element.get_or_add_rPr().append(
            r._element.makeelement(qn("w:rFonts"), {qn("w:eastAsia"): "黑体"}))
    doc.add_page_break()


def add_watermark(doc, text="中网华通"):
    """在页眉里插入 VML 斜向文字水印。"""
    sec = doc.sections[0]
    hdr = sec.header
    p = hdr.paragraphs[0]
    xml = (
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:o="urn:schemas-microsoft-com:office:office">'
        '<w:r><w:rPr><w:noProof/></w:rPr>'
        f'<w:pict><v:shapetype id="_x0001_t202" coordsize="21600,21600" '
        'o:spt="202" path="m,l,21600r21600,l21600,xe">'
        '<v:stroke joinstyle="miter"/></v:shapetype>'
        f'<v:shape id="PowerPlusWaterMarkObject" o:spid="_x0000_s2049" '
        'type="#_x0001_t202" style="position:absolute;margin-left:0;margin-top:0;'
        'width:400pt;height:200pt;rotation:315;z-index:-251654140;'
        'mso-position-horizontal:center;mso-position-horizontal-relative:margin;'
        'mso-position-vertical:center;mso-position-vertical-relative:margin"'
        ' fillcolor="#d9d9d9" stroked="f">'
        f'<v:fill opacity=".5"/><v:textpath style="font-family:&quot;宋体&quot;'
        f';font-size:1pt" string="{text}"/></v:shape></w:pict></w:r></w:p>'
    )
    from docx.oxml import parse_xml
    p._p.addnext(parse_xml(xml))


def docx_to_pdf(docx_path, pdf_path=None):
    """用本机 Word 把 docx 转成 pdf（需装 Office/WPS）。返回 pdf 路径。"""
    if pdf_path is None:
        pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
    import win32com.client as win32
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True,
                                   ConfirmConversions=False, AddToRecentFiles=False)
        doc.SaveAs2(os.path.abspath(pdf_path), FileFormat=17)  # wdFormatPDF
        doc.Close(False)
    finally:
        word.Quit()
    return pdf_path


def strip_revisions(path_in, path_out):
    """删除所有修订标记与批注（python-docx 层面：移除 w:ins/w:del/comment）。"""
    import shutil, zipfile
    shutil.copy(path_in, path_out)
    # 解压-改 XML-重打包
    tmp = path_out + ".tmp"
    with zipfile.ZipFile(path_out, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.namelist():
            data = zin.read(item)
            if item.endswith(".xml") and ("document" in item or "header" in item or "footer" in item):
                txt = data.decode("utf-8")
                # 移除插入/删除标记，保留内容
                txt = re.sub(r"<w:ins [^>]*>", "", txt)
                txt = re.sub(r"</w:ins>", "", txt)
                txt = re.sub(r"<w:del [^>]*>.*?</w:del>", "", txt, flags=re.S)
                txt = re.sub(r"<w:comment [^>]*>.*?</w:comment>", "", txt, flags=re.S)
                data = txt.encode("utf-8")
            zout.writestr(item, data)
    os.replace(tmp, path_out)
    return path_out


def qc_check(blocks, classified, template):
    """定稿体检：返回问题列表 [(级别, 描述)]。"""
    issues = []
    nonempty = [(i, b) for i, b in enumerate(blocks)
                if isinstance(b, str) and b.strip()]
    if not nonempty:
        issues.append(("WARN", "文档为空"))
        return issues
    # 1. 大标题是否存在
    if not any(classified.get(i) == "title" for i, _ in nonempty):
        issues.append(("WARN", "未识别到大标题，请检查首段"))
    # 2. 一级标题数量
    h1 = [i for i, _ in nonempty if classified.get(i) == "h1"]
    if len(h1) == 0:
        issues.append(("WARN", "未识别到一级标题（一、/第X章）"))
    # 3. 表格前后是否有表题
    for i, b in enumerate(blocks):
        if isinstance(b, list):
            prev = blocks[i-1] if i > 0 else ""
            if isinstance(prev, str) and not prev.strip().startswith(("表", "图")):
                issues.append(("INFO", f"第 {i+1} 个表格前未检测到表题"))
    # 4. 标点混用
    cn_text = "".join(b for _, b in nonempty if isinstance(b, str))
    if re.search(r"[\u4e00-\u9fff],", cn_text):
        issues.append(("INFO", "存在中文后接英文逗号，建议统一为全角"))
    # 5. 字体合规提示
    body_rule = template.get("body")
    issues.append(("OK", f"正文规范：{body_rule.font_cn} {body_rule.size}pt "
                         f"{'固定'+str(body_rule.line_exact)+'磅' if body_rule.line_exact else str(body_rule.line_multiple)+'倍行距'}"))
    return issues
