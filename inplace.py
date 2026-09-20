# -*- coding: utf-8 -*-
"""原地套用当前模板样式：Word COM 直接改原文档，不丢图片/公式/文本框。"""
import re

_CM = 28.35


def _detect_level(text, style_name):
    """根据 Word 样式名 + 文字正则判定段落级别。返回 0=body,1/2/3=h1/h2/h3,4=h4。"""
    if any(k in style_name for k in ("标题 1", "标题1", "Heading 1", "Heading1")):
        return 1
    if any(k in style_name for k in ("标题 2", "标题2", "Heading 2", "Heading2")):
        return 2
    if any(k in style_name for k in ("标题 3", "标题3", "Heading 3", "Heading3")):
        return 3
    if any(k in style_name for k in ("标题 4", "标题4", "Heading 4", "Heading4")):
        return 4
    if re.match(r"^[一二三四五六七八九十]+[、．.]\s*\S", text):
        return 1
    if re.match(r"^第[一二三四五六七八九十0-9]+[章节条款]\s*\S", text):
        return 1
    if re.match(r"^\d+\.\d+(\.\d+)?\s+\S", text):
        return 3
    if re.match(r"^\d+\s*[\.、]\s+\S", text):
        return 2
    return 0


def _apply_rule(p, rng, rule):
    """把 templates.Rule 套到一个 Word 段落上。"""
    f = rng.Font
    f.NameFarEast = rule.font_cn
    f.NameAscii = rule.font_en
    f.Size = rule.size
    f.Bold = bool(rule.bold)
    align_map = {"left": 0, "center": 1, "right": 2, "justify": 3}
    p.Alignment = align_map.get(rule.align, 3)
    p.CharacterUnitFirstLineIndent = rule.indent
    p.SpaceBefore = rule.space_before * _CM / 28.35 * 72  # 磅→磅（本来就是磅）
    p.SpaceAfter = rule.space_after
    if rule.line_exact and rule.line_exact > 0:
        p.LineSpacingRule = 4  # wdLineSpaceExactly
        p.LineSpacing = rule.line_exact
    else:
        p.LineSpacingRule = 5  # multiple
        p.LineSpacing = rule.line_multiple * 12


def apply_template_inplace(src_path, template, out_path=None, header_text=""):
    """打开原 docx，按传入的 Template 原地套样式。保留图片/公式。"""
    import os
    import win32com.client as win32
    if out_path is None:
        out_path = os.path.splitext(src_path)[0] + "_" + template.tid + ".docx"

    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        doc = word.Documents.Open(os.path.abspath(src_path), ReadOnly=False,
                                   ConfirmConversions=False, AddToRecentFiles=False)
        try:
            # 页面
            ps = doc.PageSetup
            ps.PageWidth = 21.0 * _CM
            ps.PageHeight = 29.7 * _CM
            ps.TopMargin = template.margin[0] * _CM
            ps.BottomMargin = template.margin[1] * _CM
            ps.LeftMargin = template.margin[2] * _CM
            ps.RightMargin = template.margin[3] * _CM

            level_key = {1: "h1", 2: "h2", 3: "h3", 4: "h4", 0: "body"}
            for p in doc.Paragraphs:
                rng = p.Range
                style_name = (p.Style.NameLocal or "").strip()
                text = rng.Text.strip()
                if not text:
                    continue
                level = _detect_level(text, style_name)
                ptype = level_key[level]
                rule = template.get(ptype)
                _apply_rule(p, rng, rule)
                if level == 1:
                    p.PageBreakBefore = True

            # 页眉（可选）
            if header_text:
                for sec in doc.Sections:
                    hdr = sec.Headers(1)
                    hdr.Range.Text = header_text
                    hdr.Range.Font.Size = 10.5
                    hdr.Range.ParagraphFormat.Alignment = 1

            doc.SaveAs2(os.path.abspath(out_path), FileFormat=16)
        finally:
            doc.Close(True)
    finally:
        word.Quit()
    return out_path
