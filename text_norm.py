# -*- coding: utf-8 -*-
"""文本规范化：中英文空格、全半角、空格清理、空白段清理"""
import re

_HALF_TO_FULL = {
    ',': '，', '.': '。', ';': '；', ':': '：',
    '?': '？', '!': '！', '(': '（', ')': '）',
}
_CJK = r'[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]'
_EN = r'[A-Za-z0-9]'


def normalize_text(text):
    if not text:
        return text
    text = re.sub(rf'({_CJK})({_EN})', r'\1 \2', text)
    text = re.sub(rf'({_EN})({_CJK})', r'\1 \2', text)
    def repl(m):
        return m.group(1) + _HALF_TO_FULL.get(m.group(2), m.group(2)) + m.group(3)
    for _ in range(3):
        text = re.sub(rf'({_CJK})\s*([,.;:?!()])\s*({_CJK})', repl, text)
    return text


def clean_spaces(text):
    """清理中文之间多余空格。"""
    if not text:
        return text
    text = re.sub(rf'({_CJK})\s+({_CJK})', r'\1\2', text)
    text = text.strip()
    text = re.sub(r' {2,}', ' ', text)
    return text


def remove_blank_paragraphs(blocks):
    """删除连续空段（保留一个空行）。"""
    out = []
    prev_blank = False
    for b in blocks:
        if isinstance(b, list):
            out.append(b); prev_blank = False; continue
        t = (b or '').strip()
        if not t:
            if not prev_blank:
                out.append('')
            prev_blank = True
        else:
            out.append(b); prev_blank = False
    while out and isinstance(out[0], str) and not out[0].strip():
        out.pop(0)
    while out and isinstance(out[-1], str) and not out[-1].strip():
        out.pop()
    return out


def fix_numbering(blocks, classified):
    """根据 classified 重排 h1/h2/h3 编号。"""
    cn = '一二三四五六七八九十'
    c1 = c2 = c3 = 0
    for i, b in enumerate(blocks):
        if not isinstance(b, str): continue
        text = b.strip()
        if not text: continue
        typ = classified.get(i, 'body')
        if typ == 'h1':
            c1 += 1; c2 = c3 = 0
            text = re.sub(r'^[一二三四五六七八九十]+\s*[、.．]?\s*', '', text)
            blocks[i] = f'{cn[min(c1-1,9)]}、{text}'
        elif typ == 'h2':
            c2 += 1; c3 = 0
            text = re.sub(r'^\d+\s*[.．、]\s*', '', text)
            blocks[i] = f'{c2}. {text}'
        elif typ == 'h3':
            c3 += 1
            text = re.sub(r'^\d+\.\d+\s*[.．、]?\s*', '', text)
            blocks[i] = f'{c1}.{c3} {text}'
    return blocks


def count_words(blocks):
    chars = paras = tables = 0
    for b in blocks:
        if isinstance(b, list):
            tables += 1; continue
        if isinstance(b, str) and b.strip():
            paras += 1
            chars += len(re.sub(r'\s', '', b))
    return chars, paras, tables


def check_numbering_gaps(blocks, classified):
    """检查 h1/h2 编号跳号。"""
    issues = []
    h1_seen = []
    cur = 0
    h2_seen = []
    for idx, b in enumerate(blocks):
        if not isinstance(b, str): continue
        t = b.strip()
        typ = classified.get(idx, 'body')
        if typ == 'h1':
            if h2_seen:
                for j, n in enumerate(h2_seen):
                    if n != j + 1:
                        issues.append(f"第{cur}章二级跳号：{j+1} vs {n}")
                        break
            h2_seen = []
            cur += 1
            m = re.match(r'^([一二三四五六七八九十]+)', t)
            if m: h1_seen.append(m.group(1))
        elif typ == 'h2':
            m = re.match(r'^(\d+)', t)
            if m: h2_seen.append(int(m.group(1)))
    cn = '一二三四五六七八九十'
    for i, n in enumerate(h1_seen):
        if i < len(cn) and n != cn[i]:
            issues.append(f"一级顺序异常：{cn[i]} vs {n}")
    return issues
