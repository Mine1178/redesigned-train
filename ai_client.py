# -*- coding: utf-8 -*-
"""AI 排版客户端：OpenAI 兼容 Chat Completions 接口。
负责把纯文本段落交给大模型做语义级分类（标题/正文/落款等）。
"""
import json
import re
import requests

from templates import TEMPLATES

SYS_PROMPT = """你是一名专业中文文档排版引擎。用户会给你一份文档的逐段文本（每行一段，行首为【i】序号）。
你的任务：只做段落类型判断，不改正文、不增删内容。
逐段输出 JSON 数组，每个元素形如 {"i": 序号, "t": 类型}。
类型只能取：title(文档大标题,通常是第一段短行居中题), doc_no(发文字号,如XX〔2023〕5号),
h1(一级标题), h2(二级标题), h3(三级标题), h4(四级标题),
body(正文段落), recipient(主送机关,顶格带冒号如"各部门："),
signoff(落款,文末单位名/日期), note(注释/图注/说明文字), subtitle(副标题)。
判定参考：
- 行首为"一、""二、"或"第X章/第X节/第X条"→ h1；"（一）"或"1.1"→ h2；"1."→ h3；"(1)"→ h4。
- 技术文件多级编号 1 / 1.1 / 1.1.1 / 1.1.1.1 分别对应 h1/h2/h3/h4。
- 文末出现"特此通知/特此函告/公司名/YYYY年M月D日"等 → signoff。
- 其余成段叙述 → body。
只输出 JSON 数组本身，不要解释、不要 markdown 代码块。"""


def _build_user_prompt(paragraphs, template_id):
    tpl = TEMPLATES.get(template_id)
    tname = tpl.name if tpl else "通用文档"
    lines = []
    for i, p in enumerate(paragraphs):
        p = p.strip()
        if p:
            lines.append(f"【{i}】{p}")
    head = f"文档类型：{tname}\n共 {len(lines)} 个非空段落，逐段分类如下：\n"
    return head + "\n".join(lines)


def classify_paragraphs(cfg, paragraphs, template_id, on_log=lambda m: None):
    """调用 AI 分类段落。
    返回 dict: {原段落序号: 类型}；失败抛异常。
    """
    base = cfg["api_base"].rstrip("/")
    url = base + "/chat/completions"
    headers = {
        "Authorization": f"Bearer {cfg['api_key']}",
        "Content-Type": "application/json",
    }
    body = {
        "model": cfg["model"],
        "temperature": cfg["temperature"],
        "messages": [
            {"role": "system", "content": SYS_PROMPT},
            {"role": "user", "content": _build_user_prompt(paragraphs, template_id)},
        ],
    }
    on_log(f"调用模型 {cfg['model']} ...")
    r = requests.post(url, headers=headers, json=body,
                     timeout=(10, max(90, cfg.get("timeout", 60))))
    r.raise_for_status()
    data = r.json()
    text = data["choices"][0]["message"]["content"].strip()
    # 容忍代码块
    m = re.search(r"(\[.*\])", text, re.S)
    if m:
        text = m.group(1)
    arr = json.loads(text)
    result = {}
    for item in arr:
        try:
            idx = int(item["i"])
            result[idx] = str(item["t"]).strip()
        except (KeyError, ValueError, TypeError):
            continue
    on_log(f"AI 分类完成，覆盖 {len(result)} 段。")
    return result


def test_connection(cfg):
    """测试 API 连通性与密钥有效性，返回 (ok, message)。"""
    if not cfg["api_key"]:
        return False, "未填写 API Key"
    try:
        url = cfg["api_base"].rstrip("/") + "/chat/completions"
        r = requests.post(
            url,
            headers={"Authorization": f"Bearer {cfg['api_key']}",
                     "Content-Type": "application/json"},
            json={"model": cfg["model"],
                  "messages": [{"role": "user", "content": "ping"}],
                  "max_tokens": 5},
            timeout=(10, 90),
        )
        if r.status_code == 200:
            return True, "连接成功，模型可用。"
        return False, f"HTTP {r.status_code}: {r.text[:200]}"
    except Exception as e:
        return False, f"连接失败: {e}"
