# -*- coding: utf-8 -*-
"""配置管理：API 地址、密钥、模型等，持久化到 config.json。"""
import json
import os
import sys

# 打包成 EXE 后写到 EXE 所在目录；开发时写到脚本目录
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

DEFAULTS = {
    "api_base": "https://api.openai.com/v1",
    "api_key": "",
    "model": "gpt-4o-mini",
    "provider": "自定义",
    "temperature": 0.2,
    "timeout": 60,
    "last_dir": "",
    "recent_files": [],
}

# 服务商预设：选了自动填 base/model
PROVIDERS = {
    "通义千问":   ("https://dashscope.aliyuncs.com/compatible-mode/v1", "qwen-plus"),
    "DeepSeek":  ("https://api.deepseek.com/v1", "deepseek-chat"),
    "智谱GLM":   ("https://open.bigmodel.cn/api/paas/v4", "glm-4-flash"),
    "OpenAI":    ("https://api.openai.com/v1", "gpt-4o-mini"),
    "月之暗面":   ("https://api.moonshot.cn/v1", "moonshot-v1-8k"),
    "Ollama 本地": ("http://127.0.0.1:11434/v1", "qwen2.5:7b"),
    "vLLM 本地":  ("http://127.0.0.1:8000/v1", "Qwen/Qwen2.5-7B-Instruct"),
    "自定义":     ("", ""),
}


def load_config() -> dict:
    cfg = dict(DEFAULTS)
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                saved = json.load(f)
            cfg.update({k: v for k, v in saved.items() if k in DEFAULTS})
        except Exception:
            pass
    return cfg


def save_config(cfg: dict) -> None:
    data = load_config()
    data.update({k: v for k, v in cfg.items() if k in DEFAULTS})
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
