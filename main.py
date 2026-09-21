# -*- coding: utf-8 -*-
"""智排 AI · 文档智能排版工具
- 首页选择文档类型（公文/论文/合同/报告/出版/范文）
- 本地智能排版（启发式）+ AI 智能排版（OpenAI 兼容 API）
- 设置页配置 API；导出 .docx
"""
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

import config_manager
import ai_client
import docx_engine
from templates import TEMPLATES, HOME_CARDS, TYPE_CN

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

PRIMARY = "#2B6CE6"
PRIMARY_HOVER = "#1E56C4"
SIDEBAR_BG = "#102A4C"
CARD_BG = "#ffffff"
TEXT_DARK = "#1f2937"


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("智排 AI · 文档智能排版工具")
        self.geometry("1160x740")
        self.minsize(1000, 660)

        self.cfg = config_manager.load_config()
        self.blocks = []          # str=段落 / list[list]=表格
        self.classified = {}
        self.template_id = "general"
        self.template = TEMPLATES["general"]
        self.source_path = None

        self._build_sidebar()
        self.container = ctk.CTkFrame(self, fg_color="#f2f5fa")
        self.container.pack(side="right", fill="both", expand=True)

        self.frames = {}
        self._build_home()
        self._build_workspace()
        self._build_settings()
        self.show_frame("home")

    # ---------------- 侧边栏 ----------------
    def _build_sidebar(self):
        bar = ctk.CTkFrame(self, width=190, fg_color=SIDEBAR_BG, corner_radius=0)
        bar.pack(side="left", fill="y")
        bar.pack_propagate(False)

        ctk.CTkLabel(bar, text="智排 AI", font=ctk.CTkFont(size=22, weight="bold"),
                     text_color="#ffffff").pack(pady=(28, 4))
        ctk.CTkLabel(bar, text="智能文档排版", font=ctk.CTkFont(size=12),
                     text_color="#9db4d4").pack(pady=(0, 24))

        self.nav_home = self._nav_btn(bar, "🏠  首页", lambda: self.show_frame("home"))
        self.nav_work = self._nav_btn(bar, "📝  排版工作区", lambda: self.show_frame("work"))
        self.nav_set = self._nav_btn(bar, "⚙  设置", lambda: self.show_frame("settings"))

        ctk.CTkLabel(bar, text="v1.0  ·  中网华通格式内置",
                     font=ctk.CTkFont(size=11),
                     text_color="#6b86ad").pack(side="bottom", pady=14)

    def _nav_btn(self, parent, text, cmd):
        b = ctk.CTkButton(parent, text=text, command=cmd,
                          fg_color="transparent", hover_color="#1c3d6b",
                          anchor="w", corner_radius=8, height=40,
                          font=ctk.CTkFont(size=14), text_color="#dce6f5")
        b.pack(fill="x", padx=12, pady=2)
        return b

    # ---------------- 首页 ----------------
    def _build_home(self):
        f = ctk.CTkFrame(self.container, fg_color="#f2f5fa")
        self.frames["home"] = f
        f.grid_columnconfigure((0, 1), weight=1)
        f.grid_rowconfigure(0, minsize=60)

        ctk.CTkLabel(f, text="请选择文档类型",
                     font=ctk.CTkFont(size=26, weight="bold"),
                     text_color=TEXT_DARK).grid(row=0, column=0, columnspan=2,
                                                pady=(28, 6))
        ctk.CTkLabel(f, text="将自动套用对应标准格式，一键修复排版混乱",
                     font=ctk.CTkFont(size=14),
                     text_color="#6b7280").grid(row=1, column=0, columnspan=2, pady=(0, 18))

        for i, (tid, name, icon) in enumerate(HOME_CARDS):
            r, c = 2 + i // 2, i % 2
            card = self._make_card(f, icon, name, tid)
            card.grid(row=r, column=c, padx=18, pady=10, sticky="nsew")

    def _make_card(self, parent, icon, name, tid):
        card = ctk.CTkFrame(parent, fg_color=PRIMARY if tid != "sample" else "#1658b3",
                            corner_radius=12, height=110)
        card.pack_propagate(False)
        lbl = ctk.CTkLabel(card, text=f"{icon}   {name}",
                           font=ctk.CTkFont(size=17, weight="bold"),
                           text_color="#ffffff")
        lbl.pack(expand=True)
        for w in (card, lbl):
            w.bind("<Button-1>", lambda e, t=tid: self._pick(t))
        return card

    def _pick(self, tid):
        if tid == "sample":
            # 导入范文排版：先选范文文件
            path = filedialog.askopenfilename(
                title="选择公司范文（.docx）",
                filetypes=[("Word 文档", "*.docx")])
            if not path:
                return
            try:
                sig = docx_engine.extract_style_signature(path)
                self.template = docx_engine.template_from_signature(
                    sig, name="范文复刻: " + os.path.basename(path))
                self.template_id = "__sample__"
                self._log(f"已学习范文样式: {os.path.basename(path)}")
            except Exception as e:
                messagebox.showerror("范文读取失败", str(e))
                return
        else:
            self.template_id = tid
            self.template = TEMPLATES[tid]
        self._log(f"已选择模板: {self.template.name}")
        self.show_frame("work")

    # ---------------- 工作区 ----------------
    def _build_workspace(self):
        f = ctk.CTkFrame(self.container, fg_color="#f2f5fa")
        self.frames["work"] = f

        # 工具选项状态
        self.opt_header = ctk.BooleanVar(value=False)
        self.opt_toc = ctk.BooleanVar(value=False)
        self.opt_watermark = ctk.StringVar(value="")
        self.opt_cover = ctk.BooleanVar(value=False)

        # 第一行：文件 + 排版
        bar1 = ctk.CTkFrame(f, fg_color="transparent")
        bar1.pack(fill="x", padx=14, pady=(12, 4))
        self._tool_btn(bar1, "📂 导入文档", self._open_doc, compact=True)
        self._tool_btn(bar1, "📚 批量", self._batch_layout, fg_color="#0891b2", hover="#0e7490", compact=True)
        self._tool_btn(bar1, "🧹 清脏数据", self._clean, fg_color="#e11d48", hover="#be123c", compact=True)
        self._tool_btn(bar1, "🧩 范文", self._open_sample, compact=True)
        self._tool_btn(bar1, "✨ 本地排版", self._run_local,
                       fg_color="#16a34a", hover="#15803d", compact=True)
        self._tool_btn(bar1, "🤖 AI 排版", self._run_ai,
                       fg_color="#7c3aed", hover="#6d28d9", compact=True)
        self._tool_btn(bar1, "🔧 原地套模板", self._inplace,
                       fg_color="#0369a1", hover="#075985", compact=True)

        # 第二行：预览 + 导出 + 工具
        bar2 = ctk.CTkFrame(f, fg_color="transparent")
        bar2.pack(fill="x", padx=14, pady=4)
        self._tool_btn(bar2, "👁 排版预览", self._preview,
                       fg_color="#0e7490", hover="#155e75", compact=True)
        self._tool_btn(bar2, "💾 导出 Word", self._export,
                       fg_color="#d97706", hover="#b45309", compact=True)
        self._tool_btn(bar2, "📄 导出 PDF", self._export_pdf,
                       fg_color="#1d4ed8", hover="#1e40af", compact=True)
        self._tool_btn(bar2, "🧷 清修订", self._strip_rev,
                       fg_color="#9333ea", hover="#7e22ce", compact=True)
        self._tool_btn(bar2, "🩺 体检", self._qc,
                       fg_color="#059669", hover="#047857", compact=True)

        # 第三行：选项开关
        bar_sw = ctk.CTkFrame(f, fg_color="transparent")
        bar_sw.pack(fill="x", padx=14, pady=(0, 8))
        ctk.CTkSwitch(bar_sw, text="页眉页脚", variable=self.opt_header,
                      font=ctk.CTkFont(size=12)).pack(side="left", padx=6)
        ctk.CTkSwitch(bar_sw, text="自动目录", variable=self.opt_toc,
                      font=ctk.CTkFont(size=12)).pack(side="left", padx=6)
        ctk.CTkSwitch(bar_sw, text="封面", variable=self.opt_cover,
                      font=ctk.CTkFont(size=12)).pack(side="left", padx=6)
        ctk.CTkLabel(bar_sw, text="水印：", font=ctk.CTkFont(size=12)).pack(side="left", padx=(12,0))
        ctk.CTkEntry(bar_sw, textvariable=self.opt_watermark, width=160,
                     placeholder_text="水印文字（留空不加水印）",
                     height=28).pack(side="left", padx=6)

        info = ctk.CTkLabel(f, text="当前模板：未选择", anchor="w",
                            font=ctk.CTkFont(size=12, weight="bold"))
        info.pack(fill="x", padx=16, pady=(4, 0))
        self.lbl_template = info

        body = ctk.CTkFrame(f, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=12, pady=8)
        body.grid_columnconfigure(0, weight=3)
        body.grid_columnconfigure(1, weight=2)
        body.grid_rowconfigure(0, weight=1)

        self.preview = ctk.CTkTextbox(body, wrap="word", font=ctk.CTkFont("微软雅黑", 13))
        self.preview.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self.preview.configure(state="disabled")
        self.preview.tag_config("title", foreground="#d97706")
        self.preview.tag_config("h", foreground="#2563eb")
        self.preview.tag_config("signoff", foreground="#0f9d58")
        self.preview.tag_config("note", foreground="#6b7280")

        self.log = ctk.CTkTextbox(body, wrap="word", font=ctk.CTkFont("微软雅黑", 12),
                                  fg_color="#ffffff")
        self.log.grid(row=0, column=1, sticky="nsew")
        self.log.configure(state="disabled")

    def _tool_btn(self, parent, text, cmd, fg_color=PRIMARY, hover=PRIMARY_HOVER, compact=False):
        h = 32 if compact else 36
        sz = 12 if compact else 13
        w = 110 if compact else 130
        b = ctk.CTkButton(parent, text=text, command=cmd, fg_color=fg_color,
                          hover_color=hover, corner_radius=8, height=h, width=w,
                          font=ctk.CTkFont(size=sz, weight="bold"))
        b.pack(side="left", padx=4, pady=2)
        return b

    # ---------------- 设置页 ----------------
    def _build_settings(self):
        f = ctk.CTkFrame(self.container, fg_color="#f2f5fa")
        self.frames["settings"] = f
        f.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(f, text="🤖 AI 服务配置", font=ctk.CTkFont(size=24, weight="bold"),
                     text_color=TEXT_DARK).grid(row=0, column=0, pady=(28, 4))
        ctk.CTkLabel(f, text="配置一次，AI 排版所有功能即可使用；兼容 OpenAI 协议，支持内网本地部署",
                     font=ctk.CTkFont(size=12), text_color="#6b7280").grid(row=1, column=0, pady=(0, 18))

        card = ctk.CTkFrame(f, fg_color="#ffffff", corner_radius=12)
        card.grid(row=2, column=0, padx=80, pady=6, sticky="ew")
        card.grid_columnconfigure(1, weight=1)

        self.var_provider = ctk.StringVar(value=self.cfg.get("provider", "自定义"))
        self.var_base = ctk.StringVar(value=self.cfg["api_base"])
        self.var_key = ctk.StringVar(value=self.cfg["api_key"])
        self.var_model = ctk.StringVar(value=self.cfg["model"])
        self.var_show_key = ctk.BooleanVar(value=False)

        # 服务商下拉
        ctk.CTkLabel(card, text="服务商：", font=ctk.CTkFont(size=13, weight="bold"),
                      anchor="w").grid(row=0, column=0, padx=(20, 10), pady=10, sticky="w")
        providers = list(config_manager.PROVIDERS.keys())
        self.cmb_provider = ctk.CTkOptionMenu(
            card, values=providers, variable=self.var_provider,
            command=self._on_provider_change, height=34)
        self.cmb_provider.grid(row=0, column=1, padx=(0, 20), pady=10, sticky="ew")

        # API 密钥 + 显示
        ctk.CTkLabel(card, text="API 密钥：", font=ctk.CTkFont(size=13, weight="bold"),
                      anchor="w").grid(row=1, column=0, padx=(20, 10), pady=10, sticky="w")
        key_row = ctk.CTkFrame(card, fg_color="transparent")
        key_row.grid(row=1, column=1, padx=(0, 20), pady=10, sticky="ew")
        key_row.grid_columnconfigure(0, weight=1)
        self.ent_key = ctk.CTkEntry(key_row, textvariable=self.var_key, height=34, show="*",
                                     placeholder_text="sk-...")
        self.ent_key.grid(row=0, column=0, sticky="ew")
        ctk.CTkCheckBox(key_row, text="显示", variable=self.var_show_key,
                        command=self._toggle_key_show, width=60).grid(row=0, column=1, padx=(8, 0))

        self._field(card, 2, "接口地址：", self.var_base,
                    "如 https://api.deepseek.com/v1 或 http://服务器IP:11434/v1")
        self._field(card, 3, "模型名称：", self.var_model,
                    "如 deepseek-chat / qwen2.5:7b")

        ctk.CTkLabel(card, text="💡 通义千问 bailian.console.aliyun.com 开通即送额度；本地部署见下方教程",
                     font=ctk.CTkFont(size=11), text_color="#2563eb").grid(
            row=4, column=0, columnspan=2, pady=(0, 6))

        btns = ctk.CTkFrame(card, fg_color="transparent")
        btns.grid(row=5, column=0, columnspan=2, pady=14)
        ctk.CTkButton(btns, text="🔌 测试连接", command=self._test_conn,
                      width=120).pack(side="left", padx=8)
        ctk.CTkButton(btns, text="💾 保存配置", command=self._save_cfg,
                      fg_color=PRIMARY, hover_color=PRIMARY_HOVER,
                      width=120).pack(side="left", padx=8)
        ctk.CTkButton(btns, text="🖥 内网部署教程", command=self._open_tutorial,
                      fg_color="#64748b", hover_color="#475569",
                      width=120).pack(side="left", padx=8)

        self.set_status = ctk.CTkLabel(card, text="", text_color="#0f9d58",
                                       font=ctk.CTkFont(size=12))
        self.set_status.grid(row=6, column=0, columnspan=2, pady=(0, 12))

    def _toggle_key_show(self):
        self.ent_key.configure(show="" if self.var_show_key.get() else "*")

    def _on_provider_change(self, name):
        base, model = config_manager.PROVIDERS.get(name, ("", ""))
        if base:
            self.var_base.set(base)
        if model:
            self.var_model.set(model)

    def _field(self, parent, row, label, var, placeholder, show=""):
        ctk.CTkLabel(parent, text=label, font=ctk.CTkFont(size=13, weight="bold"),
                     anchor="w").grid(row=row, column=0, padx=(20, 10), pady=10, sticky="w")
        ent = ctk.CTkEntry(parent, textvariable=var, height=34, show=show,
                           placeholder_text=placeholder)
        ent.grid(row=row, column=1, padx=(0, 20), pady=10, sticky="ew")

    def _open_tutorial(self):
        """内网部署教程弹窗。"""
        win = ctk.CTkToplevel(self)
        win.title("企业内网 AI 部署教程")
        win.geometry("760x620")
        win.transient(self)
        ctk.CTkLabel(win, text="🖥 企业内网 AI 部署教程",
                     font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(18, 4))
        ctk.CTkLabel(win, text="数据不出内网，AI 排版照常使用（OpenAI 兼容协议）",
                     font=ctk.CTkFont(size=12), text_color="#6b7280").pack(pady=(0, 10))
        box = ctk.CTkTextbox(win, wrap="word", font=ctk.CTkFont("微软雅黑", 13))
        box.pack(fill="both", expand=True, padx=20, pady=8)
        text = (
            "一、方案选择\n"
            "  · Ollama（推荐）：单机部署一条命令启动，支持 Windows/Linux，中小团队首选\n"
            "  · vLLM：GPU 服务器高并发部署，适合大型政企多人共用\n\n"
            "二、Ollama 部署步骤\n"
            "第1步：在内网服务器安装 Ollama\n"
            "  Windows：官网 ollama.com 下载安装包（可离线拷贝）\n"
            "  Linux：curl -fsSL https://ollama.com/install.sh | sh\n\n"
            "第2步：拉取模型（8G 显存或 16G 内存可跑 7B）\n"
            "  ollama pull qwen2.5:7b\n"
            "  其他推荐：qwen2.5:14b（更佳）/ deepseek-r1:7b / glm4:9b\n\n"
            "第3步：启动服务（默认端口 11434）\n"
            "  ollama serve\n"
            "  允许局域网访问：新建环境变量 OLLAMA_HOST=0.0.0.0\n\n"
            "第4步：回软件「AI 服务配置」填写\n"
            "  服务商：Ollama 本地（自动填好）\n"
            "  接口地址：http://服务器IP:11434/v1\n"
            "  API 密钥：任意填写（如 ollama）\n"
            "  模型名称：qwen2.5:7b（与 pull 的模型名完全一致）\n"
            "  点「测试连接」显示成功即对接完成。\n\n"
            "三、vLLM 部署（GPU 服务器）\n"
            "  pip install vllm\n"
            "  vllm serve Qwen/Qwen2.5-7B-Instruct --port 8000\n"
            "  接口地址填 http://服务器IP:8000/v1\n\n"
            "四、常见问题\n"
            "  · 连接失败：检查防火墙放行端口（11434 或 8000），浏览器访问 http://服务器IP:端口/v1/models 验证\n"
            "  · 显存不足：换小模型（7b→3b）或选量化版（q4）\n"
            "  · 完全离线：先在有网机器 ollama pull，再把 ~/.ollama/models 拷到内网\n"
            "  · 模型怎么选：8G 显存选 7B，24G 选 14B/32B，纯 CPU 选 3B 量化版\n"
        )
        box.insert("end", text)
        box.configure(state="disabled")
        ctk.CTkButton(win, text="关 闭", width=160, height=40,
                      command=win.destroy, fg_color=PRIMARY,
                      hover_color=PRIMARY_HOVER).pack(pady=14)

    # ---------------- 页面切换 ----------------
    def show_frame(self, name):
        for k, v in self.frames.items():
            v.pack_forget()
        self.frames[name].pack(fill="both", expand=True)
        if name == "work":
            self.lbl_template.configure(
                text=f"当前模板：{self.template.name}　|　{self.template.desc}")

    # ---------------- 业务动作 ----------------
    def _open_doc(self):
        path = filedialog.askopenfilename(
            title="选择要排版的文档",
            filetypes=[("支持格式", "*.docx *.txt"), ("Word", "*.docx"), ("文本", "*.txt")])
        if not path:
            return
        try:
            self.blocks = docx_engine.extract_paragraphs(path)
            self.source_path = path
            self.classified = {}
            n_para = sum(1 for b in self.blocks if isinstance(b, str))
            n_tab = sum(1 for b in self.blocks if isinstance(b, list))
            self._log(f"已载入 {os.path.basename(path)}：{n_para} 段、{n_tab} 个表格。")
            # 记录最近文件
            rf = [path] + [f for f in self.cfg.get("recent_files", []) if f != path][:5]
            self.cfg["recent_files"] = rf
            config_manager.save_config(self.cfg)
            self._render_preview()
        except Exception as e:
            messagebox.showerror("打开失败", str(e))

    def _open_sample(self):
        self._pick("sample")

    # ---------- 第一波新功能 ----------
    def _clean(self):
        if not self.blocks:
            messagebox.showinfo("提示", "请先导入文档。")
            return
        self.blocks, report = docx_engine.clean_blocks(self.blocks)
        self.classified = {}
        for r in report:
            self._log("🧹 " + r)
        if not report:
            self._log("🧹 未发现脏数据。")
        self._render_preview()

    def _batch_layout(self):
        """批量：多选 docx，逐个套用当前模板后导出到同目录 _已排版.docx。"""
        files = filedialog.askopenfilenames(
            title="选择多个要排版的 Word 文档",
            filetypes=[("Word 文档", "*.docx")])
        if not files:
            return
        ok, fail = 0, 0
        for fp in files:
            try:
                blocks = docx_engine.extract_paragraphs(fp)
                cls = docx_engine.local_classify(blocks)
                out = os.path.splitext(fp)[0][0] + "_已排版.docx"
                docx_engine.build_document(
                    blocks, cls, self.template, out,
                    header_text="" if not self.opt_header.get() else "",
                    page_num=self.opt_header.get(),
                    with_toc=self.opt_toc.get(),
                    watermark=self.opt_watermark.get().strip())
                ok += 1
                self._log(f"✅ {os.path.basename(fp)} -> {os.path.basename(out)}")
            except Exception as e:
                fail += 1
                self._log(f"❌ {os.path.basename(fp)}: {e}")
        messagebox.showinfo("批量完成", f"成功 {ok} 个，失败 {fail} 个。")

    def _qc(self):
        if not self.blocks:
            messagebox.showinfo("提示", "请先导入文档。")
            return
        if not self.classified:
            self.classified = docx_engine.local_classify(self.blocks)
        issues = docx_engine.qc_check(self.blocks, self.classified, self.template)
        self._log("—— 定稿体检报告 ——")
        for level, msg in issues:
            self._log(f"[{level}] {msg}")

    def _strip_rev(self):
        if not self.source_path or not self.source_path.lower().endswith(".docx"):
            messagebox.showinfo("提示", "请先导入 .docx 文档。")
            return
        out = os.path.splitext(self.source_path)[0] + "_无修订.docx"
        docx_engine.strip_revisions(self.source_path, out)
        self._log(f"🧷 已生成无修订版本: {out}")
        messagebox.showinfo("完成", f"已生成：\n{out}")

    def _run_local(self):
        if not self.blocks:
            messagebox.showinfo("提示", "请先导入文档。")
            return
        self.classified = docx_engine.local_classify(self.blocks)
        self._log(f"本地启发式分类完成，覆盖 {len(self.classified)} 段。")
        self._render_preview()

    def _run_ai(self):
        if not self.blocks:
            messagebox.showinfo("提示", "请先导入文档。")
            return
        if not self.cfg["api_key"]:
            messagebox.showwarning("未配置 API", "请先到「设置」页填写 API Key。")
            self.show_frame("settings")
            return

        # 只把段落文本送给 AI，表格块不参与分类
        para_idx = [i for i, b in enumerate(self.blocks) if isinstance(b, str)]
        para_texts = [self.blocks[i] for i in para_idx]

        def work():
            try:
                base = docx_engine.local_classify(self.blocks)
                ai_flat = ai_client.classify_paragraphs(
                    self.cfg, para_texts, self.template_id, on_log=self._log)
                # 把 AI 的扁平序号映射回块序号
                for flat_i, ptype in ai_flat.items():
                    if 0 <= flat_i < len(para_idx):
                        base[para_idx[flat_i]] = ptype
                self.classified = base
                self.after(0, lambda: (self._render_preview(),
                                       self._log("✅ AI 排版完成，可导出 Word。")))
            except Exception as e:
                self.after(0, lambda: (messagebox.showerror("AI 调用失败", str(e)),
                                       self._log(f"❌ 失败: {e}")))
        threading.Thread(target=work, daemon=True).start()
        self._log("AI 排版进行中（后台线程）...")

    def _export(self):
        if not self.blocks:
            messagebox.showinfo("提示", "请先导入并排版文档。")
            return
        if not self.classified:
            self._run_local()
        default_name = os.path.splitext(os.path.basename(self.source_path or "document"))[0] \
            + "_已排版.docx"
        path = filedialog.asksaveasfilename(
            defaultextension=".docx", initialfile=default_name,
            filetypes=[("Word 文档", "*.docx")])
        if not path:
            return
        try:
            cover = None
            if self.opt_cover.get() and self.blocks:
                title = next((b for b in self.blocks
                              if isinstance(b, str) and b.strip()), "工程名称")
                cover = {"project_name": title, "book_no": "第X册 第X分册"}
            hdr = docx_engine.guess_project_name(self.blocks)
            docx_engine.build_document(
                self.blocks, self.classified, self.template, path,
                header_text=hdr,
                page_num=self.opt_header.get(),
                with_toc=self.opt_toc.get(),
                watermark=self.opt_watermark.get().strip(),
                cover=cover)
            self.last_out = path
            self._log(f"✅ 已导出: {path}")
            messagebox.showinfo("完成", f"已导出：\n{path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def _export_pdf(self):
        """先排好版存临时 docx，再用 Word 转 PDF。"""
        if not self.blocks:
            messagebox.showinfo("提示", "请先导入并排版文档。")
            return
        if not self.classified:
            self._run_local()
        import tempfile
        tmp_docx = os.path.join(tempfile.gettempdir(), "_layout_for_pdf.docx")
        cover = None
        if self.opt_cover.get() and self.blocks:
            title = next((b for b in self.blocks if isinstance(b, str) and b.strip()), "工程名称")
            cover = {"project_name": title, "book_no": "第X册 第X分册"}
        hdr = docx_engine.guess_project_name(self.blocks)
        docx_engine.build_document(
            self.blocks, self.classified, self.template, tmp_docx,
            header_text=hdr,
            page_num=self.opt_header.get(),
            with_toc=self.opt_toc.get(),
            watermark=self.opt_watermark.get().strip(),
            cover=cover)
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=os.path.splitext(
                os.path.basename(self.source_path or "document"))[0] + ".pdf")
        if not path:
            return
        try:
            docx_engine.docx_to_pdf(tmp_docx, path)
            self._log(f"✅ 已导出 PDF: {path}")
            messagebox.showinfo("完成", f"已导出：\n{path}")
        except Exception as e:
            messagebox.showerror("导出 PDF 失败", str(e))

    def _preview(self):
        """用 Word 只读模式打开最近导出的文档，预览排版效果。"""
        path = self.last_out
        if not path or not os.path.exists(path):
            path = self.source_path
        if not path or not os.path.exists(path):
            path = filedialog.askopenfilename(
                title="选择要预览的 docx",
                filetypes=[("Word", "*.docx")])
            if not path:
                return
        try:
            import win32com.client as win32
            word = win32.Dispatch("Word.Application")
            word.Visible = True
            word.WindowState = 1  # 最大化
            doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True,
                                       AddToRecentFiles=False)
            doc.ActiveWindow.View.Type = 3  # 页面视图
            self._log("✅ 已在 Word 中打开预览")
        except Exception as e:
            messagebox.showerror("预览失败", str(e))

    def _inplace(self):
        """原地套用华通样式：不重建，保留原文档图片/公式。"""
        path = filedialog.askopenfilename(
            title="选择要原地套华通样式的文档",
            filetypes=[("Word", "*.docx")])
        if not path:
            return
        out = os.path.splitext(path)[0] + "_华通排版.docx"
        self._log("🔧 正在原地套用华通样式...")
        try:
            import inplace
            out = os.path.splitext(path)[0] + "_" + self.template.tid + ".docx"
            inplace.apply_template_inplace(path, self.template, out, header_text="")
            self._log(f"✅ 已生成: {out}")
            messagebox.showinfo("完成", f"已生成：\n{out}")
        except Exception as e:
            messagebox.showerror("失败", str(e))

    # ---------------- 设置页动作 ----------------
    def _test_conn(self):
        cfg = self._collect_cfg()
        self.set_status.configure(text="测试中...", text_color="#6b7280")
        self.update()

        def work():
            ok, msg = ai_client.test_connection(cfg)
            color = "#0f9d58" if ok else "#dc2626"
            self.after(0, lambda: self.set_status.configure(text=msg, text_color=color))
        threading.Thread(target=work, daemon=True).start()

    def _save_cfg(self):
        self.cfg.update(self._collect_cfg())
        config_manager.save_config(self.cfg)
        self.set_status.configure(text="已保存 ✓", text_color="#0f9d58")
        self._log("API 配置已保存。")

    def _collect_cfg(self):
        return {"provider": self.var_provider.get(),
                "api_base": self.var_base.get().strip(),
                "api_key": self.var_key.get().strip(),
                "model": self.var_model.get().strip()}

    # ---------------- 预览与日志 ----------------
    def _render_preview(self):
        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")
        for i, block in enumerate(self.blocks):
            if isinstance(block, list):
                nr, nc = len(block), max((len(r) for r in block), default=0)
                self.preview.insert("end", f"【表格】 {nr} 行 × {nc} 列（已按公司规范重排）\n", "note")
                continue
            text = block.strip()
            if not text:
                continue
            ptype = self.classified.get(i)
            tag = None
            if ptype in ("title", "subtitle"):
                tag = "title"
            elif ptype and ptype.startswith("h"):
                tag = "h"
            elif ptype == "signoff":
                tag = "signoff"
            elif ptype in ("note", "doc_no", "table_caption", "figure_caption"):
                tag = "note"
            cn = TYPE_CN.get(ptype, "正文" if ptype == "body" else "未排")
            line = f"【{cn}】 {text}\n"
            self.preview.insert("end", line, tag)
        self.preview.configure(state="disabled")

    def _log(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")


if __name__ == "__main__":
    app = App()
    app.mainloop()
