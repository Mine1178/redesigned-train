# -*- coding: utf-8 -*-
"""软件内置排版预览：生成临时 docx → 转 PDF → 渲染图片 → 弹窗滚动浏览。"""
import os, tempfile

def preview_in_app(parent, blocks, classified, template,
                   header_text="", page_num=True, with_toc=False,
                   watermark="", cover=None):
    import docx_engine
    import fitz
    import customtkinter as ctk
    from PIL import Image, ImageTk

    tmp_docx = os.path.join(tempfile.gettempdir(), "_preview_typeset.docx")
    tmp_pdf = os.path.join(tempfile.gettempdir(), "_preview_typeset.pdf")

    docx_engine.build_document(
        blocks, classified, template, tmp_docx,
        header_text=header_text, page_num=page_num,
        with_toc=with_toc, watermark=watermark, cover=cover)

    # Word 转 PDF
    import win32com.client as win32
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(os.path.abspath(tmp_docx), ReadOnly=True)
    doc.SaveAs(os.path.abspath(tmp_pdf), FileFormat=17)
    doc.Close(False)
    word.Quit()

    # 弹窗
    win = ctk.CTkToplevel(parent)
    win.title("排版预览（缩放滑块控制大小）")
    win.geometry("1100x800")

    # 顶部工具栏
    bar = ctk.CTkFrame(win)
    bar.pack(fill="x", padx=8, pady=6)
    ctk.CTkLabel(bar, text="缩放:").pack(side="left", padx=6)
    scale = ctk.CTkSlider(bar, from_=0.3, to=1.5, number_of_steps=12, width=200)
    scale.set(0.8)
    scale.pack(side="left", padx=6)
    lbl_zoom = ctk.CTkLabel(bar, text="80%")
    lbl_zoom.pack(side="left", padx=6)

    # 滚动区
    canvas = ctk.CTkCanvas(win, bg="#2b2b2b", highlightthickness=0)
    sb = ctk.CTkScrollbar(win, command=canvas.yview)
    canvas.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    frame = ctk.CTkFrame(canvas, fg_color="#2b2b2b")
    win_id = canvas.create_window((0, 0), window=frame, anchor="nw")

    def render():
        for w in frame.winfo_children():
            w.destroy()
        z = float(scale.get())
        lbl_zoom.configure(text=f"{int(z*100)}%")
        pdf = fitz.open(tmp_pdf)
        for i, page in enumerate(pdf):
            mat = fitz.Matrix(z, z)
            pix = page.get_pixmap(matrix=mat)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            ctk_img = ImageTk.PhotoImage(img)
            lbl = ctk.CTkLabel(frame, image=ctk_img, text="")
            lbl.image = ctk_img
            lbl.pack(padx=20, pady=10)
            ctk.CTkLabel(frame, text=f"— 第 {i+1} 页 —",
                         text_color="#999", font=ctk.CTkFont(size=11)).pack(pady=(0,14))
        pdf.close()
        frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(win_id, width=canvas.winfo_width())

    def on_scroll(e):
        canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
    canvas.bind_all("<MouseWheel>", on_scroll)
    scale.configure(command=lambda v: render())
    win.after(100, render)
