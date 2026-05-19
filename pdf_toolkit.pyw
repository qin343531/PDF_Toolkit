#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF 工具箱 — 免WPS会员，免费使用
功能：删除空白页 | 删除指定页 | 旋转页面 | 合并拆分 | 提取压缩 | 加水印
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import re
import sys
import threading
import tempfile
import shutil
from pathlib import Path
from pypdf import PdfReader, PdfWriter, PageObject, Transformation

def get_script_dir():
    """获取脚本/exe所在目录（兼容PyInstaller打包）"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


class PDFToolkit:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 工具箱 — 免WPS会员")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # 当前打开的文件
        self.current_file = None
        self.current_reader = None
        # 默认输出到同级 output 目录
        script_dir = get_script_dir()
        self.output_dir = os.path.join(script_dir, "output")
        os.makedirs(self.output_dir, exist_ok=True)
        
        self.setup_ui()
    
    def setup_ui(self):
        """搭建界面"""
        # 顶部工具栏
        toolbar = ttk.Frame(self.root, padding=5)
        toolbar.pack(fill=tk.X)
        
        ttk.Button(toolbar, text="📂 打开PDF", command=self.open_pdf, width=15).pack(side=tk.LEFT, padx=3)
        ttk.Button(toolbar, text="📁 输出目录", command=self.set_output_dir, width=15).pack(side=tk.LEFT, padx=3)
        self.file_label = ttk.Label(toolbar, text="未打开文件", foreground="gray")
        self.file_label.pack(side=tk.LEFT, padx=10)
        self.page_label = ttk.Label(toolbar, text="", foreground="blue")
        self.page_label.pack(side=tk.LEFT, padx=10)
        
        ttk.Separator(self.root, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=3)
        
        # 第二行：输出目录
        row2 = ttk.Frame(self.root, padding=3)
        row2.pack(fill=tk.X)
        ttk.Label(row2, text="📂 输出:", foreground="gray").pack(side=tk.LEFT)
        self.output_label = ttk.Label(row2, text=self.output_dir, foreground="darkgreen")
        self.output_label.pack(side=tk.LEFT, padx=5)
        
        ttk.Separator(self.root, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=3)
        
        # Notebook 选项卡
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # === 标签1：页面处理 ===
        tab_page = ttk.Frame(notebook, padding=10)
        notebook.add(tab_page, text="📄 页面处理")
        self.build_page_tab(tab_page)
        
        # === 标签2：合并拆分 ===
        tab_merge = ttk.Frame(notebook, padding=10)
        notebook.add(tab_merge, text="🔗 合并拆分")
        self.build_merge_tab(tab_merge)
        
        # === 标签3：工具集 ===
        tab_tools = ttk.Frame(notebook, padding=10)
        notebook.add(tab_tools, text="🔧 更多工具")
        self.build_tools_tab(tab_tools)
        
        # 底部状态栏
        self.status = ttk.Label(self.root, text="就绪", relief=tk.SUNKEN, anchor=tk.W, padding=3)
        self.status.pack(fill=tk.X, side=tk.BOTTOM)
    
    # ==================== 页面处理标签 ====================
    def build_page_tab(self, parent):
        # 左半部分
        left = ttk.Frame(parent)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # ---- 删除空白页 ----
        frm1 = ttk.LabelFrame(left, text="🗑️ 删除空白页", padding=10)
        frm1.pack(fill=tk.X, pady=5)
        
        ttk.Label(frm1, text="判定阈值（字符数低于此值视为空白）:").pack(anchor=tk.W)
        self.blank_threshold = tk.IntVar(value=20)
        ttk.Spinbox(frm1, from_=0, to=500, textvariable=self.blank_threshold, width=8).pack(anchor=tk.W, pady=3)
        
        btn_frm = ttk.Frame(frm1)
        btn_frm.pack(fill=tk.X, pady=5)
        ttk.Button(btn_frm, text="🔍 扫描空白页", command=self.scan_blank).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frm, text="🗑️ 一键删除全部空白页", command=self.remove_all_blank).pack(side=tk.LEFT, padx=3)
        
        self.blank_list = scrolledtext.ScrolledText(frm1, height=6, width=40)
        self.blank_list.pack(fill=tk.X, pady=5)
        
        # ---- 删除指定页面 ----
        frm2 = ttk.LabelFrame(left, text="📑 删除指定页面", padding=10)
        frm2.pack(fill=tk.X, pady=5)
        
        ttk.Label(frm2, text="页码范围（如: 7  或  3-5  或  1,3,7）:").pack(anchor=tk.W)
        self.delete_range = ttk.Entry(frm2, width=30)
        self.delete_range.pack(anchor=tk.W, pady=3)
        ttk.Button(frm2, text="🗑️ 删除指定页", command=self.delete_pages).pack(anchor=tk.W, pady=5)
        
        # ---- 提取页面 ----
        frm5 = ttk.LabelFrame(left, text="📋 提取页面（保存为新文件）", padding=10)
        frm5.pack(fill=tk.X, pady=5)
        
        ttk.Label(frm5, text="页码范围（如: 1-10）:").pack(anchor=tk.W)
        self.extract_range = ttk.Entry(frm5, width=30)
        self.extract_range.pack(anchor=tk.W, pady=3)
        ttk.Button(frm5, text="📋 提取为新PDF", command=self.extract_pages).pack(anchor=tk.W, pady=5)
        
        # 右半部分
        right = ttk.Frame(parent)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # ---- 旋转页面 ----
        frm3 = ttk.LabelFrame(right, text="🔄 旋转页面", padding=10)
        frm3.pack(fill=tk.X, pady=5)
        
        ttk.Label(frm3, text="旋转角度:").pack(anchor=tk.W)
        rot_frm = ttk.Frame(frm3)
        rot_frm.pack(fill=tk.X, pady=3)
        self.rotation_angle = tk.IntVar(value=90)
        ttk.Radiobutton(rot_frm, text="顺时针90°", variable=self.rotation_angle, value=90).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(rot_frm, text="逆时针90°", variable=self.rotation_angle, value=270).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(rot_frm, text="180°", variable=self.rotation_angle, value=180).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(frm3, text="页面范围（留空=全部）:").pack(anchor=tk.W, pady=(5, 0))
        self.rotate_range = ttk.Entry(frm3, width=25)
        self.rotate_range.pack(anchor=tk.W, pady=3)
        
        ttk.Separator(frm3, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=5)
        ttk.Label(frm3, text="🤖 智能旋转（自动识别横/竖版）:").pack(anchor=tk.W)
        smart_frm = ttk.Frame(frm3)
        smart_frm.pack(fill=tk.X, pady=3)
        self.smart_direction = tk.StringVar(value="portrait_to_landscape")
        ttk.Radiobutton(smart_frm, text="竖版→横版", variable=self.smart_direction, value="portrait_to_landscape").pack(side=tk.LEFT, padx=3)
        ttk.Radiobutton(smart_frm, text="横版→竖版", variable=self.smart_direction, value="landscape_to_portrait").pack(side=tk.LEFT, padx=3)
        
        btn_frm3 = ttk.Frame(frm3)
        btn_frm3.pack(fill=tk.X, pady=5)
        ttk.Button(btn_frm3, text="🔄 手动旋转", command=self.rotate_pages).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frm3, text="🤖 智能旋转", command=self.smart_rotate).pack(side=tk.LEFT, padx=3)
        
        # ---- 页面排序 ----
        frm4 = ttk.LabelFrame(right, text="🔢 页面排序（重新排列）", padding=10)
        frm4.pack(fill=tk.X, pady=5)
        
        ttk.Label(frm4, text="新顺序（如: 3,1,2,5,4  留空=反转全部）:").pack(anchor=tk.W)
        self.reorder_seq = ttk.Entry(frm4, width=30)
        self.reorder_seq.pack(anchor=tk.W, pady=3)
        btn_frm4 = ttk.Frame(frm4)
        btn_frm4.pack(fill=tk.X, pady=5)
        ttk.Button(btn_frm4, text="🔢 重新排序", command=self.reorder_pages).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frm4, text="↔️ 反转页面顺序", command=self.reverse_pages).pack(side=tk.LEFT, padx=3)
    
    # ==================== 合并拆分标签 ====================
    def build_merge_tab(self, parent):
        # 左：合并
        left = ttk.Frame(parent)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        frm1 = ttk.LabelFrame(left, text="🔗 合并多个PDF", padding=10)
        frm1.pack(fill=tk.X, pady=5)
        
        ttk.Label(frm1, text="按顺序合并，点击按钮添加文件:").pack(anchor=tk.W)
        
        self.merge_files = []
        self.merge_listbox = tk.Listbox(frm1, height=6, width=40)
        self.merge_listbox.pack(fill=tk.X, pady=5)
        
        btn_frm1 = ttk.Frame(frm1)
        btn_frm1.pack(fill=tk.X, pady=3)
        ttk.Button(btn_frm1, text="➕ 添加文件", command=self.add_merge_file).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frm1, text="➖ 移除选中", command=self.remove_merge_file).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frm1, text="⬆ 上移", command=lambda: self.move_merge_item(-1)).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frm1, text="⬇ 下移", command=lambda: self.move_merge_item(1)).pack(side=tk.LEFT, padx=3)
        
        ttk.Button(frm1, text="🔗 开始合并", command=self.merge_pdfs).pack(anchor=tk.W, pady=5)
        
        # 右：拆分
        right = ttk.Frame(parent)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        frm2 = ttk.LabelFrame(right, text="✂️ 拆分PDF", padding=10)
        frm2.pack(fill=tk.X, pady=5)
        
        ttk.Label(frm2, text="拆分方式:").pack(anchor=tk.W)
        self.split_mode = tk.StringVar(value="every")
        ttk.Radiobutton(frm2, text="每页一个文件", variable=self.split_mode, value="every").pack(anchor=tk.W, padx=10)
        ttk.Radiobutton(frm2, text="按页数拆分（每N页一个文件）", variable=self.split_mode, value="by_count").pack(anchor=tk.W, padx=10)
        ttk.Radiobutton(frm2, text="按页码范围拆分", variable=self.split_mode, value="by_range").pack(anchor=tk.W, padx=10)
        
        self.split_param = ttk.Entry(frm2, width=25)
        self.split_param.pack(anchor=tk.W, padx=10, pady=3)
        ttk.Label(frm2, text="↑ 按页数填数字(如3)，按范围填范围(如1-10,11-20)", foreground="gray", font=("", 8)).pack(anchor=tk.W, padx=10)
        
        ttk.Button(frm2, text="✂️ 开始拆分", command=self.split_pdf).pack(anchor=tk.W, padx=10, pady=5)
    
    # ==================== 更多工具标签 ====================
    def build_tools_tab(self, parent):
        # 左列
        left = ttk.Frame(parent)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # 水印
        frm1 = ttk.LabelFrame(left, text="💧 添加水印", padding=10)
        frm1.pack(fill=tk.X, pady=5)
        
        ttk.Label(frm1, text="水印文字:").pack(anchor=tk.W)
        self.watermark_text = ttk.Entry(frm1, width=30)
        self.watermark_text.insert(0, "机密文件")
        self.watermark_text.pack(anchor=tk.W, pady=3)
        
        ttk.Label(frm1, text="透明度(0-1):").pack(anchor=tk.W)
        self.watermark_opacity = tk.DoubleVar(value=0.15)
        ttk.Scale(frm1, from_=0.05, to=0.5, variable=self.watermark_opacity, orient=tk.HORIZONTAL).pack(fill=tk.X)
        
        ttk.Button(frm1, text="💧 添加水印", command=self.add_watermark).pack(anchor=tk.W, pady=5)
        
        # 压缩
        frm2 = ttk.LabelFrame(left, text="📦 压缩PDF", padding=10)
        frm2.pack(fill=tk.X, pady=5)
        
        ttk.Label(frm2, text="压缩级别（建议2-5，数字越大文件越小质量越低）:").pack(anchor=tk.W)
        self.compress_level = tk.IntVar(value=3)
        ttk.Scale(frm2, from_=1, to=9, variable=self.compress_level, orient=tk.HORIZONTAL).pack(fill=tk.X)
        ttk.Label(frm2, textvariable=self.compress_level).pack()
        
        ttk.Button(frm2, text="📦 压缩PDF", command=self.compress_pdf).pack(anchor=tk.W, pady=5)
        
        # 右列
        right = ttk.Frame(parent)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # 提取图片（基础版）
        frm3 = ttk.LabelFrame(right, text="🖼️ 提取文本", padding=10)
        frm3.pack(fill=tk.X, pady=5)
        
        ttk.Button(frm3, text="📝 提取全部文本到TXT", command=self.extract_text).pack(anchor=tk.W, pady=5)
        
        # 元信息
        frm4 = ttk.LabelFrame(right, text="ℹ️ 文档信息", padding=10)
        frm4.pack(fill=tk.X, pady=5)
        
        ttk.Button(frm4, text="ℹ️ 查看文档信息", command=self.show_doc_info).pack(anchor=tk.W, pady=3)
        ttk.Button(frm4, text="🔒 移除密码/加密", command=self.remove_encryption).pack(anchor=tk.W, pady=3)
    
    # ==================== 核心功能 ====================
    
    def open_pdf(self):
        filepath = filedialog.askopenfilename(filetypes=[("PDF文件", "*.pdf")])
        if not filepath:
            return
        try:
            self.current_file = filepath
            self.current_reader = PdfReader(filepath)
            self.file_label.config(text=os.path.basename(filepath)[:50])
            self.page_label.config(text=f"共 {len(self.current_reader.pages)} 页")
            self.status.config(text=f"已打开: {os.path.basename(filepath)}")
        except Exception as e:
            messagebox.showerror("错误", f"无法打开PDF: {e}")
    
    def set_output_dir(self):
        d = filedialog.askdirectory(initialdir=self.output_dir)
        if d:
            self.output_dir = d
            os.makedirs(d, exist_ok=True)
            self.output_label.config(text=d)
            self.status.config(text=f"输出目录: {d}")
    
    def get_output_path(self, suffix):
        """生成输出文件路径"""
        base = os.path.splitext(os.path.basename(self.current_file or "output"))[0]
        d = self.output_dir or os.path.dirname(self.current_file) if self.current_file else "."
        return os.path.join(d, f"{base}_{suffix}.pdf")
    
    def run_async(self, func, *args, **kwargs):
        """异步执行避免界面卡顿"""
        self.status.config(text="处理中...")
        t = threading.Thread(target=lambda: self._run_safe(func, *args, **kwargs))
        t.daemon = True
        t.start()
    
    def _run_safe(self, func, *args, **kwargs):
        try:
            func(*args, **kwargs)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
            self.root.after(0, lambda: self.status.config(text="出错"))
    
    # ---- 扫描空白页 ----
    def scan_blank(self):
        if not self.current_reader:
            messagebox.showwarning("提示", "请先打开PDF文件")
            return
        
        threshold = self.blank_threshold.get()
        blank_pages = []
        total = len(self.current_reader.pages)
        
        for i, page in enumerate(self.current_reader.pages):
            text = (page.extract_text() or "").strip()
            non_ws = re.sub(r'\s+', '', text)
            if len(non_ws) < threshold:
                blank_pages.append(i + 1)
        
        self.blank_list.delete(1.0, tk.END)
        if blank_pages:
            self.blank_list.insert(tk.END, f"找到 {len(blank_pages)} 个空白页:\n")
            self.blank_list.insert(tk.END, f"页码: {', '.join(map(str, blank_pages))}\n\n")
            self.blank_list.insert(tk.END, f"点击「一键删除全部空白页」即可清除")
            self.blank_pages_found = blank_pages
        else:
            self.blank_list.insert(tk.END, f"✅ 未发现空白页（阈值={threshold}字符）")
            self.blank_pages_found = []
        
        self.status.config(text=f"扫描完成，{len(blank_pages)}个空白页")
    
    def remove_all_blank(self):
        if not self.current_reader or not hasattr(self, 'blank_pages_found') or not self.blank_pages_found:
            messagebox.showwarning("提示", "请先扫描空白页")
            return
        
        self.remove_pages_by_list(self.blank_pages_found, "去空白页")
    
    # ---- 删除指定页面 ----
    def delete_pages(self):
        if not self.current_reader:
            messagebox.showwarning("提示", "请先打开PDF文件")
            return
        
        pages = self._parse_page_range(self.delete_range.get())
        if not pages:
            messagebox.showwarning("提示", "请输入有效的页码范围")
            return
        
        self.remove_pages_by_list(pages, "删页")
    
    def remove_pages_by_list(self, pages_to_remove, suffix):
        """通用：删除指定页码列表"""
        writer = PdfWriter()
        total = len(self.current_reader.pages)
        remove_set = set(pages_to_remove)
        removed = []
        
        for i in range(total):
            if (i + 1) not in remove_set:
                writer.add_page(self.current_reader.pages[i])
            else:
                removed.append(i + 1)
        
        out = self.get_output_path(suffix)
        writer.write(out)
        self._done(f"已删除第 {', '.join(map(str, removed))} 页 ({len(removed)}页)\n保存至: {out}")
    
    # ---- 提取页面 ----
    def extract_pages(self):
        if not self.current_reader:
            messagebox.showwarning("提示", "请先打开PDF文件")
            return
        
        pages = self._parse_page_range(self.extract_range.get())
        if not pages:
            messagebox.showwarning("提示", "请输入有效的页码范围")
            return
        
        writer = PdfWriter()
        for p in pages:
            if 1 <= p <= len(self.current_reader.pages):
                writer.add_page(self.current_reader.pages[p - 1])
        
        out = self.get_output_path("提取")
        writer.write(out)
        self._done(f"已提取 {len(pages)} 页\n保存至: {out}")
    
    # ---- 旋转页面 ----
    def rotate_pages(self):
        if not self.current_reader:
            messagebox.showwarning("提示", "请先打开PDF文件")
            return
        
        angle = self.rotation_angle.get()
        pages = self._parse_page_range(self.rotate_range.get())
        if not pages:
            pages = list(range(1, len(self.current_reader.pages) + 1))
        
        writer = PdfWriter()
        total = len(self.current_reader.pages)
        for i in range(total):
            page = self.current_reader.pages[i]
            if (i + 1) in pages:
                page.rotate(angle)
            writer.add_page(page)
        
        out = self.get_output_path(f"旋转{angle}°")
        writer.write(out)
        self._done(f"已旋转 {len(pages)} 页 ({angle}°)\n保存至: {out}")
    
    def smart_rotate(self):
        """智能旋转：自动检测页面方向并旋转"""
        if not self.current_reader:
            messagebox.showwarning("提示", "请先打开PDF文件")
            return
        
        direction = self.smart_direction.get()
        writer = PdfWriter()
        rotated_count = 0
        
        for i, page in enumerate(self.current_reader.pages):
            # 获取页面尺寸
            mediabox = page.mediabox
            w = float(mediabox.width)
            h = float(mediabox.height)
            
            is_landscape = w > h
            is_portrait = h > w
            
            if direction == "portrait_to_landscape" and is_portrait:
                page.rotate(90)
                rotated_count += 1
            elif direction == "landscape_to_portrait" and is_landscape:
                page.rotate(90)
                rotated_count += 1
            
            writer.add_page(page)
        
        out = self.get_output_path("智能旋转")
        writer.write(out)
        self._done(f"智能旋转: {rotated_count}/{len(self.current_reader.pages)} 页\n保存至: {out}")
    
    # ---- 页面排序 ----
    def reorder_pages(self):
        if not self.current_reader:
            messagebox.showwarning("提示", "请先打开PDF文件")
            return
        
        seq_str = self.reorder_seq.get().strip()
        if not seq_str:
            messagebox.showwarning("提示", "请输入新顺序")
            return
        
        try:
            new_order = [int(x.strip()) for x in seq_str.split(",")]
        except ValueError:
            messagebox.showwarning("提示", "格式错误，请用逗号分隔数字，如: 3,1,2,5,4")
            return
        
        total = len(self.current_reader.pages)
        if max(new_order) > total or min(new_order) < 1:
            messagebox.showwarning("提示", f"页码超出范围 (1-{total})")
            return
        
        writer = PdfWriter()
        for p in new_order:
            writer.add_page(self.current_reader.pages[p - 1])
        
        out = self.get_output_path("重排")
        writer.write(out)
        self._done(f"已按新顺序重排\n保存至: {out}")
    
    def reverse_pages(self):
        if not self.current_reader:
            messagebox.showwarning("提示", "请先打开PDF文件")
            return
        
        writer = PdfWriter()
        for i in range(len(self.current_reader.pages) - 1, -1, -1):
            writer.add_page(self.current_reader.pages[i])
        
        out = self.get_output_path("反转")
        writer.write(out)
        self._done(f"已反转页面顺序\n保存至: {out}")
    
    # ---- 合并PDF ----
    def add_merge_file(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF文件", "*.pdf")])
        for f in files:
            if f not in self.merge_files:
                self.merge_files.append(f)
                self.merge_listbox.insert(tk.END, os.path.basename(f))
    
    def remove_merge_file(self):
        sel = self.merge_listbox.curselection()
        if sel:
            idx = sel[0]
            self.merge_listbox.delete(idx)
            del self.merge_files[idx]
    
    def move_merge_item(self, direction):
        sel = self.merge_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        new_idx = idx + direction
        if 0 <= new_idx < len(self.merge_files):
            self.merge_files[idx], self.merge_files[new_idx] = self.merge_files[new_idx], self.merge_files[idx]
            self.merge_listbox.delete(idx)
            self.merge_listbox.insert(new_idx, os.path.basename(self.merge_files[new_idx]))
            self.merge_listbox.selection_set(new_idx)
    
    def merge_pdfs(self):
        if not self.merge_files:
            messagebox.showwarning("提示", "请先添加要合并的文件")
            return
        
        writer = PdfWriter()
        for f in self.merge_files:
            try:
                reader = PdfReader(f)
                for page in reader.pages:
                    writer.add_page(page)
            except Exception as e:
                messagebox.showerror("错误", f"无法读取 {os.path.basename(f)}: {e}")
                return
        
        out = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF文件", "*.pdf")])
        if out:
            writer.write(out)
            self._done(f"已合并 {len(self.merge_files)} 个文件\n保存至: {out}")
    
    # ---- 拆分PDF ----
    def split_pdf(self):
        if not self.current_reader:
            messagebox.showwarning("提示", "请先打开PDF文件")
            return
        
        mode = self.split_mode.get()
        total = len(self.current_reader.pages)
        out_dir = self.output_dir or os.path.dirname(self.current_file) or "."
        base = os.path.splitext(os.path.basename(self.current_file))[0]
        count = 0
        
        if mode == "every":
            for i in range(total):
                writer = PdfWriter()
                writer.add_page(self.current_reader.pages[i])
                out = os.path.join(out_dir, f"{base}_第{i+1}页.pdf")
                writer.write(out)
                count += 1
        
        elif mode == "by_count":
            try:
                n = int(self.split_param.get().strip())
            except ValueError:
                messagebox.showwarning("提示", "请输入每N页的数字")
                return
            
            for start in range(0, total, n):
                end = min(start + n, total)
                writer = PdfWriter()
                for i in range(start, end):
                    writer.add_page(self.current_reader.pages[i])
                out = os.path.join(out_dir, f"{base}_第{start+1}-{end}页.pdf")
                writer.write(out)
                count += 1
        
        elif mode == "by_range":
            ranges = re.findall(r'(\d+)\s*-\s*(\d+)', self.split_param.get())
            if not ranges:
                messagebox.showwarning("提示", "请输入有效范围，如: 1-10,11-20")
                return
            
            for a, b in ranges:
                start, end = int(a), int(b)
                writer = PdfWriter()
                for i in range(start - 1, min(end, total)):
                    writer.add_page(self.current_reader.pages[i])
                out = os.path.join(out_dir, f"{base}_第{start}-{end}页.pdf")
                writer.write(out)
                count += 1
        
        self._done(f"已拆分为 {count} 个文件\n输出至: {out_dir}")
    
    # ---- 水印 ----
    def add_watermark(self):
        if not self.current_reader:
            messagebox.showwarning("提示", "请先打开PDF文件")
            return
        
        text = self.watermark_text.get()
        opacity = self.watermark_opacity.get()
        
        # pypdf本身不直接支持文本水印，但可以通过叠加页面的方式
        # 创建一个简单的水印方法：在每页内容上叠加半透明文字
        # 这里使用一个间接方法：创建包含水印文本的临时PDF并合并
        try:
            from reportlab.pdfgen import canvas
            from reportlab.lib.units import cm
            from reportlab.lib.pagesizes import A4
            from io import BytesIO
        except ImportError:
            messagebox.showerror("缺少依赖", "水印功能需要安装reportlab库\n请运行: pip install reportlab")
            return
        
        temp_watermark = BytesIO()
        w, h = A4
        
        c = canvas.Canvas(temp_watermark, pagesize=A4)
        c.setFont("Helvetica", 40)
        c.setFillColorRGB(0.5, 0.5, 0.5, alpha=opacity)
        c.saveState()
        c.translate(w / 2, h / 2)
        c.rotate(45)
        c.drawString(-150, 0, text)
        c.restoreState()
        c.save()
        
        temp_watermark.seek(0)
        watermark_reader = PdfReader(temp_watermark)
        watermark_page = watermark_reader.pages[0]
        
        writer = PdfWriter()
        for page in self.current_reader.pages:
            page.merge_page(watermark_page)
            writer.add_page(page)
        
        out = self.get_output_path("水印")
        writer.write(out)
        self._done(f"已添加水印\n保存至: {out}")
    
    # ---- 压缩PDF ----
    def compress_pdf(self):
        if not self.current_reader:
            messagebox.showwarning("提示", "请先打开PDF文件")
            return
        
        level = self.compress_level.get()
        
        writer = PdfWriter()
        for page in self.current_reader.pages:
            # 压缩页面内容流
            page.compress_content_streams()
            writer.add_page(page)
        
        out = self.get_output_path("压缩")
        writer.write(out)
        
        orig_size = os.path.getsize(self.current_file)
        new_size = os.path.getsize(out)
        ratio = (1 - new_size / orig_size) * 100 if orig_size > 0 else 0
        
        self._done(f"压缩完成\n原始: {orig_size/1024/1024:.1f}MB → 压缩后: {new_size/1024/1024:.1f}MB\n减少: {ratio:.0f}%\n保存至: {out}")
    
    # ---- 提取文本 ----
    def extract_text(self):
        if not self.current_reader:
            messagebox.showwarning("提示", "请先打开PDF文件")
            return
        
        all_text = []
        for i, page in enumerate(self.current_reader.pages):
            text = page.extract_text() or ""
            all_text.append(f"=== 第{i+1}页 ===\n{text}\n")
        
        out = self.get_output_path("文本").replace('.pdf', '.txt')
        with open(out, 'w', encoding='utf-8') as f:
            f.write('\n'.join(all_text))
        
        self._done(f"已提取文本\n保存至: {out}")
    
    # ---- 文档信息 ----
    def show_doc_info(self):
        if not self.current_reader:
            messagebox.showwarning("提示", "请先打开PDF文件")
            return
        
        meta = self.current_reader.metadata
        info = f"文件: {os.path.basename(self.current_file)}\n"
        info += f"页数: {len(self.current_reader.pages)}\n\n"
        
        if meta:
            for key in ['/Title', '/Author', '/Subject', '/Creator', '/Producer', '/CreationDate', '/ModDate']:
                val = meta.get(key, '')
                if val:
                    # Clean the value (remove D: prefix from dates)
                    if isinstance(val, str) and val.startswith('D:'):
                        val = val[2:].replace("'", "")
                    info += f"{key.replace('/', '')} : {val}\n"
        
        # 页面尺寸统计
        sizes = {}
        for page in self.current_reader.pages:
            mb = page.mediabox
            w, h = round(float(mb.width), 1), round(float(mb.height), 1)
            key = f"{w}x{h}"
            sizes[key] = sizes.get(key, 0) + 1
        
        info += "\n页面尺寸分布:\n"
        for s, c in sizes.items():
            orientation = "横版" if float(s.split('x')[0]) > float(s.split('x')[1]) else "竖版"
            info += f"  {s} ({orientation}) : {c}页\n"
        
        messagebox.showinfo("文档信息", info)
    
    # ---- 移除加密 ----
    def remove_encryption(self):
        if not self.current_reader:
            messagebox.showwarning("提示", "请先打开PDF文件")
            return
        
        if not self.current_reader.is_encrypted:
            messagebox.showinfo("提示", "此PDF未加密")
            return
        
        pw = tk.simpledialog.askstring("密码", "请输入PDF密码:", show="*")
        if pw is None:
            return
        
        try:
            self.current_reader.decrypt(pw)
            writer = PdfWriter()
            for page in self.current_reader.pages:
                writer.add_page(page)
            
            out = self.get_output_path("解密")
            writer.write(out)
            self._done(f"已移除密码\n保存至: {out}")
        except Exception as e:
            messagebox.showerror("错误", f"解密失败（密码可能错误）: {e}")
    
    # ==================== 工具方法 ====================
    def _parse_page_range(self, text):
        """解析页码范围字符串"""
        text = text.strip()
        if not text:
            return []
        
        pages = set()
        parts = re.split(r'[,，\s]+', text)
        
        for part in parts:
            part = part.strip()
            if not part:
                continue
            
            m = re.match(r'^(\d+)\s*-\s*(\d+)$', part)
            if m:
                a, b = int(m.group(1)), int(m.group(2))
                if a <= b:
                    pages.update(range(a, b + 1))
                continue
            
            m = re.match(r'^(\d+)$', part)
            if m:
                pages.add(int(m.group(1)))
        
        return sorted(pages)
    
    def _done(self, msg):
        """完成回调"""
        def cb():
            self.status.config(text="完成")
            messagebox.showinfo("完成", msg)
        self.root.after(0, cb)


if __name__ == '__main__':
    root = tk.Tk()
    app = PDFToolkit(root)
    root.mainloop()
