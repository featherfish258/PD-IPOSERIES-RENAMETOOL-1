import os
import shutil  # 👈 新增这行，用于复制文件
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ── 配色方案 ──
COLORS = {
    "bg":          "#f5f7fa",
    "card":        "#ffffff",
    "primary":     "#4361ee",
    "primary_fg":  "#ffffff",
    "accent":      "#3a0ca3",
    "success":     "#2ec4b6",
    "danger":      "#e63946",
    "text":        "#2b2d42",
    "text_sub":    "#6c757d",
    "border":      "#dee2e6",
    "input_bg":    "#f8f9fa",
    "log_bg":      "#1e1e2e",
    "log_fg":      "#cdd6f4",
    "log_ok":      "#a6e3a1",
    "log_err":     "#f38ba8",
    "hover":       "#3651d4",
    "tag_even":    "#f0f4ff",
}

# ── 辅助函数 ──
def _round_rect(canvas, x1, y1, x2, y2, r=12, **kw):
    """在 Canvas 上画圆角矩形"""
    pts = [
        x1+r, y1, x1+r, y1, x2-r, y1, x2-r, y1, x2, y1, x2, y1+r,
        x2, y1+r, x2, y2-r, x2, y2-r, x2, y2, x2-r, y2, x2-r, y2,
        x1+r, y2, x1+r, y2, x1, y2, x1, y2-r, x1, y2-r, x1, y1+r,
        x1, y1+r, x1, y1, x1+r, y1,
    ]
    return canvas.create_polygon(pts, smooth=True, **kw)


class ModernButton(tk.Canvas):
    """自绘圆角按钮"""
    def __init__(self, parent, text="Button", command=None,
                 bg="#4361ee", fg="#ffffff", hover="#3651d4",
                 font=("Microsoft YaHei UI", 10, "bold"),
                 width=160, height=38, radius=10, **kw):
        super().__init__(parent, width=width, height=height,
                         bg=parent["bg"], highlightthickness=0, **kw)
        self._bg, self._hover, self._fg = bg, hover, fg
        self._cmd = command
        
        # 💡 修改处：避免使用 self._w 覆盖 Tkinter 内置属性
        self.btn_w, self.btn_h, self._r = width, height, radius
        
        self._font = font
        self._text = text
        self._draw(bg)
        self.bind("<Enter>",      lambda e: self._draw(self._hover))
        self.bind("<Leave>",      lambda e: self._draw(self._bg))
        self.bind("<ButtonPress-1>", lambda e: self._on_click())

    def _draw(self, fill):
        self.delete("all")
        # 💡 修改处：同步更改为 btn_w 和 btn_h
        _round_rect(self, 2, 2, self.btn_w-2, self.btn_h-2, r=self._r, fill=fill, outline="")
        self.create_text(self.btn_w/2, self.btn_h/2, text=self._text,
                         fill=self._fg, font=self._font)

    def _on_click(self):
        if self._cmd:
            self._cmd()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("回函文件重命名工具 V3.0")
        self.geometry("720x680")
        self.configure(bg=COLORS["bg"])
        self.minsize(640, 580)

        # ── 全局字体 ──
        self.fn      = ("Microsoft YaHei UI", 10)
        self.fn_bold = ("Microsoft YaHei UI", 10, "bold")
        self.fn_sm   = ("Microsoft YaHei UI", 9)
        self.fn_title = ("Microsoft YaHei UI", 14, "bold")
        self.fn_mono = ("Consolas", 9)

        # ── ttk 样式 ──
        self._setup_styles()

        # ── 页面构建 ──
        self._build_header()
        self._build_folder_card()
        self._build_table_card()
        self._build_action_bar()
        self._build_log_card()

    # ────────── 样式 ──────────
    def _setup_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("Card.TFrame", background=COLORS["card"])

        # Treeview
        style.configure("Custom.Treeview",
                         background=COLORS["card"],
                         foreground=COLORS["text"],
                         fieldbackground=COLORS["card"],
                         borderwidth=0,
                         font=self.fn_sm,
                         rowheight=30)
        style.configure("Custom.Treeview.Heading",
                         background=COLORS["primary"],
                         foreground=COLORS["primary_fg"],
                         font=("Microsoft YaHei UI", 9, "bold"),
                         borderwidth=0,
                         relief="flat",
                         padding=(8, 6))
        style.map("Custom.Treeview.Heading",
                   background=[("active", COLORS["hover"])])
        style.map("Custom.Treeview",
                   background=[("selected", "#dbe4ff")],
                   foreground=[("selected", COLORS["accent"])])

        # Scrollbar
        style.configure("Round.Vertical.TScrollbar",
                         gripcount=0,
                         background=COLORS["border"],
                         troughcolor=COLORS["card"],
                         borderwidth=0,
                         arrowsize=0)

    # ────────── 顶部标题栏 ──────────
    def _build_header(self):
        hdr = tk.Frame(self, bg=COLORS["primary"], height=56)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)

        tk.Label(hdr, text="📂  回函文件重命名工具",
                 font=("Microsoft YaHei UI", 15, "bold"),
                 bg=COLORS["primary"], fg=COLORS["primary_fg"]
                 ).pack(side=tk.LEFT, padx=20)

        tk.Label(hdr, text="v3.0",
                 font=self.fn_sm, bg=COLORS["primary"],
                 fg="#b8c4ff").pack(side=tk.RIGHT, padx=20)

    # ────────── 文件夹选择卡片 ──────────
    def _build_folder_card(self):
        wrap = tk.Frame(self, bg=COLORS["bg"])
        wrap.pack(fill=tk.X, padx=16, pady=(14, 0))

        card = tk.Frame(wrap, bg=COLORS["card"],
                        highlightbackground=COLORS["border"],
                        highlightthickness=1)
        card.pack(fill=tk.X)

        inner = tk.Frame(card, bg=COLORS["card"])
        inner.pack(fill=tk.X, padx=16, pady=14)

        tk.Label(inner, text="待处理文件夹", font=self.fn_bold,
                 bg=COLORS["card"], fg=COLORS["text"]).pack(anchor=tk.W)

        row = tk.Frame(inner, bg=COLORS["card"])
        row.pack(fill=tk.X, pady=(8, 0))

        self.folder_var = tk.StringVar()
        ent = tk.Entry(row, textvariable=self.folder_var, font=self.fn,
                       state="readonly", readonlybackground=COLORS["input_bg"],
                       fg=COLORS["text"], relief="solid", bd=1,
                       highlightcolor=COLORS["primary"])
        ent.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=5)

        ModernButton(row, text="浏览文件夹", command=self._select_folder,
                     bg=COLORS["primary"], hover=COLORS["hover"],
                     width=120, height=34, radius=8,
                     font=("Microsoft YaHei UI", 9, "bold")
                     ).pack(side=tk.RIGHT, padx=(10, 0))

    # ────────── 数据表格卡片 ──────────
    def _build_table_card(self):
        wrap = tk.Frame(self, bg=COLORS["bg"])
        wrap.pack(fill=tk.BOTH, expand=True, padx=16, pady=(10, 0))

        card = tk.Frame(wrap, bg=COLORS["card"],
                        highlightbackground=COLORS["border"],
                        highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True)

        inner = tk.Frame(card, bg=COLORS["card"])
        inner.pack(fill=tk.BOTH, expand=True, padx=16, pady=14)

        # 工具栏
        toolbar = tk.Frame(inner, bg=COLORS["card"])
        toolbar.pack(fill=tk.X, pady=(0, 10))

        tk.Label(toolbar, text="数据列表", font=self.fn_bold,
                 bg=COLORS["card"], fg=COLORS["text"]).pack(side=tk.LEFT)

        self.count_label = tk.Label(
            toolbar, text="0 行数据", font=self.fn_sm,
            bg="#fff3cd", fg="#856404", padx=10, pady=2)
        self.count_label.pack(side=tk.RIGHT)

        ModernButton(toolbar, text="📋  粘贴 Excel 数据",
                     command=self._paste_from_excel,
                     bg=COLORS["success"], hover="#25a89d",
                     width=155, height=32, radius=8,
                     font=("Microsoft YaHei UI", 9, "bold")
                     ).pack(side=tk.RIGHT, padx=(0, 10))

        tk.Label(toolbar, text="从 Excel 复制 3 列后点击粘贴 →",
                 font=self.fn_sm, bg=COLORS["card"],
                 fg=COLORS["text_sub"]).pack(side=tk.RIGHT, padx=(0, 8))

        # 表格容器
        tree_frame = tk.Frame(inner, bg=COLORS["card"])
        tree_frame.pack(fill=tk.BOTH, expand=True)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical",
                             style="Round.Vertical.TScrollbar")
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        columns = ("seq", "conf_num", "reply_num", "print_seq")
        self.tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings",
            yscrollcommand=vsb.set, height=9,
            style="Custom.Treeview")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.config(command=self.tree.yview)

        heads = [("seq", "序号", 55), ("conf_num", "函证编号", 170),
                 ("reply_num", "回函单号", 170), ("print_seq", "打印序列", 110)]
        for cid, txt, w in heads:
            self.tree.heading(cid, text=txt)
            self.tree.column(cid, width=w, anchor="center", minwidth=w)

        self.tree.tag_configure("even", background=COLORS["tag_even"])
        self.tree.bind('<Control-v>', lambda e: self._paste_from_excel())

    # ────────── 操作按钮栏 ──────────
    def _build_action_bar(self):
        bar = tk.Frame(self, bg=COLORS["bg"])
        bar.pack(fill=tk.X, padx=16, pady=10)

        ModernButton(bar, text="🚀  开始重命名", command=self._process_files,
                     bg=COLORS["accent"], hover="#2d0a80",
                     width=200, height=42, radius=10,
                     font=("Microsoft YaHei UI", 11, "bold")
                     ).pack()

    # ────────── 日志卡片（暗色终端风格） ──────────
    def _build_log_card(self):
        wrap = tk.Frame(self, bg=COLORS["bg"])
        wrap.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 14))

        card = tk.Frame(wrap, bg=COLORS["log_bg"],
                        highlightbackground="#313244",
                        highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True)

        title_bar = tk.Frame(card, bg="#313244", height=28)
        title_bar.pack(fill=tk.X)
        title_bar.pack_propagate(False)

        # 模拟终端三个小圆点
        dots = tk.Frame(title_bar, bg="#313244")
        dots.pack(side=tk.LEFT, padx=10)
        for c in ["#f38ba8", "#fab387", "#a6e3a1"]:
            tk.Canvas(dots, width=10, height=10, bg="#313244",
                      highlightthickness=0).pack(side=tk.LEFT, padx=2)
            cv = dots.winfo_children()[-1]
            cv.create_oval(1, 1, 9, 9, fill=c, outline="")

        tk.Label(title_bar, text="处理日志", font=("Microsoft YaHei UI", 9),
                 bg="#313244", fg="#cdd6f4").pack(side=tk.LEFT, padx=6)

        self.log_area = tk.Text(
            card, height=5, bg=COLORS["log_bg"], fg=COLORS["log_fg"],
            font=self.fn_mono, relief="flat", padx=12, pady=8,
            insertbackground=COLORS["log_fg"], wrap="word",
            selectbackground="#45475a")
        self.log_area.pack(fill=tk.BOTH, expand=True)
        self.log_area.tag_configure("ok",  foreground=COLORS["log_ok"])
        self.log_area.tag_configure("err", foreground=COLORS["log_err"])
        self.log_area.tag_configure("info", foreground="#89b4fa")

    # ────────── 业务逻辑 ──────────
    def _select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_var.set(folder)

    def _paste_from_excel(self):
        try:
            # 尝试强制获取 UTF-8 格式的剪贴板内容，防止系统默认 ANSI 导致中文乱码
            clipboard_data = self.selection_get(selection="CLIPBOARD")
        except tk.TclError:
            try:
                # 备用方案
                clipboard_data = self.clipboard_get()
            except tk.TclError:
                messagebox.showwarning("提示", "剪贴板为空，请先从 Excel 复制数据！")
                return "break"

        for item in self.tree.get_children():
            self.tree.delete(item)

        lines = clipboard_data.split('\n')
        count = 0
        for line in lines:
            # 深度清理 Excel 带来的不可见字符（\r, \t, 零宽空格等）
            line = line.strip('\r\n\u200b\xa0 ') 
            if not line:
                continue
            
            parts = line.split('\t')
            # 清理每个单元格的前后空格
            parts = [p.strip() for p in parts]
            
            while len(parts) < 3:
                parts.append("")
                
            count += 1
            tag = "even" if count % 2 == 0 else ""
            self.tree.insert("", tk.END,
                             values=(count, parts[0], parts[1], parts[2]),
                             tags=(tag,))

        self.count_label.config(text=f"{count} 行数据")
        if count > 0:
            self.count_label.config(bg="#d1e7dd", fg="#0f5132")
        else:
            self.count_label.config(bg="#fff3cd", fg="#856404")
        return "break"

    def _log(self, msg, tag=""):
        self.log_area.insert(tk.END, msg + "\n", tag)
        self.log_area.see(tk.END)

    def _process_files(self):
        folder = self.folder_var.get()
        if not folder:
            messagebox.showerror("错误", "请先选择待处理文件夹！")
            return

        mapping = []
        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            if len(values) >= 4:
                mapping.append({
                    'conf_num':  str(values[1]).strip(),
                    'reply_num': str(values[2]).strip(),
                    'print_seq': str(values[3]).strip()
                })

        if not mapping:
            messagebox.showerror("错误", "表格中没有数据，请先粘贴 Excel 序列！")
            return

        success_count = 0
        copy_count = 0  # 记录生成副本的数量
        self.log_area.delete("1.0", tk.END)
        self._log(f"▶ 开始处理  目标文件夹: {folder}", "info")
        self._log("─" * 56, "info")

        matched_row_indices = set() # 记录被成功消耗的 Excel 数据行索引
        processed_files = set()     # 记录已经处理过的物理文件

        all_files = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]

        for filename in all_files:
            filepath = os.path.join(folder, filename)

            # 1. 先找出当前文件包含哪个回函单号
            matched_reply_num = None
            for item in mapping:
                if item['reply_num'] and item['reply_num'] in filename:
                    matched_reply_num = item['reply_num']
                    break
            
            # 2. 如果找到了对应的单号，找出 Excel 中所有同单号且未被处理的行
            # 2. 如果找到了对应的单号，找出 Excel 中所有同单号且未被处理的行
            if matched_reply_num:
                matching_rows = [(idx, m) for idx, m in enumerate(mapping) 
                                 if m['reply_num'] == matched_reply_num and idx not in matched_row_indices]
                
                if not matching_rows:
                    continue # 如果该单号的数据已经被前面的文件消耗光了，就跳过

                name, ext = os.path.splitext(filename)
                
                # 💡 核心修改：预处理剩下的文本，剥离原有的 + 号和两端空格
                remaining = name.replace(matched_reply_num, "").strip("+ _-")

                # 3. 遍历这些同单号的行，生成副本或重命名
                for i, (idx, item) in enumerate(matching_rows):
                    
                    # 💡 核心修改：智能拼接，避免出现 ++ 或者末尾多 + 的情况
                    core_parts = [item['conf_num'], matched_reply_num]
                    if remaining:  # 如果还有剩下的文本，才加进去
                        core_parts.append(remaining)
                    
                    core_name = "+".join(core_parts) # 自动用 + 号把这几部分连起来
                    
                    # 最终组合： 序号、 + 核心部分 + 后缀
                    new_name = f"{item['print_seq']}、{core_name}{ext}"
                    new_filepath = os.path.join(folder, new_name)

                    try:
                        # 如果不是最后一个，就用 copy2 复制一份
                        if i < len(matching_rows) - 1:
                            shutil.copy2(filepath, new_filepath)
                            copy_count += 1
                            self._log(f"  ✓ [生成副本] {new_name}", "info")
                        else:
                            # 最后一个匹配项，直接 rename 重命名
                            os.rename(filepath, new_filepath)
                            self._log(f"  ✓ [重命名] {new_name}", "ok")
                        
                        success_count += 1
                        matched_row_indices.add(idx)
                    except Exception as e:
                        self._log(f"  ✗ {filename} 错误: {e}", "err")

                # 标记该物理文件已被处理
                processed_files.add(filename)

        # ── 分类汇总 ──
        matched_rows = [item for idx, item in enumerate(mapping) if idx in matched_row_indices]
        unmatched_rows = [item for idx, item in enumerate(mapping) if idx not in matched_row_indices]
        unmatched_files = [f for f in all_files if f not in processed_files]

        # ── 导出对比报告 ──
        report_path = os.path.join(folder, "回函处理对比报告.csv")
        try:
            with open(report_path, "w", encoding="utf-8-sig") as f:
                f.write("分类,函证编号,回函单号,打印序列/文件名\n")
                
                for item in matched_rows:
                    f.write(f"已成功匹配,{item['conf_num']},{item['reply_num']},{item['print_seq']}\n")
                
                for item in unmatched_rows:
                    f.write(f"【缺文件需催收】,{item['conf_num']},{item['reply_num']},{item['print_seq']}\n")
                
                for file in unmatched_files:
                    f.write(f"未识别的冗余文件,N/A,N/A,{file}\n")
                    
            self._log(f"  📝 已生成对比报告: {report_path}", "info")
        except Exception as e:
            self._log(f"  ✗ 报告生成失败: {e}", "err")

        self._log("─" * 56, "info")
        self._log(f"▶ 完成！成功处理 {success_count} 笔回函 (包含 {copy_count} 个副本)", "info")
        
        msg = (f"处理完毕！\n共匹配 {success_count} 笔函证（其中自动生成副本 {copy_count} 份）。\n\n"
               f"⚠️ 缺文件需催收：{len(unmatched_rows)} 家\n"
               f"❓ 未识别的冗余文件：{len(unmatched_files)} 个\n\n"
               f"详细名单已导出至文件夹下的《回函处理对比报告.csv》")
        messagebox.showinfo("完成", msg)


if __name__ == "__main__":
    app = App()
    app.mainloop()