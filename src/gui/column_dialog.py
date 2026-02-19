import customtkinter as ctk
import tkinter

class TableConfigDialog(ctk.CTkToplevel):
    def __init__(self, parent, current_columns, config=None):
        super().__init__(parent)
        self.title("表格数据配置")
        self.geometry("800x700")
        self.after(10, self.lift)

        self.parent = parent
        self.current_columns = current_columns or []
        print(f"DEBUG: TableConfigDialog init. columns={self.current_columns}")
        
        self.config = config or {}
        if "hidden" not in self.config: self.config["hidden"] = []
        if "new" not in self.config: self.config["new"] = []
        if "filter_formula" not in self.config: self.config["filter_formula"] = ""
        if "auto_index" not in self.config: self.config["auto_index"] = False
        if "auto_index_name" not in self.config: self.config["auto_index_name"] = "序号"
        if "sum_columns" not in self.config: self.config["sum_columns"] = []
        if "order" not in self.config: self.config["order"] = []

        self.result = None
        
        # --- Data Preparation (Unified sorting) ---
        self.column_items = [] # List of dict: {type: 'original'|'new', name: str, val: ..., var: ...}
        
        # 1. Determine load order
        load_order = self.config.get("order", [])
        
        # If no order, default is current_columns + new columns
        if not load_order:
            for col in self.current_columns:
                self.column_items.append({
                    "type": "original",
                    "name": col,
                    "visible": col not in self.config["hidden"]
                })
            for new_col in self.config["new"]:
                self.column_items.append({
                    "type": "new",
                    "name": new_col["name"],
                    "formula": new_col["formula"],
                })
        else:
            # Reconstruct based on saved order
            # We must also ensure all current_columns are present (in case source changed)
            # And all config['new'] are present
            
            orig_set = set(self.current_columns)
            new_map = {item["name"]: item for item in self.config["new"]}
            
            # 1. Add known items in order
            processed_orig = set()
            processed_new = set()
            
            for name in load_order:
                if name in orig_set:
                    self.column_items.append({
                        "type": "original",
                        "name": name,
                        "visible": name not in self.config["hidden"]
                    })
                    processed_orig.add(name)
                elif name in new_map:
                    self.column_items.append({
                        "type": "new",
                        "name": name,
                        "formula": new_map[name]["formula"]
                    })
                    processed_new.add(name)
            
            # 2. Append missing original columns
            for col in self.current_columns:
                if col not in processed_orig:
                    self.column_items.append({
                        "type": "original",
                        "name": col,
                        "visible": col not in self.config["hidden"]
                    })
                    
            # 3. Append missing new columns (rare, but possible if config out of sync)
            for new_col in self.config["new"]:
                if new_col["name"] not in processed_new:
                    self.column_items.append({
                        "type": "new",
                        "name": new_col["name"],
                        "formula": new_col["formula"]
                    })

        self._init_ui()

    def _init_ui(self):
        # Buttons
        self.frame_btns = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_btns.pack(side="bottom", fill="x", padx=20, pady=10)
        ctk.CTkButton(self.frame_btns, text="取消", fg_color="transparent", border_width=1, command=self.on_cancel).pack(side="right", padx=10)
        ctk.CTkButton(self.frame_btns, text="确定", command=self.on_save).pack(side="right")

        # Tabs
        self.tab_view = ctk.CTkTabview(self)
        self.tab_view.pack(side="top", fill="both", expand=True, padx=10, pady=(10, 5))
        
        self.tab_cols = self.tab_view.add("1. 列管理 & 排序")
        self.tab_filter = self.tab_view.add("2. 行筛选")

        self.selected_row_frame = None # For sorting

        self._setup_cols_tab()
        self._setup_filter_tab()

    def _setup_cols_tab(self):
        # Top: Auto Index
        f_idx = ctk.CTkFrame(self.tab_cols, fg_color="transparent")
        f_idx.pack(fill="x", padx=5, pady=5)
        
        self.var_auto_index = ctk.BooleanVar(value=self.config.get("auto_index", False))
        ctk.CTkCheckBox(f_idx, text="自动生成序号列", variable=self.var_auto_index, command=self.toggle_index).pack(side="left")
        
        self.ent_index_name = ctk.CTkEntry(f_idx, width=100, placeholder_text="列名")
        self.ent_index_name.pack(side="left", padx=5)
        self.ent_index_name.insert(0, self.config.get("auto_index_name", "序号"))
        
        if not self.var_auto_index.get(): self.ent_index_name.configure(state="disabled")

        # Top Tools
        f_tools = ctk.CTkFrame(self.tab_cols, fg_color="transparent")
        f_tools.pack(fill="x", padx=5, pady=5)
        
        ctk.CTkButton(f_tools, text="+ 添加计算列", width=100, fg_color="#2ECC71", command=self.add_new_compute_col).pack(side="left")
        
        ctk.CTkButton(f_tools, text="↓ 下移", width=60, command=self.move_down).pack(side="right", padx=5)
        ctk.CTkButton(f_tools, text="↑ 上移", width=60, command=self.move_up).pack(side="right", padx=5)

        # Header Row
        f_header = ctk.CTkFrame(self.tab_cols, height=30, fg_color="#333333")
        f_header.pack(fill="x", padx=5)
        ctk.CTkLabel(f_header, text="显示", width=40).pack(side="left", padx=5)
        ctk.CTkLabel(f_header, text="求和", width=40).pack(side="left", padx=5)
        ctk.CTkLabel(f_header, text="列名 (双击复制)", width=150, anchor="w").pack(side="left", padx=5)
        ctk.CTkLabel(f_header, text="计算公式 (混合排序)", anchor="w").pack(side="left", padx=5)

        # Main List
        self.scroll_cols = ctk.CTkScrollableFrame(self.tab_cols)
        self.scroll_cols.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Render initial list
        self.rendered_rows = [] # store widgets and data ref
        self._render_all_rows()

    def _render_all_rows(self):
        for widget in self.scroll_cols.winfo_children():
            widget.destroy()
        self.rendered_rows = []
        
        for item in self.column_items:
            self._create_row_widget(item)

    def _create_row_widget(self, item):
        row = ctk.CTkFrame(self.scroll_cols, fg_color="transparent")
        row.pack(fill="x", pady=2)
        
        # Click row to select
        row.bind("<Button-1>", lambda e, r=row: self.select_row(r))
        
        # 1. Checkbox / Visibility
        # For original: visibility toggle
        # For new: always visible? or can simply delete. Let's assume always visible implies 'exists'.
        
        var_vis = ctk.BooleanVar(value=True)
        if item["type"] == "original":
            var_vis.set(item.get("visible", True))
            chk = ctk.CTkCheckBox(row, text="", width=24, variable=var_vis)
            chk.pack(side="left", padx=5)
        else:
            # New cols are implicitly visible. If uncheck -> maybe disabled? 
            # Let's just put a placeholder or disabled check
            lbl_new = ctk.CTkLabel(row, text="★", text_color="yellow", width=24)
            lbl_new.pack(side="left", padx=5)
            # Bind click
            lbl_new.bind("<Button-1>", lambda e, r=row: self.select_row(r))
            
        # 1.5 Sum Checkbox
        var_sum = ctk.BooleanVar(value=False)
        # Check against config
        is_sum = item["name"] in self.config.get("sum_columns", [])
        var_sum.set(is_sum)
        
        chk_sum = ctk.CTkCheckBox(row, text="", width=24, variable=var_sum, checkbox_width=18, checkbox_height=18)
        chk_sum.pack(side="left", padx=5)
        # 2. Name
        if item["type"] == "original":
            lbl_name = ctk.CTkLabel(row, text=item["name"], width=150, anchor="w")
            lbl_name.pack(side="left", padx=5)
            lbl_name.bind("<Button-1>", lambda e, r=row: self.select_row(r))
            # Double click to copy {{ name }}
            lbl_name.bind("<Double-Button-1>", lambda e, n=item["name"], w=lbl_name: self.copy_col_name(n, w))
            
            item["widget_name"] = None # No entry
        else:
            ent_name = ctk.CTkEntry(row, width=150, placeholder_text="新列名")
            ent_name.pack(side="left", padx=5)
            ent_name.insert(0, item["name"])
            ent_name.bind("<Button-1>", lambda e, r=row: self.select_row(r))
            # For new cols, user can copy from entry.
            item["widget_name"] = ent_name

        # 3. Formula / Info
        if item["type"] == "original":
            # Just spacer or info
            f_lbl = ctk.CTkLabel(row, text="(原数据)", text_color="gray", anchor="w")
            f_lbl.pack(side="left", padx=5, fill="x", expand=True)
            f_lbl.bind("<Button-1>", lambda e, r=row: self.select_row(r))
            item["widget_formula"] = None
        else:
            ent_form = ctk.CTkEntry(row, placeholder_text="公式: {{数量}}*{{价}}")
            ent_form.pack(side="left", padx=5, fill="x", expand=True)
            ent_form.insert(0, item.get("formula", ""))
            ent_form.bind("<Button-1>", lambda e, r=row: self.select_row(r))
            item["widget_formula"] = ent_form

            # Delete button for new cols
            btn_del = ctk.CTkButton(row, text="✕", width=24, fg_color="transparent", text_color="#E74C3C",
                                    command=lambda i=item, r=row: self.delete_new_col(i, r))
            btn_del.pack(side="right", padx=5)

            btn_del.pack(side="right", padx=5)

        self.rendered_rows.append({"item": item, "frame": row, "var_vis": var_vis, "var_sum": var_sum})

    def select_row(self, row_frame):
        if self.selected_row_frame:
            try:
                self.selected_row_frame.configure(fg_color="transparent")
            except: pass
        self.selected_row_frame = row_frame
        row_frame.configure(fg_color=["#EBEBEB", "#3A3A3A"])

    def copy_col_name(self, name, widget=None):
        txt = f"{{{{ {name} }}}}"
        self.clipboard_clear()
        self.clipboard_append(txt)
        self.update()
        
        if widget:
            try:
                original_color = widget.cget("text_color")
                widget.configure(text_color="#2ECC71") # Green
                self.after(500, lambda: widget.configure(text_color=original_color))
            except: pass

    def add_new_compute_col(self):
        item = {"type": "new", "name": f"列{len(self.column_items)+1}", "formula": ""}
        self.column_items.append(item)
        self._create_row_widget(item)
        # Scroll to bottom?
        
    def delete_new_col(self, item, row_frame):
        if item in self.column_items:
            self.column_items.remove(item)
        
        # Remove from UI list
        for idx, rr in enumerate(self.rendered_rows):
            if rr["frame"] == row_frame:
                self.rendered_rows.pop(idx)
                break
        row_frame.destroy()

    def move_up(self):
        if not self.selected_row_frame: return
        idx = -1
        for i, rr in enumerate(self.rendered_rows):
            if rr["frame"] == self.selected_row_frame:
                idx = i
                break
        
        if idx > 0:
            # Swap in rendered list
            self.rendered_rows[idx], self.rendered_rows[idx-1] = self.rendered_rows[idx-1], self.rendered_rows[idx]
            
            # Repack all to ensure strict order
            for rr in self.rendered_rows:
                rr["frame"].pack_forget()
                rr["frame"].pack(fill="x", pady=2)

    def move_down(self):
        if not self.selected_row_frame: return
        idx = -1
        for i, rr in enumerate(self.rendered_rows):
            if rr["frame"] == self.selected_row_frame:
                idx = i
                break
        
        if idx != -1 and idx < len(self.rendered_rows) - 1:
            # Swap in rendered list
            self.rendered_rows[idx], self.rendered_rows[idx+1] = self.rendered_rows[idx+1], self.rendered_rows[idx]
            
            # Repack all to ensure strict order
            for rr in self.rendered_rows:
                rr["frame"].pack_forget()
                rr["frame"].pack(fill="x", pady=2)

    def _setup_filter_tab(self):
        ctk.CTkLabel(self.tab_filter, text="行筛选公式 (Python表达式):", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=10)
        self.txt_filter = ctk.CTkTextbox(self.tab_filter, height=150)
        self.txt_filter.pack(fill="x", padx=10, pady=5)
        self.txt_filter.insert("1.0", self.config.get("filter_formula", ""))

    def toggle_index(self):
        if self.var_auto_index.get():
            self.ent_index_name.configure(state="normal")
        else:
            self.ent_index_name.configure(state="disabled")

    def on_save(self):
        # Reconstruct output
        hidden_cols = []
        new_cols_cfg = []
        sum_cols = []
        ordered_cols = [] # names
        
        # Iterate based on UI order (rendered_rows)
        for rr in self.rendered_rows:
            item = rr["item"]
            
            # Determine Name
            if item["type"] == "original":
                name = item["name"]
                visible = rr["var_vis"].get()
                if not visible:
                    hidden_cols.append(name)
            else:
                name = item["widget_name"].get().strip()
                formula = item["widget_formula"].get().strip()
                new_cols_cfg.append({"name": name, "formula": formula})
            
                new_cols_cfg.append({"name": name, "formula": formula})
            
            if rr["var_sum"].get():
                sum_cols.append(name)

            ordered_cols.append(name)

        self.result = {
            "hidden": hidden_cols,
            "new": new_cols_cfg,
            "sum_columns": sum_cols,
            "order": ordered_cols,
            "filter_formula": self.txt_filter.get("1.0", "end-1c").strip(),
            "auto_index": self.var_auto_index.get(),
            "auto_index_name": self.ent_index_name.get().strip() or "序号"
        }
        self.destroy()

    def on_cancel(self):
        self.destroy()
