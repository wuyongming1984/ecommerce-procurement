import customtkinter as ctk
import tkinter

class TableConfigDialog(ctk.CTkToplevel):
    def __init__(self, parent, current_columns, config=None):
        super().__init__(parent)
        self.title("表格数据配置")
        self.geometry("700x600") # Increased size
        self.after(10, self.lift) # Bring to front

        self.parent = parent
        self.current_columns = current_columns or [] # List of strings (headers from Excel)
        
        # config structure: {"hidden": [], "new": [], "filter_formula": "", "auto_index": False, "auto_index_name": "序号"}
        self.config = config or {}
        
        # Ensure defaults
        if "hidden" not in self.config: self.config["hidden"] = []
        if "new" not in self.config: self.config["new"] = []
        if "filter_formula" not in self.config: self.config["filter_formula"] = ""
        if "auto_index" not in self.config: self.config["auto_index"] = False
        if "auto_index_name" not in self.config: self.config["auto_index_name"] = "序号"

        self.result = None

        # --- Buttons (Pack first to ensure visibility at bottom) ---
        self.frame_btns = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_btns.pack(side="bottom", fill="x", padx=20, pady=10)
        
        ctk.CTkButton(self.frame_btns, text="取消", fg_color="transparent", border_width=1, 
                      command=self.on_cancel).pack(side="right", padx=10)
        ctk.CTkButton(self.frame_btns, text="确定", command=self.on_save).pack(side="right")

        # Tabview
        self.tab_view = ctk.CTkTabview(self)
        self.tab_view.pack(side="top", fill="both", expand=True, padx=10, pady=(10, 5))
        
        self.tab_columns = self.tab_view.add("1. 列管理")
        self.tab_filter = self.tab_view.add("2. 行筛选")

        self._setup_columns_tab()
        self._setup_filter_tab()

    def _setup_columns_tab(self):
        # --- 1. Auto Index Section ---
        self.frame_index = ctk.CTkFrame(self.tab_columns, fg_color="transparent")
        self.frame_index.pack(side="top", fill="x", padx=5, pady=5)

        self.var_auto_index = ctk.BooleanVar(value=self.config.get("auto_index", False))
        cb_index = ctk.CTkCheckBox(self.frame_index, text="自动生成序号列", variable=self.var_auto_index, command=self.toggle_index_entry)
        cb_index.pack(side="left", padx=5)

        self.ent_index_name = ctk.CTkEntry(self.frame_index, width=120, placeholder_text="列名 (默认: 序号)")
        self.ent_index_name.pack(side="left", padx=5)
        self.ent_index_name.insert(0, self.config.get("auto_index_name", "序号"))
        
        if not self.var_auto_index.get():
            self.ent_index_name.configure(state="disabled")

        # --- 2. Hidden Columns Section ---
        self.frame_hidden = ctk.CTkFrame(self.tab_columns, fg_color="transparent")
        self.frame_hidden.pack(side="top", fill="x", padx=5, pady=5)
        
        ctk.CTkLabel(self.frame_hidden, text="保留列 (取消勾选以隐藏):", font=("Arial", 12, "bold")).pack(anchor="w")
        
        self.scroll_hidden = ctk.CTkScrollableFrame(self.frame_hidden, height=120, label_text=None)
        self.scroll_hidden.pack(fill="x", pady=5)
        
        self.check_vars = {}
        for col in self.current_columns:
            is_checked = col not in self.config.get("hidden", [])
            var = ctk.BooleanVar(value=is_checked)
            chk = ctk.CTkCheckBox(self.scroll_hidden, text=col, variable=var)
            chk.pack(anchor="w", pady=2)
            self.check_vars[col] = var

        # --- 3. New Columns Section ---
        # Use expand=True to fill remaining space
        self.frame_new = ctk.CTkFrame(self.tab_columns, fg_color="transparent")
        self.frame_new.pack(side="top", fill="both", expand=True, padx=5, pady=5)

        ctk.CTkLabel(self.frame_new, text="新增计算列:", font=("Arial", 12, "bold")).pack(anchor="w")
        
        # Tools
        self.frame_new_tools = ctk.CTkFrame(self.frame_new, fg_color="transparent")
        self.frame_new_tools.pack(fill="x", pady=2)
        ctk.CTkButton(self.frame_new_tools, text="+ 添加列", width=80, height=24, 
                      command=self.add_new_column_row, fg_color="#2ECC71").pack(side="right")

        # Header for new cols
        header_frame = ctk.CTkFrame(self.frame_new, height=24, fg_color="#333333")
        header_frame.pack(fill="x", pady=2)
        ctk.CTkLabel(header_frame, text="列名", width=100, anchor="w").pack(side="left", padx=5)
        ctk.CTkLabel(header_frame, text="计算公式 (可用 {{列名}})，留空则为空值", anchor="w").pack(side="left", padx=5)

        self.scroll_new = ctk.CTkScrollableFrame(self.frame_new)
        self.scroll_new.pack(fill="both", expand=True, pady=0)
        
        self.new_col_widgets = []
        
        # Load existing new columns
        for item in self.config.get("new", []):
            self.add_new_column_row(item["name"], item["formula"])

    def _setup_filter_tab(self):
        ctk.CTkLabel(self.tab_filter, text="行筛选公式 (Python表达式):", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=(10, 5))
        ctk.CTkLabel(self.tab_filter, text="仅保留计算结果为 True 的行。例如: {{金额}} > 1000", text_color="gray").pack(anchor="w", padx=10, pady=(0, 5))

        self.txt_filter = ctk.CTkTextbox(self.tab_filter, height=100)
        self.txt_filter.pack(fill="x", padx=10, pady=5)
        
        current_formula = self.config.get("filter_formula", "")
        self.txt_filter.insert("1.0", current_formula)

    def add_new_column_row(self, name="", formula=""):
        row = ctk.CTkFrame(self.scroll_new, fg_color="transparent")
        row.pack(fill="x", pady=2)
        
        ent_name = ctk.CTkEntry(row, width=100, placeholder_text="新列名")
        ent_name.pack(side="left", padx=2)
        ent_name.insert(0, name)
        
        ent_formula = ctk.CTkEntry(row, placeholder_text="例如: {{数量}} * {{单价}}")
        ent_formula.pack(side="left", padx=2, fill="x", expand=True)
        ent_formula.insert(0, formula)
        
        btn_del = ctk.CTkButton(row, text="✕", width=24, fg_color="transparent", text_color="red",
                                command=lambda r=row: self.remove_new_row(r))
        btn_del.pack(side="right", padx=2)
        
        self.new_col_widgets.append({"row": row, "name": ent_name, "formula": ent_formula})
        
    def remove_new_row(self, row):
        for i, item in enumerate(self.new_col_widgets):
            if item["row"] == row:
                self.new_col_widgets.pop(i)
                break
        row.destroy()

    def toggle_index_entry(self):
        if self.var_auto_index.get():
            self.ent_index_name.configure(state="normal", fg_color=["#F9F9FA", "#343638"]) # Default ctk colors
        else:
            self.ent_index_name.configure(state="disabled")

    def on_save(self):
        # 1. Collect Hidden
        hidden_cols = []
        for col, var in self.check_vars.items():
            if not var.get():
                hidden_cols.append(col)
        
        # 2. Collect New
        new_cols = []
        for item in self.new_col_widgets:
            name = item["name"].get().strip()
            formula = item["formula"].get().strip()
            # Allow empty formula (creates empty field)
            if name:
                new_cols.append({"name": name, "formula": formula})

        # 3. Filter
        filter_formula = self.txt_filter.get("1.0", "end-1c").strip()
        
        # 4. Auto Index
        auto_index = self.var_auto_index.get()
        auto_index_name = self.ent_index_name.get().strip() or "序号"

        self.result = {
            "hidden": hidden_cols,
            "new": new_cols,
            "filter_formula": filter_formula,
            "auto_index": auto_index,
            "auto_index_name": auto_index_name
        }
        self.destroy()

    def on_cancel(self):
        self.destroy()
