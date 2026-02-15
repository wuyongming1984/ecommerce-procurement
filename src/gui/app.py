import customtkinter as ctk
import tkinter.messagebox
from tkinter import filedialog
import os
import json
import re
from pathlib import Path
from ..core.template_mgr import TemplateManager, ExcelLoader
from ..core.generator import DocumentGenerator
from .column_dialog import TableConfigDialog

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title("Document Generator (2.0)")
        self.geometry(f"{1100}x{850}")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Top Bar (Unified Config)
        self.top_bar = ctk.CTkFrame(self, fg_color="transparent")
        self.top_bar.pack(fill="x", padx=10, pady=5)
        
        btn_load_proj = ctk.CTkButton(self.top_bar, text="📂 加载项目配置", command=self.load_project, width=120)
        btn_load_proj.pack(side="left", padx=5)
        
        btn_save_proj = ctk.CTkButton(self.top_bar, text="💾 保存项目配置", command=self.save_project, width=120)
        btn_save_proj.pack(side="left", padx=5)

        self.tab_view = ctk.CTkTabview(self, segmented_button_fg_color="#2B2B2B", 
                                       segmented_button_selected_color="#3B8ED0",
                                       segmented_button_selected_hover_color="#2D6A9F",
                                       segmented_button_unselected_color="#333333",
                                       segmented_button_unselected_hover_color="#3D3D3D")
        self.tab_view._segmented_button.configure(font=("Arial", 16, "bold"))
        self.tab_view.pack(fill="both", expand=True, padx=20, pady=(10, 20))

        # Create tabs
        self.vars_tab = self.tab_view.add("1. 变量配置")
        self.tables_tab = self.tab_view.add("2. 表格数据")
        self.files_tab = self.tab_view.add("3. 生成设置")

        # Data structures
        self.variables = [] # List of dicts {"g": g_ent, "k": k_ent, "v": v_ent, "row": row}
        self.table_configs = [] # List of dicts {"name": k_ent, "path": p_ent, "sheet": s_ent, "row": row}
        self.file_mappings = [] # List of (input_entry, output_entry, row_frame)
        self.group_collapsed = {} # group_name (str) -> is_collapsed (bool)

        # State tracking for unsaved changes
        self.last_saved_vars = []
        self.last_saved_tables = []
        self.last_saved_mappings = []

        self._setup_vars_tab()
        self._setup_tables_tab()
        self._setup_files_tab()
        
        # Override close button
        self.protocol("WM_DELETE_WINDOW", self.on_app_close)

        # Auto Load Project
        if os.path.exists("project.json"):
            print(f"检测到项目配置文件: {os.path.abspath('project.json')}")
            try:
                self.load_project("project.json")
                print("自动加载项目配置成功")
            except Exception as e:
                print(f"自动加载项目配置失败: {e}")
                import traceback
                traceback.print_exc()
        else:
            # Fallback for compatibility or fresh start
            print("未检测到 'project.json'")
            # Load defaults (old way for compatibility if project.json doesn't exist)
            if os.path.exists("变量配置.json"):
                self.load_vars_json("变量配置.json")
                
            if os.path.exists("表格配置.json"):
                print(f"检测到配置文件: {os.path.abspath('表格配置.json')}")
                try:
                    self.load_tables_json("表格配置.json")
                    print("自动加载表格配置成功")
                except Exception as e:
                    print(f"自动加载表格配置失败: {e}")
                    import traceback
                    traceback.print_exc()
            else:
                print("未检测到 '表格配置.json'，跳过自动加载。")

            if os.path.exists("生成配置.json"):
                self.load_mappings_json("生成配置.json")

    def _setup_vars_tab(self):
        """
        Page 1: Define placeholder names and values, save as JSON.
        """
        self.vars_tab.grid_columnconfigure(0, weight=1)
        self.vars_tab.grid_rowconfigure(1, weight=1)

        # Toolbar
        frame_tools = ctk.CTkFrame(self.vars_tab, fg_color="transparent")
        frame_tools.grid(row=0, column=0, padx=20, pady=(10, 5), sticky="ew")

        btn_add = ctk.CTkButton(frame_tools, text="+ 添加新变量", command=lambda: self.add_variable_row("", ""), 
                                 fg_color="#2ECC71", hover_color="#27AE60", font=("Arial", 13, "bold"))
        btn_add.pack(side="left", padx=5, pady=5)

        btn_scan = ctk.CTkButton(frame_tools, text="🔍 从文件扫描", command=self.scan_variables_from_file)
        btn_scan.pack(side="left", padx=5, pady=5)

        # Removed individual load/save buttons

        btn_clear = ctk.CTkButton(frame_tools, text="🗑 清空", fg_color="#E74C3C", hover_color="#C0392B", 
                                   command=self.clear_vars, width=80)
        btn_clear.pack(side="right", padx=5, pady=5)

        # List Area
        self.vars_scroll = ctk.CTkScrollableFrame(self.vars_tab, label_text="变量列表 (Key - Value)")
        self.vars_scroll.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

        # Help
        lbl = ctk.CTkLabel(self.vars_tab, text="提示: 定义模板中的 {{ 变量名 }} 及其对应的值。保存后可在下次直接加载。", text_color="gray")
        lbl.grid(row=2, column=0, pady=5)

    def _setup_tables_tab(self):
        """
        Page 2: Configure Excel Data Sources (Tables)
        """
        self.tables_tab.grid_columnconfigure(0, weight=1)
        self.tables_tab.grid_rowconfigure(1, weight=1) # Config list
        self.tables_tab.grid_rowconfigure(3, weight=2) # Preview area (larger)

        # Toolbar
        frame_tools = ctk.CTkFrame(self.tables_tab, fg_color="transparent")
        frame_tools.grid(row=0, column=0, padx=20, pady=(10, 5), sticky="ew")

        btn_add = ctk.CTkButton(frame_tools, text="+ 添加表格源", command=lambda: self.add_table_row("", "", ""),
                                 fg_color="#2ECC71", hover_color="#27AE60", font=("Arial", 13, "bold"))
        btn_add.pack(side="left", padx=5, pady=5)

        # Removed individual load/save buttons

        btn_clear = ctk.CTkButton(frame_tools, text="🗑 清空", fg_color="#E74C3C", hover_color="#C0392B", 
                                   command=self.clear_tables, width=80)
        btn_clear.pack(side="right", padx=5, pady=5)

        # List Area (Top Half)
        self.tables_scroll = ctk.CTkScrollableFrame(self.tables_tab, label_text="表格数据源列表 (Variable -> Excel Path)", height=200)
        self.tables_scroll.grid(row=1, column=0, padx=20, pady=5, sticky="nsew")

        # Preview Header
        preview_header = ctk.CTkFrame(self.tables_tab, height=30, fg_color="#333333", corner_radius=6)
        preview_header.grid(row=2, column=0, padx=20, pady=(10, 0), sticky="ew")
        
        self.lbl_preview_title = ctk.CTkLabel(preview_header, text="数据预览 (点击上方 '👁' 查看内容)", font=("Arial", 12, "bold"), text_color="#3B8ED0")
        self.lbl_preview_title.pack(side="left", padx=10)

        # Preview Area (Bottom Half)
        self.preview_scroll = ctk.CTkScrollableFrame(self.tables_tab, fg_color="#1E1E1E", label_text=None)
        self.preview_scroll.grid(row=3, column=0, padx=20, pady=(0, 10), sticky="nsew")

        # Help (Moved to bottom of preview or keep distinct? Let's put in preview header or toolbar? Keep at bottom row 4)
        lbl = ctk.CTkLabel(self.tables_tab, text="提示: 定义的表格变量名可在模板中使用 {% tr for item in items %} ... {% endfor %} 进行循环。Excel首行需为表头。", 
                           text_color="gray", font=("Arial", 11))
        lbl.grid(row=4, column=0, pady=5)

    def _setup_files_tab(self):
        """
        Page 2: Import Template -> Export File mapping.
        """
        self.files_tab.grid_columnconfigure(0, weight=1)
        self.files_tab.grid_rowconfigure(1, weight=1)

        # Toolbar
        frame_tools = ctk.CTkFrame(self.files_tab, fg_color="transparent")
        frame_tools.grid(row=0, column=0, padx=20, pady=10, sticky="ew")

        btn_add = ctk.CTkButton(frame_tools, text="+ 添加新映射", command=lambda: self.add_mapping_row("", ""),
                                 fg_color="#2ECC71", hover_color="#27AE60", font=("Arial", 13, "bold"))
        btn_add.pack(side="left", padx=5, pady=5)

        btn_save = ctk.CTkButton(frame_tools, text="💾 保存映射", command=self.save_mappings_json)
        btn_save.pack(side="left", padx=5, pady=5)

        btn_load = ctk.CTkButton(frame_tools, text="📂 加载映射", command=self.load_mappings_json)
        btn_load.pack(side="left", padx=5, pady=5)

        btn_run = ctk.CTkButton(frame_tools, text="🚀 开始生成 (执行所有)", command=self.run_generation, 
                                 fg_color="#3498DB", hover_color="#2980B9", font=("Arial", 14, "bold"),
                                 height=36)
        btn_run.pack(side="right", padx=5, pady=5)
        
        btn_clear = ctk.CTkButton(frame_tools, text="🗑 清空", fg_color="#E74C3C", hover_color="#C0392B", 
                                   command=self.clear_mappings, width=80)
        btn_clear.pack(side="right", padx=5, pady=5)

        # Headers
        header_frame = ctk.CTkFrame(self.files_tab, height=35, fg_color="#333333", corner_radius=6)
        header_frame.grid(row=0, column=0, sticky="ew", padx=25, pady=(60,0)) 
        
        ctk.CTkLabel(header_frame, text="📄 导入模板位置", font=("Arial", 12, "bold"), text_color="#3B8ED0").grid(row=0, column=0, sticky="w", padx=100)
        ctk.CTkLabel(header_frame, text="📁 导出目录位置", font=("Arial", 12, "bold"), text_color="#3B8ED0").grid(row=0, column=1, sticky="w", padx=200)

        # List Area
        self.files_scroll = ctk.CTkScrollableFrame(self.files_tab, fg_color="transparent")
        self.files_scroll.grid(row=1, column=0, padx=20, pady=5, sticky="nsew")
        
        # Log Area
        lbl_log = ctk.CTkLabel(self.files_tab, text="系统日志 (System Logs)", font=("Arial", 11, "bold"), text_color="gray")
        lbl_log.grid(row=2, column=0, padx=25, sticky="w")
        
        self.log_text = ctk.CTkTextbox(self.files_tab, height=180, fg_color="#1E1E1E", border_width=1, 
                                       font=("Consolas", 12), text_color="#A9B7C6")
        self.log_text.grid(row=3, column=0, padx=20, pady=(0, 10), sticky="ew")


    # --- Variables Logic ---

    def add_variable_row(self, key, val, group=""):
        row = ctk.CTkFrame(self.vars_scroll, fg_color="#2B2B2B", corner_radius=8)
        row.pack(fill="x", pady=4, padx=5)
        
        g_ent = ctk.CTkEntry(row, placeholder_text="分组", width=100, border_width=1)
        g_ent.pack(side="left", padx=(10, 5), pady=8)
        g_ent.insert(0, group)
        g_ent.bind("<KeyRelease>", lambda e: self.refresh_var_list())

        k_ent = ctk.CTkEntry(row, placeholder_text="变量名", width=180, border_width=1, font=("Arial", 12, "bold"))
        k_ent.pack(side="left", padx=5, pady=8)
        k_ent.insert(0, key)
        k_ent.bind("<KeyRelease>", lambda e: self.validate_all_keys())
        k_ent.bind("<Double-Button-1>", lambda e, w=k_ent: self.copy_to_clipboard_bracket(w))

        v_ent = ctk.CTkTextbox(row, width=380, height=35, border_width=1, fg_color="#1E1E1E")
        v_ent.pack(side="left", padx=5, pady=8, fill="x", expand=True)
        v_ent.insert("1.0", val)

        btn_img = ctk.CTkButton(row, text="🖼 选图", width=65, height=28, command=lambda e=v_ent: self.select_image_file(e))
        btn_img.pack(side="left", padx=2, pady=8)

        # Dropdown for image size
        size_options = ["默认", "全宽 (width=full)", "宽60mm", "宽80mm", "高30mm", "高50mm"]
        size_menu = ctk.CTkOptionMenu(row, values=size_options, width=110, height=28,
                                      command=lambda choice, e=v_ent: self.update_image_size(e, choice))
        
        def on_value_change(event=None):
            val = v_ent.get("1.0", "end-1c").lower().strip()
            is_img = any(ext in val for ext in ['.png', '.jpg', '.jpeg'])
            if is_img:
                size_menu.pack(side="left", padx=2, pady=8, after=btn_img)
            else:
                size_menu.pack_forget()

            try:
                line_count = v_ent._textbox.count("1.0", "end-1c", "displaylines")[0]
            except:
                line_count = len(v_ent.get("1.0", "end-1c").split('\n'))

            new_height = 35 + max(0, line_count - 1) * 20
            new_height = min(new_height, 300)
            if v_ent.cget("height") != new_height:
                v_ent.configure(height=new_height)

        v_ent.bind("<KeyRelease>", on_value_change)
        v_ent.bind("<Configure>", on_value_change)
        on_value_change()
        btn_img.configure(command=lambda e=v_ent, cb=on_value_change: self.select_image_file(e, cb))

        lbl_handle = ctk.CTkLabel(row, text="⠿", width=30, cursor="hand2", text_color="gray")
        lbl_handle.pack(side="left", padx=2, pady=8)
        lbl_handle.bind("<Button-1>", lambda e, r=row: self.start_drag(e, r))
        lbl_handle.bind("<B1-Motion>", lambda e, r=row: self.on_drag(e, r))

        btn_del = ctk.CTkButton(row, text="✕", width=30, height=28, fg_color="transparent", 
                                 text_color="#E74C3C", hover_color="#3D3D3D", 
                                 command=lambda r=row: self.remove_var_row(r))
        btn_del.pack(side="right", padx=(5, 10), pady=8)

        self.variables.append({"g": g_ent, "k": k_ent, "v": v_ent, "row": row})
        self.refresh_var_list()

    def select_image_file(self, entry, callback=None):
        f = filedialog.askopenfilename(filetypes=[("Images", "*.png;*.jpg;*.jpeg"), ("All Files", "*.*")], initialdir=os.getcwd())
        if f:
            try:
                rel_f = os.path.relpath(f, os.getcwd())
            except ValueError:
                rel_f = f
            entry.delete("1.0", "end")
            # Default to width=full as requested
            entry.insert("1.0", f"{rel_f}|width=full")
            
            if callback:
                callback()

    def update_image_size(self, entry, choice):
        # 1. Get current path (strip existing params)
        current_val = entry.get("1.0", "end-1c").strip()
        if not current_val:
            return
            
        parts = current_val.split('|')
        clean_path = parts[0].strip()
        
        # 2. Append new param base on choice
        new_val = clean_path
        if choice == "全宽 (width=full)":
            new_val += "|width=full"
        elif choice == "宽60mm":
            new_val += "|width=60"
        elif choice == "宽80mm":
            new_val += "|width=80"
        elif choice == "高30mm":
            new_val += "|height=30"
        elif choice == "高50mm":
            new_val += "|height=50"
        # "默认" does nothing, effectively removing params
            
        # 3. Update entry
        entry.delete("1.0", "end")
        entry.insert("1.0", new_val)

    def remove_var_row(self, row):
        for i, item in enumerate(self.variables):
            if item["row"] == row:
                self.variables.pop(i)
                break
        row.destroy()
        self.refresh_var_list()

    def start_drag(self, event, row):
        pass

    def on_drag(self, event, row):
        # find index of current row
        curr_idx = -1
        for i, item in enumerate(self.variables):
            if item["row"] == row:
                curr_idx = i
                break
        
        if curr_idx == -1: return

        y = event.widget.winfo_rooty() + event.y
        
        target_idx = -1
        for i, item in enumerate(self.variables):
            r = item["row"]
            r_y = r.winfo_rooty()
            r_h = r.winfo_height()
            
            if r_y < y < r_y + r_h:
                target_idx = i
                break
        
        if target_idx != -1 and target_idx != curr_idx:
            item = self.variables.pop(curr_idx)
            self.variables.insert(target_idx, item)
            self.refresh_var_list()
    
    def toggle_group(self, group_name):
        curr = self.group_collapsed.get(group_name, False)
        self.group_collapsed[group_name] = not curr
        self.refresh_var_list()

    def refresh_var_list(self):
        """Stable refresh to prevent flickering."""
        # Calculate the desired order of all elements
        desired_widgets = []
        current_group = None
        group_data = {} # To store header info if we were to recreate it
        
        for item in self.variables:
            g_name = item["g"].get().strip() or "未分组"
            if g_name != current_group:
                current_group = g_name
                group_data[g_name] = {
                    "collapsed": self.group_collapsed.get(g_name, False)
                }
                # Find or create header
                # For simplicity in this optimization, we'll still recreate headers
                # but we'll minimize the pack_forget calls.
            
        # Optimization: Clear everything once, but use update_idletasks
        # Actually, the most reliable way to avoid flicker in CTk is to not clear everything.
        # But headers are dynamic based on grouping.
        
        if hasattr(self, "_group_headers"):
            for h in self._group_headers:
                h.destroy()
        self._group_headers = []

        # Hide rows temporarily
        for item in self.variables:
            item["row"].pack_forget()

        current_group = None
        for item in self.variables:
            g_name = item["g"].get().strip() or "未分组"
            if g_name != current_group:
                current_group = g_name
                is_collapsed = self.group_collapsed.get(current_group, False)
                
                header = ctk.CTkFrame(self.vars_scroll, fg_color="#333333", height=36, corner_radius=6, cursor="hand2")
                header.pack(fill="x", pady=(12, 4), padx=2)
                
                icon = "   ▸" if is_collapsed else "   ▾"
                lbl_icon = ctk.CTkLabel(header, text=icon, font=("Arial", 14, "bold"), text_color="#3B8ED0")
                lbl_icon.pack(side="left", padx=(5, 5))
                
                lbl_name = ctk.CTkLabel(header, text=current_group.upper(), font=("Arial", 12, "bold"), text_color="#E0E0E0")
                lbl_name.pack(side="left", padx=5)
                
                btn_clone_group = ctk.CTkButton(header, text="📑 克隆组", width=70, height=24, font=("Arial", 11), 
                                               fg_color="#2ECC71", hover_color="#27AE60",
                                               command=lambda gn=current_group: self.clone_group(gn))
                btn_clone_group.pack(side="right", padx=10)

                btn_batch_rename = ctk.CTkButton(header, text="✏️ 重命名", width=80, height=24, font=("Arial", 11), 
                                                fg_color="#3498DB", hover_color="#2980B9",
                                                command=lambda gn=current_group: self.rename_group(gn))
                btn_batch_rename.pack(side="right", padx=5)
                
                header.bind("<Button-1>", lambda e, gn=current_group: self.toggle_group(gn))
                lbl_icon.bind("<Button-1>", lambda e, gn=current_group: self.toggle_group(gn))
                lbl_name.bind("<Button-1>", lambda e, gn=current_group: self.toggle_group(gn))
                
                self._group_headers.append(header)
            
            if not self.group_collapsed.get(current_group, False):
                item["row"].pack(fill="x", pady=2)
        
        # This is the key to reducing perceived flicker
        self.vars_scroll.update_idletasks()
            
    def clear_vars(self):
        for item in self.variables:
            item["row"].destroy()
        self.variables = []
        self.refresh_var_list()

    def validate_all_keys(self):
        """Check for duplicates and highlight them."""
        counts = {}
        for item in self.variables:
            k = item["k"].get().strip()
            if k:
                counts[k] = counts.get(k, 0) + 1
        
        has_duplicate = False
        for item in self.variables:
            k = item["k"].get().strip()
            if k and counts.get(k, 0) > 1:
                item["k"].configure(border_color="red")
                has_duplicate = True
            else:
                # Need to use the default color. CTk defaults vary by theme.
                # Let's reset to a standard border color.
                item["k"].configure(border_color=["#979da2", "#565b5e"]) 
        return has_duplicate
    
    def copy_to_clipboard_bracket(self, entry):
        key_raw = entry.get().strip()
        if key_raw:
            key_formatted = f"{{{{ {key_raw} }}}}"
            self.clipboard_clear()
            self.clipboard_append(key_formatted)
            self.update() # Required for some systems
            # Show a brief status or log
            if hasattr(self, "log"):
                self.log(f"已复制: {key_formatted}")
            tkinter.messagebox.showinfo("成功", f"已复制到剪贴板: {key_formatted}")

    def copy_group_keys(self, group_name):
        """Copy all keys in a group as {{key1}} {{key2}} ..."""
        keys = []
        for item in self.variables:
            g = item["g"].get().strip() or "未分组"
            if g == group_name:
                k = item["k"].get().strip()
                if k:
                    keys.append(f"{{{{ {k} }}}}")
        
        if keys:
            all_keys = "\n".join(keys) # One per line or space? Let's do newline for better pasting in some contexts
            self.clipboard_clear()
            self.clipboard_append(all_keys)
            self.update()
            self.log(f"已批量复制组 [{group_name}] 的 {len(keys)} 个变量")
            tkinter.messagebox.showinfo("成功", f"已批量复制组 [{group_name}] 的变量名到剪贴板")

    def clone_group(self, source_group_name):
        """Duplicate an entire group with auto-incremented name."""
        # 1. Identify variables in the group
        to_clone = []
        for item in self.variables:
            g = item["g"].get().strip() or "未分组"
            if g == source_group_name:
                to_clone.append({
                    "k": item["k"].get().strip(),
                    "v": item["v"].get("1.0", "end-1c").strip()
                })
        
        if not to_clone:
            return

        # 2. Determine next group name
        # Try to find a trailing number (e.g. "Unit01" -> "Unit", "01")
        match = re.search(r'^(.*?)(\d+)$', source_group_name)
        if match:
            base = match.group(1)
            num_str = match.group(2)
            width = len(num_str)
            
            # Find the highest number for this base
            max_num = int(num_str)
            all_groups = {item["g"].get().strip() for item in self.variables}
            for g in all_groups:
                m = re.match(rf'^{re.escape(base)}(\d+)$', g)
                if m:
                    max_num = max(max_num, int(m.group(1)))
            
            new_num = max_num + 1
            new_group_name = f"{base}{str(new_num).zfill(width)}"
        else:
            # Fallback for names without trailing numbers: "Group" -> "Group_2"
            base = source_group_name
            max_num = 1
            all_groups = {item["g"].get().strip() for item in self.variables}
            for g in all_groups:
                m = re.match(rf'^{re.escape(base)}_(\d+)$', g)
                if m:
                    max_num = max(max_num, int(m.group(1)))
            
            new_num = max_num + 1
            new_group_name = f"{base}_{new_num}"

        # 3. Create new variables
        # Fixed keys, cleared values as requested
        for item in to_clone:
            self.add_variable_row(item["k"], "", new_group_name)
            
        self.log(f"已克隆组 [{source_group_name}] -> [{new_group_name}] (Key不变，Val已清空)")
        self.refresh_var_list()

    def rename_group(self, old_group_name):
        """Rename all items in a group to a new group name."""
        # 1. Ask for new name using a simple dialog (assuming user wants to replace the whole name)
        # We use CTkInputDialog which prompts for a single string.
        dialog = ctk.CTkInputDialog(text=f"请输入新的组名 (原: {old_group_name}):", title="重命名组")
        new_name = dialog.get_input()
        
        if new_name is None: return # Cancelled
        new_name = new_name.strip()
        if not new_name: return # Empty check
        if new_name == old_group_name: return

        count = 0
        for item in self.variables:
            g = item["g"].get().strip() or "未分组"
            if g == old_group_name:
                item["g"].delete(0, "end")
                item["g"].insert(0, new_name)
                count += 1
        
        if count > 0:
            self.refresh_var_list()
            self.log(f"组 [{old_group_name}] 已重命名为 [{new_name}]")
        else:
            tkinter.messagebox.showinfo("提示", "没有找到该组的变量")

    def save_vars_json(self):
        if self.validate_all_keys():
            tkinter.messagebox.showerror("错误", "变量名存在重复，请修改后再保存！")
            return

        fpath = filedialog.asksaveasfilename(defaultextension=".json", initialfile="变量配置.json", filetypes=[("JSON", "*.json")], initialdir=os.getcwd())
        if not fpath:
            return

        json_dir = os.path.dirname(os.path.abspath(fpath))
        data = []
        for item in self.variables:
            key = item["k"].get().strip()
            val = item["v"].get("1.0", "end-1c").strip()
            group = item["g"].get().strip()
            if key:
                # Try to make path relative if it exists
                if val and os.path.exists(val):
                    try:
                        val = os.path.relpath(val, json_dir)
                    except ValueError:
                        pass # Keep absolute if different drive
                data.append({"key": key, "val": val, "group": group})
        
        try:
            with open(fpath, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            self.last_saved_vars = data # Sync state
            self.log("配置已保存到: " + fpath)
        except Exception as e:
            tkinter.messagebox.showerror("错误", str(e))

    def load_vars_json(self, fpath=None):
        if not fpath:
            fpath = filedialog.askopenfilename(filetypes=[("JSON", "*.json")], initialdir=os.getcwd())
        
        if not fpath:
            return

        json_dir = os.path.dirname(os.path.abspath(fpath))
        try:
            with open(fpath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.clear_vars()
            
            # Support both old flat dict and new list of dicts with groups
            if isinstance(data, dict):
                for k, v in data.items():
                    self.add_variable_row(k, v)
            elif isinstance(data, list):
                for item in data:
                    k = item.get("key", "")
                    v = item.get("val", "")
                    g = item.get("group", "")
                    
                    # Try to resolve relative path if it looks like a path
                    if isinstance(v, str) and v and not os.path.isabs(v):
                         abs_path = os.path.abspath(os.path.join(json_dir, v))
                         if os.path.exists(abs_path):
                             try:
                                 v = os.path.relpath(abs_path, os.getcwd())
                             except ValueError:
                                 v = abs_path
                                 v = abs_path
                    self.add_variable_row(k, v, g)

            # Auto-collapse all groups on initial load
            all_groups = set()
            for item in self.variables:
                g_val = item["g"].get().strip()
                if g_val:
                    all_groups.add(g_val)
            
            for g_name in all_groups:
                self.group_collapsed[g_name] = True
            
            self.refresh_var_list()
            self.last_saved_vars = self._get_current_vars_data() # Sync state correctly
            self.log(f"配置已加载: {fpath} (共 {len(self.variables)} 个变量, 分组默认收缩)")
            

        except Exception as e:
            tkinter.messagebox.showerror("错误", str(e))

    def scan_variables_from_file(self):
        fpaths = filedialog.askopenfilenames(filetypes=[("Word", "*.docx")])
        if not fpaths: return
        
        mgr = TemplateManager(".")
        found = set()
        for f in fpaths:
            try:
                found.update(mgr.get_template_variables(f))
            except Exception as e:
                print(e)
        
        # Add only new
        existing = {item["k"].get().strip() for item in self.variables}
        count = 0
        for v in found:
            if v not in existing:
                self.add_variable_row(v, "")
                count += 1
        tkinter.messagebox.showinfo("扫瞄完成", f"新增了 {count} 个变量。")


    # --- Tables Logic ---

    def add_table_row(self, name, path, sheet="", columns_config=None, source_type="file"):
        # 1. Determine Parent Container
        parent_item = None
        if source_type == "table":
            # path is parent_name
            for item in self.table_configs:
                if item["name"].get().strip() == path:
                    parent_item = item
                    break
        
        master = self.tables_scroll
        is_child = False
        if parent_item:
            master = parent_item["children_frame"]
            is_child = True

        # 2. Structure: 
        # wrapper (contains content + children) -> for Parent
        # For Child: Just content frame? 
        # To support multi-level (if ever needed), let's make every row capable of having children.
        
        wrapper = ctk.CTkFrame(master, fg_color="transparent")
        wrapper.pack(fill="x", pady=2, padx=0 if is_child else 5)
        
        # Content Row (The visual part)
        # Indent if child?
        bg_color = "#2B2B2B" if not is_child else "#252525"
        row = ctk.CTkFrame(wrapper, fg_color=bg_color, corner_radius=8)
        row.pack(fill="x", pady=0, padx=0)
        
        if is_child:
            # Add a visual connector or indent
            # Use margin on left inside the row?
            pass

        # --- Row Content ---
        # Drag Handle
        lbl_handle = ctk.CTkLabel(row, text="⠿", width=30, cursor="hand2", text_color="gray")
        lbl_handle.pack(side="left", padx=2, pady=8)
        # Only enable drag for root items for now to avoid complexity of dragging children out
        if not is_child:
            lbl_handle.bind("<Button-1>", lambda e, r=wrapper: self.start_table_drag(e, r))
            lbl_handle.bind("<B1-Motion>", lambda e, r=wrapper: self.on_table_drag(e, r))
        else:
            lbl_handle.configure(text="   ", cursor="arrow") # Hide handle for children or make non-draggable

        # Variable Name
        k_ent = ctk.CTkEntry(row, placeholder_text="表格变量名", width=150, border_width=1, font=("Arial", 12, "bold"))
        k_ent.pack(side="left", padx=(5, 5), pady=8)
        k_ent.insert(0, name)
        k_ent.bind("<Double-Button-1>", lambda e, w=k_ent: self.copy_to_clipboard_bracket(w))
        
        # Source Display/Input
        p_ent = ctk.CTkEntry(row, placeholder_text="Excel 文件路径", border_width=1, fg_color="#1E1E1E")
        s_ent = ctk.CTkEntry(row, placeholder_text="Sheet名", width=100, border_width=1)
        
        if source_type == "file":
            p_ent.pack(side="left", padx=5, pady=8, fill="x", expand=True)
            p_ent.insert(0, path)

            btn_file = ctk.CTkButton(row, text="📊 选择文件", width=80, height=28, 
                                     command=lambda e=p_ent: self.select_excel_file(e))
            btn_file.pack(side="left", padx=2, pady=8)
            
            s_ent.pack(side="left", padx=5, pady=8)
            s_ent.insert(0, sheet)
            
        else: # source_type == "table"
            # Indent indicator or icon?
            icon_label = ctk.CTkLabel(row, text="↳", font=("Arial", 16, "bold"), text_color="#27AE60", width=20)
            icon_label.pack(side="left", padx=(0, 0))

            lbl_source = ctk.CTkLabel(row, text=f"源: {path}", fg_color="#333333", corner_radius=4, width=150, anchor="w")
            lbl_source.pack(side="left", padx=5, pady=8, fill="x", expand=True)
            
            # Hidden entries
            p_ent = ctk.CTkEntry(row) 
            p_ent.insert(0, path)
            s_ent = ctk.CTkEntry(row) # Dummy

        # Actions
        btn_config = ctk.CTkButton(row, text="⚙️ 配置", width=60, height=28, 
                                     fg_color="#F39C12", hover_color="#D35400",
                                     command=lambda: self.open_column_config(row, p_ent.get() if source_type=="file" else None, source_type=source_type, source_value=p_ent.get()))
        btn_config.pack(side="left", padx=5, pady=8)

        # Removed Clone button

        # Link (Create derived table) - Only allow linking from Parent? Or chaining? 
        # Chaining is fine.
        btn_link = ctk.CTkButton(row, text="🔗", width=40, height=28, 
                                     fg_color="#27AE60", hover_color="#2ECC71",
                                     command=lambda: self.add_derived_table(k_ent.get()))
        btn_link.pack(side="left", padx=5, pady=8)

        btn_preview = ctk.CTkButton(row, text="👁", width=40, height=28, 
                                     fg_color="#3498DB", hover_color="#2980B9",
                                     command=lambda: self.preview_table_data(p_ent.get(), s_ent.get(), k_ent.get(), row, source_type=source_type))
        btn_preview.pack(side="left", padx=5, pady=8)

        btn_del = ctk.CTkButton(row, text="✕", width=30, height=28, fg_color="transparent", 
                                 text_color="#E74C3C", hover_color="#3D3D3D", 
                                 command=lambda r=wrapper: self.remove_table_row(r))
        btn_del.pack(side="right", padx=(5, 10), pady=8)

        # Children Container
        children_frame = ctk.CTkFrame(wrapper, fg_color="transparent", height=0)
        children_frame.pack(fill="x", padx=(40, 0), pady=0) # Indent children

        self.table_configs.append({
            "name": k_ent, 
            "path": p_ent, 
            "sheet": s_ent, 
            "row": row, # Config dialog uses this to identify config. We should use 'row' (content frame) for that.
            "wrapper": wrapper, # For removal and dragging
            "children_frame": children_frame,
            "columns_config": columns_config or {},
            "source_type": source_type
        })

    def start_table_drag(self, event, row):
        pass # Placeholder for start drag visualization if needed

    def on_table_drag(self, event, row):
        # row passed here is now the wrapper for top-level items
        
        # find index of current wrapper in top level
        # This is tricky because self.table_configs contains ALL items (flat list).
        # We need to find the subset of top-level items. Or just drag based on visual list?
        
        # Simpler approach: Iterate through widgets in tables_scroll.
        # But we need to update table_configs order to match visual order for persistence.
        
        # Let's find index in table_configs
        curr_idx = -1
        for i, item in enumerate(self.table_configs):
            if item["wrapper"] == row:
                curr_idx = i
                break
        
        if curr_idx == -1: return

        y = event.widget.winfo_rooty() + event.y
        
        target_idx = -1
        
        # Scan only other top-level wrappers?
        # If we just swap in table_configs, it might mess up if table_configs includes children.
        # Actually table_configs order determines load order. Visual order is just pack order.
        
        # Current implementation of on_table_drag relied on flat list. 
        # With nesting, we can only drag root items.
        
        # Check collision with other items in table_configs that are ROOTS (source_type="file" or orphan)
        # Actually checking widget collision is easier.
        
        for i, item in enumerate(self.table_configs):
            r = item["wrapper"] # Check against wrappers
            # Only consider items that are siblings in the UI (same master)
            if r.master != row.master: continue
            
            r_y = r.winfo_rooty()
            r_h = r.winfo_height()
            
            if r_y < y < r_y + r_h:
                target_idx = i
                break
        
        if target_idx != -1 and target_idx != curr_idx:
            # Swap in data list
            item = self.table_configs.pop(curr_idx)
            self.table_configs.insert(target_idx, item)
            
            # Re-pack visually
            # We need to re-pack ALL siblings in the new order.
            master = row.master
            siblings = [item for item in self.table_configs if item["wrapper"].master == master]
            
            for sib in siblings:
                sib["wrapper"].pack_forget()
                sib["wrapper"].pack(fill="x", pady=2, padx=5 if master == self.tables_scroll else 0)
            
            self.tables_scroll.update_idletasks()

    def remove_table_row(self, row_wrapper):
        # row_wrapper is the wrapper frame
        # Find item
        idx = -1
        for i, item in enumerate(self.table_configs):
            if item["wrapper"] == row_wrapper:
                idx = i
                break
        
        if idx != -1:
            # Also need to remove children?
            # If we remove a parent, children are visually destroyed (inside wrapper), 
            # but we must remove them from table_configs too.
            parent_item = self.table_configs[idx]
            parent_name = parent_item["name"].get().strip()
            
            # Find children
            to_remove = [idx]
            for i, item in enumerate(self.table_configs):
                 if item.get("source_type") == "table" and item["path"].get().strip() == parent_name:
                     to_remove.append(i)
            
            # Remove indices in reverse order to stay valid
            for i in sorted(to_remove, reverse=True):
                self.table_configs.pop(i)
                
        row_wrapper.destroy()

    def add_derived_table(self, parent_name):
        if not parent_name:
            tkinter.messagebox.showwarning("提示", "请先定义父表名称")
            return
        
        new_name = parent_name + "_sub"
        self.add_table_row(new_name, parent_name, "", {}, source_type="table")
        self.log(f"已创建关联表: {new_name} (源于 {parent_name})")

    def open_column_config(self, row_frame, excel_path, source_type="file", source_value=None):
        # source_value is path or parent_name
        
        # Load headers
        headers = []
        try:
            if source_type == "file":
                 if not excel_path or not os.path.exists(excel_path):
                     tkinter.messagebox.showerror("错误", "请先选择有效的 Excel 文件")
                     return
                 
                 loader = ExcelLoader(excel_path)
                 
                 # Need sheet name
                 target_config = None
                 for item in self.table_configs:
                     if item["row"] == row_frame:
                         target_config = item
                         break
                 if not target_config: return
                 
                 sheet = target_config["sheet"].get().strip()
                 data = loader.load_data(sheet_name=sheet if sheet else None)
                 headers = list(data[0].keys()) if data else []
                 
            else:
                 # Table source
                 parent_name = source_value
                 # Use the recursive loader to get parent data
                 data = self._load_table_data_recursive(parent_name)
                 headers = list(data[0].keys()) if data else []

        except Exception as e:
            tkinter.messagebox.showerror("错误", f"读取数据失败: {e}")
            return
            
        # Find config config for this row (duplicated search, but ok)
        target_config = None
        current_config_data = {}
        for item in self.table_configs:
            if item["row"] == row_frame:
                target_config = item
                current_config_data = item.get("columns_config", {})
                break
            
        if not target_config: return

        dialog = TableConfigDialog(self, headers, current_config_data)
        self.wait_window(dialog)
        
        if dialog.result is not None:
            target_config["columns_config"] = dialog.result
            self.log(f"已更新表格配置: 隐藏 {len(dialog.result.get('hidden', []))} 列, 新增 {len(dialog.result.get('new', []))} 计算列, 包含筛选: {bool(dialog.result.get('filter_formula'))}")

    def clone_table_config(self, row_frame):
        # Find config to clone
        source = None
        for item in self.table_configs:
            if item["row"] == row_frame:
                source = item
                break
        if not source: return

        # Get values
        old_name = source["name"].get().strip()
        path = source["path"].get().strip()
        sheet = source["sheet"].get().strip()
        config = source.get("columns_config", {}).copy() # Deep copy if needed? simple dict so shallow copy ok for now if structure simple.
        # Deep copy config just in case
        import copy
        config = copy.deepcopy(config)
        
        # New name
        new_name = old_name + "_copy"
        
        self.add_table_row(new_name, path, sheet, config)
        self.log(f"已克隆表格配置: {old_name} -> {new_name}")

    def process_table_data(self, data, config):
        """Apply column hiding, new column calculations, and ROW FILTERING."""
        if not data: return []
        if not config: return data
        
        hidden = set(config.get("hidden", []))
        new_cols = config.get("new", [])
        filter_formula = config.get("filter_formula", "").strip()
        
        processed_data = []
        for i, row in enumerate(data):
            new_row = row.copy()
            
            # 0. Auto Index
            if config.get("auto_index"):
                idx_name = config.get("auto_index_name", "序号")
                new_row[idx_name] = i + 1
            
            # 1. Calculate new columns
            for col_def in new_cols:
                name = col_def["name"]
                formula = col_def["formula"]
                
                if not formula:
                    new_row[name] = ""
                    continue

                # Replace {{ key }} with row.get('key', 0)
                def replacer(match):
                    key = match.group(1).strip()
                    val = new_row.get(key, 0)
                    try:
                        return str(float(val))
                    except:
                        return f"'{str(val)}'" 
                        
                expression = re.sub(r'\{\{(.*?)\}\}', replacer, formula)
                try:
                    res = eval(expression, {"__builtins__": None}, {})
                    new_row[name] = res
                except Exception:
                    new_row[name] = "Calc Error"

            # 2. Row Filter
            if filter_formula:
                def replacer_filter(match):
                    key = match.group(1).strip()
                    val = new_row.get(key, 0)
                    try:
                        return str(float(val))
                    except:
                        return f"'{str(val)}'"
                
                filter_expr = re.sub(r'\{\{(.*?)\}\}', replacer_filter, filter_formula)
                try:
                    is_keep = eval(filter_expr, {"__builtins__": None}, {})
                    if not is_keep:
                        continue # Skip this row
                except Exception:
                    # On error, maybe keep? or skip? Let's skip and maybe log?
                    # For UI simplicity, we skip on error or false.
                    continue

            # 3. Hide columns
            final_row = {}
            for k, v in new_row.items():
                if k not in hidden:
                    final_row[k] = v
            
            processed_data.append(final_row)
            
        return processed_data


    def _get_table_item_by_name(self, name):
        for item in self.table_configs:
             if item["name"].get().strip() == name:
                 return item
        return None

    def _load_table_data_recursive(self, name, seen=None):
        if seen is None: seen = set()
        if name in seen:
            raise Exception(f"检测到循环依赖: {name}")
        seen.add(name)

        item = self._get_table_item_by_name(name)
        if not item:
            raise Exception(f"找不到表格变量: {name}")

        src_type = item.get("source_type", "file")
        cols_cfg = item.get("columns_config", {})
        
        raw_data = []
        if src_type == "file":
             path = item["path"].get().strip()
             sheet = item["sheet"].get().strip()
             if not path or not os.path.exists(path):
                 raise Exception(f"文件不存在: {path}")
             
             loader = ExcelLoader(path)
             raw_data = loader.load_data(sheet_name=sheet if sheet else None)
        else:
             # Derived table
             parent_name = item["path"].get().strip()
             raw_data = self._load_table_data_recursive(parent_name, seen)
        
        # Apply processing (Calc + Filter + Hide)
        return self.process_table_data(raw_data, cols_cfg)

    def preview_table_data(self, path, sheet, name, row_frame=None, source_type="file"):
        # 1. Clear existing preview
        for widget in self.preview_scroll.winfo_children():
            widget.destroy()
            
        display_name = name if name else "当前表格"
        self.lbl_preview_title.configure(text=f"数据预览: {display_name}")

        try:
            # Determine which table we are previewing
            # If row_frame is passed, use it to identify the table accurately in config (though not strictly needed if name is unique)
            # Actually better to rely on name since our recursive loader uses name
            # But name in Entry might be edited?
            # Safe to assume name in text field is what user intends.
            
            target_name = name
            if not target_name:
                ctk.CTkLabel(self.preview_scroll, text="⚠️ 请先输入表格变量名", text_color="orange").pack(pady=20)
                return

            # Use recursive loader
            data = self._load_table_data_recursive(target_name)

            if not data:
                ctk.CTkLabel(self.preview_scroll, text="⚠️ 表格为空或数据被全部过滤", text_color="orange").pack(pady=20)
                return

            # 2. Render Grid
            # Limit to 50 rows for performance
            display_data = data[:50]
            headers = list(data[0].keys()) if data else []
            
            if not headers:
                ctk.CTkLabel(self.preview_scroll, text="⚠️ 未找到表头 (或全部被隐藏)", text_color="orange").pack(pady=20)
                return

            # Render Headers
            for col, header in enumerate(headers):
                h_lbl = ctk.CTkLabel(self.preview_scroll, text=str(header), font=("Arial", 12, "bold"), 
                                     fg_color="#2B2B2B", corner_radius=4, width=100, anchor="w")
                h_lbl.grid(row=0, column=col, padx=1, pady=1, sticky="ew")

            # Render Rows
            for r, row_data in enumerate(display_data):
                for c, header in enumerate(headers):
                    val = row_data.get(header, "")
                    if isinstance(val, str): val = val.strip()
                    
                    cell = ctk.CTkEntry(self.preview_scroll, width=100, border_width=0, fg_color="transparent")
                    cell.grid(row=r+1, column=c, padx=1, pady=1, sticky="ew")
                    cell.insert(0, str(val))
                    cell.configure(state="readonly")
            
            if len(data) > 50:
                 ctk.CTkLabel(self.preview_scroll, text=f"... 仅展示前 50 行 (共 {len(data)} 行) ...", text_color="gray").grid(row=51, column=0, columnspan=len(headers), pady=10)

        except Exception as e:
            ctk.CTkLabel(self.preview_scroll, text=f"❌ 预览错误: {str(e)}", text_color="#E74C3C").pack(pady=20)

    def select_excel_file(self, entry):
        f = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")], initialdir=os.getcwd())
        if f:
            try:
                rel_f = os.path.relpath(f, os.getcwd())
            except ValueError:
                rel_f = f
            entry.delete(0, "end")
            entry.insert(0, rel_f)

    def remove_table_row(self, row_wrapper):
        # row_wrapper is the wrapper frame (for parent) or row frame?
        # In add_table_row, we bound remove_table_row(wrapper).
        
        # 1. Find the item corresponding to this wrapper
        target_idx = -1
        for i, item in enumerate(self.table_configs):
            if item["wrapper"] == row_wrapper:
                target_idx = i
                break
        
        if target_idx == -1: return

        # 2. Find all descendants recursively
        # We need to find items where path == name of parent (source_type="table")
        parent_name = self.table_configs[target_idx]["name"].get().strip()
        
        # Iterative or recursive search for indices to remove
        indices_to_remove = {target_idx}
        
        # Simple approach: Loop until no new children found (since depth is small)
        # or just recursion.
        
        def find_children_indices(p_name):
            children = []
            for i, item in enumerate(self.table_configs):
                if i in indices_to_remove: continue
                if item.get("source_type") == "table" and item["path"].get().strip() == p_name:
                    children.append(i)
            return children

        queue = [parent_name]
        while queue:
            curr_p = queue.pop(0)
            children_idxs = find_children_indices(curr_p)
            for child_i in children_idxs:
                indices_to_remove.add(child_i)
                # Add child name to queue to find grandchildren
                child_name = self.table_configs[child_i]["name"].get().strip()
                queue.append(child_name)

        # 3. Remove from config list (Reverse order to keep indices valid)
        for i in sorted(list(indices_to_remove), reverse=True):
            self.table_configs.pop(i)

        # 4. Destroy UI
        # Destroying the wrapper kills children UI automatically if they are nested.
        row_wrapper.destroy()
        
        # 5. Refresh Layout (Important!)
        self.tables_scroll.update_idletasks()


    def clear_tables(self):
        for item in self.table_configs:
            item["wrapper"].destroy()
        self.table_configs = []

    def save_tables_json(self):
        fpath = filedialog.asksaveasfilename(defaultextension=".json", initialfile="表格配置.json", filetypes=[("JSON", "*.json")], initialdir=os.getcwd())
        if not fpath: return

        json_dir = os.path.dirname(os.path.abspath(fpath))
        data = []
        for item in self.table_configs:
            name = item["name"].get().strip()
            path = item["path"].get().strip() # path or parent name
            sheet = item["sheet"].get().strip()
            cols_cfg = item.get("columns_config", {})
            src_type = item.get("source_type", "file")
            
            if name: # allow empty path for validation later, but here we just save
                # Relative path logic ONLY for file type
                path_val = path
                if src_type == "file" and path and os.path.exists(path):
                    try:
                        path_val = os.path.relpath(path, json_dir)
                    except ValueError:
                        path_val = path
                    
                data.append({
                    "name": name, 
                    "path": path_val, 
                    "sheet": sheet, 
                    "columns_config": cols_cfg,
                    "source_type": src_type
                })
        
        try:
            with open(fpath, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            self.last_saved_tables = data
            self.log("表格配置已保存: " + fpath)
        except Exception as e:
            tkinter.messagebox.showerror("错误", str(e))

    def load_tables_json(self, fpath=None):
        if not fpath:
            fpath = filedialog.askopenfilename(filetypes=[("JSON", "*.json")], initialdir=os.getcwd())
        if not fpath: return

        print(f"正在加载表格配置: {fpath}")
        json_dir = os.path.dirname(os.path.abspath(fpath))
        try:
            with open(fpath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Print data snippet for debugging
            print(f"读取到 {len(data)} 个表格配置项")
            
            self.clear_tables()
            
            for item in data:
                name = item.get("name", "")
                path = item.get("path", "")
                sheet = item.get("sheet", "")
                cols_cfg = item.get("columns_config", {})
                src_type = item.get("source_type", "file")
                
                # Resolve relative path only for files
                if src_type == "file" and path and not os.path.isabs(path):
                     abs_path = os.path.abspath(os.path.join(json_dir, path))
                     if os.path.exists(abs_path):
                         try:
                             path = os.path.relpath(abs_path, os.getcwd())
                         except ValueError:
                             path = abs_path
                             
                self.add_table_row(name, path, sheet, cols_cfg, src_type)

            self.last_saved_tables = self._get_current_tables_data()
            self.log(f"成功加载配置: {fpath} ({len(data)} 个表)")
        except Exception as e:
            err_msg = f"加载配置失败: {str(e)}"
            print(err_msg)
            self.log(err_msg)
            tkinter.messagebox.showerror("错误", err_msg)
            
    def _get_current_vars_data(self):
        # Helper for state tracking
        data = []
        for item in self.variables:
             data.append({
                 "key": item["k"].get().strip(),
                 "val": item["v"].get("1.0", "end-1c").strip(),
                 "group": item["g"].get().strip()
             })
        return data

    def _get_current_tables_data(self):
        data = []
        for item in self.table_configs:
             data.append({
                 "name": item["name"].get().strip(),
                 "path": item["path"].get().strip(),
                 "sheet": item["sheet"].get().strip(),
                 "columns_config": item.get("columns_config", {}),
                 "source_type": item.get("source_type", "file")
             })
        return data
    
    def _get_current_mappings_data(self):
        # Helper for state tracking
        data = []
        for in_ent, out_ent, _ in self.file_mappings:
             data.append({"input": in_ent.get().strip(), "output": out_ent.get().strip()})
        return data


    # --- File Mapping Logic ---

    def add_mapping_row(self, input_path, output_path):
        row = ctk.CTkFrame(self.files_scroll, fg_color="#2B2B2B", corner_radius=8)
        row.pack(fill="x", pady=4, padx=5)
        
        # Input
        in_ent = ctk.CTkEntry(row, placeholder_text="模板文件路径 (.docx/.doc)", height=32, border_width=1)
        in_ent.pack(side="left", padx=(10, 5), pady=8, fill="x", expand=True)
        in_ent.insert(0, input_path)
        
        btn_in = ctk.CTkButton(row, text="🔍 选择模板", width=90, height=30, command=lambda e=in_ent: self.select_input_file(e),
                                fg_color="#3D3D3D", hover_color="#4D4D4D")
        btn_in.pack(side="left", padx=2, pady=8)

        # Arrow icon
        ctk.CTkLabel(row, text="➡", font=("Arial", 14), text_color="#3B8ED0").pack(side="left", padx=8, pady=8)

        # Output
        out_ent = ctk.CTkEntry(row, placeholder_text="导出目录路径", height=32, border_width=1)
        out_ent.pack(side="left", padx=5, pady=8, fill="x", expand=True)
        out_ent.insert(0, output_path)

        btn_out_file = ctk.CTkButton(row, text="📂 导出目录", width=90, height=30, command=lambda e=out_ent: self.select_output_dir(e),
                                      fg_color="#3D3D3D", hover_color="#4D4D4D")
        btn_out_file.pack(side="left", padx=2, pady=8)

        # Remove
        btn_del = ctk.CTkButton(row, text="✕", width=30, height=30, fg_color="transparent", 
                                 text_color="#E74C3C", hover_color="#3D3D3D", 
                                 command=lambda r=row: self.remove_mapping_row(r))
        btn_del.pack(side="right", padx=(5, 10), pady=8)

        self.file_mappings.append((in_ent, out_ent, row))

    def select_input_file(self, entry):
        f = filedialog.askopenfilename(filetypes=[("Word Docs", "*.docx;*.doc")], initialdir=os.getcwd())
        if f:
            try:
                rel_f = os.path.relpath(f, os.getcwd())
            except ValueError:
                rel_f = f
            entry.delete(0, "end")
            entry.insert(0, rel_f)
            # Auto-suggest output? 
            # Could replace .docx with _generated.docx
            # But let's leave flexible.

    def select_output_dir(self, entry):
        f = filedialog.askdirectory(initialdir=os.getcwd())
        if f:
            try:
                rel_f = os.path.relpath(f, os.getcwd())
            except ValueError:
                rel_f = f
            entry.delete(0, "end")
            entry.insert(0, rel_f)

    def remove_mapping_row(self, row):
        for i, item in enumerate(self.file_mappings):
            if item[2] == row:
                self.file_mappings.pop(i)
                break
        row.destroy()

    def clear_mappings(self):
        for item in self.file_mappings:
            item[2].destroy()
        self.file_mappings = []

    def save_mappings_json(self):
        fpath = filedialog.asksaveasfilename(defaultextension=".json", initialfile="生成配置.json", filetypes=[("JSON", "*.json")], initialdir=os.getcwd())
        if not fpath:
            return
            
        json_dir = os.path.dirname(os.path.abspath(fpath))
        data = []
        for in_ent, out_ent, _ in self.file_mappings:
            in_path = in_ent.get().strip()
            out_path = out_ent.get().strip()
            
            if in_path or out_path:
                # Convert to relative path if possible
                try:
                    rel_in = os.path.relpath(in_path, json_dir)
                except ValueError:
                    rel_in = in_path # Fallback to absolute if different drive

                try:
                    rel_out = os.path.relpath(out_path, json_dir)
                except ValueError:
                    rel_out = out_path

                data.append({"input": rel_in, "output": rel_out})
        
        try:
            with open(fpath, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            self.last_saved_mappings = data # Sync state
            self.log("映射配置已保存到: " + fpath)
        except Exception as e:
            tkinter.messagebox.showerror("错误", str(e))

    def copy_to_clipboard_bracket(self, widget):
        val = widget.get().strip()
        if val:
            template = f"{{{{ {val} }}}}"
            self.clipboard_clear()
            self.clipboard_append(template)
            self.update() # Keep clipboard
            # Optional: visual feedback?
            print(f"Copied to clipboard: {template}")
            # Flash the widget?
            orig_color = widget.cget("fg_color")
            widget.configure(fg_color="#3B8ED0")
            self.after(200, lambda: widget.configure(fg_color=orig_color))

    # --- Unified Check Logic for Save Prompt ---
    def has_unsaved_changes(self):
         # Check all
         return (self._get_current_vars_data() != self.last_saved_vars or
                 self._get_current_tables_data() != self.last_saved_tables or
                 self._get_current_mappings_data() != self.last_saved_mappings)

    def save_project(self, fpath=None):
        if self.validate_all_keys():
            tkinter.messagebox.showerror("错误", "变量名存在重复，请修改后再保存！")
            return

        if not fpath:
            fpath = filedialog.asksaveasfilename(defaultextension=".json", initialfile="project.json", filetypes=[("Project JSON", "*.json")], initialdir=os.getcwd())
        if not fpath: return

        json_dir = os.path.dirname(os.path.abspath(fpath))
        
        # 1. Prepare Data
        # We reuse the helper methods but need to adjust relative paths based on the new json_dir
        
        # Variables (Paths in value)
        vars_data = []
        for item in self.variables:
            key = item["k"].get().strip()
            val = item["v"].get("1.0", "end-1c").strip()
            group = item["g"].get().strip()
            if key:
                if val and os.path.exists(val):
                    try:
                         # Handle image path with params like |width=full
                         real_path = val.split("|")[0]
                         params = val[len(real_path):]
                         if os.path.isabs(real_path):
                             rel = os.path.relpath(real_path, json_dir)
                             val = rel + params
                    except ValueError: pass
                vars_data.append({"key": key, "val": val, "group": group})

        # Tables
        tables_data = []
        for item in self.table_configs:
            name = item["name"].get().strip()
            path = item["path"].get().strip() 
            sheet = item["sheet"].get().strip()
            cols_cfg = item.get("columns_config", {})
            src_type = item.get("source_type", "file")
            
            if name:
                path_val = path
                if src_type == "file" and path and os.path.exists(path):
                    try:
                        path_val = os.path.relpath(path, json_dir)
                    except ValueError: pass
                
                tables_data.append({
                    "name": name, 
                    "path": path_val, 
                    "sheet": sheet, 
                    "columns_config": cols_cfg,
                    "source_type": src_type
                })

        # Mappings
        mappings_data = []
        for in_ent, out_ent, _ in self.file_mappings:
            in_path = in_ent.get().strip()
            out_path = out_ent.get().strip()
            if in_path or out_path:
                try:
                    rel_in = os.path.relpath(in_path, json_dir)
                except ValueError: rel_in = in_path
                try:
                    rel_out = os.path.relpath(out_path, json_dir)
                except ValueError: rel_out = out_path
                mappings_data.append({"input": rel_in, "output": rel_out})

        project_data = {
            "version": "1.0",
            "variables": vars_data,
            "tables": tables_data,
            "mappings": mappings_data
        }

        try:
            with open(fpath, 'w', encoding='utf-8') as f:
                json.dump(project_data, f, indent=4, ensure_ascii=False)
            
            # Sync state
            self.last_saved_vars = vars_data
            self.last_saved_tables = tables_data
            self.last_saved_mappings = mappings_data
            
            self.log("项目配置已保存: " + fpath)
        except Exception as e:
            tkinter.messagebox.showerror("错误", f"保存失败: {str(e)}")

    def load_project(self, fpath=None):
        if not fpath:
             fpath = filedialog.askopenfilename(filetypes=[("Project JSON", "*.json")], initialdir=os.getcwd())
        if not fpath: return

        print(f"正在加载项目配置: {fpath}")
        json_dir = os.path.dirname(os.path.abspath(fpath))
        
        try:
            with open(fpath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Clear all
            self.clear_vars()
            self.clear_tables()
            self.clear_mappings()

            # Load Variables
            vars_data = data.get("variables", [])
            for item in vars_data:
                # Key maps to new structure or old? 
                # Old structure: {"key":..., "val":..., "group":...}
                # Check format
                k = item.get("key", "")
                v = item.get("val", "")
                g = item.get("group", "")
                
                # Resolve relative paths in v (if it looks like an image path)
                clean_v = v.split("|")[0]
                if clean_v and not os.path.isabs(clean_v):
                    abs_p = os.path.abspath(os.path.join(json_dir, clean_v))
                    if os.path.exists(abs_p):
                        try:
                            # We store relative to CWD in UI usually
                            final_v = os.path.relpath(abs_p, os.getcwd())
                            if "|" in v: final_v += v[len(clean_v):]
                            v = final_v
                        except ValueError: v = abs_p
                
                self.add_variable_row(k, v, g)
            
            # Load Tables
            tables_data = data.get("tables", [])
            for item in tables_data:
                 name = item.get("name", "")
                 path = item.get("path", "")
                 sheet = item.get("sheet", "")
                 cols_cfg = item.get("columns_config", {})
                 src_type = item.get("source_type", "file")
                 
                 if src_type == "file" and path and not os.path.isabs(path):
                     abs_path = os.path.abspath(os.path.join(json_dir, path))
                     if os.path.exists(abs_path):
                         try:
                             path = os.path.relpath(abs_path, os.getcwd())
                         except ValueError:
                             path = abs_path
                 
                 self.add_table_row(name, path, sheet, cols_cfg, src_type)
            
            # Load Mappings
            map_data = data.get("mappings", [])
            for item in map_data:
                in_p = item.get("input", "")
                out_p = item.get("output", "")
                
                if in_p and not os.path.isabs(in_p):
                    in_p = os.path.abspath(os.path.join(json_dir, in_p))
                if out_p and not os.path.isabs(out_p):
                    out_p = os.path.abspath(os.path.join(json_dir, out_p))
                
                try: in_p = os.path.relpath(in_p, os.getcwd())
                except: pass
                try: out_p = os.path.relpath(out_p, os.getcwd())
                except: pass
                
                self.add_mapping_row(in_p, out_p)
                
            self.log(f"成功加载项目配置: {fpath}")
            
            # Sync state (Using the getters to get what's in UI, which is what we just loaded)
            self.last_saved_vars = self._get_current_vars_data()
            self.last_saved_tables = self._get_current_tables_data()
            self.last_saved_mappings = self._get_current_mappings_data()

        except Exception as e:
            err_msg = f"加载项目配置失败: {str(e)}"
            print(err_msg)
            self.log(err_msg)
            tkinter.messagebox.showerror("错误", err_msg)

    def load_mappings_json(self, fpath=None):
        if not fpath:
            fpath = filedialog.askopenfilename(filetypes=[("JSON", "*.json")], initialdir=os.getcwd())
            
        if not fpath:
            return

        json_dir = os.path.dirname(os.path.abspath(fpath))
        try:
            with open(fpath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.clear_mappings()
            
            for item in data:
                in_p = item.get("input", "")
                out_p = item.get("output", "")
                
                # Resolve relative path to absolute
                if in_p and not os.path.isabs(in_p):
                    in_p = os.path.abspath(os.path.join(json_dir, in_p))
                if out_p and not os.path.isabs(out_p):
                    out_p = os.path.abspath(os.path.join(json_dir, out_p))
                
                # Convert back to relative to CWD for display
                try:
                    in_p = os.path.relpath(in_p, os.getcwd())
                except ValueError: pass
                try:
                    out_p = os.path.relpath(out_p, os.getcwd())
                except ValueError: pass

                self.add_mapping_row(in_p, out_p)
            
            self.last_saved_mappings = self._get_current_mappings_data() # Initialize state
        except Exception as e:
            tkinter.messagebox.showerror("错误", str(e))

    def log(self, msg):
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")

    def run_generation(self):
        if self.validate_all_keys():
            tkinter.messagebox.showerror("错误", "变量名存在重复，请修改后再执行生成！")
            return
            
        self.log("开始任务...")
        
        # 1. Collect Variables
        context = {}
        for item in self.variables:
            key = item["k"].get().replace("{{", "").replace("}}", "").strip()
            if key:
                val = item["v"].get("1.0", "end-1c").strip()
                context[key] = val
        
        if not context:
            if not tkinter.messagebox.askyesno("警告", "变量列表为空，确定要继续吗？"):
                return

        self.log(f"当前上下文包含变量: {list(context.keys())}")


        # 1.5 Load Tables (Sequential Processing)
        for item in self.table_configs:
            name = item["name"].get().strip()
            path = item["path"].get().strip()
            sheet = item["sheet"].get().strip()
            cols_cfg = item.get("columns_config", {})
            src_type = item.get("source_type", "file")
            
            if not name: continue

            try:
                raw_data = []
                if src_type == "file":
                    if path and os.path.exists(path):
                        loader = ExcelLoader(path)
                        raw_data = loader.load_data(sheet_name=sheet if sheet else None)
                        self.log(f"已加载表格 '{name}': {len(raw_data)} 行数据 (来自文件)")
                    else:
                        self.log(f"⚠️ 跳过表格 '{name}': 路径无效")
                        continue
                else: 
                    # Derived from table
                    parent_name = path # path holds parent name
                    if parent_name in context:
                        raw_data = context[parent_name]
                        self.log(f"已加载关联表格 '{name}': {len(raw_data)} 行数据 (源于 {parent_name})")
                    else:
                        self.log(f"❌ 加载关联表格 '{name}' 失败: 父表 '{parent_name}' 未找到或未加载")
                        continue

                # Process
                data = self.process_table_data(raw_data, cols_cfg)
                context[name] = data
                print(f"DEBUG_APP: Loaded table '{name}' into context. Rows: {len(data)}")
                
            except Exception as e:
                self.log(f"❌ 处理表格 '{name}' 失败: {str(e)}")
                print(f"DEBUG_APP: Failed to load table '{name}': {e}")

        self.log(f"上下文更新后包含: {list(context.keys())}")


        # 2. Process Mappings
        mgr = TemplateManager(".")
        success = 0
        total = 0

        for in_ent, out_ent, _ in self.file_mappings:
            in_path_raw = in_ent.get().strip()
            out_dir = out_ent.get().strip()
            
            if not in_path_raw or not out_dir:
                continue
                
            # Dynamic filename resolution: Replace {{key}} in in_path
            in_path = in_path_raw
            for k, v in context.items():
                placeholder = f"{{{{ {k} }}}}"
                placeholder_no_space = f"{{{{{k}}}}}"
                # Clean value for filename: remove image parameters like |width=full
                v_clean = str(v).split("|")[0].strip()
                
                if placeholder in in_path:
                    in_path = in_path.replace(placeholder, v_clean)
                if placeholder_no_space in in_path:
                    in_path = in_path.replace(placeholder_no_space, v_clean)

            # 2. Decide which template file to use as source
            # Priority 1: Use raw path if it exists (supporting {{key}} literality on disk)
            # Priority 2: Use resolved path (supporting dynamic template selection)
            actual_in_path = None
            if os.path.exists(in_path_raw):
                actual_in_path = in_path_raw
            elif os.path.exists(in_path):
                actual_in_path = in_path
            
            if not actual_in_path:
                self.log(f"❌ 错误: 找不到模板文件: {in_path_raw} (替换后: {in_path})")
                continue

            total += 1
            try:
                # Output filename is ALWAYS the resolved one
                filename_resolved = os.path.basename(in_path)
                out_path = os.path.join(out_dir, filename_resolved)

                # Debug: Check variables in template
                template_vars = mgr.get_template_variables(actual_in_path)
                self.log(f"[{filename_resolved}] 模板源: {os.path.basename(actual_in_path)}")
                
                # Check for missing variables
                missing = [v for v in template_vars if v not in context]
                if missing:
                    self.log(f"⚠️ 警告: 模板中存在未定义的变量: {missing}")

                self.log(f"正在生成: {out_path}")
                mgr.render_and_save(actual_in_path, context, out_path)
                success += 1
            except Exception as e:
                self.log(f"ERROR: {str(e)}")
        
        self.log(f"完成! 成功: {success}/{total}")
        tkinter.messagebox.showinfo("完成", f"成功生成 {success} 个文件。")

    def _get_current_vars_data(self):
        data = []
        for item in self.variables:
            key = item["k"].get().strip()
            val = item["v"].get("1.0", "end-1c").strip()
            group = item["g"].get().strip()
            if key:
                data.append({"key": key, "val": val, "group": group})
        return data

    def _get_current_mappings_data(self):
        data = []
        for in_ent, out_ent, _ in self.file_mappings:
            in_path = in_ent.get().strip()
            out_dir = out_ent.get().strip()
            if in_path or out_dir:
                data.append({"input": in_path, "output": out_dir})
        return data

    def on_app_close(self):
        """Check for unsaved changes before exiting."""
        vars_changed = self._get_current_vars_data() != self.last_saved_vars
        mappings_changed = self._get_current_mappings_data() != self.last_saved_mappings
        
        if vars_changed or mappings_changed:
            msg = "检测到配置已修改，是否在退出前保存？\n\n"
            if vars_changed: msg += "- 变量配置表 已修改\n"
            if mappings_changed: msg += "- 生成配置表 已修改"
            
            res = tkinter.messagebox.askyesnocancel("退出提示", msg)
            if res is True: # Yes
                # We can't easily call save dialogs here because they might be cancelled.
                # Let's prompt specifically for each if changed.
                if vars_changed: self.save_vars_json()
                if mappings_changed: self.save_mappings_json()
                self.destroy()
            elif res is False: # No
                self.destroy()
            else: # Cancel
                pass
        else:
            self.destroy()

if __name__ == "__main__":
    app = App()
    app.mainloop()
