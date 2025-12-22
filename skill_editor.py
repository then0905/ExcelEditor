# skill_editor_ctk.py
import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import os
from typing import Dict
from PIL import Image, ImageTk, ImageDraw, ImageFont

# ---------- 常數與欄位定義（沿用你原始的） ----------
SKILL_COLUMNS = [
    "SkillID", "Job", "Name", "NeedLv", "Characteristic", "CastMage",
    "CD", "ChantTime", "AnimaTrigger", "Type", "EffectTarget", "Distance",
    "Width", "Height", "CircleDistance", "Damage", "AdditionMode", "Intro"
]

OP_COLUMNS = [
    "SkillID", "SkillComponentID", "DependCondition", "EffectValue",
    "InfluenceStatus", "AddType", "ConditionOR$", "ConditionAND$",
    "EffectDurationTime", "EffectRecive", "TargetCount", "Bonus$"
]

SKILL_COMPONENTS = [
    "Damage", "CrowdControl", "Debuff", "ContinuanceBuff", "PassiveBuff",
    "ElementDamage", "DotDamage", "Channeled", "Utility", "Teleportation",
    "Lunge", "Charge", "UpgradeSkill", "EnhanceSkill", "InheritDamage",
    "AdditiveBuff", "MultipleDamage", "Health"
]

SKILL_TYPE_MAP = {
    "Characteristic": 'bool',
    "NeedLv": 'int',
    "CastMage": 'int',
    "AnimaTrigger": 'int',
    "EffectRecive": 'int',
    "TargetCount": 'int',
    "CD": 'float',
    "ChantTime": 'float',
    "Distance": 'float',
    "Width": 'float',
    "Height": 'float',
    "CircleDistance": 'float',
    "Damage": 'float',
    "SkillID": 'str', "Job": 'str', "Name": 'str', "Type": 'str',
    "EffectTarget": 'str', "AdditionMode": 'str', "Intro": 'str',
}
for col in SKILL_COLUMNS:
    if col not in SKILL_TYPE_MAP:
        SKILL_TYPE_MAP[col] = 'str'

# ---------- App ----------
class SkillEditorApp:
    def __init__(self, root: ctk.CTk):
        self.root = root
        self.root.title("Game Skill Data Editor (customtkinter)")
        self.root.geometry("1400x820")
        # appearance
        ctk.set_appearance_mode("System")  # "Dark", "Light", or "System"
        ctk.set_default_color_theme("dark-blue")

        # data
        self.job_data: Dict[str, Dict[str, pd.DataFrame]] = {}
        self.current_job: str = None
        self.current_skill_id: str = None

        # ui refs
        self.variables: Dict[str, tk.Variable] = {}
        self.entries: Dict[str, ctk.CTkEntry] = {}
        self.current_icon = None
        self.ICON_SIZE_PIXELS = 128

        self._init_ui()

    # ---------- validation ----------
    def _validate_int(self, P):
        if P == "": return True
        try:
            int(P)
            return True
        except ValueError:
            return False

    def _validate_float(self, P):
        if P == "": return True
        if P.count('.') > 1: return False
        try:
            float(P)
            return True
        except ValueError:
            return False

    # ---------- UI init ----------
    def _init_ui(self):
        # top toolbar
        toolbar = ctk.CTkFrame(self.root, height=56, corner_radius=0)
        toolbar.pack(side="top", fill="x")

        btn_frame = ctk.CTkFrame(toolbar, fg_color="transparent")
        btn_frame.pack(side="left", padx=10, pady=8)

        ctk.CTkButton(btn_frame, text="📂 讀取 Master Excel", command=self.load_master_excel, width=180).grid(row=0, column=0, padx=6)
        ctk.CTkButton(btn_frame, text="➕ 新增職業", command=self.add_new_job_dialog, width=120).grid(row=0, column=1, padx=6)
        ctk.CTkButton(btn_frame, text="💾 保存當前職業(單檔)", command=self.save_current_job, width=180).grid(row=0, column=2, padx=6)

        ctk.CTkButton(toolbar, text="🚀 重建並導出 Master Excel (合併排序)", command=self.merge_and_export,
                      width=260, fg_color="#ff5c5c", hover_color="#ff8080").pack(side="right", padx=12, pady=6)

        # main paned layout using frames
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=(8,10))

        # left: jobs (使用 CTkScrollableFrame 來替代 listbox, 綁定點擊)
        left_frame = ctk.CTkFrame(main_frame, width=240, corner_radius=8)
        left_frame.pack(side="left", fill="y", padx=(0,8), pady=4)
        ctk.CTkLabel(left_frame, text="職業列表", anchor="w", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", padx=10, pady=(10,4))

        self.job_scroll = ctk.CTkScrollableFrame(left_frame, height=600, corner_radius=8)
        self.job_scroll.pack(fill="both", expand=True, padx=10, pady=(0,10))

        # job items container
        self.job_buttons = {}  # job_name -> button

        # middle: skill tree (仍用 ttk.Treeview 因為 Treeview 功能完整)
        mid_frame = ctk.CTkFrame(main_frame, width=420, corner_radius=8)
        mid_frame.pack(side="left", fill="both", expand=False, padx=(0,8), pady=4)
        ctk.CTkLabel(mid_frame, text="技能清單", anchor="w", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", padx=10, pady=(10,4))

        search_frame = ctk.CTkFrame(mid_frame, fg_color="transparent")
        search_frame.pack(fill="x", padx=10)

        self.search_var = tk.StringVar()
        self.search_var.trace("w", lambda *a: self.filter_skill_list())
        ctk.CTkEntry(search_frame, placeholder_text="搜尋 SkillID 或 Name...", textvariable=self.search_var).pack(side="left", fill="x", expand=True)
        ctk.CTkButton(search_frame, text="➕ 新增技能", command=self.add_new_skill, width=120).pack(side="left", padx=(8,0))

        # Treeview for skills
        tree_container = ctk.CTkFrame(mid_frame)
        tree_container.pack(fill="both", expand=True, padx=10, pady=8)
        self.skill_tree = ttk.Treeview(tree_container, columns=("SkillID", "Name", "Type"), show="headings",
                                       selectmode="browse", height=18)
        for col in ("SkillID", "Name", "Type"):
            self.skill_tree.heading(col, text=col)
            self.skill_tree.column(col, width=120)
        self.skill_tree.bind("<<TreeviewSelect>>", self.on_skill_select)

        # 技能清單卷軸區 (修正)

        # 1. 使用 ctk.CTkScrollbar 替換 ttk.Scrollbar
        tree_v = ctk.CTkScrollbar(tree_container, orientation="vertical", command=self.skill_tree.yview)
        tree_h = ctk.CTkScrollbar(tree_container, orientation="horizontal", command=self.skill_tree.xview)

        # 2. 綁定
        self.skill_tree.configure(yscrollcommand=tree_v.set,xscrollcommand=tree_h.set)

        # 3. 調整 pack 順序：先 pack 捲軸，後 pack Treeview
        # 垂直捲軸先佔據右邊
        tree_v.pack(side="right", fill="y")
        tree_h.pack(side="bottom", fill="x")

        # Treeview 填充剩下的空間
        self.skill_tree.pack(fill="both", expand=True, side="left")

        # right: detail and operations
        right_frame = ctk.CTkFrame(main_frame, corner_radius=8)
        right_frame.pack(side="left", fill="both", expand=True, padx=(0,4), pady=4)

        # top detail: icon + attributes
        top_detail = ctk.CTkFrame(right_frame, height=220)
        top_detail.pack(fill="x", padx=10, pady=(10,6))

        # icon
        self.icon_frame = ctk.CTkFrame(top_detail, width=150, corner_radius=8)
        self.icon_frame.pack(side="left", padx=(5,5), pady=10)
        self.icon_frame.pack_propagate(False)

        # placeholder icon generation
        placeholder_img = Image.new('RGB', (self.ICON_SIZE_PIXELS, self.ICON_SIZE_PIXELS), color='#d0d0d0')
        draw = ImageDraw.Draw(placeholder_img)
        text = "無圖示"
        try:
            font = ImageFont.truetype("msyh.ttc", 24)
        except:
            try:
                font = ImageFont.truetype("arial.ttf", 20)
            except:
                font = ImageFont.load_default()
        try:
            bbox = draw.textbbox((0, 0), text, font=font)
            w = bbox[2] - bbox[0]
            h = bbox[3] - bbox[1]
            draw.text(((self.ICON_SIZE_PIXELS-w)//2, (self.ICON_SIZE_PIXELS-h)//2), text, fill="#666666", font=font)
        except:
            draw.text((35,55), text, fill="#666666", font=font)
        self.placeholder_icon = ImageTk.PhotoImage(placeholder_img)
        self.icon_label = ctk.CTkLabel(self.icon_frame, text="", image=self.placeholder_icon)
        self.icon_label.pack(expand=True, pady=8)

        # 外層容器，負責垂直排版 ScrollableFrame + Button
        attr_container = ctk.CTkFrame(top_detail, fg_color="#d0d0d0")
        attr_container.pack(side="left", fill="both", expand=True, pady=6)
        # attributes area with scroll
        attr_frame_outer = ctk.CTkScrollableFrame(attr_container, corner_radius=8,fg_color="#d0d0d0")
        attr_frame_outer.pack( fill="both", expand=True, pady=6)


        self.detail_frame = ctk.CTkFrame(attr_frame_outer, fg_color="transparent")
        self.detail_frame.pack(fill="both", expand=True, padx=6, pady=6)

        # register validation with underlying tk root
        vcmd_int = (self.root.register(self._validate_int), '%P')
        vcmd_float = (self.root.register(self._validate_float), '%P')

        # dynamic fields (two columns)
        row_idx = 0
        col_idx = 0
        for field in SKILL_COLUMNS:
            label = ctk.CTkLabel(self.detail_frame, text=field, anchor="e")
            label.grid(row=row_idx, column=col_idx * 2, sticky="e", padx=8, pady=8)

            ui_type = SKILL_TYPE_MAP.get(field, 'str')
            widget = None
            var = None

            if ui_type == 'bool':
                var = tk.BooleanVar()
                widget = ctk.CTkCheckBox(self.detail_frame, variable=var, text="")
                self.variables[field] = var
            elif ui_type == 'int':
                var = tk.StringVar()
                widget = ctk.CTkEntry(self.detail_frame, textvariable=var)
                # attach validation via underlying tk Entry
                widget._entry.config(validate='key', validatecommand=vcmd_int)
                self.variables[field] = var
                self.entries[field] = widget
            elif ui_type == 'float':
                var = tk.StringVar()
                widget = ctk.CTkEntry(self.detail_frame, textvariable=var)
                widget._entry.config(validate='key', validatecommand=vcmd_float)
                self.variables[field] = var
                self.entries[field] = widget
            else:
                var = tk.StringVar()
                widget = ctk.CTkEntry(self.detail_frame, textvariable=var)
                self.variables[field] = var
                self.entries[field] = widget

            if widget:
                widget.grid(row=row_idx, column=col_idx * 2 + 1, sticky="we", padx=8, pady=8)

            col_idx += 1
            if col_idx > 1:
                col_idx = 0
                row_idx += 1

        # update attributes button
        update_btn = ctk.CTkButton(
            attr_container,
            text="💾 更新屬性",
            command=self.update_skill_attributes
        )
        update_btn.pack(fill="x", pady=9, padx=24)
        #ctk.CTkButton(attr_frame_outer, text="💾 更新屬性", command=self.update_skill_attributes).grid(row=row_idx+1, column=0, columnspan=4, sticky="we", pady=8, padx=6)

        # operations area (下方)
        op_frame = ctk.CTkFrame(right_frame, corner_radius=8)
        op_frame.pack(fill="both", expand=True, padx=10, pady=(6,10))

        top_op_btns = ctk.CTkFrame(op_frame, fg_color="transparent")
        top_op_btns.pack(fill="x", pady=(8,6), padx=8)
        ctk.CTkButton(top_op_btns, text="➕ 新增 Operation", command=self.add_operation, width=160).pack(side="left", padx=6)
        ctk.CTkButton(top_op_btns, text="❌ 刪除選中 Op", command=self.delete_operation, width=160).pack(side="left", padx=6)
        ctk.CTkButton(top_op_btns, text="💾 保存 Operation 變更", command=self.save_operations, width=200).pack(side="right", padx=6)

        # OP區 卷軸與Treeview
        op_table_frame = ctk.CTkFrame(op_frame)
        op_table_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.op_tree = ttk.Treeview(op_table_frame, columns=OP_COLUMNS, show='headings', height=10)
        op_h = ctk.CTkScrollbar(op_table_frame, orientation="horizontal", command=self.op_tree.xview)
        op_v = ctk.CTkScrollbar(op_table_frame, orientation="vertical", command=self.op_tree.yview)

        for col in OP_COLUMNS:
            self.op_tree.heading(col, text=col)
            width = 150 if col in ["ConditionOR$", "ConditionAND$", "Bonus$"] else 100
            self.op_tree.column(col, width=width, stretch=False)
        self.op_tree.bind("<Button-1>", self.on_op_single_click)
        self.op_tree.configure(xscrollcommand=op_h.set, yscrollcommand=op_v.set)
        op_v.pack(side="right", fill="y")
        op_h.pack(side="bottom", fill="x")
        self.op_tree.pack(side="left", fill="both", expand=True)

        # final: status bar
        self.status = ctk.CTkLabel(self.root, text="Ready", anchor="w")
        self.status.pack(side="bottom", fill="x")

    # ---------- helper: job list rendering ----------
    def refresh_job_list(self):
        # clear scroll frame
        for child in self.job_scroll.winfo_children():
            child.destroy()
        self.job_buttons.clear()

        for job in sorted(self.job_data.keys()):
            b = ctk.CTkButton(self.job_scroll, text=job, width=200, corner_radius=6,
                              command=lambda j=job: self._on_job_button_clicked(j))
            b.pack(fill="x", padx=8, pady=6)
            self.job_buttons[job] = b

    def _on_job_button_clicked(self, job_name):
        self.current_job = job_name
        self.current_skill_id = None
        self.clear_skill_detail()
        # visually highlight selected
        for j, btn in self.job_buttons.items():
            if j == job_name:
                btn.configure(fg_color="#00345C")
            else:
                btn.configure(fg_color="#1f6aa5")
        self.filter_skill_list()

    # ---------- file load ----------
    def load_master_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not file_path:
            return
        try:
            xls = pd.ExcelFile(file_path)
            # accept either sheet names "SkillData" or "SkillData.json" to be tolerant
            if "SkillData" in xls.sheet_names:
                skill_sheet = "SkillData"
            elif "SkillData.json" in xls.sheet_names:
                skill_sheet = "SkillData.json"
            else:
                messagebox.showerror("錯誤", "Excel 缺少 SkillData 工作表")
                return

            if "SkillOperation" in xls.sheet_names:
                op_sheet = "SkillOperation"
            elif "SkillData.json#Operation" in xls.sheet_names:
                op_sheet = "SkillData.json#Operation"
            else:
                messagebox.showerror("錯誤", "Excel 缺少 SkillOperation 工作表")
                return

            df_skill = pd.read_excel(xls, skill_sheet, dtype=str)
            df_op = pd.read_excel(xls, op_sheet, dtype=str)

            for col in SKILL_COLUMNS:
                if col not in df_skill.columns: df_skill[col] = None
            df_skill = df_skill[SKILL_COLUMNS]

            for col in OP_COLUMNS:
                if col not in df_op.columns: df_op[col] = None
            df_op = df_op[OP_COLUMNS]

            # split by job
            self.job_data = {}
            df_skill['Job'] = df_skill['Job'].astype(str).str.strip()
            unique_jobs = df_skill['Job'].unique()
            for job in unique_jobs:
                if pd.isna(job) or job == "": continue
                job_skills = df_skill[df_skill['Job'] == job].copy()
                skill_ids = job_skills['SkillID'].astype(str).tolist()
                job_ops = df_op[df_op['SkillID'].astype(str).isin(skill_ids)].copy()
                self.job_data[job] = {'info': job_skills, 'ops': job_ops}

            self.refresh_job_list()
            messagebox.showinfo("成功", f"已讀取 {len(self.job_data)} 個職業資料")
            self.status.configure(text=f"已載入 {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取失敗: {e}")

    # ---------- skill list ----------
    def filter_skill_list(self, *args):
        if not self.current_job: return
        for item in self.skill_tree.get_children():
            self.skill_tree.delete(item)
        df = self.job_data[self.current_job]['info']
        search_txt = self.search_var.get().lower()
        for _, row in df.iterrows():
            s_id = str(row.get('SkillID', ''))
            s_name = str(row.get('Name', ''))
            s_type = str(row.get('Type', ''))
            if search_txt in s_id.lower() or search_txt in s_name.lower():
                self.skill_tree.insert('', tk.END, values=(s_id, s_name, s_type))

    def on_skill_select(self, event):
        selection = self.skill_tree.selection()
        if not selection: return
        item = self.skill_tree.item(selection[0])
        skill_id = item['values'][0]
        self.current_skill_id = str(skill_id)
        self.load_skill_detail(self.current_skill_id)

    # ---------- operations editing ----------
    def on_op_single_click(self, event):
        if not self.current_skill_id: return
        item_id = self.op_tree.identify_row(event.y)
        column_identifier = self.op_tree.identify_column(event.x)
        if not item_id: return
        col_idx = int(column_identifier.replace('#', '')) - 1
        col_name = OP_COLUMNS[col_idx]
        if col_name == "SkillComponentID":
            self._edit_op_cell_dropdown(item_id, column_identifier, col_idx, SKILL_COMPONENTS)
        elif col_name != "SkillID":
            self._edit_op_cell_text(item_id, column_identifier, col_idx)

    def _edit_op_cell_text(self, item_id, column_identifier, col_idx):
        current_vals = list(self.op_tree.item(item_id, 'values'))
        bbox = self.op_tree.bbox(item_id, column_identifier)
        if not bbox: return
        x, y, w, h = bbox
        entry = tk.Entry(self.op_tree)
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, current_vals[col_idx])
        entry.focus()

        def save_edit(ev):
            new_val = entry.get()
            current_vals[col_idx] = new_val
            self.op_tree.item(item_id, values=current_vals)
            entry.destroy()

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", lambda e: entry.destroy())

    def _edit_op_cell_dropdown(self, item_id, column_identifier, col_idx, options):
        current_vals = list(self.op_tree.item(item_id, 'values'))
        current_value = current_vals[col_idx]
        bbox = self.op_tree.bbox(item_id, column_identifier)
        if not bbox: return
        x, y, w, h = bbox
        combobox = ttk.Combobox(self.op_tree, values=options, state='readonly')
        combobox.place(x=x, y=y, width=w, height=h)
        try:
            combobox.set(current_value)
        except:
            if options:
                combobox.current(0)
            else:
                combobox.set("")
        combobox.focus()

        def save_edit(ev):
            new_val = combobox.get()
            current_vals[col_idx] = new_val
            self.op_tree.item(item_id, values=current_vals)
            combobox.destroy()

        combobox.bind("<<ComboboxSelected>>", save_edit)
        combobox.bind("<FocusOut>", lambda e: combobox.destroy())

    # ---------- clear / icon ----------
    def clear_skill_detail(self):
        for var in self.variables.values():
            if isinstance(var, tk.BooleanVar):
                var.set(False)
            else:
                var.set("")
        for item in self.op_tree.get_children():
            self.op_tree.delete(item)
        self.icon_label.configure(image=self.placeholder_icon)
        self.current_icon = None

    def load_skill_icon(self, job_name, skill_id):
        icon_path_png = f"skill_icon/Icon/{job_name}/{skill_id}.png"
        icon_path_jpg = f"skill_icon/Icon/{job_name}/{skill_id}.jpg"
        file_path = None
        if os.path.exists(icon_path_png):
            file_path = icon_path_png
        elif os.path.exists(icon_path_jpg):
            file_path = icon_path_jpg
        if file_path:
            try:
                img = Image.open(file_path)
                img = img.resize((self.ICON_SIZE_PIXELS, self.ICON_SIZE_PIXELS), Image.Resampling.LANCZOS)
                self.current_icon = ImageTk.PhotoImage(img)
                self.icon_label.configure(image=self.current_icon)
            except Exception:
                self.icon_label.configure(image=self.placeholder_icon)
                self.current_icon = None
        else:
            self.icon_label.configure(image=self.placeholder_icon)
            self.current_icon = None

    def load_skill_detail(self, skill_id):
        self.load_skill_icon(self.current_job, skill_id)
        df_info = self.job_data[self.current_job]['info']
        row = df_info[df_info['SkillID'].astype(str) == skill_id]
        if row.empty:
            self.clear_skill_detail()
            return
        for field in SKILL_COLUMNS:
            val = row.iloc[0].get(field, None)
            display_val = val if pd.notna(val) else None
            ui_type = SKILL_TYPE_MAP.get(field, 'str')
            var = self.variables.get(field)
            if var is None: continue
            if ui_type == 'bool':
                is_checked = str(display_val).lower() in ['true', '1', 't']
                var.set(is_checked)
            else:
                var.set(str(display_val) if display_val is not None else "")
        # fill ops
        df_ops = self.job_data[self.current_job]['ops']
        op_rows = df_ops[df_ops['SkillID'].astype(str) == skill_id]
        for item in self.op_tree.get_children():
            self.op_tree.delete(item)
        for _, op_row in op_rows.iterrows():
            vals = [op_row.get(c, "") for c in OP_COLUMNS]
            vals = ["" if pd.isna(v) or v is None else str(v) for v in vals]
            self.op_tree.insert('', tk.END, values=vals)

    # ---------- update / save ----------
    def update_skill_attributes(self):
        if not self.current_job or not self.current_skill_id: return
        df = self.job_data[self.current_job]['info']
        idx = df[df['SkillID'].astype(str) == self.current_skill_id].index
        if len(idx) == 0: return
        for field in SKILL_COLUMNS:
            var = self.variables.get(field)
            if var is None: continue
            ui_type = SKILL_TYPE_MAP.get(field, 'str')
            raw_val = var.get()
            final_val = None
            try:
                if ui_type == 'bool':
                    final_val = raw_val
                elif ui_type == 'int':
                    final_val = int(raw_val.strip()) if raw_val.strip() != "" else None
                elif ui_type == 'float':
                    final_val = float(raw_val.strip()) if raw_val.strip() != "" else None
                else:
                    final_val = raw_val.strip() if raw_val.strip() != "" else None
            except ValueError:
                messagebox.showerror("數據錯誤", f"欄位 '{field}' 的值 '{raw_val}' 格式不正確 ({ui_type})。請修正。")
                return
            df.at[idx[0], field] = final_val
        messagebox.showinfo("Info", "技能屬性已暫存 (尚未寫入檔案)")
        self.filter_skill_list()

    def save_operations(self):
        if not self.current_job or not self.current_skill_id: return
        df_ops = self.job_data[self.current_job]['ops']
        df_ops = df_ops[df_ops['SkillID'].astype(str) != self.current_skill_id]
        new_rows = []
        for item in self.op_tree.get_children():
            vals = self.op_tree.item(item)['values']
            row_dict = {col: str(val) if pd.notna(val) else None for col, val in zip(OP_COLUMNS, vals)}
            row_dict['SkillID'] = self.current_skill_id
            new_rows.append(row_dict)
        if new_rows:
            for row in new_rows:
                if row.get('SkillComponentID') not in SKILL_COMPONENTS:
                    pass
            df_new = pd.DataFrame(new_rows, columns=OP_COLUMNS)
            df_ops = pd.concat([df_ops, df_new], ignore_index=True)
        self.job_data[self.current_job]['ops'] = df_ops
        messagebox.showinfo("Info", "Operation 資料已暫存")

    # ---------- add / delete ----------
    def add_new_job_dialog(self):
        top = ctk.CTkToplevel(self.root)
        top.title("新增職業")
        top.geometry("360x140")
        ctk.CTkLabel(top, text="職業代號 (Job Name):").pack(pady=(12,6))
        e = ctk.CTkEntry(top)
        e.pack(padx=12, fill="x")

        def confirm():
            name = e.get().strip()
            if not name: return
            if name in self.job_data:
                messagebox.showerror("Error", "職業已存在")
                return
            self.job_data[name] = {
                'info': pd.DataFrame(columns=SKILL_COLUMNS),
                'ops': pd.DataFrame(columns=OP_COLUMNS)
            }
            self.refresh_job_list()
            top.destroy()

        ctk.CTkButton(top, text="確定", command=confirm).pack(pady=12)

    def add_new_skill(self):
        if not self.current_job:
            messagebox.showwarning("Warning", "請先選擇一個職業")
            return
        df = self.job_data[self.current_job]['info']
        current_count = len(df)
        new_id = f"{self.current_job}_New_{current_count + 1}"
        new_row = {col: None for col in SKILL_COLUMNS}
        new_row['SkillID'] = new_id
        new_row['Job'] = self.current_job
        new_row['Name'] = "New Skill"
        new_row['NeedLv'] = 1
        new_row['Characteristic'] = False
        df_new_row = pd.DataFrame([new_row], columns=SKILL_COLUMNS)
        df = pd.concat([df, df_new_row], ignore_index=True)
        self.job_data[self.current_job]['info'] = df
        self.filter_skill_list()

    def add_operation(self):
        if not self.current_skill_id:
            messagebox.showwarning("Warning", "請先選擇一個技能")
            return
        default_component = SKILL_COMPONENTS[0] if SKILL_COMPONENTS else "Damage"
        default_vals_map = {
            "SkillID": self.current_skill_id, "SkillComponentID": default_component,
            "DependCondition": "None", "EffectValue": "1.0",
            "InfluenceStatus": "None", "AddType": "Add",
            "ConditionOR$": "[]", "ConditionAND$": "[]",
            "EffectDurationTime": "0", "EffectRecive": "-1",
            "TargetCount": "1", "Bonus$": "[]"
        }
        default_vals = [default_vals_map.get(col, '') for col in OP_COLUMNS]
        self.op_tree.insert('', tk.END, values=default_vals)

    def delete_operation(self):
        selected = self.op_tree.selection()
        if not selected: return
        if messagebox.askyesno("確認刪除", f"確定刪除選中的 {len(selected)} 個 Operation 嗎？刪除後請點擊 '保存 Operation 變更' 生效。"):
            for item in selected:
                self.op_tree.delete(item)

    # ---------- file save / merge ----------
    def save_current_job(self):
        if not self.current_job: return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"Job_{self.current_job}.xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not file_path: return
        data = self.job_data[self.current_job]
        try:
            with pd.ExcelWriter(file_path) as writer:
                data['info'][SKILL_COLUMNS].to_excel(writer, sheet_name="SkillData", index=False)
                data['ops'][OP_COLUMNS].to_excel(writer, sheet_name="SkillOperation", index=False)
            messagebox.showinfo("成功", f"職業 {self.current_job} 已保存")
        except Exception as e:
            messagebox.showerror("失敗", str(e))

    def merge_and_export(self):
        if not self.job_data: return
        all_skills = []
        all_ops = []
        for job, data in self.job_data.items():
            all_skills.append(data['info'][SKILL_COLUMNS])
            all_ops.append(data['ops'][OP_COLUMNS])
        if not all_skills: return
        final_skill_df = pd.concat(all_skills, ignore_index=True)
        final_op_df = pd.concat(all_ops, ignore_index=True)
        final_skill_df['NeedLv_sort'] = pd.to_numeric(final_skill_df['NeedLv'], errors='coerce').fillna(0).astype(int)
        final_skill_df.sort_values(by=['Job', 'NeedLv_sort', 'SkillID'], inplace=True)
        final_skill_df.drop(columns=['NeedLv_sort'], inplace=True)
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile="Master_Skill_Data.xlsx",
            title="導出完整 Master Excel"
        )
        if not file_path: return
        try:
            with pd.ExcelWriter(file_path) as writer:
                final_skill_df.to_excel(writer, sheet_name="SkillData", index=False)
                final_op_df.to_excel(writer, sheet_name="SkillOperation", index=False)
            messagebox.showinfo("完成", "所有職業資料已合併並導出！")
        except Exception as e:
            messagebox.showerror("失敗", str(e))

# ---------- demo helpers ----------
def create_dummy_excel():
    file_name = "Demo_SkillData.xlsx"
    if os.path.exists(file_name): return
    skill_data = {
        "SkillID": ["Sword_1", "Mage_1"],
        "Job": ["Sword", "Mage"],
        "Name": ["斬擊", "火球"],
        "NeedLv": [1, 1],
        "Characteristic": [True, False],
        "CastMage": [0, 10],
        "CD": [1, 5],
        "ChantTime": [0, 1.5],
        "AnimaTrigger": [0, 1],
        "Type": ["Target", "Location"],
        "EffectTarget": ["Enemy", "Enemy"],
        "Distance": [1, 10],
        "Width": [0.5, 3],
        "Height": [0, 0],
        "CircleDistance": [None, 3],
        "Damage": [1.1, 2.5],
        "AdditionMode": [None, "Explode"],
        "Intro": ["基礎劍術", "基礎魔法"]
    }
    op_data = {
        "SkillID": ["Sword_1", "Mage_1", "Mage_1"],
        "SkillComponentID": ["Damage", "CrowdControl", "DotDamage"],
        "DependCondition": ["None", "None", "EnemyAlive"],
        "EffectValue": ["1.1", "2.5", "50"],
        "InfluenceStatus": ["Physical", "Magic", "Fire"],
        "AddType": ["Add", "Add", "Set"],
        "ConditionOR$": ["[]", "[]", "['IsWet']"],
        "ConditionAND$": ["[]", "[]", "['IsNear']"],
        "EffectDurationTime": ["0", "0", "3"],
        "EffectRecive": ["-1", "5", "1"],
        "TargetCount": ["1", "5", "5"],
        "Bonus$": ["[]", "['CritRate:0.1']", "[]"]
    }
    df_s = pd.DataFrame(skill_data, columns=SKILL_COLUMNS)
    df_o = pd.DataFrame(op_data, columns=OP_COLUMNS)
    with pd.ExcelWriter(file_name) as writer:
        df_s.to_excel(writer, sheet_name="SkillData", index=False)
        df_o.to_excel(writer, sheet_name="SkillOperation", index=False)

def create_dummy_icons():
    os.makedirs("skill_icon/Icon/Sword", exist_ok=True)
    os.makedirs("skill_icon/Icon/Mage", exist_ok=True)
    try:
        dummy_img = Image.new('RGB', (64, 64), color='red')
        dummy_img.save("skill_icon/Icon/Sword/Sword_1.png")
        dummy_img_2 = Image.new('RGB', (64, 64), color='blue')
        dummy_img_2.save("skill_icon/Icon/Mage/Mage_1.png")
    except Exception as e:
        print(f"無法建立測試圖示: {e}")

# ---------- run ----------
if __name__ == "__main__":
    #測試方法 有需要再反註解
    #create_dummy_excel()
    #create_dummy_icons()
    root = ctk.CTk()
    try:
        app = SkillEditorApp(root)
        root.mainloop()
    except Exception as e:
        print("啟動失敗:", e)
