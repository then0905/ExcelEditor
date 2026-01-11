import customtkinter as ctk
from tkinter import messagebox, filedialog
from data_manager import DataManager
import os
from PIL import Image
import pandas as pd
import tkinter.font as tkfont

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class SheetEditor(ctk.CTkFrame):
    """ 單一母表的編輯介面 (包含左中右佈局) """
    def __init__(self, parent, sheet_name, manager):
        super().__init__(parent)
        self.sheet_name = sheet_name
        self.manager = manager
        self.df = manager.master_dfs[sheet_name]
        self.cfg = manager.config.get(sheet_name, {})
        
        # 取得關鍵欄位
        self.cls_key = self.cfg.get("classification_key", self.df.columns[0])
        self.pk_key = self.cfg.get("primary_key", self.df.columns[0])
        
        self.current_cls_val = None      # 目前點選的分類
        self.current_master_idx = None   # 目前點選的母表 index
        self.current_master_pk = None    # 目前母表的 Primary Key (用於子表新增)
        self.current_image_ref = None # 防止圖片被回收

        self.setup_layout()
        self.load_classification_list()

    def setup_layout(self):
        self.columnconfigure(2, weight=1) # 右側編輯區權重最大
        self.rowconfigure(0, weight=1)

        # --- 左區：分類 (如職業) ---
        self.frame_left = ctk.CTkFrame(self, width=150)
        self.frame_left.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        
        ctk.CTkLabel(self.frame_left, text=f"分類: {self.cls_key}", font=("微軟正黑體", 12, "bold")).pack(pady=5)
        self.scroll_cls = ctk.CTkScrollableFrame(self.frame_left)
        self.scroll_cls.pack(fill="both", expand=True)
        
        # 左側操作按鈕
        btn_box_left = ctk.CTkFrame(self.frame_left, height=40, fg_color="transparent")
        btn_box_left.pack(fill="x", pady=5, padx=2)
        ctk.CTkButton(btn_box_left, text="+", width=50, fg_color="green", command=self.add_classification).pack(side="left", padx=2, expand=True)
        ctk.CTkButton(btn_box_left, text="-", width=50, fg_color="darkred", command=self.delete_classification).pack(side="right", padx=2, expand=True)

        # --- 中區：項目清單 (如技能) ---
        self.frame_mid = ctk.CTkFrame(self, width=200)
        self.frame_mid.grid(row=0, column=1, sticky="nsew", padx=2, pady=2)
        
        ctk.CTkLabel(self.frame_mid, text="清單", font=("微軟正黑體", 12, "bold")).pack(pady=5)
        self.scroll_items = ctk.CTkScrollableFrame(self.frame_mid)
        self.scroll_items.pack(fill="both", expand=True)

        # 中間操作按鈕
        btn_box_mid = ctk.CTkFrame(self.frame_mid, height=40, fg_color="transparent")
        btn_box_mid.pack(fill="x", pady=5, padx=2)
        ctk.CTkButton(btn_box_mid, text="新增項目", width=80, fg_color="green", command=self.add_master_item).pack(side="left", padx=2, fill="x", expand=True)
        ctk.CTkButton(btn_box_mid, text="刪除", width=60, fg_color="darkred", command=self.delete_master_item).pack(side="right", padx=2)

        # --- 右區：編輯區 (上:母表 / 下:子表) ---
        self.frame_right = ctk.CTkFrame(self)
        self.frame_right.grid(row=0, column=2, sticky="nsew", padx=2, pady=2)
        
        # 右上：母表資料
        ctk.CTkLabel(self.frame_right, text="[母表資料]", font=("微軟正黑體", 12, "bold")).pack(pady=2)
        self.top_container = ctk.CTkFrame(self.frame_right, height=100, fg_color="transparent")
        self.top_container.pack(fill="x", expand=False, padx=5, pady=5)

        # 右下：子表資料 (標題區含新增按鈕)
        sub_header_frame = ctk.CTkFrame(self.frame_right, fg_color="transparent")
        sub_header_frame.pack(fill="x", pady=2, padx=5)
        
        ctk.CTkLabel(sub_header_frame, text="[子表資料]", font=("微軟正黑體", 12, "bold")).pack(pady=2)
        # 子表新增按鈕
        ctk.CTkButton(sub_header_frame, text="+ 新增子表資料", width=100, height=24, fg_color="green", 
                      command=self.add_sub_item).pack(side="right")

        # 建立 TabView 用於子表切換
        self.sub_tables_tabs = ctk.CTkTabview(self.frame_right)
        self.sub_tables_tabs.pack(fill="both", expand=True, padx=5, pady=5)

        # (原本的 scroll_container 邏輯移到 load_sub_tables 內部處理，因為 TabView 結構改變)

    # ================= 資料操作邏輯區域 =================

    def add_classification(self):
        """ 新增分類 """
        dialog = ctk.CTkInputDialog(text="請輸入新分類名稱:", title="新增分類")
        new_cls = dialog.get_input()
        if not new_cls: return
        
        dialog_id = ctk.CTkInputDialog(text="請輸入第一筆資料的 ID (Key):", title="初始資料")
        new_id = dialog_id.get_input()
        if not new_id: return

        if new_id in self.df[self.pk_key].astype(str).values:
            messagebox.showerror("錯誤", "此 ID 已存在")
            return

        # 建立新列
        new_row = {col: "" for col in self.df.columns}
        new_row[self.cls_key] = new_cls
        new_row[self.pk_key] = new_id
        
        self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
        self.manager.master_dfs[self.sheet_name] = self.df
        self.load_classification_list()
        self.load_items_by_group(new_cls)

    def delete_classification(self):
        """ 刪除選取的分類 """
        if not self.current_cls_val: return
        if not messagebox.askyesno("刪除確認", f"確定要刪除分類 [{self.current_cls_val}] 及其下所有資料嗎？"): return
        
        self.df = self.df[self.df[self.cls_key] != self.current_cls_val]
        self.df.reset_index(drop=True, inplace=True) # 重置索引
        self.manager.master_dfs[self.sheet_name] = self.df
        
        self.current_cls_val = None
        self.current_master_idx = None
        self.load_classification_list()
        # 清空右側
        for w in self.scroll_items.winfo_children(): w.destroy()
        self.load_sub_tables(None)

    def add_master_item(self):
        """ 新增項目到當前分類（插在該分類最後） """
        if not self.current_cls_val:
            messagebox.showwarning("提示", "請先選擇左側分類")
            return

        dialog = ctk.CTkInputDialog(text="請輸入新項目 ID:", title="新增項目")
        new_id = dialog.get_input()
        if not new_id:
            return

        if new_id in self.df[self.pk_key].astype(str).values:
            messagebox.showerror("錯誤", "ID 已存在")
            return

        # 建立新列
        new_row = {col: "" for col in self.df.columns}
        new_row[self.cls_key] = self.current_cls_val
        new_row[self.pk_key] = new_id

        # 找出該分類最後一筆的位置
        cls_rows = self.df[self.df[self.cls_key] == self.current_cls_val]

        if cls_rows.empty:
            insert_idx = len(self.df)
        else:
            insert_idx = cls_rows.index.max() + 1

        # 插入資料（不是 concat）
        top = self.df.iloc[:insert_idx]
        bottom = self.df.iloc[insert_idx:]

        self.df = pd.concat(
            [top, pd.DataFrame([new_row]), bottom],
            ignore_index=True
        )

        self.manager.master_dfs[self.sheet_name] = self.df

        # 重新載入該分類
        self.load_items_by_group(self.current_cls_val)

        # 選取新增那筆
        new_idx = self.df[self.df[self.pk_key].astype(str) == str(new_id)].index[0]
        self.load_editor(new_idx)

    def delete_master_item(self):
        """ 刪除選取的項目 """
        if self.current_master_idx is None: 
            messagebox.showwarning("提示", "請先選擇要刪除的項目")
            return
        
        if not messagebox.askyesno("刪除確認", "確定要刪除此筆資料嗎？"): return

        self.df.drop(self.current_master_idx, inplace=True)
        self.df.reset_index(drop=True, inplace=True)
        self.manager.master_dfs[self.sheet_name] = self.df
        
        self.current_master_idx = None
        self.load_items_by_group(self.current_cls_val)
        self.load_sub_tables(None)

    def add_sub_item(self):
        """ 新增子表資料（插在該母表最後） """
        if self.current_master_pk is None:
            messagebox.showwarning("提示", "請先選擇母表資料")
            return

        try:
            current_tab = self.sub_tables_tabs.get()
        except:
            return

        if current_tab == "無子表":
            return

        full_sub_name = f"{self.sheet_name}#{current_tab}"
        sub_df = self.manager.sub_dfs.get(full_sub_name)
        if sub_df is None:
            return

        sub_cfg = self.cfg.get("sub_sheets", {}).get(current_tab, {})
        fk_key = sub_cfg.get("foreign_key", self.pk_key)

        # 建立新列
        new_row = {col: "" for col in sub_df.columns}
        new_row[fk_key] = self.current_master_pk

        # 找出該母表最後一筆子資料的位置
        siblings = sub_df[sub_df[fk_key] == self.current_master_pk]

        if siblings.empty:
            insert_idx = len(sub_df)
        else:
            insert_idx = siblings.index.max() + 1

        # 插入資料
        top = sub_df.iloc[:insert_idx]
        bottom = sub_df.iloc[insert_idx:]

        sub_df = pd.concat(
            [top, pd.DataFrame([new_row]), bottom],
            ignore_index=True
        )

        self.manager.sub_dfs[full_sub_name] = sub_df

        # 重新載入子表，停留在原 tab
        self.load_sub_tables(self.current_master_pk)
        self.sub_tables_tabs.set(current_tab)

    def delete_sub_item(self, sheet_full_name, row_idx):
        """ 刪除子表資料 (由每一列的 X 按鈕觸發) """
        if not messagebox.askyesno("確認", "刪除此列子表資料？"): return

        sub_df = self.manager.sub_dfs[sheet_full_name]
        sub_df.drop(row_idx, inplace=True)
        sub_df.reset_index(drop=True, inplace=True)
        self.manager.sub_dfs[sheet_full_name] = sub_df

        try:
            current_tab = self.sub_tables_tabs.get()
        except: current_tab = None

        self.load_sub_tables(self.current_master_pk)
        
        if current_tab: 
            try: self.sub_tables_tabs.set(current_tab)
            except: pass

    # ================= 原有邏輯區域 (含微調) =================

    def load_classification_list(self):
        """ 讀取分類欄位的唯一值 """
        for w in self.scroll_cls.winfo_children(): w.destroy()
        
        groups = self.df[self.cls_key].unique()
        for g in groups:
            # [微調] 增加選取的高亮邏輯
            fg_color = "transparent"
            if str(g) == str(self.current_cls_val):
                fg_color = ("#3B8ED0", "#1F6AA5") # Blue-ish

            btn = ctk.CTkButton(self.scroll_cls, text=str(g), fg_color=fg_color, border_width=1, text_color=("black", "white"),
                                command=lambda val=g: self.load_items_by_group(val))
            btn.pack(fill="x", pady=2)

    def load_items_by_group(self, group_val):
        """ 根據分類顯示中間清單 """
        self.current_cls_val = group_val
        self.load_classification_list() # 刷新左側以更新高亮

        for widget in self.scroll_items.winfo_children(): widget.destroy()
        
        # 篩選資料
        filter_df = self.df[self.df[self.cls_key] == group_val]
        
        for idx, row in filter_df.iterrows():
            display_name = f"{row[self.pk_key]}"
            if "Name" in row: display_name += f" | {row['Name']}"
            
            # [微調] 增加選取的高亮邏輯
            fg_color = "gray"
            if idx == self.current_master_idx:
                 fg_color = ("#3B8ED0", "#1F6AA5")

            btn = ctk.CTkButton(self.scroll_items, text=display_name, anchor="w", fg_color=fg_color,
                                command=lambda i=idx: self.load_editor(i))
            btn.pack(fill="x", pady=2)

    def load_editor(self, row_idx):
            self.current_master_idx = row_idx
        
            # 1. 刷新清單高亮
            if self.current_cls_val:
                for child in self.scroll_items.winfo_children(): child.destroy() # 簡單暴力刷新
                self.load_items_by_group(self.current_cls_val)

            # 2. 清空頂部容器
            for w in self.top_container.winfo_children(): w.destroy()
        
            if row_idx not in self.df.index: return 
            row_data = self.df.loc[row_idx]
            self.current_master_pk = row_data[self.pk_key]

            # 3. 讀取 Config 設定
            # 注意：這裡假設 ConfigWindow 已經把 'use_icon' 和 'image_path' 寫入 manager.config[sheet_name]
            # 或者如果是全域設定，可能是 manager.config['global']，這裡暫設為與 sheet 同級或上一層
            # 為了保險，先從 sheet config 找，沒有再找全域 (如果有設計全域的話)
        
            use_icon = self.cfg.get("use_icon", False)
            img_base_path = self.cfg.get("image_path", "")

            # 4. 建立編輯區容器 (scroll_master_edit)
            # 如果要顯示圖片，把 top_container 分成左右兩邊
        
            edit_target_frame = None

            if use_icon:
                # --- 啟用圖片模式：左圖右文 ---
            
                # 左側圖片區塊
                img_frame = ctk.CTkFrame(self.top_container, width=150, height=100)
                img_frame.pack(side="left", fill="y", padx=(0, 5))
                img_frame.pack_propagate(False) # 固定大小
            
                # 組出圖片路徑規則
                # 分類是最後一層資料夾名稱
                # 詳細清單名稱是檔名
                img_folder = f"{img_base_path}/{self.current_cls_val}"
                img_file = f"{self.current_master_pk}.png"
            
                img_label = ctk.CTkLabel(img_frame, text="No Image")
                img_label.pack(expand=True)

                if img_file and img_folder:
                    full_path = os.path.join(img_folder, img_file)
                
                    if os.path.exists(full_path):
                        try:
                            pil_img = Image.open(full_path)
                            # 保持比例縮放
                            pil_img.thumbnail((128, 128)) 
                            ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=pil_img.size)
                        
                            img_label.configure(image=ctk_img, text="")
                            self.current_image_ref = ctk_img # 必須保留引用
                        except Exception as e:
                            img_label.configure(text="Error")
                            print(f"Load Image Error: {e}")
                    else:
                        # 圖片第二規則: 沒有加上分類的路徑
                        full_path = os.path.join(f"{img_base_path}", img_file)
                        if os.path.exists(full_path):
                            try:
                                pil_img = Image.open(full_path)
                                # 保持比例縮放
                                pil_img.thumbnail((128, 128))
                                ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=pil_img.size)

                                img_label.configure(image=ctk_img, text="")
                                self.current_image_ref = ctk_img  # 必須保留引用
                            except Exception as e:
                                img_label.configure(text="Error")
                                print(f"Load Image Error: {e}")
                        else:
                            img_label.configure(text=f"File not found\n{img_file}")
                else:
                    img_label.configure(text="No Icon Col")

                # 右側編輯區
                edit_target_frame = ctk.CTkScrollableFrame(self.top_container, height=100)
                edit_target_frame.pack(side="right", fill="both", expand=True)
            
            else:
                # --- 原始模式：全寬編輯 ---
                edit_target_frame = ctk.CTkScrollableFrame(self.top_container, height=100)
                edit_target_frame.pack(fill="both", expand=True)

            # 5. 生成欄位輸入框 (填入 edit_target_frame)
            cols_cfg = self.cfg.get("columns", {})
            for col in self.df.columns:
                f = ctk.CTkFrame(edit_target_frame, fg_color="transparent")
                f.pack(fill="x", pady=2)
                ctk.CTkLabel(f, text=col, width=100, anchor="w").pack(side="left")
            
                val = row_data[col]
                col_type = cols_cfg.get(col, {}).get("type", "string")
            
                if col_type == "bool":
                    var = ctk.BooleanVar(value=bool(val))
                    ctk.CTkCheckBox(f, text="", variable=var, 
                                    command=lambda c=col, v=var: self.manager.update_cell(False, self.sheet_name, row_idx, c, v.get())).pack(side="left")
                elif col_type == "enum":
                    opts = cols_cfg.get(col, {}).get("options", [])
                    w = ctk.CTkOptionMenu(f, values=opts, 
                                          command=lambda v, c=col: self.manager.update_cell(False, self.sheet_name, row_idx, c, v))
                    w.set(str(val))
                    w.pack(side="left", fill="x", expand=True)
                else:
                    var = ctk.StringVar(value=str(val))
                    ctk.CTkEntry(f, textvariable=var).pack(side="left", fill="x", expand=True)
                    var.trace_add("write", lambda *args, c=col, v=var: self.manager.update_cell(False, self.sheet_name, row_idx, c, v.get()))

            # 6. 載入子表
            self.load_sub_tables(self.current_master_pk)

    def load_sub_tables(self, master_id):
        # 1. 刪除所有舊的 Tab
        for tab_name in list(self.sub_tables_tabs._tab_dict.keys()):
            self.sub_tables_tabs.delete(tab_name)
    
        related_sheets = [s for s in self.manager.sub_dfs if s.startswith(self.sheet_name + "#")]
    
        # 如果沒有子表，顯示提示
        if not related_sheets:
            self.sub_tables_tabs.add("無子表")
            return
    
        # 2. 為每個子表建立一個 Tab
        for sheet in related_sheets:
            short_name = sheet.split("#")[1]
            self.sub_tables_tabs.add(short_name)
            tab_frame = self.sub_tables_tabs.tab(short_name)
            
            sub_df = self.manager.sub_dfs[sheet]
            sub_cfg = self.cfg.get("sub_sheets", {}).get(short_name, {})
            sub_cols_cfg = sub_cfg.get("columns", {})
            fk = sub_cfg.get("foreign_key", self.pk_key)
            
            if fk not in sub_df.columns:
                ctk.CTkLabel(tab_frame, text=f"錯誤: 子表找不到關聯欄位 {fk}").pack()
                continue
            
            try:
                # 確保 master_id 與欄位型態一致比較
                mask = sub_df[fk].astype(str) == str(master_id)
                filtered_rows = sub_df[mask]
            except Exception as e:
                filtered_rows = sub_df.head(0)

            headers = list(sub_df.columns)

            # 預先計算欄位寬度
            header_font = tkfont.Font(family="微軟正黑體", size=12, weight="bold")
            # 這裡假設一個 scaling，若無可設為 1.0 或使用你的 _get_widget_scaling()
            try: scaling = self._get_widget_scaling()
            except: scaling = 1.0 
            
            column_widths = {}
            # 總寬度初始值要加上刪除按鈕的寬度 (例如 50px)
            total_width = 50 

            for col in headers:
                text_width_pixels = header_font.measure(col)
                text_width_scaled = text_width_pixels / scaling
                target_width = text_width_scaled
                column_widths[col] = max(120, int(target_width))
                total_width += column_widths[col] + 10

            # ========== 3. 標題列容器 ==========
            header_scroll_container = ctk.CTkFrame(tab_frame)
            header_scroll_container.pack(fill="x", pady=(0, 5))
            
            header_canvas = ctk.CTkCanvas(header_scroll_container, bg="#2b2b2b", highlightthickness=0, height=25)
            header_canvas.pack(side="top", fill="x")
            
            header_content = ctk.CTkFrame(header_canvas, fg_color="transparent")
            header_canvas.create_window((0, 0), window=header_content, anchor="nw")
            
            h_frame = ctk.CTkFrame(header_content, fg_color="gray25", width=total_width, height=25)
            h_frame.pack(anchor="w")
            h_frame.pack_propagate(False)
            
            # 標題列最左側增加 "Del" 欄位
            ctk.CTkLabel(h_frame, text="Del", width=40, font=("Arial", 10, "bold"), text_color="red").pack(side="left", padx=5, pady=5)

            for col in headers:
                label = ctk.CTkLabel(h_frame, text=col, width=column_widths[col], font=("微軟正黑體", 12, "bold"))
                label.pack(side="left", padx=5, pady=5)
            
            header_content.update_idletasks()
            header_canvas.configure(scrollregion=header_canvas.bbox("all"))

            # ========== 4. 資料區容器 ==========
            data_scroll_container = ctk.CTkFrame(tab_frame)
            data_scroll_container.pack(fill="both", expand=True)

            data_canvas = ctk.CTkCanvas(data_scroll_container, bg="#2b2b2b", highlightthickness=0)
            data_scroll_v = ctk.CTkScrollbar(data_scroll_container, orientation="vertical", command=data_canvas.yview)
            data_scroll_h = ctk.CTkScrollbar(data_scroll_container, orientation="horizontal")
            
            data_canvas.configure(yscrollcommand=data_scroll_v.set, xscrollcommand=data_scroll_h.set)
            
            data_scroll_v.pack(side="right", fill="y")
            data_scroll_h.pack(side="bottom", fill="x")
            data_canvas.pack(side="left", fill="both", expand=True)
            
            data_content = ctk.CTkFrame(data_canvas, fg_color="transparent")
            data_canvas.create_window((0, 0), window=data_content, anchor="nw")
            
            # 更新捲動範圍
            def make_update_scroll(canvas):
                def update_scroll(event=None):
                    canvas.update_idletasks()
                    bbox = canvas.bbox("all")
                    if bbox: canvas.configure(scrollregion=bbox)
                return update_scroll
            data_content.bind("<Configure>", make_update_scroll(data_canvas))
            
            # 同步捲動
            def make_sync_scroll_command(header_c, data_c):
                def sync_command(*args):
                    data_c.xview(*args)
                    header_c.xview(*args)
                return sync_command
            data_scroll_h.configure(command=make_sync_scroll_command(header_canvas, data_canvas))
            
            # 滑鼠滾輪綁定
            def make_mousewheel_handler(canvas):
                def handler(event): canvas.yview_scroll(int(-1*(event.delta/120)), "units")
                return handler
            def make_shift_mousewheel_handler(header_c, data_c):
                def handler(event): 
                    data_c.xview_scroll(int(-1*(event.delta/120)), "units")
                    header_c.xview_scroll(int(-1*(event.delta/120)), "units")
                return handler

            data_canvas.bind("<MouseWheel>", make_mousewheel_handler(data_canvas))
            data_canvas.bind("<Shift-MouseWheel>", make_shift_mousewheel_handler(header_canvas, data_canvas))
            data_content.bind("<MouseWheel>", make_mousewheel_handler(data_canvas))
            data_content.bind("<Shift-MouseWheel>", make_shift_mousewheel_handler(header_canvas, data_canvas))

            # ========== 7. 繪製資料列 (含刪除按鈕) ==========
            if filtered_rows.empty:
                ctk.CTkLabel(data_content, text="(此項目無資料)", text_color="gray").pack(anchor="w", pady=10)
                continue

            for s_idx, s_row in filtered_rows.iterrows():
                r_frame = ctk.CTkFrame(data_content, width=total_width, height=25)
                r_frame.pack(anchor="w", pady=5)
                r_frame.pack_propagate(False)
                
                # 刪除按鈕 (放在每一列最前面)
                del_btn = ctk.CTkButton(r_frame, text="X", width=40, height=25, fg_color="darkred", hover_color="#800000",
                                        command=lambda full_n=sheet, r=s_idx: self.delete_sub_item(full_n, r))
                del_btn.pack(side="left", padx=5)

                for col in headers:
                    val = s_row[col]
                    c_info = sub_cols_cfg.get(col, {"type": "string"})
                    col_type = c_info.get("type", "string")
                    target_width = column_widths[col]

                    if col_type == "enum":
                        menu = ctk.CTkOptionMenu(r_frame, values=c_info.get("options", ["None"]), width=target_width, height=25,
                                                 command=lambda v, r=s_idx, c=col, s=sheet: self.manager.update_cell(True, s, r, c, v))
                        menu.set(str(val))
                        menu.pack(side="left", padx=5, pady=0)
                    elif col_type == "bool":
                        var = ctk.BooleanVar(value=bool(val) if val != "" else False)
                        chk = ctk.CTkCheckBox(r_frame, text="", variable=var, width=target_width, height=25,
                                              command=lambda r=s_idx, c=col, s=sheet, v=var: self.manager.update_cell(True, s, r, c, v.get()))
                        chk.pack(side="left", padx=5, pady=0)
                    else:
                        var = ctk.StringVar(value=str(val))
                        entry = ctk.CTkEntry(r_frame, textvariable=var, width=target_width, height=25)
                        entry.pack(side="left", padx=5, pady=0)
                        var.trace_add("write", lambda *args, s=sheet, r=s_idx, c=col, v=var: self.manager.update_cell(True, s, r, c, v.get()))
            
            # 更新範圍
            data_content.update_idletasks()
            data_canvas.update_idletasks()
            bbox = data_canvas.bbox("all")
            if bbox: data_canvas.configure(scrollregion=bbox)

class ConfigEditorWindow(ctk.CTkToplevel):
    def __init__(self, parent, manager):
        super().__init__(parent)
        self.title("配置詳細設定")
        self.manager = manager
        self.grab_set()

        # ===== 視窗 =====
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        win_w = int(screen_w * 0.60)
        win_h = int(screen_h * 0.50)
        self.geometry(f"{win_w}x{win_h}")
        self.resizable(False, False)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # ========= Header =========
        header = ctk.CTkFrame(self)
        header.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text="偵測到資料表變動，請確認各表配置",
            font=("微軟正黑體", 16, "bold")
        ).grid(row=0, column=0, sticky="w", pady=(0, 5))

        # ========= 圖片設定（固定在上方，不進 Scroll） =========
        self.var_use_icon = ctk.BooleanVar(value=False)

        icon_block = ctk.CTkFrame(header, fg_color="transparent")
        icon_block.grid(row=1, column=0, sticky="w", pady=(0, 5))

        self.chk_use_icon = ctk.CTkCheckBox(
            icon_block,
            text="啟用圖示顯示 (需有 'Icon' 或 'Image' 欄位)",
            variable=self.var_use_icon,
            command=self.toggle_icon_input
        )
        self.chk_use_icon.grid(row=0, column=0, sticky="w")

        self.frame_img_path = ctk.CTkFrame(icon_block, fg_color="transparent")
        self.frame_img_path.grid(row=1, column=0, sticky="w", padx=20, pady=5)

        ctk.CTkLabel(
            self.frame_img_path,
            text="圖片讀取路徑 (資料夾):"
        ).grid(row=0, column=0, sticky="w")

        self.entry_img_path = ctk.CTkEntry(self.frame_img_path, width=320)
        self.entry_img_path.grid(row=0, column=1, padx=5)
        self.entry_img_path.bind("<KeyRelease>", self.on_image_path_change)

        # ========= Tabs =========
        center = ctk.CTkFrame(self)
        center.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        center.grid_rowconfigure(0, weight=1)
        center.grid_columnconfigure(0, weight=1)

        self.tab_view = ctk.CTkTabview(
            center,
            command=self.on_tab_changed
        )
        self.tab_view.grid(row=0, column=0, sticky="nsew")

        # ===== 初始化 config + Tabs =====
        for sheet_name in self.manager.master_dfs.keys():
            if sheet_name not in self.manager.config:
                self.manager.config[sheet_name] = {
                    "use_icon": False,
                    "image_path": "",
                    "classification_key": self.manager.master_dfs[sheet_name].classification_key,
                    "primary_key": self.manager.master_dfs[sheet_name].primary_key,
                    "columns": {
                        col: {"type": "string"}
                        for col in self.manager.master_dfs[sheet_name].columns
                    },
                    "sub_sheets": {}
                }

            tab = self.tab_view.add(sheet_name)
            self.build_tab_content(tab, sheet_name)

        self.after(10, self.sync_icon_setting_from_tab)

        # ========= Footer =========
        footer = ctk.CTkFrame(self)
        footer.grid(row=3, column=0, sticky="ew", padx=10, pady=(5, 10))

        ctk.CTkButton(
            footer,
            text="儲存配置並刷新介面",
            fg_color="#28a745",
            hover_color="#218838",
            command=self.save_and_close
        ).pack(pady=5)

    # ================== 圖片設定同步 ==================

    def on_tab_changed(self):
        self.sync_icon_setting_from_tab()

    def sync_icon_setting_from_tab(self):
        sheet = self.tab_view.get()
        cfg = self.manager.config.get(sheet, {})

        self.var_use_icon.set(cfg.get("use_icon", False))

        self.entry_img_path.delete(0, "end")
        self.entry_img_path.insert(0, cfg.get("image_path", ""))

        self.toggle_icon_input()

    def toggle_icon_input(self):
        sheet = self.tab_view.get()
        self.manager.config[sheet]["use_icon"] = self.var_use_icon.get()

        if self.var_use_icon.get():
            self.frame_img_path.grid()
        else:
            self.frame_img_path.grid_remove()

    def on_image_path_change(self, event=None):
        sheet = self.tab_view.get()
        self.manager.config[sheet]["image_path"] = self.entry_img_path.get()

    # ================== Tab 內容 ==================

    def build_tab_content(self, tab, sheet_name):
        main_scroll = ctk.CTkScrollableFrame(tab)
        main_scroll.pack(fill="both", expand=True)

        cfg = self.manager.config[sheet_name]
        all_cols = list(self.manager.master_dfs[sheet_name].columns)

        # --- 母表設定 ---
        base_frame = ctk.CTkFrame(main_scroll)
        base_frame.pack(fill="x", padx=5, pady=5)

        ctk.CTkLabel(
            base_frame,
            text="母表分類參數 (Classification):",
            font=("微軟正黑體", 12, "bold")
        ).grid(row=0, column=0, padx=5, pady=5)

        cls_menu = ctk.CTkOptionMenu(
            base_frame,
            values=all_cols,
            command=lambda v: cfg.update({"classification_key": v})
        )
        cls_menu.set(cfg.get("classification_key", all_cols[0]))
        cls_menu.grid(row=0, column=1, padx=5, pady=5)

        # --- 母表欄位 ---
        ctk.CTkLabel(
            main_scroll,
            text="母表欄位類型設定",
            font=("微軟正黑體", 13, "bold"),
            text_color="#3B8ED0"
        ).pack(pady=5)

        for col in all_cols:
            line = ctk.CTkFrame(main_scroll, fg_color="transparent")
            line.pack(fill="x", padx=20, pady=1)

            ctk.CTkLabel(line, text=col, width=150, anchor="w").pack(side="left")

            if col not in cfg["columns"]:
                cfg["columns"][col] = {"type": "string"}

            t_menu = ctk.CTkOptionMenu(
                line,
                values=["string", "float", "int", "bool", "enum"],
                width=100,
                command=lambda v, c=col: self.set_col_type(sheet_name, c, v)
            )
            t_menu.set(cfg["columns"][col]["type"])
            t_menu.pack(side="right")

        # --- 子表 ---
        related_subs = [s for s in self.manager.sub_dfs if s.startswith(sheet_name + "#")]
        if related_subs:
            ctk.CTkLabel(
                main_scroll,
                text="子表欄位類型設定",
                font=("微軟正黑體", 13, "bold"),
                text_color="#E38D2D"
            ).pack(pady=10)

            for sub_full in related_subs:
                short = sub_full.split("#")[1]
                sub_group = ctk.CTkFrame(main_scroll, border_width=1, border_color="gray")
                sub_group.pack(fill="x", padx=10, pady=5)

                ctk.CTkLabel(
                    sub_group, text=f"子表: {short}",
                    font=("微軟正黑體", 12, "bold")
                ).pack(anchor="w", padx=5)

                if short not in cfg["sub_sheets"]:
                    cfg["sub_sheets"][short] = {
                        "foreign_key": cfg["primary_key"],
                        "columns": {}
                    }

                sub_cols = list(self.manager.sub_dfs[sub_full].columns)
                for s_col in sub_cols:
                    s_line = ctk.CTkFrame(sub_group, fg_color="transparent")
                    s_line.pack(fill="x", padx=15, pady=1)

                    ctk.CTkLabel(
                        s_line, text=s_col, width=150, anchor="w"
                    ).pack(side="left")

                    if s_col not in cfg["sub_sheets"][short]["columns"]:
                        cfg["sub_sheets"][short]["columns"][s_col] = {"type": "string"}

                    st_menu = ctk.CTkOptionMenu(
                        s_line,
                        values=["string", "float", "int", "bool", "enum"],
                        width=100,
                        command=lambda v, sn=short, sc=s_col:
                        self.set_sub_col_type(sheet_name, sn, sc, v)
                    )
                    st_menu.set(cfg["sub_sheets"][short]["columns"][s_col]["type"])
                    st_menu.pack(side="right")

    # ================== Config 操作 ==================

    def set_col_type(self, sheet_name, col, val):
        self.manager.config[sheet_name]["columns"][col]["type"] = val

    def set_sub_col_type(self, m, s, col, val):
        self.manager.config[m]["sub_sheets"][s]["columns"][col]["type"] = val

    def save_and_close(self):
        self.manager.save_config()
        self.destroy()
        if hasattr(self.master, "refresh_ui"):
            self.master.refresh_ui()

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Game Data Editor (Config Driven)")

        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()

        win_w = int(screen_w * 0.75)
        win_h = int(screen_h * 0.75)

        self.geometry(f"{win_w}x{win_h}")
        # self.geometry("1280x720")
        
        self.manager = DataManager()
        
        # 頂部操作列
        self.top_bar = ctk.CTkFrame(self, height=40)
        self.top_bar.pack(fill="x", padx=5, pady=5)
        
        ctk.CTkButton(self.top_bar, text="讀取 Excel", command=self.load_file).pack(side="left", padx=5)
        ctk.CTkButton(self.top_bar, text="儲存 Excel", command=self.save_file, fg_color="green").pack(side="left", padx=5)
        ctk.CTkButton(self.top_bar, text="配置設定", command=self.open_configwnd, fg_color="gray").pack(side="right", padx=5)

        # 內容區 (Tabview 存放不同的母表)
        self.main_tabs = ctk.CTkTabview(self)
        self.main_tabs.pack(fill="both", expand=True, padx=5, pady=5)

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not path: return
    
        try:
            self.manager.load_excel(path)
        
            # 關鍵：判斷是否需要提示使用者
            if self.manager.need_config_alert:
                messagebox.showinfo("提示", "偵測到新資料表，請先設定【分類參數】與【欄位格式】")
                self.open_configwnd()
            else:
                self.refresh_ui()
            
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取失敗: {str(e)}")

    def save_file(self):
        try:
            self.manager.save_excel()
            messagebox.showinfo("成功", "存檔完成")
        except Exception as e:
            messagebox.showerror("存檔失敗", str(e))

    def refresh_ui(self):
        # 根據母表數量建立 Tabs
        for tab_name in list(self.main_tabs._tab_dict.keys()):
            self.main_tabs.delete(tab_name)
            
        for sheet_name in self.manager.master_dfs:
            self.main_tabs.add(sheet_name)
            parent = self.main_tabs.tab(sheet_name)
            # 實例化單一 Sheet 編輯器
            editor = SheetEditor(parent, sheet_name, self.manager)
            editor.pack(fill="both", expand=True)

    def open_configwnd(self):
        if not self.manager.master_dfs:
            messagebox.showinfo("提示", "請先匯入Excel後再進行參數的配置")
            return
        _ = ConfigEditorWindow(self, self.manager)

if __name__ == "__main__":
    app = App()
    app.mainloop()