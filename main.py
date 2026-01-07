import customtkinter as ctk
from tkinter import messagebox, filedialog
from data_manager import DataManager

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
        
        self.setup_layout()
        self.load_classification_list()

    def setup_layout(self):
        self.columnconfigure(2, weight=1) # 右側編輯區權重最大
        self.rowconfigure(0, weight=1)

        # --- 左區：分類 (如職業) ---
        self.frame_left = ctk.CTkFrame(self, width=150)
        self.frame_left.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        ctk.CTkLabel(self.frame_left, text=f"分類: {self.cls_key}").pack(pady=5)
        self.scroll_cls = ctk.CTkScrollableFrame(self.frame_left)
        self.scroll_cls.pack(fill="both", expand=True)

        # --- 中區：項目清單 (如技能) ---
        self.frame_mid = ctk.CTkFrame(self, width=200)
        self.frame_mid.grid(row=0, column=1, sticky="nsew", padx=2, pady=2)
        ctk.CTkLabel(self.frame_mid, text="清單").pack(pady=5)
        self.scroll_items = ctk.CTkScrollableFrame(self.frame_mid)
        self.scroll_items.pack(fill="both", expand=True)

        # --- 右區：編輯區 (上:母表 / 下:子表) ---
        self.frame_right = ctk.CTkFrame(self)
        self.frame_right.grid(row=0, column=2, sticky="nsew", padx=2, pady=2)
        
        # 右上：母表資料
        ctk.CTkLabel(self.frame_right, text="[母表資料]").pack(pady=2)
        self.scroll_master_edit = ctk.CTkScrollableFrame(self.frame_right, height=150)
        self.scroll_master_edit.pack(fill="x", expand=False, padx=5, pady=5)

        # 右下：子表資料
        ctk.CTkLabel(self.frame_right, text="[子表資料]").pack(pady=2)

        # 建立容器來放置捲軸和內容
        scroll_container = ctk.CTkFrame(self.frame_right)
        scroll_container.pack(fill="both", expand=True, padx=5, pady=5)

        # 建立 Canvas 作為可捲動區域
        self.sub_tables_canvas = ctk.CTkCanvas(scroll_container, bg="#2b2b2b", highlightthickness=0)

        # 建立垂直和水平捲軸
        self.scroll_v = ctk.CTkScrollbar(
            scroll_container, 
            orientation="vertical", 
            command=self.sub_tables_canvas.yview
        )
        self.scroll_h = ctk.CTkScrollbar(
            scroll_container, 
            orientation="horizontal", 
            command=self.sub_tables_canvas.xview
        )

        # 配置 Canvas 捲軸
        self.sub_tables_canvas.configure(
            yscrollcommand=self.scroll_v.set,
            xscrollcommand=self.scroll_h.set
        )

        # Pack 捲軸和 Canvas（順序很重要）
        self.scroll_v.pack(side="right", fill="y")
        self.scroll_h.pack(side="bottom", fill="x")
        self.sub_tables_canvas.pack(side="left", fill="both", expand=True)

        # 在 Canvas 內建立 Frame 來放置內容
        self.scroll_tab_sub_tables = ctk.CTkFrame(self.sub_tables_canvas, fg_color="transparent")
        self.canvas_frame = self.sub_tables_canvas.create_window(
            (0, 0), 
            window=self.scroll_tab_sub_tables, 
            anchor="nw"
        )

        # 更新捲動區域的函數
        def update_scroll_region(event=None):
            self.sub_tables_canvas.update_idletasks()
            # 取得內容的實際大小
            bbox = self.sub_tables_canvas.bbox("all")
            if bbox:
                self.sub_tables_canvas.configure(scrollregion=bbox)

        # 綁定配置事件
        self.scroll_tab_sub_tables.bind("<Configure>", update_scroll_region)

        # 綁定滑鼠滾輪事件
        # def _on_mousewheel(event):
        #     self.sub_tables_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

        # def _on_shift_mousewheel(event):
        #     self.sub_tables_canvas.xview_scroll(int(-1*(event.delta/120)), "units")

        # self.sub_tables_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        # self.sub_tables_canvas.bind_all("<Shift-MouseWheel>", _on_shift_mousewheel)

    def load_classification_list(self):
        """ 讀取分類欄位的唯一值 """
        groups = self.df[self.cls_key].unique()
        for g in groups:
            btn = ctk.CTkButton(self.scroll_cls, text=str(g), fg_color="transparent", border_width=1,
                                command=lambda val=g: self.load_items_by_group(val))
            btn.pack(fill="x", pady=2)

    def load_items_by_group(self, group_val):
        """ 根據分類顯示中間清單 """
        for widget in self.scroll_items.winfo_children(): widget.destroy()
        
        # 篩選資料
        filter_df = self.df[self.df[self.cls_key] == group_val]
        
        for idx, row in filter_df.iterrows():
            display_name = f"{row[self.pk_key]}"
            if "Name" in row: display_name += f" | {row['Name']}"
            
            btn = ctk.CTkButton(self.scroll_items, text=display_name, anchor="w", fg_color="gray",
                                command=lambda i=idx: self.load_editor(i))
            btn.pack(fill="x", pady=2)

    def load_editor(self, row_idx):
        """ 載入右側編輯區 """
        # 1. 清空舊 UI
        for w in self.scroll_master_edit.winfo_children(): w.destroy()
        
        # 2. 生成母表欄位
        row_data = self.df.loc[row_idx]
        cols_cfg = self.cfg.get("columns", {})
        
        for col in self.df.columns:
            f = ctk.CTkFrame(self.scroll_master_edit, fg_color="transparent")
            f.pack(fill="x", pady=2)
            ctk.CTkLabel(f, text=col, width=100, anchor="w").pack(side="left")
            
            val = row_data[col]
            col_type = cols_cfg.get(col, {}).get("type", "string")
            
            # 建立輸入元件並綁定更新事件
            if col_type == "bool":
                var = ctk.BooleanVar(value=bool(val))
                widget = ctk.CTkCheckBox(f, text="", variable=var, 
                                         command=lambda c=col, v=var: self.manager.update_cell(False, self.sheet_name, row_idx, c, v.get()))
                widget.pack(side="left")
            elif col_type == "enum":
                opts = cols_cfg.get(col, {}).get("options", [])
                widget = ctk.CTkOptionMenu(f, values=opts, 
                                           command=lambda v, c=col: self.manager.update_cell(False, self.sheet_name, row_idx, c, v))
                widget.set(str(val))
                widget.pack(side="left", fill="x", expand=True)
            else:
                var = ctk.StringVar(value=str(val))
                widget = ctk.CTkEntry(f, textvariable=var)
                widget.pack(side="left", fill="x", expand=True)
                # 使用 trace 監聽輸入變更
                var.trace_add("write", lambda *args, c=col, v=var: self.manager.update_cell(False, self.sheet_name, row_idx, c, v.get()))

        # 3. 載入子表 (根據當前母表 ID)
        current_pk = row_data[self.pk_key]
        self.load_sub_tables(current_pk)

    def load_sub_tables(self, master_id):
        # 安全刪除舊內容
        for w in self.scroll_tab_sub_tables.winfo_children(): 
            w.destroy()
    
        related_sheets = [s for s in self.manager.sub_dfs if s.startswith(self.sheet_name + "#")]
    
        for sheet in related_sheets:
            short_name = sheet.split("#")[1]
    
            sub_df = self.manager.sub_dfs[sheet]
            sub_cfg = self.cfg.get("sub_sheets", {}).get(short_name, {})
            sub_cols_cfg = sub_cfg.get("columns", {})
    
            fk = sub_cfg.get("foreign_key", self.pk_key)
    
            if fk not in sub_df.columns:
                ctk.CTkLabel(
                    self.scroll_tab_sub_tables,
                    text=f"錯誤: 子表找不到關聯欄位 {fk}", 
                    text_color="red"
                ).pack(anchor="w")  # 改用 anchor
                continue
    
            try:
                mask = sub_df[fk].astype(str) == str(master_id)
                filtered_rows = sub_df[mask]
            except Exception as e:
                print(f"篩選錯誤: {e}")
                filtered_rows = sub_df.head(0)

            headers = list(sub_df.columns)

            # 預先計算欄位寬度
            import tkinter.font as tkfont
            header_font = tkfont.Font(family="微軟正黑體", size=12, weight="bold")
            scaling = self._get_widget_scaling()
    
            column_widths = {}
            total_width = 0  # 計算總寬度
            for col in headers:
                text_width_pixels = header_font.measure(col)
                text_width_scaled = text_width_pixels / scaling
                target_width = text_width_scaled * 1.1 + 10
                column_widths[col] = max(120, int(target_width))
                total_width += column_widths[col] + 10  # 加上 padx

            # 標題列 - 不使用 fill="x"，改用固定寬度
            h_frame = ctk.CTkFrame(
                self.scroll_tab_sub_tables,
                fg_color="gray25",
                width=total_width,  # 設定固定寬度
                height=40
            )
            h_frame.pack(anchor="w", pady=(0, 5))  # 改用 anchor="w"
            h_frame.pack_propagate(False)  # 防止自動調整大小
    
            for col in headers:
                label = ctk.CTkLabel(
                    h_frame, 
                    text=col, 
                    width=column_widths[col],
                    font=("微軟正黑體", 12, "bold")
                )
                label.pack(side="left", padx=5)

            # 若無資料
            if filtered_rows.empty:
                ctk.CTkLabel(
                    self.scroll_tab_sub_tables,
                    text="(此項目無子表資料)", 
                    text_color="gray"
                ).pack(anchor="w", pady=10)
                continue

            # 資料列 - 同樣不使用 fill="x"
            for s_idx, s_row in filtered_rows.iterrows():
                r_frame = ctk.CTkFrame(
                    self.scroll_tab_sub_tables,
                    width=total_width,  # 設定固定寬度
                    height=40
                )
                r_frame.pack(anchor="w", pady=5)  # 改用 anchor="w"
                r_frame.pack_propagate(False)  # 防止自動調整大小
        
                for col in headers:
                    val = s_row[col]
                    c_info = sub_cols_cfg.get(col, {"type": "string"})
                    col_type = c_info.get("type", "string")
            
                    target_width = column_widths[col]

                    if col_type == "enum":
                        menu = ctk.CTkOptionMenu(
                            r_frame, 
                            values=c_info.get("options", ["None"]), 
                            width=target_width,
                            command=lambda v, r=s_idx, c=col, s=sheet: self.manager.update_cell(True, s, r, c, v)
                        )
                        menu.set(str(val))
                        menu.pack(side="left", padx=5)
                    elif col_type == "bool":
                        var = ctk.BooleanVar(value=bool(val) if val != "" else False)
                        chk = ctk.CTkCheckBox(
                            r_frame, text="", variable=var, width=target_width,
                            command=lambda r=s_idx, c=col, s=sheet, v=var: self.manager.update_cell(True, s, r, c, v.get())
                        )
                        chk.pack(side="left", padx=5)
                    else:
                        var = ctk.StringVar(value=str(val))
                        entry = ctk.CTkEntry(r_frame, textvariable=var, width=target_width)
                        entry.pack(side="left", padx=5)
                        var.trace_add("write", lambda *args, s=sheet, r=s_idx, c=col, v=var: self.manager.update_cell(True, s, r, c, v.get()))
    
        # 強制更新捲動區域
        self.scroll_tab_sub_tables.update_idletasks()
        self.sub_tables_canvas.update_idletasks()
    
        # 手動設定 scrollregion
        bbox = self.sub_tables_canvas.bbox("all")
        if bbox:
            self.sub_tables_canvas.configure(scrollregion=bbox)

class ConfigEditorWindow(ctk.CTkToplevel):
    def __init__(self, parent, manager):
        super().__init__(parent)
        self.title("配置詳細設定")
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()

        # 視窗用螢幕的 85%
        win_w = int(screen_w * 0.60)
        win_h = int(screen_h * 0.50)

        self.geometry(f"{win_w}x{win_h}")
        self.manager = manager
        self.grab_set() 
        
        # 標題
        ctk.CTkLabel(self, text="偵測到資料表變動，請確認各表配置", font=("微軟正黑體", 16, "bold")).pack(pady=10)

        # 建立 Tabview
        self.tab_view = ctk.CTkTabview(self)
        self.tab_view.pack(fill="both", expand=True, padx=10, pady=10)

        # 遍歷目前 DataManager 裡所有的母表資料
        for sheet_name in self.manager.master_dfs.keys():
            # 確保 config 字典裡有這個 key，如果沒有才初始化
            if sheet_name not in self.manager.config:
                self.manager.config[sheet_name] = {
                    "classification_key": self.manager.master_dfs[sheet_name].columns[0],
                    "primary_key": self.manager.master_dfs[sheet_name].columns[0],
                    "columns": {col: {"type": "string"} for col in self.manager.master_dfs[sheet_name].columns},
                    "sub_sheets": {}
                }
            
            tab = self.tab_view.add(sheet_name)
            self.build_tab_content(tab, sheet_name)

        # 底部按鈕
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", pady=10)
        ctk.CTkButton(btn_frame, text="儲存配置並刷新介面", fg_color="#28a745", hover_color="#218838",
                      command=self.save_and_close).pack(side="bottom", pady=5)

    def build_tab_content(self, tab, sheet_name):
        # 加入全域卷軸，解決你提到的「沒有卷軸」問題
        main_scroll = ctk.CTkScrollableFrame(tab)
        main_scroll.pack(fill="both", expand=True)

        cfg = self.manager.config[sheet_name]
        all_cols = list(self.manager.master_dfs[sheet_name].columns)

        # --- 母表基本設定 ---
        base_frame = ctk.CTkFrame(main_scroll)
        base_frame.pack(fill="x", padx=5, pady=5)
        
        ctk.CTkLabel(base_frame, text="母表分類參數 (Classification):", font=("微軟正黑體", 12, "bold")).grid(row=0, column=0, padx=5, pady=5)
        cls_menu = ctk.CTkOptionMenu(base_frame, values=all_cols,
                                     command=lambda v: cfg.update({"classification_key": v}))
        cls_menu.set(cfg.get("classification_key", all_cols[0]))
        cls_menu.grid(row=0, column=1, padx=5, pady=5)

        # --- 母表欄位格式設定 ---
        ctk.CTkLabel(main_scroll, text="母表欄位類型設定", font=("微軟正黑體", 13, "bold"), text_color="#3B8ED0").pack(pady=5)
        
        for col in all_cols:
            line = ctk.CTkFrame(main_scroll, fg_color="transparent")
            line.pack(fill="x", padx=20, pady=1)
            ctk.CTkLabel(line, text=col, width=150, anchor="w").pack(side="left")
            
            # 確保該欄位在 config 裡有配置
            if col not in cfg["columns"]: cfg["columns"][col] = {"type": "string"}
            
            t_var = cfg["columns"][col].get("type", "string")
            t_menu = ctk.CTkOptionMenu(line, values=["string", "float", "int", "bool", "enum"], width=100,
                                       command=lambda v, c=col: self.set_col_type(sheet_name, c, v))
            t_menu.set(t_var)
            t_menu.pack(side="right")

        # --- 子表欄位格式設定 ---
        related_subs = [s for s in self.manager.sub_dfs if s.startswith(sheet_name + "#")]
        if related_subs:
            ctk.CTkLabel(main_scroll, text="子表欄位類型設定", font=("微軟正黑體", 13, "bold"), text_color="#E38D2D").pack(pady=10)
            
            for sub_full_name in related_subs:
                short_name = sub_full_name.split("#")[1]
                sub_group = ctk.CTkFrame(main_scroll, border_width=1, border_color="gray")
                sub_group.pack(fill="x", padx=10, pady=5)
                
                ctk.CTkLabel(sub_group, text=f"子表: {short_name}", font=("微軟正黑體", 12, "bold")).pack(anchor="w", padx=5)
                
                # 初始化子表配置結構
                if short_name not in cfg["sub_sheets"]:
                    cfg["sub_sheets"][short_name] = {"foreign_key": cfg["primary_key"], "columns": {}}
                
                sub_cols = list(self.manager.sub_dfs[sub_full_name].columns)
                for s_col in sub_cols:
                    s_line = ctk.CTkFrame(sub_group, fg_color="transparent")
                    s_line.pack(fill="x", padx=15, pady=1)
                    ctk.CTkLabel(s_line, text=s_col, width=150, anchor="w").pack(side="left")
                    
                    if s_col not in cfg["sub_sheets"][short_name]["columns"]:
                        cfg["sub_sheets"][short_name]["columns"][s_col] = {"type": "string"}
                    
                    st_var = cfg["sub_sheets"][short_name]["columns"][s_col].get("type", "string")
                    st_menu = ctk.CTkOptionMenu(s_line, values=["string", "float", "int", "bool", "enum"], width=100,
                                                command=lambda v, sn=short_name, sc=s_col: self.set_sub_col_type(sheet_name, sn, sc, v))
                    st_menu.set(st_var)
                    st_menu.pack(side="right")

    def set_col_type(self, sheet_name, col, val):
        self.manager.config[sheet_name]["columns"][col]["type"] = val
        if val == "enum":
            res = ctk.CTkInputDialog(text=f"請輸入 {col} 的選項(逗號隔開):", title="Enum設定").get_input()
            if res: self.manager.config[sheet_name]["columns"][col]["options"] = [x.strip() for x in res.split(",")]

    def set_sub_col_type(self, m_name, s_name, col, val):
        self.manager.config[m_name]["sub_sheets"][s_name]["columns"][col]["type"] = val
        if val == "enum":
            res = ctk.CTkInputDialog(text=f"子表 {s_name} 欄位 {col} 選項:", title="Enum設定").get_input()
            if res: self.manager.config[m_name]["sub_sheets"][s_name]["columns"][col]["options"] = [x.strip() for x in res.split(",")]

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