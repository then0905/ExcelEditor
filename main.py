import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog
from data_manager import DataManager
import os
import sys
from PIL import Image
import pandas as pd
import tkinter.font as tkfont

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# Dark theme 色彩常數
_BG = "#2b2b2b"
_BG_HEADER = "#404040"
_ROW_EVEN = "#2b2b2b"
_ROW_ODD = "#3a3a3a"


class LightScrollableFrame(tk.Frame):
    """
    輕量化可捲動框架 — 替代 ctk.CTkScrollableFrame。
    內部全部使用原生 tk 元件，避免 CTk 雙層 canvas 在快速捲動時產生殘影。
    子元件請放到 .interior 屬性中。
    """

    def __init__(self, parent, height=None, **kwargs):
        super().__init__(parent, bg=_BG)

        self._canvas = tk.Canvas(self, bg=_BG, highlightthickness=0, bd=0)
        self._scrollbar = ctk.CTkScrollbar(self, orientation="vertical",
                                           command=self._canvas.yview)
        self._scrollbar.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)
        self._canvas.configure(yscrollcommand=self._scrollbar.set)

        self.interior = tk.Frame(self._canvas, bg=_BG)
        self._win_id = self._canvas.create_window((0, 0), window=self.interior, anchor="nw")

        self.interior.bind("<Configure>", self._on_interior_cfg)
        self._canvas.bind("<Configure>", self._on_canvas_cfg)

        if height:
            self._canvas.configure(height=height)

        # 供 App 層級滾輪路由識別
        self._canvas._is_light_scrollable = True

    def _update_scroll_region(self):
        bbox = self._canvas.bbox("all")
        if bbox:
            canvas_h = self._canvas.winfo_height()
            region_h = max(bbox[3], canvas_h)
            self._canvas.configure(scrollregion=(bbox[0], bbox[1], bbox[2], region_h))

    def _on_interior_cfg(self, _event=None):
        self._update_scroll_region()

    def _on_canvas_cfg(self, event):
        self._canvas.itemconfig(self._win_id, width=event.width)
        self._update_scroll_region()


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

        self.current_cls_val = None
        self.current_master_idx = None
        self.current_master_pk = None
        self.current_image_ref = None

        # 母表UI緩存
        self.cls_buttons = {}  # {分類值: 按鈕widget}
        self.item_buttons = {}  # {row_idx: 按鈕widget}
        self.master_fields = {}  # {欄位名: Entry/CheckBox等widget}
        self.master_field_vars = {}  # {欄位名: StringVar/BooleanVar}
        self.trace_ids = {}  # {欄位名: trace_id} 用於清理舊的 trace

        # 子表UI緩存
        self.sub_table_frames = {}  # {tab_name: 容器frame}
        self.sub_table_headers = {}  # {tab_name: 標題frame}
        self.sub_table_row_pools = {}  # {tab_name: [可重用的row_frame列表]}
        self.sub_table_active_rows = {}  # {tab_name: [正在使用的row_frame列表]}

        self.setup_layout()
        self.load_classification_list()

    def setup_layout(self):
        """佈局設置"""
        self.columnconfigure(0, weight=0)
        self.columnconfigure(1, weight=0)
        self.columnconfigure(2, weight=1)
        self.rowconfigure(0, weight=1)

        # --- 左側：分類 ---
        self.frame_left = ctk.CTkFrame(self, width=150)
        self.frame_left.grid(row=0, column=0, sticky="nsew")

        ctk.CTkLabel(self.frame_left, text=f"分類: {self.cls_key}", font=("微軟正黑體", 12, "bold")).pack(pady=5)
        self.scroll_cls = LightScrollableFrame(self.frame_left)
        self.scroll_cls.pack(fill="both", expand=True)

        # 左側操作按鈕
        btn_box_left = ctk.CTkFrame(self.frame_left, height=40, fg_color="transparent")
        btn_box_left.pack(fill="x", pady=5, padx=2)
        ctk.CTkButton(btn_box_left, text="+", width=50, fg_color="green", command=self.add_classification).pack(side="left", padx=2, expand=True)
        ctk.CTkButton(btn_box_left, text="-", width=50, fg_color="darkred", command=self.delete_classification).pack(side="right", padx=2, expand=True)

        # --- 中間：項目清單 ---
        self.frame_mid = ctk.CTkFrame(self, width=200)
        self.frame_mid.grid(row=0, column=1, sticky="nsew")

        ctk.CTkLabel(self.frame_mid, text="清單", font=("微軟正黑體", 12, "bold")).pack(pady=5)
        self.scroll_items = LightScrollableFrame(self.frame_mid)
        self.scroll_items.pack(fill="both", expand=True)

        # 中間操作按鈕
        btn_box_mid = ctk.CTkFrame(self.frame_mid, height=40, fg_color="transparent")
        btn_box_mid.pack(fill="x", pady=5, padx=2)
        ctk.CTkButton(btn_box_mid, text="新增項目", width=80, fg_color="green", command=self.add_master_item).pack(side="left", padx=2, fill="x", expand=True)
        ctk.CTkButton(btn_box_mid, text="刪除", width=60, fg_color="darkred", command=self.delete_master_item).pack(side="right", padx=2)

        # --- 右區：編輯區 (上:母表 / 下:子表) ---
        self.frame_right = ctk.CTkFrame(self)
        self.frame_right.grid(row=0, column=2, sticky="nsew")

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

    def load_classification_list(self):
        """載入分類列表 """
        groups = self.df[self.cls_key].unique()
        current_groups = set(groups)
        cached_groups = set(self.cls_buttons.keys())

        # 1. 移除已不存在的分類按鈕
        for group in (cached_groups - current_groups):
            if group in self.cls_buttons:
                self.cls_buttons[group].destroy()
                del self.cls_buttons[group]

        # 2. 新增或更新分類按鈕
        for g in groups:
            if g in self.cls_buttons:
                # 已存在：只更新顏色（高亮狀態）
                btn = self.cls_buttons[g]
                if str(g) == str(self.current_cls_val):
                    btn.configure(fg_color=("#3B8ED0", "#1F6AA5"))
                else:
                    btn.configure(fg_color="transparent")
            else:
                # 不存在：創建新按鈕
                fg_color = ("#3B8ED0", "#1F6AA5") if str(g) == str(self.current_cls_val) else "transparent"
                btn = ctk.CTkButton(
                    self.scroll_cls.interior,
                    text=str(g),
                    fg_color=fg_color,
                    border_width=1,
                    text_color=("black", "white"),
                    command=lambda val=g: self.load_items_by_group(val)
                )
                btn.pack(fill="x", pady=2)
                self.cls_buttons[g] = btn

    def load_items_by_group(self, group_val):
        """載入項目清單 """
        self.current_cls_val = group_val

        # 更新左側分類按鈕的高亮狀態（不重建）
        for g, btn in self.cls_buttons.items():
            if str(g) == str(group_val):
                btn.configure(fg_color=("#3B8ED0", "#1F6AA5"))
            else:
                btn.configure(fg_color="transparent")

        # 篩選該分類的資料
        filter_df = self.df[self.df[self.cls_key] == group_val]
        current_indices = set(filter_df.index)
        cached_indices = set(self.item_buttons.keys())

        # 1. 移除已不存在的項目按鈕
        for idx in (cached_indices - current_indices):
            if idx in self.item_buttons:
                self.item_buttons[idx].destroy()
                del self.item_buttons[idx]

        # 2. 新增或更新項目按鈕
        for idx, row in filter_df.iterrows():
            # 取得顯示名稱
            text_dict_name = self.manager.text_dict.get(row['Name'])
            display_name = f"{row[self.pk_key]}"
            if text_dict_name:
                display_name = text_dict_name["value"]

            if idx in self.item_buttons:
                # 已存在：只更新文字和顏色
                btn = self.item_buttons[idx]
                btn.configure(text=display_name)
                if idx == self.current_master_idx:
                    btn.configure(fg_color=("#3B8ED0", "#1F6AA5"))
                else:
                    btn.configure(fg_color="gray")
            else:
                # 不存在：創建新按鈕
                fg_color = ("#3B8ED0", "#1F6AA5") if idx == self.current_master_idx else "gray"
                btn = ctk.CTkButton(
                    self.scroll_items.interior,
                    text=display_name,
                    anchor="w",
                    fg_color=fg_color,
                    command=lambda i=idx: self.load_editor(i)
                )
                btn.pack(fill="x", pady=2)
                self.item_buttons[idx] = btn

    def load_editor(self, row_idx):
        """載入編輯器 """
        self.current_master_idx = row_idx

        # 1. 更新中間清單的高亮（不重建）
        for idx, btn in self.item_buttons.items():
            if idx == row_idx:
                btn.configure(fg_color=("#3B8ED0", "#1F6AA5"))
            else:
                btn.configure(fg_color="gray")

        if row_idx not in self.df.index:
            return

        row_data = self.df.loc[row_idx]
        self.current_master_pk = row_data[self.pk_key]

        # 2. 如果是第一次載入，建立 UI 結構
        if not self.master_fields:
            self._build_editor_ui(row_data)
        else:
            # 已有 UI，只更新數據
            self._update_editor_data(row_data)

        # 3. 更新圖片（如果有）
        self._update_image()

        # 4. 載入子表
        self.load_sub_tables(self.current_master_pk)

    def _build_editor_ui(self, row_data):
        """首次建立編輯器 UI（只執行一次）"""
        # 清空容器
        for w in self.top_container.winfo_children():
            w.destroy()

        use_icon = self.cfg.get("use_icon", False)
        img_base_path = self.cfg.get("image_path", "")

        # 建立圖片框架（如果需要）
        if use_icon:
            self.img_frame = ctk.CTkFrame(self.top_container, width=150, height=100)
            self.img_frame.pack(side="left", fill="y", padx=(0, 5))
            self.img_frame.pack_propagate(False)

            self.img_label = ctk.CTkLabel(self.img_frame, text="No Image")
            self.img_label.pack(expand=True)

            edit_target_frame = LightScrollableFrame(self.top_container, height=100)
            edit_target_frame.pack(side="right", fill="both", expand=True)
        else:
            self.img_frame = None
            self.img_label = None
            edit_target_frame = LightScrollableFrame(self.top_container, height=100)
            edit_target_frame.pack(fill="both", expand=True)

        # 建立欄位 UI
        cols_cfg = self.cfg.get("columns", {})
        self.master_fields = {}
        self.master_field_vars = {}
        self.trace_ids = {}

        for col in self.df.columns:
            f = tk.Frame(edit_target_frame.interior, bg=_BG)
            f.pack(fill="x", pady=2)

            ctk.CTkLabel(f, text=col, width=100, anchor="w").pack(side="left")

            col_conf = cols_cfg.get(col, {})
            col_type = col_conf.get("type", "string")
            is_linked = col_conf.get("link_to_text", False)

            if is_linked:
                # Key (唯讀)
                key_entry = ctk.CTkEntry(f, width=80, text_color="gray")
                key_entry.configure(state="disabled")
                key_entry.pack(side="left", padx=(0, 5))

                # Text (可編輯)
                var = ctk.StringVar()
                text_entry = ctk.CTkEntry(f, textvariable=var)
                text_entry.pack(side="left", fill="x", expand=True)

                self.master_fields[col] = (key_entry, text_entry)  # 保存兩個
                self.master_field_vars[col] = var

                # 綁定更新事件
                trace_id = var.trace_add("write",
                                         lambda *args, c=col, v=var: self._on_linked_field_change(c, v))
                self.trace_ids[col] = trace_id

            elif col_type == "bool":
                var = ctk.BooleanVar()
                chk = ctk.CTkCheckBox(f, text="", variable=var,
                                      command=lambda c=col, v=var: self._on_field_change(c, v.get()))
                chk.pack(side="left")

                self.master_fields[col] = chk
                self.master_field_vars[col] = var

            elif col_type == "enum":
                opts = col_conf.get("options", [])
                menu = ctk.CTkOptionMenu(f, values=opts,
                                         command=lambda v, c=col: self._on_field_change(c, v))
                menu.pack(side="left", fill="x", expand=True)

                self.master_fields[col] = menu

            else:  # string, int, float
                var = ctk.StringVar()
                entry = ctk.CTkEntry(f, textvariable=var)
                entry.pack(side="left", fill="x", expand=True)

                self.master_fields[col] = entry
                self.master_field_vars[col] = var

                trace_id = var.trace_add("write",
                                         lambda *args, c=col, v=var: self._on_field_change(c, v.get()))
                self.trace_ids[col] = trace_id

        # 首次載入數據
        self._update_editor_data(row_data)

    def _update_editor_data(self, row_data):
        """只更新欄位的數據值（不重建 UI）"""
        cols_cfg = self.cfg.get("columns", {})

        for col in self.df.columns:
            if col not in self.master_fields:
                continue

            val = row_data[col]
            col_conf = cols_cfg.get(col, {})
            col_type = col_conf.get("type", "string")
            is_linked = col_conf.get("link_to_text", False)

            # 暫時移除 trace 避免觸發更新
            if col in self.trace_ids:
                var = self.master_field_vars[col]
                var.trace_remove("write", self.trace_ids[col])

            if is_linked:
                key_entry, text_entry = self.master_fields[col]

                # 更新 Key
                key_entry.configure(state="normal")
                key_entry.delete(0, "end")
                key_entry.insert(0, str(val))
                key_entry.configure(state="disabled")

                # 更新 Text
                real_text = self.manager.get_text_value(val)
                var = self.master_field_vars[col]
                var.set(str(real_text))

            elif col_type == "bool":
                var = self.master_field_vars[col]
                var.set(bool(val))

            elif col_type == "enum":
                menu = self.master_fields[col]
                menu.set(str(val))

            else:
                var = self.master_field_vars[col]
                var.set(str(val))

            # 重新綁定 trace
            if col in self.trace_ids and col in self.master_field_vars:
                var = self.master_field_vars[col]
                if is_linked:
                    trace_id = var.trace_add("write",
                                             lambda *args, c=col, v=var: self._on_linked_field_change(c, v))
                else:
                    trace_id = var.trace_add("write",
                                             lambda *args, c=col, v=var: self._on_field_change(c, v.get()))
                self.trace_ids[col] = trace_id

    def _update_image(self):
        """只更新圖片（不重建 UI）"""
        if not self.img_label:
            return

        use_icon = self.cfg.get("use_icon", False)
        if not use_icon:
            return

        img_base_path = self.cfg.get("image_path", "")
        img_folder = f"{img_base_path}/{self.current_cls_val}"
        img_file = f"{self.current_master_pk}.png"

        full_path = os.path.join(img_folder, img_file)
        if not os.path.exists(full_path):
            full_path = os.path.join(img_base_path, img_file)

        if os.path.exists(full_path):
            try:
                pil_img = Image.open(full_path)
                pil_img.thumbnail((128, 128))
                ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=pil_img.size)
                self.img_label.configure(image=ctk_img, text="")
                self.current_image_ref = ctk_img
            except Exception as e:
                self.img_label.configure(text="Error", image=None)
        else:
            self.img_label.configure(text=f"File not found\n{img_file}", image=None)
            self.current_image_ref = None

    def _on_field_change(self, col_name, value):
        """欄位變更回調"""
        if self.current_master_idx is not None:
            self.manager.update_cell(False, self.sheet_name, self.current_master_idx, col_name, value)

    def _on_linked_field_change(self, col_name, var):
        """連結文字欄位變更回調"""
        if self.current_master_idx is not None:
            # 取得 Key
            key_entry, _ = self.master_fields[col_name]
            key_entry.configure(state="normal")
            key = key_entry.get()
            key_entry.configure(state="disabled")

            # 更新文字表
            self.manager.update_linked_text(key, var.get())

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

        new_row = {col: "" for col in self.df.columns}
        new_row[self.cls_key] = new_cls
        new_row[self.pk_key] = new_id

        self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
        self.manager.master_dfs[self.sheet_name] = self.df
        self.manager.dirty = True
        self.load_classification_list()
        self.load_items_by_group(new_cls)

    def delete_classification(self):
        """ 刪除選取的分類 """
        if not self.current_cls_val: return
        if not messagebox.askyesno("刪除確認", f"確定要刪除分類 [{self.current_cls_val}] 及其下所有資料嗎？"): return

        self.df = self.df[self.df[self.cls_key] != self.current_cls_val]
        self.df.reset_index(drop=True, inplace=True)
        self.manager.master_dfs[self.sheet_name] = self.df
        self.manager.dirty = True

        self.current_cls_val = None
        self.current_master_idx = None

        # 清空項目列表
        for btn in self.item_buttons.values():
            btn.destroy()
        self.item_buttons.clear()

        self.load_classification_list()
        self.load_sub_tables(None)

    def add_master_item(self):
        """ 新增項目到當前分類（插在該分類最後） """
        if not self.current_cls_val:
            messagebox.showwarning("提示", "請先選擇左側分類")
            return

        dialog = ctk.CTkInputDialog(text="請輸入新項目 ID:", title="新增項目")
        new_id = dialog.get_input()
        if not new_id: return

        if new_id in self.df[self.pk_key].astype(str).values:
            messagebox.showerror("錯誤", "ID 已存在")
            return

        new_row = {col: "" for col in self.df.columns}
        new_row[self.cls_key] = self.current_cls_val
        new_row[self.pk_key] = new_id

        cls_rows = self.df[self.df[self.cls_key] == self.current_cls_val]
        insert_idx = cls_rows.index.max() + 1 if not cls_rows.empty else len(self.df)

        top = self.df.iloc[:insert_idx]
        bottom = self.df.iloc[insert_idx:]
        self.df = pd.concat([top, pd.DataFrame([new_row]), bottom], ignore_index=True)
        self.manager.master_dfs[self.sheet_name] = self.df
        self.manager.dirty = True

        self.load_items_by_group(self.current_cls_val)

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
        self.manager.dirty = True

        # 從緩存中移除
        if self.current_master_idx in self.item_buttons:
            self.item_buttons[self.current_master_idx].destroy()
            del self.item_buttons[self.current_master_idx]

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

        new_row = {col: "" for col in sub_df.columns}
        new_row[fk_key] = self.current_master_pk

        siblings = sub_df[sub_df[fk_key] == self.current_master_pk]
        insert_idx = siblings.index.max() + 1 if not siblings.empty else len(sub_df)

        top = sub_df.iloc[:insert_idx]
        bottom = sub_df.iloc[insert_idx:]
        sub_df = pd.concat([top, pd.DataFrame([new_row]), bottom], ignore_index=True)
        self.manager.sub_dfs[full_sub_name] = sub_df
        self.manager.dirty = True

        # 重新載入該 Tab（會自動重用行）
        self._update_sub_table_data(current_tab, full_sub_name, self.current_master_pk)

    def delete_sub_item(self, sheet_full_name, row_idx):
        """ 刪除子表資料 (由每一列的 X 按鈕觸發) """
        if not messagebox.askyesno("確認", "刪除此列子表資料？"): return

        sub_df = self.manager.sub_dfs[sheet_full_name]
        sub_df.drop(row_idx, inplace=True)
        sub_df.reset_index(drop=True, inplace=True)
        self.manager.sub_dfs[sheet_full_name] = sub_df
        self.manager.dirty = True

        try:
            current_tab = self.sub_tables_tabs.get()
        except:
            current_tab = None

        self.load_sub_tables(self.current_master_pk)

        if current_tab:
            try:
                self.sub_tables_tabs.set(current_tab)
            except:
                pass

    def load_sub_tables(self, master_id):
        """載入子表（保持原版實現或使用優化版）"""
        # 取得相關子表
        related_sheets = [s for s in self.manager.sub_dfs if s.startswith(self.sheet_name + "#")]

        if not related_sheets:
            # 清空所有 Tab
            for tab_name in list(self.sub_tables_tabs._tab_dict.keys()):
                self.sub_tables_tabs.delete(tab_name)
            self.sub_tables_tabs.add("無子表")
            return

        # 取得現有和需要的 Tab
        existing_tabs = set(self.sub_tables_tabs._tab_dict.keys())
        needed_tabs = {s.split("#")[1] for s in related_sheets}

        # 1. 移除不需要的 Tab
        for tab in (existing_tabs - needed_tabs):
            self.sub_tables_tabs.delete(tab)
            # 清理緩存
            if tab in self.sub_table_frames:
                del self.sub_table_frames[tab]
            if tab in self.sub_table_headers:
                del self.sub_table_headers[tab]
            if tab in self.sub_table_row_pools:
                del self.sub_table_row_pools[tab]
            if tab in self.sub_table_active_rows:
                del self.sub_table_active_rows[tab]

        # 2. 為每個子表更新內容
        for sheet in related_sheets:
            short_name = sheet.split("#")[1]

            # 如果 Tab 不存在，創建它
            if short_name not in existing_tabs:
                self.sub_tables_tabs.add(short_name)
                self._create_sub_table_structure(short_name)

            # 更新該 Tab 的資料（重用行）
            self._update_sub_table_data(short_name, sheet, master_id)

    def _create_sub_table_structure(self, tab_name):
        """創建子表的固定結構 (標題固定在頂部，資料區域獨立捲動)"""
        tab_frame = self.sub_tables_tabs.tab(tab_name)

        # 1. 外層容器
        scroll_container = ctk.CTkFrame(tab_frame)
        scroll_container.pack(fill="both", expand=True)

        # 2. 固定標題區 (header_canvas — 只做水平捲動，不做垂直捲動)
        header_canvas = tk.Canvas(scroll_container, bg=_BG_HEADER,
                                  highlightthickness=0, bd=0, height=30)
        header_canvas.pack(fill="x", side="top")

        header_container = tk.Frame(header_canvas, bg=_BG_HEADER, height=30)
        header_canvas_window = header_canvas.create_window((0, 0), window=header_container, anchor="nw")

        def on_header_configure(_event=None):
            header_canvas.configure(scrollregion=header_canvas.bbox("all"))
            # 更新 header_canvas 的高度以符合內容
            h = header_container.winfo_reqheight()
            header_canvas.configure(height=h)

        header_container.bind("<Configure>", on_header_configure)

        # 3. 資料區容器 (canvas + scrollbars)
        body_frame = tk.Frame(scroll_container, bg=_BG)
        body_frame.pack(fill="both", expand=True)

        canvas = ctk.CTkCanvas(body_frame, bg=_BG, highlightthickness=0)

        v_scrollbar = ctk.CTkScrollbar(body_frame, orientation="vertical",
                                       command=canvas.yview)
        v_scrollbar.pack(side="right", fill="y")

        # 水平捲軸同步驅動 data canvas 和 header canvas
        def xview_sync(*args):
            canvas.xview(*args)
            header_canvas.xview(*args)

        h_scrollbar = ctk.CTkScrollbar(body_frame, orientation="horizontal",
                                       command=xview_sync)
        h_scrollbar.pack(side="bottom", fill="x")

        canvas.pack(side="left", fill="both", expand=True)

        # data canvas 的 xscrollcommand 同步更新 scrollbar 和 header 位置
        def on_data_xscroll(*args):
            h_scrollbar.set(*args)
            header_canvas.xview_moveto(args[0])

        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=on_data_xscroll)

        # 4. 資料框架 (使用原生 tk.Frame 避免殘影)
        data_container = tk.Frame(canvas, bg=_BG)
        canvas_window = canvas.create_window((0, 0), window=data_container, anchor="nw")

        # === 視窗縮放邏輯 ===
        def sync_header_width():
            """同步 header canvas 的寬度與 scrollregion 以配合資料 canvas"""
            req_w = max(data_container.winfo_reqwidth(),
                        header_container.winfo_reqwidth(),
                        canvas.winfo_width())
            header_canvas.itemconfig(header_canvas_window, width=req_w)
            header_canvas.configure(scrollregion=(0, 0, req_w,
                                                  header_container.winfo_reqheight()))

        def _update_widths():
            """統一更新 data canvas window 寬度 + header 同步"""
            req_width = max(data_container.winfo_reqwidth(),
                            header_container.winfo_reqwidth())
            canvas_w = canvas.winfo_width()
            target_w = max(canvas_w, req_width)
            canvas.itemconfig(canvas_window, width=target_w)
            bbox = canvas.bbox("all")
            if bbox:
                canvas_h = canvas.winfo_height()
                region_h = max(bbox[3], canvas_h)
                canvas.configure(scrollregion=(bbox[0], bbox[1], bbox[2], region_h))
            sync_header_width()

        def on_frame_configure(_event=None):
            _update_widths()

        def on_canvas_configure(_event=None):
            _update_widths()

        data_container.bind("<Configure>", on_frame_configure)
        canvas.bind("<Configure>", on_canvas_configure)

        # 標記 Canvas 用途 (供 App 層級滾輪路由識別)
        canvas._is_sub_table_canvas = True
        header_canvas._is_sub_table_header_canvas = True
        header_canvas._linked_data_canvas = canvas

        # 緩存容器
        self.sub_table_frames[tab_name] = {
            'header': header_container,
            'header_canvas': header_canvas,
            'data': data_container,
            'canvas': canvas,
            'scroll_container': scroll_container
        }
        self.sub_table_row_pools[tab_name] = []
        self.sub_table_active_rows[tab_name] = []

    def _update_sub_table_data(self, tab_name, sheet_full_name, master_id):
        """更新子表資料（智能重用行）"""

        # 取得資料
        sub_df = self.manager.sub_dfs[sheet_full_name]
        sub_cfg = self.cfg.get("sub_sheets", {}).get(tab_name, {})
        sub_cols_cfg = sub_cfg.get("columns", {})
        fk = sub_cfg.get("foreign_key", self.pk_key)

        if fk not in sub_df.columns:
            self._show_error_in_tab(tab_name, f"錯誤: 找不到關鍵欄位 {fk}")
            return

        # 篩選資料
        try:
            mask = sub_df[fk].astype(str) == str(master_id)
            filtered_rows = sub_df[mask]
        except:
            filtered_rows = sub_df.head(0)

        # 取得容器
        frames = self.sub_table_frames.get(tab_name)
        if not frames:
            return

        header_frame = frames['header']
        data_frame = frames['data']

        # 更新標題（只在需要時）
        headers = list(sub_df.columns)
        if not header_frame.winfo_children():
            self._build_sub_table_header(header_frame, headers, sub_cols_cfg)

        # === 行池機制 ===
        # 1. 回收正在使用的行到池中
        active_rows = self.sub_table_active_rows.get(tab_name, [])
        row_pool = self.sub_table_row_pools.get(tab_name, [])

        for row_frame in active_rows:
            row_frame.pack_forget()  # 隱藏但不銷毀
            row_pool.append(row_frame)

        self.sub_table_active_rows[tab_name] = []

        # 2. 處理資料
        if filtered_rows.empty:
            # 顯示空資料提示
            if not hasattr(data_frame, '_empty_label'):
                data_frame._empty_label = ctk.CTkLabel(data_frame, text="(此項目無資料)", text_color="gray")
            data_frame._empty_label.pack(pady=10)
            return
        else:
            # 隱藏空資料提示
            if hasattr(data_frame, '_empty_label'):
                data_frame._empty_label.pack_forget()

        # 3. 為每一行資料分配或創建 row_frame
        needed_rows = len(filtered_rows)
        available_rows = len(row_pool)

        new_active_rows = []

        for i, (idx, row) in enumerate(filtered_rows.iterrows()):
            row_bg = _ROW_EVEN if i % 2 == 0 else _ROW_ODD
            if i < available_rows:
                # 重用現有的行
                row_frame = row_pool[i]
                self._update_sub_table_row(row_frame, headers, row, idx, sheet_full_name, sub_cols_cfg)
                row_frame.configure(bg=row_bg)
                row_frame.pack(fill="x", pady=1, padx=2, ipady=3)
            else:
                # 創建新行
                row_frame = self._create_sub_table_row(data_frame, headers, row, idx, sheet_full_name, sub_cols_cfg)
                row_frame.configure(bg=row_bg)
                row_frame.pack(fill="x", pady=1, padx=2, ipady=3)

            new_active_rows.append(row_frame)

        # 4. 更新活躍行列表和池
        self.sub_table_active_rows[tab_name] = new_active_rows
        self.sub_table_row_pools[tab_name] = row_pool[needed_rows:]  # 剩餘未使用的保留在池中

    def _build_sub_table_header(self, header_frame, headers, cols_cfg):
        """建立子表標題（只執行一次）"""
        # 操作欄
        ctk.CTkLabel(header_frame, text="操作", width=60,
                     font=("微軟正黑體", 10, "bold")).pack(side="left", padx=2)

        # 資料欄
        for col in headers:
            col_info = cols_cfg.get(col, {})
            is_linked = col_info.get("link_to_text", False)

            # 如果是連結欄位，標題加上標記
            label_text = f"{col} 🔗" if is_linked else col
            width = 180 if is_linked else 120

            ctk.CTkLabel(header_frame, text=label_text, width=width,
                         font=("微軟正黑體", 10, "bold")).pack(side="left", padx=2)

    def _create_sub_table_row(self, parent, headers, row_data, row_idx, sheet_name, cols_cfg):
        """創建新的資料行（當池中沒有可用行時）"""
        row_frame = tk.Frame(parent, bg=_BG)

        # 儲存元數據（用於後續更新）
        row_frame._widgets = {}
        row_frame._vars = {}
        row_frame._trace_ids = {}

        # 刪除按鈕
        del_btn = ctk.CTkButton(row_frame, text="X", width=60, height=25,
                                fg_color="darkred", hover_color="#800000")
        del_btn.pack(side="left", padx=2)
        row_frame._widgets['delete_btn'] = del_btn

        # 資料欄位
        for col in headers:
            col_info = cols_cfg.get(col, {"type": "string"})
            col_type = col_info.get("type", "string")
            is_linked = col_info.get("link_to_text", False)

            if is_linked:
                # Key + Text 兩個欄位
                key_entry = ctk.CTkEntry(row_frame, width=60, height=25, text_color="gray")
                key_entry.configure(state="disabled")
                key_entry.pack(side="left", padx=1)

                text_var = ctk.StringVar()
                text_entry = ctk.CTkEntry(row_frame, textvariable=text_var, width=120, height=25)
                text_entry.pack(side="left", padx=1)

                row_frame._widgets[col] = (key_entry, text_entry)
                row_frame._vars[col] = text_var

            elif col_type == "enum":
                var = ctk.StringVar()
                menu = ctk.CTkOptionMenu(row_frame, values=col_info.get("options", ["None"]),
                                         variable=var, width=120, height=25)
                menu.pack(side="left", padx=2)

                row_frame._widgets[col] = menu
                row_frame._vars[col] = var

            elif col_type == "bool":
                var = ctk.BooleanVar()
                chk = ctk.CTkCheckBox(row_frame, text="", variable=var, width=120, height=25)
                chk.pack(side="left", padx=2)

                row_frame._widgets[col] = chk
                row_frame._vars[col] = var

            else:  # string, int, float
                var = ctk.StringVar()
                entry = ctk.CTkEntry(row_frame, textvariable=var, width=120, height=25)
                entry.pack(side="left", padx=2)

                row_frame._widgets[col] = entry
                row_frame._vars[col] = var

        # 填充初始數據
        self._update_sub_table_row(row_frame, headers, row_data, row_idx, sheet_name, cols_cfg)

        return row_frame

    def _update_sub_table_row(self, row_frame, headers, row_data, row_idx, sheet_name, cols_cfg):
        """更新資料行的內容（重用時調用）"""

        # 1. 清理舊的 trace
        for col, trace_id in row_frame._trace_ids.items():
            if col in row_frame._vars:
                try:
                    row_frame._vars[col].trace_remove("write", trace_id)
                except:
                    pass
        row_frame._trace_ids.clear()

        # 2. 更新刪除按鈕的命令
        del_btn = row_frame._widgets.get('delete_btn')
        if del_btn:
            del_btn.configure(command=lambda: self.delete_sub_item(sheet_name, row_idx))

        # 3. 更新每個欄位的值
        for col in headers:
            val = row_data[col]
            col_info = cols_cfg.get(col, {"type": "string"})
            col_type = col_info.get("type", "string")
            is_linked = col_info.get("link_to_text", False)

            if col not in row_frame._widgets:
                continue

            if is_linked:
                key_entry, text_entry = row_frame._widgets[col]

                # 更新 Key
                key_entry.configure(state="normal")
                key_entry.delete(0, "end")
                key_entry.insert(0, str(val))
                key_entry.configure(state="disabled")

                # 更新 Text
                real_text = self.manager.get_text_value(val)
                var = row_frame._vars[col]
                var.set(str(real_text))

                # 重新綁定 trace
                trace_id = var.trace_add("write",
                                         lambda *args, k=val, v=var: self.manager.update_linked_text(k, v.get()))
                row_frame._trace_ids[col] = trace_id

            elif col_type == "enum":
                var = row_frame._vars[col]
                var.set(str(val))

                # 更新命令
                menu = row_frame._widgets[col]
                menu.configure(command=lambda v, r=row_idx, c=col, s=sheet_name:
                self.manager.update_cell(True, s, r, c, v))

            elif col_type == "bool":
                var = row_frame._vars[col]
                var.set(bool(val) if val != "" else False)

                # 更新命令
                chk = row_frame._widgets[col]
                chk.configure(command=lambda r=row_idx, c=col, s=sheet_name, v=var:
                self.manager.update_cell(True, s, r, c, v.get()))

            else:  # string, int, float
                var = row_frame._vars[col]
                var.set(str(val))

                # 重新綁定 trace
                trace_id = var.trace_add("write",
                                         lambda *args, s=sheet_name, r=row_idx, c=col, v=var:
                                         self.manager.update_cell(True, s, r, c, v.get()))
                row_frame._trace_ids[col] = trace_id

    def _show_error_in_tab(self, tab_name, message):
        """在 Tab 中顯示錯誤訊息"""
        frames = self.sub_table_frames.get(tab_name)
        if not frames:
            return

        data_frame = frames['data']

        # 清空內容
        for widget in data_frame.winfo_children():
            widget.pack_forget()

        # 顯示錯誤
        ctk.CTkLabel(data_frame, text=message, text_color="red").pack(pady=20)

    def _render_simple_sub_table(self, parent, sheet_name, filtered_df, cols_cfg):
        """簡化版子表渲染（減少卡頓）"""
        if filtered_df.empty:
            ctk.CTkLabel(parent, text="(此項目無資料)", text_color="gray").pack(pady=10)
            return

        # 使用簡單的滾動框架
        scroll_frame = LightScrollableFrame(parent)
        scroll_frame.pack(fill="both", expand=True)

        headers = list(filtered_df.columns)

        # 標題列
        header_frame = tk.Frame(scroll_frame.interior, bg=_BG_HEADER)
        header_frame.pack(fill="x", pady=(0, 5))

        ctk.CTkLabel(header_frame, text="操作", width=60, font=("微軟正黑體", 10, "bold")).pack(side="left", padx=2)
        for col in headers:
            ctk.CTkLabel(header_frame, text=col, width=120, font=("微軟正黑體", 10, "bold")).pack(side="left", padx=2)

        # 資料列
        for i, (idx, row) in enumerate(filtered_df.iterrows()):
            row_bg = _ROW_EVEN if i % 2 == 0 else _ROW_ODD
            row_frame = tk.Frame(scroll_frame.interior, bg=row_bg)
            row_frame.pack(fill="x", pady=1, ipady=3)

            # 刪除按鈕
            ctk.CTkButton(row_frame, text="X", width=60, height=25, fg_color="darkred",
                          command=lambda s=sheet_name, r=idx: self.delete_sub_item(s, r)).pack(side="left", padx=2)

            # 資料欄位
            for col in headers:
                val = row[col]
                c_info = cols_cfg.get(col, {"type": "string"})
                col_type = c_info.get("type", "string")
                is_linked = c_info.get("link_to_text", False)

                if is_linked:
                    # Key + Text
                    key_entry = ctk.CTkEntry(row_frame, width=60, height=25)
                    key_entry.insert(0, str(val))
                    key_entry.configure(state="disabled")
                    key_entry.pack(side="left", padx=1)

                    real_text = self.manager.get_text_value(val)
                    var = ctk.StringVar(value=str(real_text))
                    text_entry = ctk.CTkEntry(row_frame, textvariable=var, width=100, height=25)
                    text_entry.pack(side="left", padx=1)
                    var.trace_add("write", lambda *args, k=val, v=var: self.manager.update_linked_text(k, v.get()))

                elif col_type == "enum":
                    var = ctk.StringVar(value=str(val))
                    menu = ctk.CTkOptionMenu(row_frame, values=c_info.get("options", ["None"]),
                                             variable=var, width=120, height=25,
                                             command=lambda v, r=idx, c=col, s=sheet_name: self.manager.update_cell(
                                                 True, s, r, c, v))
                    menu.pack(side="left", padx=2)
                elif col_type == "bool":
                    var = ctk.BooleanVar(value=bool(val) if val != "" else False)
                    chk = ctk.CTkCheckBox(row_frame, text="", variable=var, width=120, height=25,
                                          command=lambda r=idx, c=col, s=sheet_name, v=var: self.manager.update_cell(
                                              True, s, r, c, v.get()))
                    chk.pack(side="left", padx=2)
                else:
                    var = ctk.StringVar(value=str(val))
                    entry = ctk.CTkEntry(row_frame, textvariable=var, width=120, height=25)
                    entry.pack(side="left", padx=2)
                    var.trace_add("write",
                                  lambda *args, s=sheet_name, r=idx, c=col, v=var: self.manager.update_cell(True, s, r,
                                                                                                            c, v.get()))

    def cleanup(self):
        """清理資源"""
        # 清理母表 trace
        for col, trace_id in self.trace_ids.items():
            if col in self.master_field_vars:
                try:
                    self.master_field_vars[col].trace_remove("write", trace_id)
                except:
                    pass

        # 清理子表 trace
        for tab_name, active_rows in self.sub_table_active_rows.items():
            for row_frame in active_rows:
                for col, trace_id in row_frame._trace_ids.items():
                    if col in row_frame._vars:
                        try:
                            row_frame._vars[col].trace_remove("write", trace_id)
                        except:
                            pass

        # 清空所有緩存
        self.cls_buttons.clear()
        self.item_buttons.clear()
        self.master_fields.clear()
        self.master_field_vars.clear()
        self.trace_ids.clear()
        self.sub_table_frames.clear()
        self.sub_table_headers.clear()
        self.sub_table_row_pools.clear()
        self.sub_table_active_rows.clear()
        self.current_image_ref = None

        # 銷毀所有子 widget
        for widget in self.winfo_children():
            widget.destroy()

        import gc
        gc.collect()

    def destroy(self):
        """重寫 destroy"""
        self.cleanup()
        super().destroy()

class ConfigEditorWindow(ctk.CTkToplevel):
    """配置設定視窗"""
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

        # ========= 圖片與文字表設定 =========
        # 將原本只有 Icon 的區塊擴充
        setting_block = ctk.CTkFrame(header, fg_color="transparent")
        setting_block.grid(row=1, column=0, sticky="w", pady=(0, 5))

        # 1. 文字表路徑設定 (新增)
        self.frame_text_path = ctk.CTkFrame(setting_block, fg_color="transparent")
        self.frame_text_path.pack(anchor="w", padx=10, pady=2)
        ctk.CTkLabel(self.frame_text_path, text="外部文字表路徑 (.xlsx):").pack(side="left")

        self.entry_text_path = ctk.CTkEntry(self.frame_text_path, width=300)
        self.entry_text_path.pack(side="left", padx=5)
        self.entry_text_path.insert(0, self.manager.config.get("global_text_path", ""))

        ctk.CTkButton(self.frame_text_path, text="瀏覽", width=50,
                      command=self.browse_text_file).pack(side="left")

        # 2. 圖片設定 (原本的)
        self.var_use_icon = ctk.BooleanVar(value=False)
        self.chk_use_icon = ctk.CTkCheckBox(
            setting_block,
            text="啟用圖示顯示 (需有 'Icon' 或 'Image' 欄位)",
            variable=self.var_use_icon,
            command=self.toggle_icon_input
        )
        self.chk_use_icon.pack(anchor="w", padx=10, pady=(10, 0))

        self.frame_img_path = ctk.CTkFrame(setting_block, fg_color="transparent")
        self.frame_img_path.pack(anchor="w", padx=30, pady=2)

        ctk.CTkLabel(self.frame_img_path, text="圖片資料夾:").pack(side="left")
        self.entry_img_path = ctk.CTkEntry(self.frame_img_path, width=300)
        self.entry_img_path.pack(side="left", padx=5)
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
                        col: {
                            "type": "string",
                            "link_to_text": False
                        }
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
            self.frame_img_path.pack(anchor="w", padx=30, pady=2)
        else:
            self.frame_img_path.pack_forget()

    def on_image_path_change(self, event=None):
        sheet = self.tab_view.get()
        self.manager.config[sheet]["image_path"] = self.entry_img_path.get()

    def browse_text_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if path:
            self.entry_text_path.delete(0, "end")
            self.entry_text_path.insert(0, path)
            # 暫存到 manager config (尚未存檔)
            self.manager.config["global_text_path"] = path
            # 立即嘗試載入，以便後續編輯使用
            success = self.manager.load_external_text(path)
            if success:
                messagebox.showinfo("成功", "已載入外部文字表")
            else:
                messagebox.showerror("失敗", "載入失敗，請檢查格式")

    # ================== Tab 內容 ==================

    def build_tab_content(self, tab, sheet_name):
        main_scroll = LightScrollableFrame(tab)
        main_scroll.pack(fill="both", expand=True)

        cfg = self.manager.config[sheet_name]
        all_cols = list(self.manager.master_dfs[sheet_name].columns)

        # --- 母表設定 ---
        base_frame = ctk.CTkFrame(main_scroll.interior)
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
            main_scroll.interior,
            text="母表欄位類型設定",
            font=("微軟正黑體", 13, "bold"),
            text_color="#3B8ED0"
        ).pack(pady=5)

        for col in all_cols:
            line = tk.Frame(main_scroll.interior, bg=_BG)
            line.pack(fill="x", padx=20, pady=1)

            ctk.CTkLabel(
                line,
                text=col,
                width=150,
                anchor="w"
            ).pack(side="left")

            # 確保 config 結構完整
            if col not in cfg["columns"]:
                cfg["columns"][col] = {
                    "type": "string",
                    "link_to_text": False
                }
            else:
                cfg["columns"][col].setdefault("link_to_text", False)

            # ---------- link_to_text 勾選 ----------
            link_var = ctk.BooleanVar(
                value=cfg["columns"][col]["link_to_text"]
            )

            def on_toggle_link(c=col, v=link_var):
                cfg["columns"][c]["link_to_text"] = v.get()

            ctk.CTkCheckBox(
                line,
                text="連結文字",
                variable=link_var,
                command=on_toggle_link
            ).pack(side="right", padx=6)

            # ---------- type 選單 ----------
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
                main_scroll.interior,
                text="子表欄位類型設定",
                font=("微軟正黑體", 13, "bold"),
                text_color="#E38D2D"
            ).pack(pady=10)

            for sub_full in related_subs:
                short = sub_full.split("#")[1]
                sub_group = ctk.CTkFrame(main_scroll.interior, border_width=1, border_color="gray")
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
                    s_line = tk.Frame(sub_group, bg=_BG)
                    s_line.pack(fill="x", padx=15, pady=1)

                    ctk.CTkLabel(
                        s_line, text=s_col, width=150, anchor="w"
                    ).pack(side="left")

                    if s_col not in cfg["sub_sheets"][short]["columns"]:
                        cfg["sub_sheets"][short]["columns"][s_col] = {
                            "type": "string",
                            "link_to_text": False
                        }
                    else:
                        cfg["sub_sheets"][short]["columns"][s_col].setdefault("link_to_text", False)

                    # ---------- link_to_text 勾選 ----------
                    link_var = ctk.BooleanVar(
                        value=cfg["sub_sheets"][short]["columns"][s_col]["link_to_text"]
                    )

                    def on_toggle_link(c=s_col, v=link_var, s=short):
                        cfg["sub_sheets"][s]["columns"][c]["link_to_text"] = v.get()

                    ctk.CTkCheckBox(
                        s_line,
                        text="連結文字",
                        variable=link_var,
                        command=on_toggle_link
                    ).pack(side="right", padx=6)

                    # ---------- type 選單 ----------
                    st_menu = ctk.CTkOptionMenu(
                        s_line,
                        values=["string", "float", "int", "bool", "enum"],
                        width=100,
                        command=lambda v, sn=short, sc=s_col: self.set_sub_col_type(sheet_name, sn, sc, v)
                    )
                    st_menu.set(cfg["sub_sheets"][short]["columns"][s_col]["type"])
                    st_menu.pack(side="right")

    # ================== Config 操作 ==================

    def set_col_type(self, sheet_name, col, val):
        self.manager.config[sheet_name]["columns"][col]["type"] = val
        if val == "enum":
            self._ask_enum_options(self.manager.config[sheet_name]["columns"][col])

    def set_sub_col_type(self, m, s, col, val):
        self.manager.config[m]["sub_sheets"][s]["columns"][col]["type"] = val
        if val == "enum":
            self._ask_enum_options(self.manager.config[m]["sub_sheets"][s]["columns"][col])

    def _ask_enum_options(self, col_conf):
        """彈出視窗讓使用者輸入 enum 選項（逗號分隔）"""
        current = col_conf.get("options", [])
        current_str = ", ".join(current) if current else ""

        dialog = ctk.CTkToplevel(self)
        dialog.title("設定 Enum 選項")
        dialog.geometry("400x150")
        dialog.transient(self)
        dialog.grab_set()

        ctk.CTkLabel(dialog, text="請輸入選項（以逗號分隔）：").pack(padx=10, pady=(10, 5), anchor="w")
        var = ctk.StringVar(value=current_str)
        entry = ctk.CTkEntry(dialog, textvariable=var, width=360)
        entry.pack(padx=10, pady=5)
        entry.focus_set()

        def on_confirm():
            raw = var.get().strip()
            opts = [o.strip() for o in raw.split(",") if o.strip()] if raw else []
            col_conf["options"] = opts
            dialog.destroy()

        entry.bind("<Return>", lambda e: on_confirm())
        ctk.CTkButton(dialog, text="確認", command=on_confirm).pack(pady=10)

    def save_and_close(self):
        self.manager.save_config()
        self.destroy()
        if hasattr(self.master, "refresh_ui"):
            self.master.refresh_ui()

class App(ctk.CTk):
    """ 主畫面 """
    def __init__(self):
        super().__init__()
        self.title("Game Data Editor (Config Driven)")
        icon_path = os.path.join(getattr(sys, '_MEIPASS', os.path.abspath(".")), "icon.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)
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

        self.sheet_editors = []
        # 全域滑鼠滾輪路由 (根據游標位置決定捲動目標)
        self.bind_all("<MouseWheel>", self._route_mousewheel)
        self.bind_all("<Shift-MouseWheel>", self._route_shift_mousewheel)

        # 關閉視窗攔截
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        """關閉視窗前檢查未儲存的變更"""
        if self.manager.dirty:
            result = messagebox.askyesnocancel("資料未儲存", "有尚未儲存的變更，是否先儲存再關閉？")
            if result is None:  # Cancel
                return
            if result:  # Yes
                self.save_file()
        self.destroy()

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not path: return

        try:
            self.manager.load_excel(path)

            # 判斷是否需要強制彈出配置視窗
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
        self.sheet_editors = []
        for tab_name in list(self.main_tabs._tab_dict.keys()):
            self.main_tabs.delete(tab_name)

        for sheet_name in self.manager.master_dfs:
            self.main_tabs.add(sheet_name)
            parent = self.main_tabs.tab(sheet_name)
            # 實例化單一 Sheet 編輯器
            editor = SheetEditor(parent, sheet_name, self.manager)
            editor.pack(fill="both", expand=True)
            self.sheet_editors.append(editor)

    def open_configwnd(self):
        if not self.manager.master_dfs:
            messagebox.showinfo("提示", "請先匯入Excel後再進行參數的配置")
            return
        _ = ConfigEditorWindow(self, self.manager)

    def _route_mousewheel(self, event):
        """將滑鼠滾輪事件路由到游標所在的可捲動區域"""
        widget = self.winfo_containing(event.x_root, event.y_root)
        if not widget:
            return
        scroll_units = int(-1 * (event.delta / 120))
        w = widget
        while w:
            # 子表 Canvas
            if getattr(w, '_is_sub_table_canvas', False):
                top, bottom = w.yview()
                if bottom - top < 1.0:
                    w.yview_scroll(scroll_units, "units")
                return "break"
            # LightScrollableFrame 的內部 Canvas
            if getattr(w, '_is_light_scrollable', False):
                top, bottom = w.yview()
                if bottom - top < 1.0:
                    w.yview_scroll(scroll_units, "units")
                return "break"
            # CTkScrollableFrame (備用相容)
            if isinstance(w, ctk.CTkScrollableFrame):
                w._parent_canvas.yview_scroll(scroll_units, "units")
                return "break"
            try:
                w = w.master
            except:
                break

    def _route_shift_mousewheel(self, event):
        """將 Shift+滑鼠滾輪事件路由到游標所在的子表 Canvas (橫向捲動)"""
        widget = self.winfo_containing(event.x_root, event.y_root)
        if not widget:
            return
        scroll_units = int(-1 * (event.delta / 120))
        w = widget
        while w:
            if getattr(w, '_is_sub_table_canvas', False):
                w.xview_scroll(scroll_units, "units")
                return "break"
            # 支援在 header 區域也能觸發水平捲動 (找到相鄰的 data canvas)
            if getattr(w, '_is_sub_table_header_canvas', False):
                w._linked_data_canvas.xview_scroll(scroll_units, "units")
                return "break"
            try:
                w = w.master
            except:
                break

if __name__ == "__main__":
    app = App()
    app.mainloop()
