import pandas as pd
import json
import os
import warnings
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


class DataManager:
    def __init__(self, config_path="config.json"):
        self.config_path = config_path
        self.config = self._load_config(config_path)
        self.excel_path = None
        self.master_dfs = {}  # 存放母表 DataFrame
        self.sub_dfs = {}  # 存放子表 DataFrame
        self.need_config_alert = False  # 標記是否需要彈出配置視窗

        # --- 外部文字表相關變數 ---
        self.text_df = None  # 存放文字表的完整 DataFrame
        self.text_dict = {}  # 快速查找用字典 {Key: Value}
        self.text_file_path = ""  # 文字表路徑
        self.text_key_col = "Key"  # 假設文字表的 Key 欄位名
        self.text_val_col = "Text"  # 假設文字表的 Value 欄位名
        self.text_modified = False  # 標記是否修改過
        self.text_sheetnames = []  # 文字表的工作表
        self.text_modifications = {}  # 記錄修改: {sheet_name: {key: new_value}}

        # 初始化時嘗試載入文字表 (如果 config 有寫)
        if "global_text_path" in self.config:
            self.load_external_text(self.config["global_text_path"])

    def _load_config(self, path):
        if not os.path.exists(path):
            return {}
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}

    def load_excel(self, file_path):
        self.excel_path = file_path
        xl = pd.ExcelFile(file_path)
        self.need_config_alert = False

        self.master_dfs = {}
        self.sub_dfs = {}

        # 讀取 Sheet 名稱
        for sheet in xl.sheet_names:
            if sheet.endswith(".json"):
                # 母表：強制轉字串
                self.master_dfs[sheet] = pd.read_excel(xl, sheet, dtype=str).fillna("")

                # 初始化配置
                if sheet not in self.config:
                    self.config[sheet] = {
                        "use_icon": "False",
                        "image_path": "",
                        "classification_key": self.master_dfs[sheet].columns[0],
                        "primary_key": self.master_dfs[sheet].columns[0],
                        "columns": {col: {"type": "string"} for col in self.master_dfs[sheet].columns},
                        "sub_sheets": {}
                    }
                    self.need_config_alert = True

            elif "#" in sheet:
                self.sub_dfs[sheet] = pd.read_excel(xl, sheet, dtype=str).fillna("")

        if self.need_config_alert:
            self.save_config()

    def load_external_text(self, path, key_col="TextID", val_col="TextContent"):
        if not os.path.exists(path):
            print(f"文字表路徑不存在: {path}")
            return False

        try:
            self.text_file_path = path
            self.text_sheetnames = []
            xls = pd.ExcelFile(path)

            self.text_dict = {}  # 重點：只要一個 dict

            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str).fillna("")
                self.text_sheetnames.append(sheet_name)

                # 檢查欄位
                if key_col not in df.columns or val_col not in df.columns:
                    print(f"[{sheet_name}] 缺少欄位，改用前兩欄")
                    if len(df.columns) < 2:
                        continue
                    k = df.columns[0]
                    v = df.columns[1]
                else:
                    k = key_col
                    v = val_col

                # 合併進同一個 dict，並記錄來源工作表和欄位
                for key, val in zip(df[k], df[v]):
                    # 儲存格式: {key: {"value": val, "sheet": sheet_name, "key_col": k, "val_col": v}}
                    self.text_dict[key] = {
                        "value": val,
                        "sheet": sheet_name,
                        "key_col": k,
                        "val_col": v
                    }

            return True

        except Exception as e:
            print(f"載入文字表失敗: {e}")
            return False

    def save_config(self):
        with open(self.config_path, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, indent=4, ensure_ascii=False)

    def save_excel(self):
        """
        儲存正在編輯的Excel
        """
        if not self.excel_path:
            return

        # 檔案不存在，直接用 Pandas 建立新檔 (因為沒有格式需要保留)
        try:
            if not os.path.exists(self.excel_path):
                with pd.ExcelWriter(self.excel_path, engine='openpyxl', mode='w') as writer:
                    for sheet, df in self.master_dfs.items():
                        df.to_excel(writer, sheet_name=sheet, index=False)
                    for sheet, df in self.sub_dfs.items():
                        df.to_excel(writer, sheet_name=sheet, index=False)
            else:
                wb = load_workbook(self.excel_path)
                for sheet_name, df in self.master_dfs.items():
                    self._update_sheet_content(wb, sheet_name, df)
                for sheet_name, df in self.sub_dfs.items():
                    self._update_sheet_content(wb, sheet_name, df)
                wb.save(self.excel_path)

            # 2. 存外部文字表 (如果有修改)
            if getattr(self, "text_modified", False) and getattr(self, "text_file_path", None):
                try:
                    print(f"正在同步儲存外部文字表至: {self.text_file_path}")
                    self._save_external_text()
                    self.text_modified = False
                    self.text_modifications = {}
                    print("外部文字表儲存成功")

                except Exception as e:
                    print(f"外部文字表儲存失敗: {e}")


        except Exception as e:
            print(f"儲存失敗: {e}")
            raise e  # 拋出錯誤讓 UI 顯示

    def _save_external_text(self):
        """
        儲存外部文字表：只修改有變動的儲存格，保留格式
        """
        if not self.text_file_path or not os.path.exists(self.text_file_path):
            return

        wb = load_workbook(self.text_file_path)

        # 遍歷所有修改記錄
        for key, new_value in self.text_modifications.items():
            if key not in self.text_dict:
                continue

            text_info = self.text_dict[key]
            sheet_name = text_info["sheet"]
            key_col_name = text_info["key_col"]
            val_col_name = text_info["val_col"]

            if sheet_name not in wb.sheetnames:
                print(f"警告: 工作表 {sheet_name} 不存在於文字表中")
                continue

            ws = wb[sheet_name]

            # 找到 key 和 value 欄位的索引
            header_row = 1
            key_col_idx = None
            val_col_idx = None

            for col_idx, cell in enumerate(ws[header_row], 1):
                if cell.value == key_col_name:
                    key_col_idx = col_idx
                if cell.value == val_col_name:
                    val_col_idx = col_idx

            if key_col_idx is None or val_col_idx is None:
                print(f"警告: 在 {sheet_name} 中找不到欄位 {key_col_name} 或 {val_col_name}")
                continue

            # 尋找對應的 key 所在的列
            found = False
            for row_idx in range(2, ws.max_row + 1):
                cell_key = ws.cell(row=row_idx, column=key_col_idx).value
                if str(cell_key) == str(key):
                    # 找到了，只更新 value 欄位的內容
                    ws.cell(row=row_idx, column=val_col_idx).value = new_value
                    found = True
                    break

            if not found:
                print(f"警告: 在 {sheet_name} 中找不到 key={key}")

        wb.save(self.text_file_path)

    def _update_sheet_content(self, wb, sheet_name, df):
        """
        核心邏輯：
        1. 找到工作表 (如果沒有就建立)
        2. 填入標題 (Header)
        3. 填入數據 (Body) - 只修改 value，不動 style
        4. 清除多餘的舊資料 (若新資料比舊資料短)
        """
        # 1. 取得或建立 Worksheet
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.create_sheet(sheet_name)
            # 新建立的表沒有格式問題，可以直接用 append 或是逐格寫入

        # 2. 寫入標題 (Row 1)
        # 註：這會保留標題列原本的顏色，只改文字
        for col_idx, col_name in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.value = col_name

        # 3. 寫入數據 (Row 2 ~ N)
        # 使用 itertuples 效能較好
        current_row_idx = 2
        for row in df.itertuples(index=False):
            for col_idx, value in enumerate(row, 1):
                cell = ws.cell(row=current_row_idx, column=col_idx)
                cell.value = value
            current_row_idx += 1

        # 4. 清理殘留數據
        # 如果原本 Excel 有 100 行，現在只有 80 行，必須把第 81~100 行清空
        max_row_in_excel = ws.max_row
        if max_row_in_excel >= current_row_idx:
            # 刪除多餘的列內容
            for r in range(current_row_idx, max_row_in_excel + 1):
                for c in range(1, ws.max_column + 1):
                    ws.cell(row=r, column=c).value = None

            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
            else:
                ws = wb.create_sheet(sheet_name)

            for col_idx, col_name in enumerate(df.columns, 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.value = col_name

            current_row_idx = 2
            for row in df.itertuples(index=False):
                for col_idx, value in enumerate(row, 1):
                    cell = ws.cell(row=current_row_idx, column=col_idx)
                    cell.value = value
                current_row_idx += 1

            max_row_in_excel = ws.max_row
            if max_row_in_excel >= current_row_idx:
                for r in range(current_row_idx, max_row_in_excel + 1):
                    for c in range(1, ws.max_column + 1):
                        ws.cell(row=r, column=c).value = None

    def update_cell(self, is_sub, sheet_name, row_idx, col_name, value):
        """
        當 UI 修改時呼叫此方法更新 DataFrame
        修正：依照 Config 的設定來轉型，而不是依賴 DataFrame 原本的 type
        """
        target_dict = self.sub_dfs if is_sub else self.master_dfs

        if sheet_name in target_dict:
            df = target_dict[sheet_name]

            # 依照 Config 決定儲存型別
            try:
                # 取得該欄位的目標類型
                col_type = "string"

                if not is_sub:
                    # 母表配置
                    if sheet_name in self.config:
                        col_type = self.config[sheet_name]["columns"].get(col_name, {}).get("type", "string")
                else:
                    # 子表配置 (需解析 sheet_name: "Master#Sub")
                    master_name = sheet_name.split("#")[0]
                    sub_name = sheet_name.split("#")[1]
                    if master_name in self.config:
                        col_type = self.config[master_name]["sub_sheets"].get(sub_name, {}) \
                            .get("columns", {}).get(col_name, {}).get("type", "string")

                # 進行轉型
                if col_type == "int":
                    value = int(value)
                elif col_type == "float":
                    value = float(value)
                elif col_type == "bool":
                    # 處理布林值的多種輸入可能
                    if isinstance(value, str):
                        value = value.lower() in ['true', '1', 'yes']
                    else:
                        value = bool(value)

            except Exception as e:
                # 轉型失敗就維持字串，避免崩潰
                pass  # print(f"轉型失敗: {col_name} -> {value} ({e})")

            # 檢查是否為聯動欄位
            col_conf = self.config.get(sheet_name, {}).get("columns", {}).get(col_name, {})
            if col_conf.get("link_to_text"):
                raw_key = self.master_dfs[sheet_name].at[row_idx, col_name]
                # 更新外部文字緩存，並標記『文字表需儲存』
                self._update_external_text(raw_key, value)
            else:
                # 正常的 DataFrame 更新
                df.at[row_idx, col_name] = value

            # 更新 DataFrame
            df.at[row_idx, col_name] = value

    def _update_external_text(self, key, new_value):
        """
        內部方法：記錄文字表的修改
        """
        if not hasattr(self, "text_modifications"):
            self.text_modifications = {}

        self.text_modifications[str(key)] = str(new_value)

        # 同時更新記憶體中的 text_dict
        if str(key) in self.text_dict:
            self.text_dict[str(key)]["value"] = str(new_value)

        self.text_modified = True

    def get_text_value(self, key):
        """ 給 UI 呼叫：透過 Key 拿文字 """
        if not self.text_dict:
            return key  # 沒載入文字表就直接回傳 Key

        text_info = self.text_dict.get(str(key))
        if text_info:
            return text_info["value"]
        else:
            return f"<{key} Missing>"

    def update_linked_text(self, key, new_text):
        """ 給 UI 呼叫：更新文字內容 (不改 Key) """
        if not hasattr(self, "text_dict") or self.text_dict is None:
            return

        self._update_external_text(key, new_text)