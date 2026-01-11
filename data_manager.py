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
        self.sub_dfs = {}     # 存放子表 DataFrame
        self.need_config_alert = False # 標記是否需要彈出配置視窗

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
                        "use_icon":"False",
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

    def save_config(self):
        with open(self.config_path, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, indent=4, ensure_ascii=False)

    def save_excel(self):
        """
        儲存正在編輯的Excel
        """
        if not self.excel_path:
            return

        # 情況 A: 檔案不存在，直接用 Pandas 建立新檔 (因為沒有格式需要保留)
        if not os.path.exists(self.excel_path):
            with pd.ExcelWriter(self.excel_path, engine='openpyxl', mode='w') as writer:
                # 這裡為了方便，還是可以用原本的 helper，或者直接在這裡寫
                for sheet, df in self.master_dfs.items():
                    df.to_excel(writer, sheet_name=sheet, index=False)
                for sheet, df in self.sub_dfs.items():
                    df.to_excel(writer, sheet_name=sheet, index=False)
            return

        # 情況 B: 檔案存在，載入它並修改內容
        try:
            wb = load_workbook(self.excel_path)

            # 更新母表
            for sheet_name, df in self.master_dfs.items():
                self._update_sheet_content(wb, sheet_name, df)

            # 更新子表
            for sheet_name, df in self.sub_dfs.items():
                self._update_sheet_content(wb, sheet_name, df)

            wb.save(self.excel_path)

        except Exception as e:
            print(f"儲存失敗: {e}")
            # 如果因為檔案被開啟而鎖定，這裡可以做錯誤處理

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

            # 或者，如果你不在意多餘列的格式，想直接刪除列 (會影響列高設定)：
            # ws.delete_rows(current_row_idx, max_row_in_excel - current_row_idx + 1)
    
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
                pass # print(f"轉型失敗: {col_name} -> {value} ({e})")
            
            # 更新 DataFrame
            df.at[row_idx, col_name] = value