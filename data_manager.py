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
            return {} # 回傳空字典，後續讀取 Excel 時會補上
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)

    def load_excel(self, file_path):
        self.excel_path = file_path
        xl = pd.ExcelFile(file_path)
        self.need_config_alert = False # 重置標記
        
        # 讀取 Sheet 名稱
        for sheet in xl.sheet_names:
            if sheet.endswith(".json"):
                self.master_dfs[sheet] = pd.read_excel(xl, sheet).fillna("")
                # 如果這個母表不在配置檔裡，建立預設配置
                if sheet not in self.config:
                    self.config[sheet] = {
                        "classification_key": self.master_dfs[sheet].columns[0], # 預設第一欄
                        "primary_key": self.master_dfs[sheet].columns[0],
                        "columns": {col: {"type": "string"} for col in self.master_dfs[sheet].columns},
                        "sub_sheets": {}
                    }
                    self.need_config_alert = True # 觸發提醒
            elif "#" in sheet:
                self.sub_dfs[sheet] = pd.read_excel(xl, sheet).fillna("")

        if self.need_config_alert:
            self.save_config() # 存下生成的預設檔

    def save_config(self):
        with open(self.config_path, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, indent=4, ensure_ascii=False)

    def _write_df_to_sheet(self, book, sheet_name, df):
        # 如果 Sheet 不存在則建立，存在則清空內容重寫
        if sheet_name not in book.sheetnames:
            ws = book.create_sheet(sheet_name)
        else:
            ws = book[sheet_name]
            ws.delete_rows(2, ws.max_row) # 保留標題列，刪除數據

        # 寫入數據
        for row in dataframe_to_rows(df, index=False, header=False):
            ws.append(row)

    def update_cell(self, is_sub, sheet_name, row_idx, col_name, value):
        """ 當 UI 修改時呼叫此方法更新 DataFrame """
        target_dict = self.sub_dfs if is_sub else self.master_dfs
        if sheet_name in target_dict:
            df = target_dict[sheet_name]
            # 確保型別正確 (簡單處理)
            try:
                if df[col_name].dtype == 'int64': value = int(value)
                elif df[col_name].dtype == 'float64': value = float(value)
            except:
                pass 
            df.at[row_idx, col_name] = value