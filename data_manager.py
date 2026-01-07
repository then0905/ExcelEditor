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
        """ 儲存檔案 """
        if not self.excel_path: return
        
        # 建立一個新的 Excel Writer
        with pd.ExcelWriter(self.excel_path, engine='openpyxl') as writer:
            # 寫入母表
            for sheet, df in self.master_dfs.items():
                df.to_excel(writer, sheet_name=sheet, index=False)
            
            # 寫入子表
            for sheet, df in self.sub_dfs.items():
                df.to_excel(writer, sheet_name=sheet, index=False)
    
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