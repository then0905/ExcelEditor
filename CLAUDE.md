# SkillExcelEditor (JsonEditor) — 開發規範

## 專案概述
- **PySide6 (Qt)** 製作的暗色主題遊戲資料編輯器；後端編輯的是 **JSON** 檔（早期的 CTk/tk + Excel 版本已淘汰）。
- 關鍵檔案：
  - `main.py` — 全部 UI（App 視窗、TableEditor、欄位編輯器、子表 model/view、配置視窗、驗證規則視窗、打包入口）。
  - `json_data_manager.py` — 資料層 `JsonDataManager`（載入/儲存 JSON、欄位/子表增刪改、健檢）。
  - `validation.py` — 資料驗證規則引擎（無 Qt 依賴；規則模型、AST 白名單表達式、增量重驗）。
  - `data_manager.py` — 舊的 Excel/openpyxl 資料層（目前主程式未使用，保留參考）。
  - `config.json` — 以「JSON 檔的正規化絕對路徑」為 key 的設定檔（每個資料檔一份設定）。
  - `assets/tabler-icons.ttf` — 內建的 Tabler 圖示字型（工具列圖示用，須隨打包帶上）。
- 開發目標：穩定、流暢、生產品質（零 bug、零 UI 卡頓）。

---

## 資料模型（兩層，務必理解）
載入 `.json` 後（見 `JsonDataManager.load_json`）：
- 頂層是「物件陣列」→ 每筆是一列母表資料，存進 `tables[表名]`（pandas DataFrame，全字串）。
- 母表某欄若是 **array-of-objects** → 抽成子表，存進 `sub_tables["母表名.欄位名"]`，並注入 FK 欄（= 母表 `primary_key`）。
- 母表某欄若是 **array-of-primitives** → 以逗號合併成字串顯示。
- **目前只支援兩層**（母表 + 一層子表）；子表列裡再包 array-of-objects 不被支援。
- 儲存（`save_json`）會把子表依 FK 重組回各母表記錄的巢狀陣列；注入的 FK 欄預設不輸出（見 `_sub_fk_is_data`）。
- 沒有資料的子表靠 config 定義在重載時重建（`load_json` 尾段），避免空子表消失。

### config 結構（每個資料檔）
```
{ "<json路徑>": { "<表名>": {
    "primary_key": ..., "classification_key": ...,
    "use_icon": bool, "image_path"/"image_preview": ...,
    "text_ref_source": {json_path,key_col,val_col},
    "columns": { "<欄>": {"type","note","options","suggest_from"} },
    "sub_tables": { "<子表>": {"foreign_key","note","columns":{...}} }
} } }
```
欄位 `type`：`string / int / float / bool / enum / array / text_ref`。
`note` 會在欄位標題 / 子表分頁標題 hover 時顯示；子表的 `note` 在「配置設定」視窗每張子表底下編輯。

### 資料驗證規則（validations）
- 每張母表 config 下的 `validations` 陣列（規則結構見 `validation.normalize_rule` 與
  `docs/superpowers/specs/2026-07-12-validation-rules-design.md`）。
- 語意：`when` 成立而 `then` 不成立 → 違規；`when` 留空 = 每列都檢查 `then`。
  `scope: ""`=母表、`"子表名"`=子表；子表規則可用 `master.欄位`，母表規則可下子表聚合條件
  （builder 的 `{"agg": …}` 或 expr 的 `any_sub()/count_sub()`）。
- 引擎掛在 `manager.validator`（`ValidationEngine`）：`load_json` 尾端 `reload()` 全量驗證、
  `update_cell` 觸發 `on_cell_edited` 增量重驗（改 PK/FK 會整表重驗）；列增刪/搬移後 UI 端呼叫
  `validate_table()`（見 `_reload_all` / `_refresh_sub_tables(revalidate=True)`）。
- 上色：子表格背景（規則色 alpha 120，優先於 dirty 黃）、母表欄位編輯器邊框＋標籤、
  項目卡片右上小點（紅=error/黃=warn）、子表分頁標題轉紅＋tooltip 計數。
- 存檔閘門：`save_file` 先 `validate_all()`；error 擋存檔、warn 可「仍要儲存」，
  清單點擊經 `navigate_to` 跳到違規儲存格。關閉程式的批次儲存**不**過閘門。
- 表達式安全：`SafeExpr` 只允許比較/布林/四則/白名單函式，`ast` 驗證後自行遞迴求值，
  絕不走 eval/exec；新增白名單函式要同時改 `_ALLOWED_FUNC_NAMES` 與 `_RowCtx.call`。
- 測試：`tests/test_validation.py`（引擎、免 Qt）、`tests/test_validation_ui.py`（offscreen UI）。

### 欄位綁定（field_bindings）
- 每母表 config 下 `field_bindings:{scope:{enabled,driver,groups:{值:[欄位]}}}`（scope ""=母表）；
  引擎在 `field_binding.py`（`manager.binding`，無 Qt）。設定 UI 在驗證視窗第二分頁（勾選矩陣）。
- 語意：沒被任何 group 提到的欄位＝共用（永遠顯示）；driver/FK/PK 永遠顯示；
  資料值沒設定 group → 該列全顯示。**只控制顯示/鎖定，值一律不動**（使用者決策 B）。
- 呈現：母表表單隱藏欄位區塊；子表「聯集隱欄」（目前這筆母資料所有列的相關欄位聯集，
  `SubTablePanel.apply_column_binding`）＋不相關格灰化鎖定（model flags/data）。
- 驗證整合：違規標記跳過不相關格（`_validate_rule_row` 過濾），全不相關則整筆不記。
- 測試：`tests/test_field_binding.py`、`tests/test_field_binding_ui.py`。

---

## UI 結構
- `App(QMainWindow)`：頂列（🔍搜尋 / 🩺資料健檢 / ⚙配置 / 🕓最近）＋ 每個母表一個 `QTabWidget` 分頁。
- `TableEditor`：左＝分類清單（依 `classification_key`）、中＝項目清單、右＝欄位編輯器（上）＋子表 `QTabWidget`（下，含「跳到子表」下拉）。
- 子表 model/view：`SubTableModel` / `QTableView`（`headerData` 回傳 `note` 當 ToolTip）。

### 樣式與按鈕慣例
- 全域樣式 `APP_QSS`，配色字典 `_C`。按鈕用 `_mk_btn(text, role, icon=...)`；`role` ∈ `primary/success/danger/ghost`。
- 圖示用 `_ti_icon(name, color)`：把 Tabler 字型字符渲染成 `QIcon`（標籤仍用一般 CJK 字型）。新圖示要先在 `_TI_HEX` 補 codepoint。
- 慣例配色：**綠＝新增、紅＝刪除、藍＝編輯/主要**；上下移/複製等次要動作用純圖示＋tooltip。
- 結構操作收進 `欄位 ▾`、`子表 ▾` 下拉，避免工具列無限變長。

---

## 開發注意
- 重建子表分頁（`_build_sub_tabs`）後會跳回第一個分頁；增/改欄位後要 `_select_sub_tab(tab_name)` 留在原分頁。
- 改 config 結構後呼叫 `manager.save_config()`；改完套用呼叫 `editor.reload_after_config()`。
- 載入/儲存走背景 `QThread`（`_LoadWorker`），完成後回主線程更新 UI。
- 查找一律用 dict/set 做 O(1)，避免迴圈線性掃描（見 `_search_index`、`_ref_dict`）。

---

## 打包與部署
- 雙擊 `JsonEditor.bat`：Nuitka `--standalone` → `nuitka_dist\main.dist\` → robocopy 部署到 `C:\Users\USER\Desktop\JsonEditor\JsonEditor`。
- **務必**帶上資料檔：`--include-data-files=assets/tabler-icons.ttf=assets/tabler-icons.ttf`（沒帶圖示會消失）。
- 專案路徑含中文「工具」，**必須**加 `--experimental=force-dependencies-pefile`，否則 Nuitka 依賴偵測會因編碼崩潰。
- 部署用 `robocopy /MIR /XF config.json`（config 是使用者設定，不覆蓋）；exit code < 8 皆為成功。

---

## 測試要求
- 新功能用生產級資料量（14+ sheets、1000+ rows）測試，確認無 UI 凍結。
- 切分類 → 選項目 → 子表行顯示正確、enum 下拉、bool checkbox 可用。
- 連結（text_ref）修改文字 → key 正確回寫。
- 存檔 → 重新載入 → 資料無損。
- 視窗縮放 → 母表/子表高度自適應。
