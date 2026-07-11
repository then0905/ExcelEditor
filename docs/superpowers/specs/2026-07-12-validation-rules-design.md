# 資料驗證規則（Data Validation）設計文件

日期：2026-07-12
狀態：使用者已核准

## 目標

在 JsonEditor 加入類似 Excel 資料驗證的功能：使用者可自由定義條件式驗證規則
（例：Operation 子表的 SkillComponent = ContinueBuff 時 EffectDurationTime 必須有值），
違規儲存格即時上使用者指定的顏色，儲存前列出未通過項目並依嚴重度攔截。

## 已確認需求

1. **規則形式**：結構化規則編輯器（下拉組條件）為主，另有「進階」表達式模式。
2. **範圍**：同列欄位互驗＋母表↔子表互看（子表規則可參照母表欄位；母表規則可下子表聚合條件）。
3. **存檔行為**：每條規則可設 `error`（擋存檔）或 `warn`（僅提醒，可仍要儲存）。
4. **時機**：即時驗證——載入後全表驗一次，之後每次編輯增量重驗受影響的列，違規格立即上色。

## 方案選擇

採**方案 A**：新增獨立 `validation.py` 模組（規則模型＋引擎＋自製 `ast` 白名單安全表達式求值器），
不引入第三方套件（Nuitka 打包零風險）。否決 pandas query/eval（跨表做不到、字串語法彆扭）
與第三方驗證庫（依賴打包風險、規則模型對不上兩層母子表＋上色需求）。

## 1. 規則資料結構（存 config.json）

規則存於每個 JSON 檔、每張母表設定下的 `validations` 陣列：

```json
"SkillData": {
  "validations": [
    {
      "id": "v1", "name": "ContinueBuff 必填持續時間",
      "enabled": true,
      "severity": "error",
      "color": "#E5484D",
      "scope": "Operation",
      "mode": "builder",
      "when":  { "logic": "and", "conds": [
                  {"field": "SkillComponent", "op": "eq", "value": "ContinueBuff"} ] },
      "then":  [ {"field": "EffectDurationTime", "op": "not_empty"} ],
      "expr": "",
      "mark":  []
    }
  ]
}
```

- `severity`：`error`=擋存檔、`warn`=只提醒。
- `color`：違規格底色，每條規則自選。
- `scope`：`""`=母表本身、`"Operation"`=子表名。
- `mode`：`builder`（結構化）或 `expr`（表達式）。
- `mark`：要上色的欄位；空 = 自動用 `then` 的欄位（expr 模式空則標整列第一欄）。
- 運算子集合：`eq / ne / empty / not_empty / contains / not_contains / in_list / not_in /
  gt / ge / lt / le / between / regex`。
- **母表↔子表互看**：子表規則的欄位可選 `母表.欄位`（存成 `master.欄位名`）；
  母表規則可下聚合條件：`{"agg": {"sub": "Operation", "field": "...", "op": "...", "value": "...",
  "count_op": "ge", "count": 1}}`（「子表符合〈欄位 op 值〉的列數 count_op count」；field 留空=數全部列）。
- 語意：**when 全成立而 then 任一不成立 → 違規**；when 留空 = 每列都檢查 then（必填類規則）。

## 2. 規則編輯器 UI

- 頂列工具列加「✓ 驗證規則」按鈕（與 🩺 健檢並排，作用於當前分頁母表）。
- 視窗左側：規則清單（啟用勾選、名稱、嚴重度標籤、顏色色塊）＋「新增/複製/刪除」。
- 視窗右側：名稱、scope 下拉（本表/各子表）、嚴重度、顏色選擇器（QColorDialog）、
  「條件（當…）」與「要求（則…）」兩區，各為可增刪的「欄位｜運算子｜值」條件列，條件區含 AND/OR 切換。
- 「一般／進階」切換：進階=表達式輸入框（回傳 True=通過）＋語法說明
  （本列欄位直接用欄名、`master.欄位`、`subs("子表名")`、`empty()`、`len()` 等白名單函式）
  ＋「▶ 測試」按鈕：立即對現有資料試跑並顯示違規列數。

## 3. 驗證引擎（validation.py）

- `ValidationEngine(manager)`：讀 config `validations` 編譯成 predicate；
  維護 `violations = {(表全名, row_idx, 欄位) → [rule, …]}`。
- builder 規則直接組 Python 函式；expr 規則走 `ast` 白名單求值器
  （只允許比較/布林/算術/常數/欄位名/白名單函式呼叫，杜絕任意程式碼執行）。
- 值型別依 config 欄位 `type` 自動轉換；int/float 空白視為 None；`empty()` 對 None 與 "" 都成立。
- 跨表解析用現有 FK 機制建 dict 索引（O(1)）：子表列 → FK 對應母表列；母表列 → 名下各子表列集合。

## 4. 即時驗證與上色

- 載入後於背景執行緒全表驗證一次建立違規映射。
- `update_cell` 後只重驗受影響的列：改子表格 → 該列＋其母表列；改母表格 → 該列＋其名下子表列。
- 子表：`SubTableModel.data()` `BackgroundRole` 查違規映射，命中回傳規則色（優先於 dirty 黃）。
- 母表：右側欄位編輯器違規欄位輸入框加規則色邊框＋淡底；中間項目清單違規項加色點。
- 子表分頁標題含違規列時加「⚠ N」計數。

## 5. 儲存前檢查

`save_file()` 丟給背景存檔前先全量驗證：

- 有 error 級違規 → 「驗證未通過」對話框按規則分組列出（表/主鍵/欄位/規則名），
  點擊條目跳到該儲存格；只能取消，不能存。
- 只有 warn 級 → 同樣列出但多「仍要儲存」按鈕。
- 全過 → 直接存。

## 6. 測試與部署

- `QT_QPA_PLATFORM=offscreen` 起 QApplication 寫自動化測試：builder 規則、expr 規則、
  母↔子互看、增量重驗、存檔攔截。
- 以生產資料量（14+ 表、1000+ 列）實測載入驗證耗時與編輯流暢度。
- 完成後以 `JsonEditor.bat` 重新打包部署（部署版 config.json 不被覆蓋，規則設定保留）。
