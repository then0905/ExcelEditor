# 欄位綁定（Conditional Field Visibility）設計文件

日期：2026-07-12
狀態：使用者已核准（子表=聯集隱欄＋灰化；值處理=只隱藏、保留舊值不重置）

## 目標

依「驅動欄位」的值決定每列哪些欄位相關——例：Operation 子表選 `SkillComponent = Damage`
只顯示傷害相關欄位，選 `PassiveBuff` 才顯示狀態/加成/持續時間欄位。不相關欄位自動
隱藏/灰化，避免每筆都面對 15 個欄位。

## 已確認決策

1. **子表格呈現＝聯集隱欄＋灰化格子**：看目前選中母資料的所有子表列用到哪些 driver 值，
   取相關欄位聯集；完全沒列用到的欄位整欄 `setColumnHidden`。留下的欄位中，對某列
   不相關的格子灰化、不可編輯（每列各藏各的在格狀表格做不到）。
2. **值處理＝B（只隱藏，保留舊值）**：換 driver 值不重置任何欄位；舊值留在資料裡照常存檔。
   不做自動填預設值（新列本來就是空值，存檔時型別強制轉換已給 int 0/float 0.0 等預設）。
3. **設定 UI 放 ✅ 驗證規則視窗**：視窗改為兩分頁「驗證規則｜欄位綁定」。

## 1. 綁定資料結構（config.json，每母表一個 key）

```json
"field_bindings": {
  "Operation": {                     //  "" = 母表本身，其他 = 子表名
    "enabled": true,
    "driver": "SkillComponent",
    "groups": {
      "Damage":      ["EffectValue", "TargetCount", "Bonus"],
      "PassiveBuff": ["InfluenceStatus", "BonusType", "EffectDurationTime"]
    }
  }
}
```

- **永遠相關（不受綁定影響）**：driver 欄本身、FK 欄（子表）、主鍵欄（母表）、
  以及**沒被任何 group 提到的欄位**（共用欄位不用逐值列舉）。
- driver 值在 `groups` 沒有對應（新元件還沒設定）→ 該列**全欄位相關**，不誤藏。
- 每表最多一條綁定（一個 driver）。

## 2. 行為

- **母表欄位編輯器**（表單）：不相關欄位的整個區塊（標題＋輸入框＋分隔線）隱藏；
  改 driver 值立即重排。
- **子表格**：
  - 聯集隱欄——只針對「目前顯示的列」（= 選中母資料的列）計算。
  - 灰化格子——保留欄位中對該列不相關的格子：底色更暗、文字 txt3、不可編輯；
    若留有舊值照常顯示（灰字），tooltip 註明「與 <driver值> 無關（保留舊值）」。
- **驗證整合**：驗證規則照常運算，但違規標記（上色）**跳過不相關的儲存格**；
  若一條違規的所有標記欄位都不相關，該筆違規不記（不強迫使用者填看不到的欄位）。
- 存檔輸出完全不變（B 決策：值不動）。

## 3. 綁定編輯器 UI（驗證規則視窗第二分頁）

- 上：表選擇（母表/各子表）＋「啟用」勾選＋驅動欄位下拉（該表欄位，預設挑 enum 欄）。
- 中：**勾選矩陣**（QTableWidget）——列＝欄位（排除 driver/FK/PK），欄＝driver 值；
  勾選＝該值時此欄位相關。整列全空＝共用欄位（永遠顯示），列尾以「共用」標示。
- driver 值來源：該欄 config options（enum）∪ 資料中出現過的值；「＋新增值」可手動補。
- 套用：寫回 config、save_config、立即刷新編輯器。

## 4. 實作落點

- **`field_binding.py`（新模組，無 Qt）**：`FieldBindingEngine(manager)`——
  `binding_for(master, scope)`、`relevant_fields(master, scope, row_idx)`（None=全相關）、
  `is_relevant(sheet, row_idx, col)`、`visible_columns(master, scope, row_idxs)`（聯集）。
  即時讀 config，不做快取（欄位數小）。掛在 `manager.binding`。
- `validation.py`：`_validate_rule_row` 標記前先用 `manager.binding.is_relevant` 過濾。
- `SubTableModel`：`flags()` 對不相關格拿掉 editable/checkable；`data()` 灰化＋tooltip。
- `SubTablePanel.reload`：依 `visible_columns` `setColumnHidden`。
- `FieldEditorWidget`：記錄每欄位的區塊/分隔線 widget，`load_row` 與 driver 值變更時
  show/hide；由既有 `_refresh_validation_visuals` 順路觸發。
- `ValidationRulesDialog`：外層改 QTabWidget，第二分頁為綁定編輯器。

## 5. 測試

- 引擎單測（無 Qt）：relevant_fields、共用欄位、未綁定值全相關、聯集、驗證 mark 過濾。
- offscreen UI：model flags/灰化、reload 後隱欄正確、母表欄位隱藏、矩陣編輯 roundtrip。
- windows 平台截圖驗排版；正式 SkillData 實測；打包部署。
