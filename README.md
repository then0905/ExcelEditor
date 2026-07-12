# JsonEditor（遊戲資料 JSON 編輯器）

![示意圖](https://github.com/user-attachments/assets/8ef0ca77-3b8e-4ff3-8693-67ab0120cafa)

> 目前版本直接編輯 **JSON** 資料檔（母表＋子表兩層結構），早期的 Excel 版本已淘汰；
> 下方使用說明中與 Excel 相關的段落為舊版文件，待更新。

## 快速開始（使用原始碼）

```bash
git clone https://github.com/then0905/JsonEditor.git
cd JsonEditor

# 建立虛擬環境並安裝依賴（Python 3.11+）
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt

# 執行
.venv\Scripts\python main.py
```

- 開啟後用「開檔」載入任一遊戲資料 `.json`（頂層為物件陣列，或 {表名: 陣列}）。
- 每個 JSON 檔的欄位型別、驗證規則、欄位綁定等設定存於自動建立的 `config.json`（不入版控）。

### 跑測試

```bash
set QT_QPA_PLATFORM=offscreen
.venv\Scripts\python tests\test_validation.py
.venv\Scripts\python tests\test_field_binding.py
# 其餘 tests/*.py 同理
```

### 打包成執行檔

雙擊 `JsonEditor.bat`（Nuitka standalone，含部署步驟，詳見 `打包與發布.md`）。
路徑含非 ASCII 字元時必須保留 bat 內的 `--experimental=force-dependencies-pefile` 參數。

### 主要功能

- 母表／子表兩層編輯、enum 下拉、bool、陣列 chips 編輯（含建議值快速選填）
- 資料驗證規則（類 Excel 資料驗證：條件式規則＋違規上色＋存檔攔截）
- 欄位綁定（依驅動欄位值條件式顯示欄位）
- 範本列（從範本深拷貝母列＋子表列）、資料健檢、跨檔文字參照、全域搜尋、比較視窗、筆記、多開

開發規範與資料模型細節見 `CLAUDE.md`。

---

##  前言

![前言](https://github.com/user-attachments/assets/352bc0aa-dcd0-4f64-9373-87d89ad35eb1)

撰寫的靈感源自於某天打開遊戲的 DB (用 Excel 製作),密密麻麻的資料讓人眼花撩亂,想到日後需要打開來調整數值,即使使用搜尋功能應該也會看得很煩躁。

於是我決定嘗試開發一個編輯器來修改 Excel 的內容。

---

##  使用說明

### 表格的格式

![表格的格式](https://github.com/user-attachments/assets/2813b301-89b3-4082-a911-0a9fa8669191)

- **標題欄請放在第一行**
- **工作表的命名必須以 `.json` 結尾**才會被編輯器讀取
- 其他用於筆記的表格,正常命名即可,不要以 `.json` 結尾

### Excel 工作表的命名方式

![說明圖](https://github.com/user-attachments/assets/e4d1ff46-9540-44f6-839d-e28ecbc708c8)

資料採用**主從結構**:

- 以示意圖為例,`skill.json` 是主表,以 `SkillID` 為主鍵
- 相關的一對多資料會拆分成具名子表
- **子表命名規則**為 `skill.json#XXX`,其中 `XXX` 可自行命名,但格式必須遵守此規則
- **子表第一欄必須是 `SkillID` 作為 foreign key**

### 編輯器的配置設定

![編輯器配置](https://github.com/user-attachments/assets/7c0d8a0e-ca21-46e2-962b-484949a5a5a8)

首次匯入未記錄過的 Excel 時會彈出配置提示視窗。

**主畫面右上角的「配置設定」按鈕也可開啟此視窗**

#### 外部文字表

- 通常不需要設定(此功能是針對遊戲中使用多語系 Excel 的需求)
- ![文字表](https://github.com/user-attachments/assets/97fb48c1-bb1d-4e2f-9b7e-5221b85e34ce)
- 如果需要此設定,請將文字表格式設為:第一列為文字 ID,第二列為文字內容(標題列命名不限)
- 表內資料需與外部文字表的文字 ID 對應才能讀取到文字內容
- 勾選後,修改並儲存 Excel 時會同步更新外部文字表的資料

#### 圖片設定

- 啟用後,圖片路徑請與 exe 放在同一層
- 圖片路徑可以是:
  - `{你指定的資料夾}/{分類ID}/{清單選取的內容}.png`
  - `{你指定的資料夾}/{清單選取的內容}.png`

#### 母表分類設定

母表是否需要分類,取決於資料的使用與管理需求:

- **需要分類的情況**  
  例如「技能資料」:
  - 技能本身是一筆筆獨立的資料
  - 但實務上會依「職業」進行管理
  - 因此母表會先以「職業」作為分類,再顯示該職業底下的技能清單

- **不需要分類的情況**  
  例如「怪物資料」:
  - 不需要額外的分類層級
  - 每筆資料僅以 MonsterID 作為唯一識別
  - 因此母表採用單層結構,不再進一步分類

#### 欄位的資料類別

母表與子表的設定大致相同,欄位可選類別有:

- **String**: 單純字串
- **int**: 整數
- **float**: 小數
- **bool**: 一個 CheckBox,打勾即為 True
- **enum**: 下拉式選單,選擇後需手動輸入選單內容,並以逗號 `,` 分隔

---

### 多個主表與子表的 Excel

![多個主表與子表的Excel](https://github.com/user-attachments/assets/c7f92e33-e07c-470f-81ad-0738d0735c8d)

- 當 Excel 包含多個主表與子表時,配置設定視窗和主畫面都會提供 Tab 進行切換(如示意圖)
- **請注意確認已切換到正確的 Tab**

---

### 資料的新增與刪除

- 除了編輯功能外,分類、子表與清單都提供刪除與新增按鈕
- 讓您能更靈活地操作資料以符合需求

---

##  結語

希望這個編輯器能幫助到有需要的人 (例如:我)

---
