# 智慧排班系統 2.0 — 開發歷史紀錄

## 目前最新檔案
- 前端：`index_v4_3__30_.html`
- 後端：`Code_v4_3__15_.gs`

---

## 已完成功能清單

### 核心排班演算法
- [x] 單月自動排班（`runAutoSchedule`）
- [x] 整年一鍵排班（`autoScheduleFullYear`）
- [x] 排序鍵：① 此欄此日=0優先 ② 月總次數少優先 ③ 跨月歷史 ④ 輪序
- [x] 門診系列 max-min ≤ 1（per slot hard cap = 2）
- [x] 混合日欄（ci=2,5,6,7）依 dow 使用日別月計（週二/週四獨立）
- [x] 停班2線：假日/平日分組，衝突時 swap，swapCount 追蹤
- [x] 卡介苗奇偶月交替輪替
- [x] 跨月指針接續（平日/假日各自獨立）
- [x] 排除規則（請假、特定星期、到/離職日過濾）

### 版次字母系統
- [x] `wcToLetter(n)`：A~Z → AA~AZ → BA~BZ → ... → ZZ → AAA（無上限）
- [x] GAS 和 HTML 各一份，邏輯相同
- [x] 所有舊 `String.fromCharCode(64 + Math.min(wc, 26))` 已全數替換

### 全頁預覽系統
- [x] 全螢幕預覽 Modal（`showFullPreview`）
- [x] 縮放（＋/－）
- [x] 列印（開新視窗，含浮水印）
- [x] 統計面板（值班/支援/門診各別週二/週四分欄）

### 拖曳/替換功能
- [x] 排班者：僅 ci=2~9（門診類）可拖曳/點擊替換
- [x] 審核者（`_fpAuditMode=true`）：全部欄位可操作
- [x] 空格顯示 `+`（新增），有人名格顯示 `✎`（替換）
- [x] `fpSlotActive()` 依星期判斷有效格
- [x] 每格獨立追蹤（`_fpCellMap`），多次拖曳不斷鏈
- [x] 拖曳後即時更新統計面板

### 異動紀錄面板（排班已異動）
- [x] 標題「✏️ 排班已異動（N筆）」只計新增異動
- [x] 色塊圖例：🟡排班者 / 🔵審核者
- [x] Preloaded items（排班者已有異動）：淡黃底虛線框，唯讀
- [x] 新增異動：白底實線框，可輸入理由，可刪除
- [x] `isAuditLog` 旗標從日誌解析，驅動 badge 顏色（🟡/🔵）

### 日誌系統（去重）
- [x] GAS `writeDragShiftLog`：`排班-` 或 `審核-` 前綴由 `isAudit` 決定
- [x] 去重 Set：key = `(排班|審核)-姓名 日期 欄位 原A→新B`（忽略時間/備註）
- [x] 同批次寫入也去重（existingKeys 在同次 forEach 中累積）
- [x] GAS `getScheduleChanges`：解析 `排班-` 和 `審核-` 兩種前綴，帶 `isAuditLog` 旗標
- [x] `hasDrag` 只計 non-preloaded 的新異動，避免審核者無修改時重複寫入

### 審核流程（單一入口）
- [x] 待審核班表區：鈴鐺動畫橫幅
- [x] 「🔄 更新待審核班表」按鈕移至 pendingReviewBox **外部**（永遠可見）
- [x] 版次標籤（藍底白字）+ 時間標籤（🕐 X分鐘前排定）
- [x] 只剩「👁 預覽班表 →」一個按鈕（移除獨立「驗算審核」按鈕）
- [x] 審核者員編用 **CSS Overlay**（非 browser prompt），Enter 可確認
- [x] 員編在預覽開始前取得，存入 `_fpArrangerEmpId`，後續全程沿用
- [x] 確認寫入成功後：自動關閉全頁預覽 → 200ms 後開啟驗算審核視窗

### 驗算審核視窗
- [x] 員編欄改為 `<input type="hidden">`（從 `_fpArrangerEmpId` 自動帶入）
- [x] `openAuditModal` 開頭自動填入員編（不再清空）
- [x] `showAuditConfirm` 讀取 `auditEmpId.value || _fpArrangerEmpId`
- [x] 4項自動驗算：跨月輪序接續 / 連續同人 / 值班停班2線衝突 / 門診公平度(±2)
- [x] 核准：`approved`，退回：刪除工作表 + 發送 email

### 一鍵排班 Modal（qkModal）
- [x] 無班表：`🆕 新建排班`；已有班表：`🔄 重新排班`（橘色）
- [x] 寫入完成後：`✅ 完成，關閉視窗` + `🔄 重新排班`
- [x] 月份選單載入後自動觸發 `onchange`
- [x] 已有班表時顯示警告橫幅，含「📄 先列印備份此版本」連結
- [x] 「核章列印」按鈕（`qkPrintBtn`）寫入成功後出現

### 列印功能
- [x] `renderForPrint`：先同步開新視窗，再非同步填入（解決 popup blocker）
- [x] A4 直向，flex:1 讓表格撐滿，字體 7pt，浮水印
- [x] 版次字母顯示在副標題
- [x] 頁尾：製表鄭兆鑫 / 人事 / 護理長 / 所長

### 排班統計面板（全頁預覽內）
- [x] 門診相關職務表格：週二/週四分欄（12欄）
- [x] 值班/支援/停班2線獨立表格
- [x] 差異備註：目前平均 + 少者列表

### 點數看板 / 門診均勻看板
- [x] 全年累積次數、每月平均、期望次數、補償值
- [x] 優先順序前三名獎牌標示
- [x] 新進人員標籤（到職月自動推算）
- [x] 個人月份進程表（含補償值趨勢）
- [x] 公開點數看板（不需密碼）

### Line Bot
- [x] 班表查詢（模糊/精確）
- [x] 月份後綴篩選（「王小明 三月」）
- [x] AQI 空氣品質查詢（善化站）
- [x] Gemini AI 問答（`問 XXX`）
- [x] 日期時間查詢
- [x] Flex Message 卡片樣式

### 系統設定
- [x] 接班人設定（一鍵換人，自動同步所有名單）
- [x] 排班排除清單（UI：姓名/起日/迄日/星期 選單）
- [x] 各職務排班星期規則（N3:N13 儲存）
- [x] 通知名單管理（Z欄 TRUE/FALSE）
- [x] 操作紀錄（M40:M460）

---

## 最近 Session 重要修正（本輪 index_v4_3__28_.html → 30_.html）

| 版本 | 修正內容 |
|------|---------|
| v30 | 版次字母改用 wcToLetter()，A→Z→AA→AB... 無上限 |
| v30 | 審核者員編改用 CSS Overlay（不再用 browser prompt）|
| v30 | 待審核清單只剩「👁 預覽班表 →」一個大按鈕 |
| v30 | 「🔄 更新待審核班表」按鈕移至容器外，永遠可見 |
| v30 | 確認寫入後自動關閉預覽，直接進驗算視窗 |
| v30 | 驗算員編改為 hidden input，從 _fpArrangerEmpId 帶入 |
| v30 | 日誌去重支援 `審核-` 前綴 |
| v30 | hasDrag 只計 non-preloaded 異動，避免審核重複寫入 |
| v30 | badge 顏色：排班者🟡 / 審核者🔵（由 isAuditLog 驅動）|
| v30 | _renderPendingList 共用函式，loadPending 和 refresh 共用 |

---

## 已知待驗證項目

- [ ] 審核者異動底色（藍色）在實際操作中是否正確顯示
- [ ] `isAuditLog` 旗標在長日誌序列中的解析正確性
- [ ] 版次超過 Z（writeCount > 26）時 AA 顯示驗證

---

## 檔案版本歷史（主要里程碑）

```
Code_v4_3__1_.gs   → 基礎排班架構
Code_v4_3__10_.gs  → 拖曳日誌系統
Code_v4_3__14_.gs  → 審核流程整合，去重邏輯
Code_v4_3__15_.gs  → 版次字母無上限，審核-前綴，hasDrag 修正

index_v4_3__1_.html  → 基礎前端
index_v4_3__20_.html → 全頁預覽、拖曳系統
index_v4_3__28_.html → 審核單一入口基礎版
index_v4_3__29_.html → CSS Overlay、badge 顏色、刷新按鈕
index_v4_3__30_.html → 版次字母、流程整合完成版
```
