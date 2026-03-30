# 智慧排班系統 2.0 — 臺南市佳里區衛生所

## 專案基本資訊
- **系統名稱**：智慧排班系統 2.0
- **部署平台**：Google Apps Script (GAS) + Google Sheets
- **試算表 ID 結尾**：`OUQ`（完整 ID 見 Code.gs 的 `SHEET_ID` 常數）
- **設定工作表名稱**：`班表設定`（`EMAIL_SHEET_NAME`）
- **製表人**：鄭兆鑫（硬編碼於列印頁尾，勿改）

---

## 目前版本
| 檔案 | 版本 |
|------|------|
| 前端 | `index_v4_3__30_.html` |
| 後端 | `Code_v4_3__15_.gs` |
| 內部版號 | ver3.23+ |

---

## 試算表欄位架構（班表設定工作表）

### 人員名單
| 欄 | 用途 |
|----|------|
| B1:B11 | 停班2線輪序（登革熱二線）|
| E1:E11 | 值班輪序 |
| H1:H11 | 員工編號（備用） |
| I1:I11 | 全體人員姓名 |
| J1:J3  | 支援候選人（協助掛號）|
| K1:K8  | 門診系列候選護理師 |
| L1:L11 | Email |
| M1:M11 | 員工編號（主要，驗證用）|
| P1:P11 | 點數 |
| Q1:Q2  | 卡介苗候選 |

### 排班規則設定
| 欄 | 用途 |
|----|------|
| N2     | 管理員密碼 |
| N3:N13 | 各職務排班星期規則（colIdx\|day,day 格式）|
| O1:O12 | 可用工作表清單（空白=自動取當年）|
| T:W    | 排除規則（姓名/起日/迄日/星期）|
| X1:X11 | 到職日 M/D |
| Y1:Y11 | 離職日 M/D |
| Z1:Z11 | 是否接收通知 TRUE/FALSE |

### Line Bot 設定
| 欄 | 用途 |
|----|------|
| R1  | Line Channel Access Token |
| R2  | 搜尋關鍵字前綴 |
| R3  | 模糊搜尋 TRUE/FALSE |
| R4:R15 | 指定搜尋工作表 |
| R16 | Line Channel ID |
| R17 | Line Channel Secret |
| R18 | Gemini API Key |
| S1:S13 | 搜尋欄位代號 |

### 日誌區域
| 欄 | 用途 |
|----|------|
| A40:A460 | 換班 / 拖曳 / 排班日誌 |
| M40:M460 | 操作紀錄（writeOpLog）|

---

## 班表工作表欄位對照（C1:M1）

| colIdx | 欄 | 職務名稱 | 排班日 |
|--------|-----|---------|-------|
| 0 | C | 值班 | 每日（含假日），平日/假日分開輪序 |
| 1 | D | 支援 | 週四工作日 |
| 2 | E | 門診 | 週二＋週四（混合輪序）|
| 3 | F | 掛號 | 週四 |
| 4 | G | 前台 | 週四 |
| 5 | H | 預登1 | 週二＋週四（混合）|
| 6 | I | 預登2注 | 週二＋週四（混合）|
| 7 | J | 注射1 | 週二＋週四（混合）|
| 8 | K | 注射2 | 週二 |
| 9 | L | 卡介苗 | 每月第一個週二工作日 |
| 10 | M | 停班2線 | 每日（平日/假日分開輪序）|

**混合日欄（TUE_THU_CIS）**：ci = 2, 5, 6, 7（週二/週四分開月計）

---

## N1 備註格式（班表工作表）

```
排定時間: yyyy/MM/dd HH:mm
審核狀態: pending | approved
writeCount: N
swapCount: 姓名1=N,姓名2=N,...
```

---

## 日誌格式

### 一般換班（updateShift）
```
M/D 週X 班別 原A→新B (員工編號) 更換時間: ... 備註:...
```

### 排班者拖曳異動
```
排班-姓名 M/D 班別 原A→新B 更換時間: yyyy/MM/dd HH:mm 備註:...
```

### 審核者拖曳異動
```
審核-姓名 M/D 班別 原A→新B 更換時間: yyyy/MM/dd HH:mm 備註:...
```

### 去重 key（忽略時間和備註）
```
排班-姓名 M/D 班別 原A→新B
審核-姓名 M/D 班別 原A→新B
```

---

## 前端全域變數

```javascript
var _fpArrangerEmpId = '';    // 排班者/審核者員編（全程沿用）
var _fpAuditMode = false;     // 審核者模式（允許拖曳全部欄位）
var _fpDragChanges = [];      // 拖曳異動紀錄
var _fpPreviewData = null;    // 當前預覽資料
var _fpWriteCallback = null;  // 確認寫入回呼
var _fpCellMap = {};          // 每格獨立追蹤 key='rowIdx:ci'
var adminPwd = '';            // 管理員密碼（登入後儲存）
var selSheet = '';            // 目前選擇的班表
```

---

## 關鍵函式對照

### GAS 後端
| 函式 | 用途 |
|------|------|
| `getScheduleData(sheetName)` | 取得班表資料（含 changes/holidayRows/writeCount）|
| `quickSchedule(pw,mode,scope,month,year)` | 一鍵排班（preview/execute）|
| `runAutoSchedule(sheet,pw,opts)` | 核心排班演算法 |
| `writeDraggedPreview(sheet,pw,rows)` | 寫入拖曳後的預覽資料 |
| `writeDragShiftLog(sheet,empId,logs)` | 寫入排班/審核異動日誌（含去重）|
| `auditSchedule(empId,sheet,action,note,items)` | 核准或退回班表 |
| `autoValidateSchedule(sheetName)` | 系統自動驗算（4項檢查）|
| `getAuditHints(sheetName)` | 取得上月末位/本月首位值班資訊 |
| `getPendingSheets()` | 取得所有審核中班表（含版次/時間）|
| `getScheduleChanges(sheet,headers)` | 解析日誌回傳 changes 物件 |
| `wcToLetter(wc)` | 版次字母（A~Z→AA~AZ→...）|

### 前端
| 函式 | 用途 |
|------|------|
| `showFullPreview(r,title,sub,cb,letter)` | 開啟全頁預覽 Modal |
| `previewPendingSheet(sheetName)` | 審核者預覽（先輸入員編）|
| `openAuditModal(sheetName)` | 開啟驗算審核視窗 |
| `saveFpDragLog(sheet,eid,changes,isAudit)` | 前端呼叫 writeDragShiftLog |
| `wcToLetter(wc)` | 版次字母（與 GAS 同邏輯）|
| `renderFpCell(td,name,isDragged,isAuditCell)` | 更新格子顯示（含顏色）|
| `updateFpDragPanel()` | 更新「排班已異動」面板 |
| `buildFpStats()` | 建立排班統計面板 |
| `_renderPendingList(r,list)` | 渲染待審核班表清單 |
| `refreshPendingReview(btn)` | 即時刷新待審核清單 |

---

## 排班/審核流程

### 排班者流程
```
管理員後台 → 一鍵排班（單月/整年）
  → 輸入員編
  → 預覽班表（全頁 Modal，可拖曳 ci=2~9）
  → 確認寫入 → 儲存至試算表
  → 驗算審核視窗自動開啟
  → ✅ 檢視無誤 / ❌ 排班錯誤
```

### 審核者流程
```
管理員後台 → 建立班表 → 待審核班表區
  → 👁 預覽班表 →
  → 🔐 CSS Overlay 輸入員編
  → 全頁預覽（可拖曳全部欄位，_fpAuditMode=true）
  → 確認寫入 → 自動關閉預覽
  → 200ms 後自動開啟驗算審核視窗
  → ✅ 核准 → approved / ❌ 退回刪除 → 發送 email
```

---

## 異動顏色規則

| 角色 | 底色 | 說明 |
|------|------|------|
| 排班者異動 | `#fef9c3`（黃） | inline style |
| 審核者異動 | `#dbeafe`（藍） | inline style |
| Preloaded（唯讀展示）| 黃底虛線框 | isPreloaded: true |

---

## 版次字母規則

```javascript
// A=1, B=2...Z=26, AA=27, AB=28...AZ=52, BA=53...ZZ=702, AAA=703
function wcToLetter(wc) {
  if (!wc || wc <= 0) return '';
  let n = wc, result = '';
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}
```
**注意**：GAS 和 HTML 各有一份，邏輯完全相同。

---

## 開發規範

1. **彈窗一律用 CSS Overlay**，禁止使用 `window.prompt` / `window.alert` / `window.confirm`
2. **每次修正給完整前後端代碼**（不給 diff）
3. 日誌去重在 GAS `writeDragShiftLog` 執行，前端 `_fpDragChanges` 過濾 `isPreloaded`
4. 排班者異動：`isAudit: false`；審核者異動：`isAudit: true`
5. `hasDrag` 判斷只計 non-preloaded 的新異動
6. 試算表 `A2:A32` 強制設為純文字格式（`@`），避免日期被轉型
7. 班表工作表名稱格式：`一百一十五年四月班表`（民國漢字年月）

---

## 待辦 / 已知問題

- [ ] 確認審核者異動底色（藍色）在各瀏覽器正確顯示
- [ ] `isAuditLog` 旗標從 GAS `getScheduleChanges` 正確傳遞到前端 badge
- [ ] 2027 年（116年）班表年份支援（`parseYearMonthFromSheetName` 已支援 roc114/115）
