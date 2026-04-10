# 智慧排班系統 2.0 — 臺南市佳里區衛生所

## 專案基本資訊
- **系統名稱**：智慧排班系統 2.0
- **部署平台**：Google Apps Script (GAS) + Google Sheets
- **試算表 ID 結尾**：`OUQ`（完整 ID 見 Code.gs 的 `SHEET_ID` 常數）
- **設定工作表名稱**：`班表設定`（`EMAIL_SHEET_NAME`）
- **製表人**：鄭兆鑫（硬編碼於列印頁尾，勿改）
- **GitHub**：https://github.com/za869765/CLASS

---

## 目前版本
| 檔案 | 版本 |
|------|------|
| 前端 | `index.html` |
| 後端 | `Code.gs` |
| 內部版號 | ver4.4+ |

---

## 已修正 BUG 清單（共 25 項，三輪測試）

### 第一輪（BUG1~16）
| # | 類別 | 摘要 |
|---|------|------|
| 1 | UI | `fpActionRow` 雙 display 屬性 → 按鈕永遠可見 |
| 2 | UI | `auditEmpOverlay` 重複 display:none → flex 置中失效 |
| 3 | UI | `rocNumCh()` 缺「一百」前綴 → 工作表名稱全錯 |
| 4 | UI | `pubStatsSel` 死代碼移除 |
| 6 | 換班 | `logShiftChange` 不補星期 → regex 解析失敗 |
| 8 | 換班 | `updateShift` 換班無身份驗證 |
| 9 | 排班 | `quickSchedule` scope 死參數（整年排班已移除）|
| 10 | 排班 | `trulyNewStaff` reduce 運算子優先序 `||` 低於 `+` |
| 11 | 排班 | `isStaffActiveForMonth` 只比月份不含年份 → 跨年誤判 |
| 14 | 衝突 | `assignSlotsWithPointer` 單輪掃描 → 雙重衝突未解 |
| 15 | 衝突 | `buildOriginalScheduleMap` + 內部 `buildOriginalMap` 不解析拖曳日誌 |
| 16 | 衝突 | `simulateFullYearValidation` chkConsec 空值不重置 pv |

### 第二輪（BUG17~21）
| # | 類別 | 摘要 |
|---|------|------|
| 17 | UI 迴歸 | `'一百'+rocNumCh()` 雙重前綴「一百一百…」|
| 18 | 換班 | 管理員判斷用密碼比 empId → 改用 adminPw 參數 |
| 19 | UI | 預覽 Modal 關閉未清理全域變數 → 拖曳資料殘留 |
| 20 | 架構 | `parseYearMonthFromSheetName` 硬編碼 114/115 → 2027 失效 |
| 21 | 架構 | `parseDateFromSheet` 硬編碼 year=2026 |

### 第三輪（BUG22~25）
| # | 類別 | 摘要 |
|---|------|------|
| 22 | 換班 | 前端 `updateShift` 未傳 adminPwd → 管理員無法幫他人換班 |
| 23 | UI | `closeFpModal` 未清理 _fpArrangerEmpId/_fpWriteLetter |
| 24 | 架構 | 3 處殘留硬編碼「一百一十五」|
| 25 | 架構 | `HOLIDAYS_2026` 改為 `GOV_HOLIDAYS` 多年度 Map |

---

## 架構升級摘要
- **整年排班功能已移除**（人員變動大，只用單月排班）
- **年份全面動態化**：`parseYearMonthFromSheetName` 通用解析 + `rocStrToNum` 反向函式
- **假日多年度支援**：`GOV_HOLIDAYS` Map 結構，每年初需補充放假日資料
- **換班身份驗證**：`updateShift` 加入 `adminPw` 可選參數，管理員可幫他人換班

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

---

## 關鍵函式對照

### GAS 後端
| 函式 | 用途 |
|------|------|
| `getScheduleData(sheetName)` | 取得班表資料（含 changes/holidayRows/writeCount）|
| `quickSchedule(pw,mode,scope,month,year)` | 一鍵排班（僅 scope='month'）|
| `runAutoSchedule(sheet,pw,opts)` | 核心排班演算法 |
| `autoValidateSchedule(sheetName)` | 系統自動驗算（4項檢查）|
| `buildOriginalScheduleMap(prevSheetName)` | 從日誌重建原始排班人 map |
| `parseYearMonthFromSheetName(name)` | 通用民國漢字年月解析 |
| `rocStrToNum(str)` | 民國漢字→數字（一百一十五→115）|
| `rocNumToStr(n)` | 數字→民國漢字（115→一百一十五）|
| `isStaffActiveForMonth(obj,year,month)` | 判斷人員是否在職（完整年月日比較）|
| `GOV_HOLIDAYS` | 多年度政府假日 Map（每年初需更新）|

### 前端
| 函式 | 用途 |
|------|------|
| `showFullPreview(r,title,sub,cb,letter)` | 開啟全頁預覽 Modal |
| `closeFpModal()` | 關閉預覽並清理全部全域變數 |
| `rocNumCh(n)` | 民國漢字（含一百前綴，與 GAS rocNumToStr 同邏輯）|
| `openQkModal(mode)` | 開啟單月排班 Modal（已移除整年模式）|
| `confirmShift()` | 換班確認（傳入 adminPwd 支援管理員代換）|

---

## 開發規範

1. **彈窗一律用 CSS Overlay**，禁止使用 `window.prompt` / `window.alert` / `window.confirm`
2. **每次修正給完整前後端代碼**（不給 diff）
3. 日誌去重在 GAS `writeDragShiftLog` 執行
4. 排班者異動：`isAudit: false`；審核者異動：`isAudit: true`
5. 試算表 `A2:A32` 強制設為純文字格式（`@`）
6. 班表工作表名稱格式：`一百一十五年四月班表`（民國漢字年月）
7. **年份不可硬編碼**：一律用 `new Date().getFullYear()` 或 `parseYearMonthFromSheetName` 動態取得
8. **GOV_HOLIDAYS 每年初需更新**：新增下一年度放假日 Set

---

## 待辦 / 年度維護
- [ ] 每年 12 月前：補充下一年度 `GOV_HOLIDAYS[20XX]` 放假日資料
- [ ] 確認 `rocNumCh`（前端）與 `rocNumToStr`（後端）邏輯一致
