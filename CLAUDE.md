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
| 內部版號 | ver5.3（2026-08-20 ver5.2全功能＋卡介苗輪序接續上月原始站班者） |

---

## 已修正 BUG 清單（共 35 項，六輪測試）

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

### 第四輪（BUG26）
| # | 類別 | 摘要 |
|---|------|------|
| 26 | 排班 | `assignSlotsWithPointer` swap 選 bestJ 時未保護 slot 0 → 跨月接續首位被換走，輪序自檢失敗 |

### 第五輪（BUG34~38，ver4.6）
| # | 類別 | 摘要 |
|---|------|------|
| 34 | 寫入 | 一鍵排班預覽中拖曳挪移後，使用者關閉預覽 Modal 時 `closeFpModal` 清空 `_fpDragChanges`，再點外部 `qkConfirmBtn` 走 `quickSchedule('execute')` 寫入會用原版排班，挪移完全遺失 → **將「確認寫入」按鈕從外部移進預覽 Modal**，使用者不必關閉就能寫入 |
| 35 | UI | 預覽 Modal 右上角紅色 X 移除；下方按鈕重排為 [取消] [確認寫入]；寫入成功後變身為 [✅ 完成關閉] [🖨️ 核章列印]；「完成關閉」直接導向待審核分頁並刷新清單 |
| 36 | UI | 後台分頁重構：「⚡ 一鍵排班」獨立成最左側分頁、「📋 待審核班表」獨立成第二分頁、移除「建立班表」與「自動排班」分頁避免誤按；左側獨立容器移除，改為單欄佈局 |
| 37 | 權限 | 拖曳挪移權限放寬：原本排班者僅能挪門診相關職務，審核者能挪所有；改為**排班者與審核者皆可挪移所有職務**（含值班/支援/停班2線） |
| 38 | 識別 | 停班2線系統自動挪移加識別：(a) 後端 N1 note 加 `dengSwapRows:rIdx=M/d,...` 紀錄，(b) M 欄 cell note 已寫 `swap:M/d`（既有），(c) 前端預覽橙底格 hover 顯示「⚙️ 系統自動挪移」+原日期，(d) 圖例新增「⚙️ 系統挪移（停班2線）」說明，(e) `getPrevLastPtr` 跳過 swap 列，避免下月接續找錯起點 |

### 第六輪（BUG39~42，ver4.7）
| # | 類別 | 摘要 |
|---|------|------|
| 39 | UI | 移除預覽 Modal 右上角列印按鈕（fpPrint）；列印改由寫入後變身的 fpWriteBtn（核章列印）執行 |
| 40 | UI | 寫入成功後彈出 CSS 慶祝視窗（fpWriteSuccessOverlay）+ popBounce 動畫 + 4 秒自動關閉，明確讓使用者意識下一階段 |
| 41 | UI | tab-quick 一鍵排班按鈕改 200×200 正方形 + qkPulseRing 光圈 + qkFingerTap 食指 👆 動畫；tab-pending 查詢按鈕同款設計（文字「查詢待審核班表」），吸引使用者點擊 |
| 42 | BUG | 預覽 Modal 重新開啟時 fpWriteBtn / fpCancelBtn 仍保持寫入後變身狀態（核章列印 / 完成關閉）→ showFullPreview 開頭強制重置為 [✅ 確認寫入] [取消] 並重綁 onclick |

### 第七輪（ver4.8，2026-07-20 效能＋穩定性＋UI，Codex 三輪交叉驗證）
| 類別 | 摘要 |
|------|------|
| 效能 | 56 處 `SpreadsheetApp.openById` 統一改用 `getSpreadsheet()` 快取；`getScheduleData` 讀取合併（M1:N32 notes 一次讀、getScheduleChanges 的 I1:M11 名單一次讀） |
| 效能 | 前端：onload 預抓 `getAvailableSheets`（`_preSheets` TTL 10 分）、移除 onload 無效 `loadRecs`、`loadHdrs` 改前端常數、`viewSched` 結果存 `_schedCache`（TTL 2 分）供一鍵列印重用；所有寫入/建表成功路徑呼叫 `_invalidateCaches()` |
| 穩定性 | 新增 `withScriptLock` 共用鎖（tryLock＋深度計數防重入；搶不到鎖 throw／回忙碌，絕不無鎖執行）：套用 `updateShift`、`logShiftChange`、`writeOpLog`、`writeDragShiftLog`、`writeDraggedPreview`、`runAutoSchedule` 寫入段 |
| 穩定性 | `confirmShift` 只認「換班成功」開頭為成功（原本後端失敗字串被前端當成功，畫面已改但試算表沒變）；`updateShift` headers 先讀再寫、日誌失敗回「換班成功!（提醒…）」不誤報整體失敗 |
| 穩定性 | `writeDragShiftLog` 改收集後批次寫回（原日誌滿載時同批多筆互蓋只剩一筆）；`getScheduleChanges` 姓名不 filter，保持與員編列對齊 |
| UI | `showLoad(txt)` 改全域 `#gLoad` overlay（原 `#lt` 藏於登入面板，Modal 流程不可見）；`button:disabled` 全域視覺；換班/驗證/預覽/儲存/統計流程鎖按鈕防連點；`#ale` shake、清單與班表格按壓回饋、`fullPreviewModal` 進場動畫 |
| UI | floatTip 由 `td[data-tip]` 放寬為 `[data-tip]` 並支援觸控點按；`#tep`/fpZoom/異動 ✕ 鈕補 title；月份選單/換班視窗/預覽 Modal 補操作步驟說明 |
| UI | 月份選單「📁 過去班表」收納：前端 `_rocStrToNum`＋`_sheetIsPast` 動態分組（與後端 `isSheetLocked` 同邏輯），過去月份摺疊、點擊展開、無當期班表時自動展開 |

### 第八輪（ver4.9→ver5.1，2026-08-12 高齡認知＋雙帳本公平制＋預覽互動強化）
| 類別 | 摘要 |
|------|------|
| 新關卡 | L 欄改「卡介苗/高齡認知」時間共用：首個工作週二=卡介苗（Q欄3人池）、其餘工作週二=高齡認知（G1:G8 5人池），2026/9 起生效不追溯；`getLTypeForDate`/`lDisplayName` 為全系統唯一判定來源 |
| 公平制 | 雙帳本資格加權：卡介苗/高齡認知 1 次計 F1/F2 係數（預設1.5）進公平分數；引擎排序鍵②=月內公平分（一般班1分+資格班W分）、③=滾動近 F3 個月（預設3）每在職月平均公平分（取代跨年累積總量，修掉新人隱性超排）；受訓者自動少排一般關卡=受訓獎勵 |
| 計數基準 | 公平計數（引擎歷史視窗＋getYearlyClinicStats 看板）一律以 `buildOriginalScheduleMap` 反推的原始排班為準，換班不進公平帳 |
| 名稱翻譯 | LINE 查詢、Email 通知、換班/拖曳日誌、換班選人池（getShiftOptions 增 dateStr/sheetName 參數）、自檢，一律按「是否首個工作週二」翻譯 卡介苗/高齡認知；LEGACY_HDR 增 卡介苗/高齡認知/卡介苗/高齡認知(舊名)→卡介苗/高齡認知 對映 |
| 自檢 | check4 生效後改「加權公平分」比較；新增 check5 卡介苗/高齡認知欄規則（池別/非業務日空白/同日不兼任）；拖曳資格檢查（前端 `_lPoolOk`） |
| 看板 | getYearlyClinicStats：wdCounts 13 欄（卡介苗/高齡認知分列）＋bcgCnt/cogCnt/qualBonus/fairScore；排序改公平分；前端明細表增「獎勵分/公平分」欄＋制度說明註記 |
| 設定 | 系統設定新增：高齡認知名單 picklist（G欄）＋F1/F2/F3 設定；setSuccessor 名單替換含 G 欄 |
| 修正 | `getFirstTuesdayWorkday` 改掃整月（原只掃1~7日，首週二遇假日回 -1 → 整月無卡介苗且週二全被誤判高齡認知）；卡介苗首週二正選改用 pre-pass 嚴格 ABC 輪序（原 in-loop 奇偶排序與 pre-pass 可能不一致） |
| 版號 | ver4.9（meta/lx-ver/#vi 三處） |

---

## 架構升級摘要
- **ver5.3 卡介苗輪序改接續制**：pre-pass 改查「上月卡介苗日（首個工作週二）原始站班者」（`buildOriginalScheduleMap` 反推，換班/拖曳不進輪序帳，與值班/停班2線跨月接續同基準），Q 欄順位 +1 為本月人選；查無起點（上月無表/空白/不在 Q 欄）才回退原 v4.8.1 固定日曆公式（BASE 2026/6=Q1）
- **整年排班功能已移除**（人員變動大，只用單月排班）
- **年份全面動態化**：`parseYearMonthFromSheetName` 通用解析 + `rocStrToNum` 反向函式
- **假日多年度支援**：`GOV_HOLIDAYS` Map 結構，每年初需補充放假日資料
- **換班身份驗證**：`updateShift` 加入 `adminPw` 可選參數，管理員可幫他人換班
- **ver4.6 後台單欄分頁**：移除左側獨立容器，所有功能集中於 `adminTabs`；`tab-quick`（一鍵排班入口+進階工具）+ `tab-pending`（待審核班表清單）+ 既有 5 個 tab；移除 `tab-create`/`tab-sched`
- **ver4.6 預覽 Modal 寫入流程**：`showFullPreview` 一律帶 callback；`fpWriteBtn` 寫入成功後變身為「🖨️ 核章列印」、`fpCancelBtn` 變身為「✅ 完成關閉」並導向 `tab-pending`；右上角 `fpClose` 已移除
- **ver4.6 拖曳權限放寬**：`makeFpTableDraggable`/`makeFpTableAddable` 不再依 `_fpAuditMode` 區分排班者/審核者，所有人皆可挪移所有職務
- **ver4.6 停班2線系統挪移識別**：N1 note 加 `dengSwapRows:rIdx=M/d`；前端 hover 顯示原日期；`getPrevLastPtr` 讀 M 欄 swap note 跳過挪移列

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
| Q1:Q11 | 卡介苗候選（3人池：首週二輪值） |
| G1:G8  | 高齡認知候選（5人池：其餘週二輪值，ver4.9） |
| F1/F2  | 卡介苗/高齡認知 加權係數（預設1.5，ver4.9） |
| F3     | 公平基準滾動月數（預設3，ver4.9） |

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
| 9 | L | 卡介苗/高齡認知 | 首個工作週二=卡介苗(Q欄)；其餘工作週二=高齡認知(G欄，2026/9起不追溯) |
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
