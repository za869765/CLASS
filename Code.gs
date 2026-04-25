// =============================================
// 智慧排班系統 2.0 - Code.gs (含自動排班模組)
// ver3.23 - 修正：卡介苗月份奇偶交替，全年公平輪替
// =============================================

const SHEET_ID = '1NMiyJr0p6Vq6J2ubZy8xr3UArJhO-Vp3s4UXLOeqOUQ';
const EMAIL_SHEET_NAME = '班表設定';

// ★ 快取 Spreadsheet 物件（同一次 GAS 執行中重複使用，避免多次 openById）
let _ss = null;
function getSpreadsheet() {
  if (!_ss) _ss = SpreadsheetApp.openById(SHEET_ID);
  return _ss;
}

const GLOBAL_CONFIG = {
  AVAILABLE_SHEETS_RANGE: 'O1:O12',
  SCHEDULE_DATA_RANGE: 'C2:M32',
  SHIFT_OPTIONS: {
    'C': 'E1:E11',  // 值班：E欄順序（值班專屬輪序）
    'D': 'J1:J3',   // 協助掛號：限定人員
    'E': 'K1:K8',   // 門診
    'F': 'K1:K8',   // 流注1
    'G': 'K1:K8',   // 流注2
    'H': 'K1:K8',   // 預登1
    'I': 'K1:K8',   // 預登2
    'J': 'K1:K8',   // 預注1
    'K': 'K1:K8',   // 預注2
    'L': 'Q1:Q2',   // 卡介苗
    'M': 'B1:B11'   // 登革熱二線：B欄順序（登革熱專屬輪序）
  },
  SELECT_TYPE: {
    'C':'SC','D':'SC','E':'SC','F':'SC','G':'SC','H':'SC',
    'I':'SC','J':'SC','K':'SC','L':'SC','M':'SC'
  },
  SHIFT_HEADERS: [
    "值班","支援","門診","掛號","前台","預登1","預登2注","注射1","注射2","卡介苗","停班2線"
  ],
  AUTO_SCHEDULE: {
    STAFF_RANGE:    'I1:I11',
    EMPID_RANGE:    'M1:M11',
    EMAIL_RANGE:    'L1:L11',
    POINTS_RANGE:   'P1:P11',
    ADMIN_PASSWORD_RANGE: 'N2',
    LOG_RANGE:      'A40:A460',
  },
  // ── Line Bot 設定（存放於「班表設定」工作表 R / S 欄）──
  // R1  = Line Channel Access Token
  // R2  = 啟動搜尋關鍵字（前綴；空白=回應所有訊息）
  // R3  = 模糊搜尋 TRUE/FALSE（空白預設 TRUE）
  // R4:R15 = 指定搜尋工作表名稱（空白=自動全年115年班表）
  // S1:S13 = 搜尋欄位代號 A~M（空白=搜尋全部欄位）
  LINE_TOKEN_RANGE:      'R1',
  LINE_CHANNEL_ID:       'R16',
  LINE_CHANNEL_SECRET:   'R17',
  GEMINI_API_KEY_RANGE:  'R18',  // Google AI Studio API Key
  LINE_KEYWORD_RANGE: 'R2',
  LINE_FUZZY_RANGE:   'R3',
  LINE_SHEETS_RANGE:  'R4:R15',
  LINE_COLS_RANGE:    'S1:S13',
  // ── 排班排除設定（T欄：姓名, U欄：開始日期 M/D, V欄：結束日期 M/D 空白=持續，W欄：排除星期 如"六,日,假日"）──
  EXCLUSION_RANGE:    'T1:W30',
  JOIN_DATE_RANGE:    'X1:X11',   // 到職日 M/D（空白=1/1，視為全年在職）
  LEAVE_DATE_RANGE:   'Y1:Y11',   // 離職日 M/D（空白=12/31，視為全年在職）
  NOTIFY_RANGE:       'Z1:Z11',   // 是否接收通知（TRUE/FALSE）
  OP_LOG_RANGE:       'M40:M460'  // 操作紀錄（排班/設定變更）
};

// =============================================
// 政府行政機關放假日（多年度支援）
// BUG25 修正：改為 Map 結構，支援跨年度查詢
// 每年初需新增下一年度資料（或由外部維護）
// =============================================
const GOV_HOLIDAYS = {
  2025: new Set([
    // 114年（若需要跨年查詢上年12月，可在此補充）
  ]),
  2026: new Set([
    "2026-01-01","2026-01-03","2026-01-04","2026-01-10","2026-01-11",
    "2026-01-17","2026-01-18","2026-01-24","2026-01-25","2026-01-31",
    "2026-02-01","2026-02-07","2026-02-08","2026-02-14","2026-02-15",
    "2026-02-16","2026-02-17","2026-02-18","2026-02-19","2026-02-20",
    "2026-02-21","2026-02-22","2026-02-27","2026-02-28",
    "2026-03-01","2026-03-07","2026-03-08","2026-03-14","2026-03-15",
    "2026-03-21","2026-03-22","2026-03-28","2026-03-29",
    "2026-04-03","2026-04-04","2026-04-05","2026-04-06",
    "2026-04-11","2026-04-12","2026-04-18","2026-04-19",
    "2026-04-25","2026-04-26",
    "2026-05-01","2026-05-02","2026-05-03","2026-05-09","2026-05-10",
    "2026-05-16","2026-05-17","2026-05-23","2026-05-24","2026-05-30","2026-05-31",
    "2026-06-06","2026-06-07","2026-06-13","2026-06-14",
    "2026-06-19","2026-06-20","2026-06-21","2026-06-27","2026-06-28",
    "2026-07-04","2026-07-05","2026-07-11","2026-07-12",
    "2026-07-18","2026-07-19","2026-07-25","2026-07-26",
    "2026-08-01","2026-08-02","2026-08-08","2026-08-09",
    "2026-08-15","2026-08-16","2026-08-22","2026-08-23",
    "2026-08-29","2026-08-30",
    "2026-09-05","2026-09-06","2026-09-12","2026-09-13",
    "2026-09-19","2026-09-20","2026-09-25","2026-09-26","2026-09-27","2026-09-28",
    "2026-10-03","2026-10-04","2026-10-09","2026-10-10","2026-10-11",
    "2026-10-17","2026-10-18","2026-10-24","2026-10-25","2026-10-26",
    "2026-10-31",
    "2026-11-01","2026-11-07","2026-11-08","2026-11-14","2026-11-15",
    "2026-11-21","2026-11-22","2026-11-28","2026-11-29",
    "2026-12-05","2026-12-06","2026-12-12","2026-12-13",
    "2026-12-19","2026-12-20","2026-12-25","2026-12-26","2026-12-27"
  ]),
  2027: new Set([
    // 116年 — 待每年初更新
    // ★ 未填入時 isHoliday 僅判斷週六日，補假/國定假日不會生效
  ])
};

function isHoliday(dateObj) {
  const d = new Date(dateObj);
  const dow = d.getDay();
  if (dow === 0 || dow === 6) return true;
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  const key = `${y}-${m}-${day}`;
  const yearSet = GOV_HOLIDAYS[y];
  return yearSet ? yearSet.has(key) : false;
}

function isTuesdayWorkday(dateObj) {
  const d = new Date(dateObj);
  return d.getDay() === 2 && !isHoliday(d);
}

function isThursdayWorkday(dateObj) {
  const d = new Date(dateObj);
  return d.getDay() === 4 && !isHoliday(d);
}

function getFirstTuesdayWorkday(year, month) {
  for (let day = 1; day <= 7; day++) {
    const d = new Date(year, month - 1, day);
    if (isTuesdayWorkday(d)) return d.getDate();
  }
  return -1;
}

function parseDateFromSheet(dateStr, sheetName) {
  // BUG21 修正：從 sheetName 解析年份，不再硬編碼 2026
  if (!dateStr) return null;
  const match = dateStr.match(/(\d+)\/(\d+)/);
  if (!match) return null;
  const month = parseInt(match[1]);
  const day = parseInt(match[2]);
  const parsed = sheetName ? parseYearMonthFromSheetName(sheetName) : null;
  const year = (parsed && parsed.valid) ? parsed.year : new Date().getFullYear();
  return new Date(year, month - 1, day);
}

// =============================================
// 取得所有可用工作表
// =============================================
function getAvailableSheets() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const settingSheet = spreadsheet.getSheetByName(EMAIL_SHEET_NAME);
  const allSheetNames = spreadsheet.getSheets().map(s => s.getName());

  // ── 計算當年度民國年前綴 ──────────────────────────────────────
  // 例：2026年 → rocYear=115 → prefix=「一百一十五年」
  const now       = new Date();
  const rocYear   = now.getFullYear() - 1911;
  const rocPrefix = rocNumToStr(rocYear) + '年';   // 「一百一十五年」

  // ── O1:O12 有手動填寫 → 使用手動清單，但仍限定當年度 ─────────
  const allowedRaw = settingSheet.getRange(GLOBAL_CONFIG.AVAILABLE_SHEETS_RANGE)
    .getValues().flat()
    .filter(s => s && s.toString().trim() !== '');

  if (allowedRaw.length > 0) {
    return allowedRaw
      .map(s => s.toString().trim())
      .filter(n => allSheetNames.includes(n) && n.includes(rocPrefix));
  }

  // ── O1:O12 全空白 → 自動抓當年度民國格式班表 ─────────────────
  const yearSheets = allSheetNames.filter(n =>
    n !== EMAIL_SHEET_NAME && n.includes(rocPrefix)
  );

  // 依月份排序（一月→十二月）
  const monthOrder = ['一月','二月','三月','四月','五月','六月',
                      '七月','八月','九月','十月','十一月','十二月'];
  yearSheets.sort((a, b) => {
    const ia = monthOrder.findIndex(m => a.includes(m));
    const ib = monthOrder.findIndex(m => b.includes(m));
    return ia - ib;
  });

  // 若真的連民國年格式都找不到，退回「含班表字樣且不是設定表」
  if (yearSheets.length > 0) return yearSheets;
  return allSheetNames.filter(n =>
    n !== EMAIL_SHEET_NAME &&
    n.includes('班表') &&
    !n.includes('設定') &&
    !n.includes('副本')
  );
}

// 西元年轉民國漢字（供 getAvailableSheets 使用）
function rocNumToStr(n) {
  // 支援 100~199（即西元 2011~2110）
  // 注意：中文數字 110~119 寫作「一百一十x」，tens[1] 必須是「一十」而非「十」
  const units = ['','一','二','三','四','五','六','七','八','九'];
  const tens  = ['','一十','二十','三十','四十','五十','六十','七十','八十','九十'];
  if (n < 100 || n > 199) return String(n);
  const hundred = '一百';
  const rem = n - 100;               // 115 → 15
  const t   = Math.floor(rem / 10);  // 15 → 1
  const u   = rem % 10;              // 15 → 5
  let str = hundred;
  if (t > 0) str += tens[t];         // '一百' + '一十' → '一百一十'
  if (u > 0) str += units[u];        // '一百一十' + '五' → '一百一十五'
  return str;
  // 驗證：115 → 一百一十五　120 → 一百二十　100 → 一百
}

function getAllScheduleSheets() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const allSheetNames = spreadsheet.getSheets().map(s => s.getName());
  const scheduleSheets = allSheetNames.filter(n => n !== EMAIL_SHEET_NAME && n.includes('班表'));
  // 回傳「本年度全部」班表（含過去月份），依月份排序
  const rocYear   = new Date().getFullYear() - 1911;
  const rocPrefix = rocNumToStr(rocYear) + '年';
  const current   = scheduleSheets.filter(n => n.includes(rocPrefix));
  current.sort((a, b) => {
    const pa = parseYearMonthFromSheetName(a);
    const pb = parseYearMonthFromSheetName(b);
    return (pa.year * 100 + pa.month) - (pb.year * 100 + pb.month);
  });
  return current.length > 0 ? current : scheduleSheets;
}

// ── 補寄用：取得指定月份每日有哪些人排班（預覽用）────────────────
function getScheduleDatesWithDuties(sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const timezone    = spreadsheet.getSpreadsheetTimeZone();
    const sched = spreadsheet.getSheetByName(sheetName);
    if (!sched) return { success: false, dates: [] };
    const dateCol = sched.getRange('A2:A32').getValues();
    const headers = sched.getRange('C1:M1').getValues()[0].map(h => h ? h.toString().trim() : '');
    const data    = sched.getRange('C2:M32').getValues();
    const result  = [];
    dateCol.forEach((r, i) => {
      const v = r[0];
      if (!v) return;
      const dateStr = v instanceof Date ? Utilities.formatDate(v, timezone, 'M/d') : v.toString().trim();
      const row = data[i] || [];
      // 收集當日有值的班別與人員
      const duties = [];
      row.forEach((cell, ci) => {
        if (cell && cell.toString().trim()) {
          duties.push({ shift: headers[ci] || '', person: cell.toString().trim() });
        }
      });
      if (duties.length > 0) {
        // 彙整所有人員去重
        const persons = [...new Set(duties.map(d => d.person).filter(Boolean))];
        result.push({ date: dateStr, duties, persons });
      }
    });
    return { success: true, dates: result };
  } catch(e) { return { success: false, dates: [], message: e.message }; }
}


function listAllSheets() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  return spreadsheet.getSheets().map(s => s.getName());
}


// =============================================
// 取得當年度全部班表（含1-12月，不限制過去/未來）
// BUG20 連帶修正：不再硬編碼 115/2026，動態取當年度
// =============================================
function getAllYear115Sheets() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const allNames = spreadsheet.getSheets().map(s => s.getName());
  const currentYear = new Date().getFullYear();
  const sheetsThisYear = allNames.filter(n => {
    if (n === EMAIL_SHEET_NAME) return false;
    const ym = parseYearMonthFromSheetName(n);
    return ym.valid && ym.year === currentYear;
  });
  sheetsThisYear.sort((a, b) => {
    const pa = parseYearMonthFromSheetName(a);
    const pb = parseYearMonthFromSheetName(b);
    return (pa.year * 100 + pa.month) - (pb.year * 100 + pb.month);
  });
  return sheetsThisYear;
}

// =============================================
// 建立單月班表工作表
// =============================================
function createScheduleSheet(adminPassword, year, month, sheetName) {
  if (!verifyAdminPassword(adminPassword)) return { success: false, message: '管理員密碼錯誤。' };
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    if (spreadsheet.getSheetByName(sheetName)) {
      return { success: false, message: '工作表「' + sheetName + '」已存在。' };
    }
    const newSheet = spreadsheet.insertSheet(sheetName, 0);
    const weekdays = ['日','一','二','三','四','五','六'];
    newSheet.getRange('A1').setValue('日期');
    newSheet.getRange('B1').setValue('星期');
    const hdrs = GLOBAL_CONFIG.SHIFT_HEADERS;
    for (let i = 0; i < hdrs.length; i++) newSheet.getRange(1, 3 + i).setValue(hdrs[i]);
    const headerRange = newSheet.getRange('A1:M1');
    headerRange.setBackground('#e0eafc');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    let row = 2;
    const daysInMonth = new Date(year, month, 0).getDate();
    // ★ 先將 A2:A32 設為純文字，避免 Google Sheets 自動轉為 Date 物件
    newSheet.getRange('A2:A32').setNumberFormat('@');
    for (let d = 1; d <= daysInMonth && row <= 32; d++) {
      const date = new Date(year, month - 1, d);
      const dow = date.getDay();
      newSheet.getRange(row, 1).setValue(month + '/' + d);
      newSheet.getRange(row, 2).setValue('週' + weekdays[dow]);
      const isWknd = dow === 0 || dow === 6;
      const dateKey = year + '-' + String(month).padStart(2,'0') + '-' + String(d).padStart(2,'0');
      const yearSet2 = GOV_HOLIDAYS[year];
      const isHoliday2 = yearSet2 ? yearSet2.has(dateKey) : false;
      if (isWknd) {
        newSheet.getRange(row, 1, 1, 2).setFontColor('#d32f2f');
      } else if (isHoliday2) {
        newSheet.getRange(row, 1, 1, 2).setFontColor('#e65100').setBackground('#fff3e0');
      }
      row++;
    }
    newSheet.setColumnWidth(1, 70); newSheet.setColumnWidth(2, 50);
    for (let c = 3; c <= 13; c++) newSheet.setColumnWidth(c, 65);
    newSheet.setFrozenRows(1);
    return { success: true, message: '工作表「' + sheetName + '」建立完成！' };
  } catch(e) {
    return { success: false, message: '建立失敗：' + e.message };
  }
}

// =============================================
// 批次建立：下個月到115年12月
// =============================================
// createYearlySheets 已移除（整年排班功能已簡化為單月連排，不再需要）

// =============================================
// 原有功能
// =============================================

// =============================================
// 工具：取得某日期的班表資料（{date, dateStr, duties}）
// duties = [{shift, person}]
// =============================================
function getDaySchedule(spreadsheet, dateObj, timezone) {
  const rocYear  = dateObj.getFullYear() - 1911;
  const rocMonth = ['','一','二','三','四','五','六','七','八','九','十','十一','十二'][dateObj.getMonth()+1];
  const sheetName = rocNumToStr(rocYear) + '年' + rocMonth + '月班表';
  const sched = spreadsheet.getSheetByName(sheetName);
  if (!sched) return null;
  const dateStr = Utilities.formatDate(dateObj, timezone, 'M/d');
  const weekDay = ['日','一','二','三','四','五','六'][dateObj.getDay()];
  const dateCol = sched.getRange('A2:A32').getValues();
  let targetRow = -1;
  for (let i = 0; i < dateCol.length; i++) {
    const v = dateCol[i][0];
    const s = v instanceof Date ? Utilities.formatDate(v, timezone, 'M/d') : (v ? v.toString().trim() : '');
    if (s === dateStr) { targetRow = i + 2; break; }
  }
  if (targetRow === -1) return null;
  const headers = sched.getRange('C1:M1').getValues()[0];
  const rowData = sched.getRange(targetRow, 3, 1, 11).getValues()[0];
  const duties  = [];
  rowData.forEach((val, ci) => {
    if (val && val.toString().trim())
      duties.push({ shift: headers[ci] || '', person: val.toString().trim() });
  });
  return { date: dateObj, dateStr, weekDay, duties };
}

// =============================================
// 建立美化 HTML 通知信件
// =============================================
function buildScheduleHtml(personName, daySchedules, isAuto) {
  const intro = isAuto
    ? `您好，以下是您的即將到來的班表提醒。`
    : `您好，以下是您被補寄的班表資訊。`;

  const tableRows = daySchedules.map(ds => {
    const duties = ds.duties || [];
    if (duties.length === 0) return '';
    const myDuties = duties.filter(d => d.person && d.person.includes(personName));
    if (myDuties.length === 0) return '';

    const isHol  = isHoliday(ds.date);
    const dayBg  = isHol ? '#fff3e0' : '#e3f2fd';
    const dayCol = isHol ? '#e65100' : '#1565c0';
    let html = `<tr><td colspan="2" style="background:${dayBg};padding:8px 14px;font-weight:700;color:${dayCol};font-size:.95rem">
      📅 ${ds.dateStr}（週${ds.weekDay}）${isHol?'　🏖 假日/休假':''}
    </td></tr>`;
    myDuties.forEach(d => {
      html += `<tr style="background:#fff9c4"><td style="padding:7px 14px;border:1px solid #ddd;font-weight:700;color:#e65100">${d.shift}</td>
               <td style="padding:7px 14px;border:1px solid #ddd;font-weight:700;color:#e65100">${d.person}</td></tr>`;
    });
    // 其他非本人的班別（灰色）
    duties.filter(d => !myDuties.includes(d)).forEach(d => {
      html += `<tr><td style="padding:6px 14px;border:1px solid #eee;color:#666">${d.shift}</td>
               <td style="padding:6px 14px;border:1px solid #eee;color:#666">${d.person}</td></tr>`;
    });
    return html;
  }).filter(Boolean).join('');

  if (!tableRows) return null;

  return `<!DOCTYPE html><html><body style="font-family:'Helvetica Neue',Arial,sans-serif;color:#333;max-width:600px;margin:0 auto;padding:16px">
  <div style="background:linear-gradient(135deg,#1565c0,#0288d1);padding:18px 24px;border-radius:12px 12px 0 0;text-align:center">
    <div style="font-size:1.6rem;color:#fff;font-weight:900">佳里區衛生所</div>
    <div style="font-size:.9rem;color:rgba(255,255,255,.85);margin-top:4px">班表小幫手 — 班表提醒</div>
  </div>
  <div style="background:#fff;border:1px solid #e0e0e0;border-radius:0 0 12px 12px;padding:20px 24px">
    <p style="font-size:1rem;margin:0 0 14px">${intro}</p>
    <table style="border-collapse:collapse;width:100%;font-size:.88rem">
      <tr style="background:#1565c0;color:#fff">
        <th style="padding:8px 14px;text-align:left">職務</th>
        <th style="padding:8px 14px;text-align:left">人員</th>
      </tr>
      ${tableRows}
    </table>
    <p style="font-size:.78rem;color:#999;margin-top:20px;border-top:1px solid #eee;padding-top:10px">
     ⚠️ 此信件由系統自動發送，請勿回覆。<br>
      本信件為班表提醒通知，非社交工程郵件。
    </p>
  </div>
  </body></html>`;
}

// =============================================
// 每日觸發：前一個工作日 16:00 寄出
// 若明天是假日，則一併寄出連續假日+下一個工作日的班表
// =============================================
function checkTomorrowScheduleAndSendEmail() {
  Logger.log('--- 開始執行每日班表提醒 ---');
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const timezone    = spreadsheet.getSpreadsheetTimeZone();
    const today       = new Date();

    // ── 核心邏輯：只在「工作日」執行寄送 ─────────────────────────
    // 若今天是假日（週末或國定假日），代表前一個工作日已預先涵蓋今天，直接跳過
    if (isHoliday(today)) {
      Logger.log('今日為假日，跳過寄送（前一工作日已預先涵蓋）');
      return;
    }

    // ── 計算需要寄送的日期清單 ────────────────────────────────────
    // 從明天起，依序加入：所有連續假日 + 下一個工作日（含）
    const datesToNotify = [];
    let offset = 1;
    let foundWorkday = false;
    for (let i = 0; i < 10 && !foundWorkday; i++) {
      const dd = new Date(today);
      dd.setDate(today.getDate() + offset + i);
      datesToNotify.push(new Date(dd));
      if (!isHoliday(dd)) foundWorkday = true;
    }

    // 讀取人員清單與通知設定
    const settingSheet = spreadsheet.getSheetByName(EMAIL_SHEET_NAME);
    const names    = settingSheet.getRange('I1:I11').getValues().flat();
    const emails   = settingSheet.getRange('L1:L11').getValues().flat();
    const notifyRaw = settingSheet.getRange(GLOBAL_CONFIG.NOTIFY_RANGE).getValues().flat();
    const staffNotify = []; // [{name, email}]
    for (let i = 0; i < 11; i++) {
      if (!names[i]) continue;
      const notify = notifyRaw[i] === true || notifyRaw[i] === 'TRUE' || notifyRaw[i] === 1;
      if (!notify) continue;
      const nm = names[i].toString().trim();
      const em = emails[i] ? emails[i].toString().trim() : '';
      if (em && em.includes('@')) staffNotify.push({ name: nm, email: em });
    }
    if (!staffNotify.length) { Logger.log('無需通知人員'); return; }

    // 取得每個日期的班表
    const daySchedules = datesToNotify.map(dd => getDaySchedule(spreadsheet, dd, timezone)).filter(Boolean);
    if (!daySchedules.length) { Logger.log('查無班表資料'); return; }

    // 整理主旨：列出日期範圍 + 班別
    const dateRange = daySchedules.length === 1
      ? daySchedules[0].dateStr
      : daySchedules[0].dateStr + '～' + daySchedules[daySchedules.length-1].dateStr;

    let sent = 0;
    staffNotify.forEach(({ name, email }) => {
      const html = buildScheduleHtml(name, daySchedules, true);
      if (!html) return; // 這個人沒有排到
      // 找出本人有哪些職務（for subject）
      const myDutyNames = [];
      daySchedules.forEach(ds => {
        ds.duties.forEach(d => {
          if (d.person && d.person.includes(name) && !myDutyNames.includes(d.shift))
            myDutyNames.push(d.shift);
        });
      });
      const subject = `佳里區衛生所班表通知(非社交工程) ${dateRange} ${myDutyNames.join('、')}`;
      try {
        MailApp.sendEmail({ to: email, subject, htmlBody: html, name: '班表小幫手' });
        sent++;
        Logger.log(`已寄出：${name} <${email}> 主旨：${subject}`);
      } catch(e) { Logger.log('寄信失敗 ' + name + ': ' + e.message); }
    });
    Logger.log(`班表提醒共寄出 ${sent} 封，通知日期：${dateRange}`);
  } catch(e) { Logger.log('執行時發生錯誤: ' + e.toString()); }
}

// =============================================
// 手動補寄：指定人員 + 指定日期
// recipients = [{name, email}]，targetDate = 'M/d' 格式
// =============================================
function sendManualScheduleNotice(adminPassword, recipientNames, targetDate) {
  if (!verifyAdminPassword(adminPassword)) return { success: false, message: '管理員密碼錯誤' };
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const timezone    = spreadsheet.getSpreadsheetTimeZone();
    const settingSheet = spreadsheet.getSheetByName(EMAIL_SHEET_NAME);
    const names   = settingSheet.getRange('I1:I11').getValues().flat();
    const emails  = settingSheet.getRange('L1:L11').getValues().flat();

    // 解析目標日期
    const [mon, day] = targetDate.split('/').map(Number);
    const now   = new Date();
    const dObj  = new Date(now.getFullYear(), mon - 1, day);
    const ds    = getDaySchedule(spreadsheet, dObj, timezone);
    if (!ds || !ds.duties.length) return { success: false, message: `${targetDate} 查無班表資料` };

    let sent = 0, skipped = 0;
    recipientNames.forEach(rname => {
      const idx = names.findIndex(n => n && n.toString().trim() === rname);
      const email = idx !== -1 && emails[idx] ? emails[idx].toString().trim() : '';
      if (!email || !email.includes('@')) { skipped++; return; }
      const html = buildScheduleHtml(rname, [ds], false);
      if (!html) { skipped++; return; }
      const myDuties = ds.duties.filter(d => d.person && d.person.includes(rname)).map(d => d.shift);
      const subject  = `佳里區衛生所班表通知(非社交工程) ${ds.dateStr}（週${ds.weekDay}） ${myDuties.join('、')||'班表資訊'}`;
      MailApp.sendEmail({ to: email, subject, htmlBody: html, name: '班表小幫手' });
      sent++;
    });
    writeOpLog('手動補寄', `${targetDate} 寄出 ${sent} 封，略過 ${skipped} 位`);
    return { success: true, message: `${targetDate} 班表已補寄 ${sent} 封${skipped?'，略過 '+skipped+' 位':''}` };
  } catch(e) { return { success: false, message: '補寄失敗：' + e.message }; }
}



function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 注：「進入班表」按鈕不需要密碼，verifyPassword(N1) 已移除。

function verifyAdminPassword(inputPassword) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    return inputPassword === sheet.getRange(GLOBAL_CONFIG.AUTO_SCHEDULE.ADMIN_PASSWORD_RANGE).getValue().toString();
  } catch(e) { return false; }
}

// ─── 判斷工作表是否為過去月份（應鎖定）────────────────────────
// 規則：工作表對應的年月 < 今天的年月 → 鎖定
function isSheetLocked(sheetName) {
  const ym = parseYearMonthFromSheetName(sheetName);
  if (!ym.valid) return false;   // 無法解析 → 不鎖定
  const now = new Date();
  const thisYear  = now.getFullYear();
  const thisMonth = now.getMonth() + 1;
  return (ym.year < thisYear) || (ym.year === thisYear && ym.month < thisMonth);
}

function getScheduleData(sheetName) {
  const spreadsheet = getSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error('找不到選擇的工作表');

  // ★ 一次讀取 A1:M32 全部資料，減少 API 呼叫次數
  const allData = sheet.getRange('A1:M32').getValues();
  const tz = spreadsheet.getSpreadsheetTimeZone();

  const headers  = allData[0].slice(2);            // C1:M1 (index 2~12)
  const datesRange = allData.slice(1);             // A2:B32 (rows 1~31)
  const dataRange  = datesRange.map(r => r.slice(2)); // C2:M32

  const combinedDates = datesRange.map(row => {
    const rawA = row[0];
    let a = '';
    if (rawA instanceof Date) {
      a = Utilities.formatDate(rawA, tz, 'M/d');
    } else {
      // 嘗試解析 YYYY/M/D 格式
      const s = rawA ? rawA.toString().trim() : '';
      const ymMatch = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
      a = ymMatch ? parseInt(ymMatch[2]) + '/' + parseInt(ymMatch[3]) : s.split(' ')[0].trim();
    }
    const b = row[1] ? row[1].toString().trim() : '';
    return (a + (b ? ' ' + b : '')).trim();
  });

  // ★ 備註與換班紀錄：只在有需要時讀取
  // 登革熱挪移備註（M欄）- 批次讀取
  let dengSwapDateMap = {};
  try {
    const mNotes = sheet.getRange('M2:M32').getNotes();
    mNotes.forEach((row, i) => {
      const note = row[0] ? row[0].toString().trim() : '';
      if (note.startsWith('swap:')) dengSwapDateMap[i] = note.replace('swap:', '');
    });
  } catch(e) {}

  // ★ 讀取系統排定時間 + 審核狀態 + 版次（N1 備註）
  let schedTime = '';
  let reviewStatus = ''; // 'pending'=審核中(僅檢視) | 'approved'=已核准 | ''=未設定
  let writeCount = 0;    // 排定次數（1=A, 2=B, ...）
  try {
    const note = sheet.getRange('N1').getNote();
    if (note) {
      note.split('\n').forEach(line => {
        if (line.startsWith('排定時間:'))  schedTime    = line.replace('排定時間:', '').trim();
        if (line.startsWith('審核狀態:'))  reviewStatus = line.replace('審核狀態:', '').trim();
        if (line.startsWith('writeCount:')) writeCount  = parseInt(line.replace('writeCount:', '').trim()) || 0;
      });
    }
  } catch(e) {}

  // ★ 計算每列是否為假日（供前端 fullPreview 正確標色，含清明等非六日假日）
  const holidayRows = datesRange.map(row => {
    const rawA = row[0];
    if (!rawA) return false;
    let d = null;
    if (rawA instanceof Date) { d = rawA; }
    else {
      const s = rawA.toString().trim();
      const ymMatch = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
      if (ymMatch) { try { d = new Date(parseInt(ymMatch[1]), parseInt(ymMatch[2])-1, parseInt(ymMatch[3])); } catch(e){} }
      else {
        const mdMatch = s.match(/^(\d{1,2})\/(\d{1,2})$/);
        if (mdMatch) {
          const _ym = parseYearMonthFromSheetName(sheetName);
          const guessYear = _ym.valid ? _ym.year : new Date().getFullYear();
          try { d = new Date(guessYear, parseInt(mdMatch[1])-1, parseInt(mdMatch[2])); } catch(e){}
        }
      }
    }
    return d ? isHoliday(d) : false;
  });

  return {
    dates:           combinedDates.map(d => [d]),
    headers,
    schedule:        dataRange,
    locked:          isSheetLocked(sheetName),
    changes:         getScheduleChanges(sheetName, headers),
    dengSwapped:     Object.keys(dengSwapDateMap).map(Number),
    dengSwapDateMap: dengSwapDateMap,
    schedTime:       schedTime,
    reviewStatus:    reviewStatus,
    writeCount:      writeCount,
    holidayRows:     holidayRows
  };
}

function getShiftChangeRecords(sheetName, empId) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
  return sheet.getRange('A40:A460').getValues().filter(r => r[0] !== "").map(r => r[0]);
}

function getShiftHeaders() { return GLOBAL_CONFIG.SHIFT_HEADERS; }

// ── 審核狀態管理 ──────────────────────────────────────────────────────
// 取得所有審核中（pending）的班表
// ★ 版次字母（A~Z, AA~AZ, BA~BZ, ...）
function wcToLetter(wc) {
  if (!wc || wc <= 0) return '';
  let n = wc;
  let result = '';
  while (n > 0) {
    n--;  // 0-based
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

function getPendingSheets() {
  try {
    const spreadsheet = getSpreadsheet();
    const sheets = spreadsheet.getSheets();
    const pending = [];   // 舊格式（向後相容）
    const pendingInfo = []; // 新格式，含版次/時間
    sheets.forEach(sh => {
      const name = sh.getName();
      if (name === EMAIL_SHEET_NAME) return;
      try {
        const note = sh.getRange('N1').getNote() || '';
        let isPending = false, schedTime = '', writeCount = 0;
        note.split('\n').forEach(line => {
          const t = line.trim();
          if (t === '審核狀態:pending') isPending = true;
          if (t.startsWith('排定時間:')) schedTime = t.replace('排定時間:', '').trim();
          if (t.startsWith('writeCount:')) writeCount = parseInt(t.replace('writeCount:', '').trim()) || 0;
        });
        if (isPending) {
          pending.push(name);
          const wcLetter = writeCount > 0 ? wcToLetter(writeCount) : '';
          pendingInfo.push({ name, schedTime, writeCount, wcLetter });
        }
      } catch(e) {}
    });
    return { success: true, sheets: pending, pendingInfo };
  } catch(e) { return { success: false, sheets: [], pendingInfo: [] }; }
}

// status: 'pending'=審核中 | 'approved'=已核准 | ''=清除
function setReviewStatus(adminPassword, sheetName, status) {
  if (!verifyAdminPassword(adminPassword)) return { success: false, message: '管理員密碼錯誤。' };
  try {
    const sheet = getSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return { success: false, message: '找不到工作表：' + sheetName };
    const oldNote = sheet.getRange('N1').getNote() || '';
    // 保留排定時間，替換審核狀態
    const lines = oldNote.split('\n').filter(l => l && !l.startsWith('審核狀態:'));
    if (status) lines.push('審核狀態:' + status);
    sheet.getRange('N1').setNote(lines.join('\n'));
    const labels = { pending:'審核中', approved:'已核准', '':'已清除' };
    return { success: true, message: '「' + sheetName + '」已設為' + (labels[status]||status) };
  } catch(e) { return { success: false, message: e.message }; }
}

function getShiftOptions(column) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
  const columnLetter = String.fromCharCode(64 + column);
  const range = GLOBAL_CONFIG.SHIFT_OPTIONS[columnLetter];
  let options = range ? sheet.getRange(range).getValues().flat().filter(o => o) : [];
  return { options, selectType: GLOBAL_CONFIG.SELECT_TYPE[columnLetter] || 'SC' };
}

function updateShift(sheetName, date, shiftColumn, newShifts, sendEmail, empId, remark, adminPw) {
  if (!sheetName || typeof sheetName !== 'string' || !date || !shiftColumn || !Array.isArray(newShifts))
    return '參數錯誤，請聯絡管理員。';
  // ── 過去月份鎖定，拒絕寫入 ──────────────────────────────────
  if (isSheetLocked(sheetName)) return '🔒 此月份班表已鎖定，無法換班。';
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const sheet = spreadsheet.getSheetByName(sheetName);
    const row = findRowByDate(sheet, date.trim());
    if (row === null) return '找不到日期，請確認班表資料。';
    const oldShift = sheet.getRange(row, shiftColumn).getValue();
    const newShift = newShifts.join(', ');
    if (oldShift === newShift) return '班別未改變，無需更新。';

    // 授權人員（verifyEmpId 已驗證）均可更換任何人的班次，無自身限制
    sheet.getRange(row, shiftColumn).setValue(newShift);
    const headers = sheet.getRange('C1:M1').getValues()[0];
    const shiftType = headers[shiftColumn - 3] || '未知班別';
    logShiftChange(sheetName, date.trim(), oldShift, newShift, shiftType, empId, remark);
    if (sendEmail) {
      try { sendEmailNotification(oldShift, newShift, date.trim(), shiftType); } catch(e) {}
    }
    return '換班成功!';
  } catch (err) { return '系統錯誤，請聯絡管理員。'; }
}

function findRowByDate(sheet, targetDate) {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const tz = spreadsheet.getSpreadsheetTimeZone();
  const dateRange = sheet.getRange('A2:B32').getValues();

  // 前端傳入的日期部分（去掉星期），例如 "4/6 週一" → "4/6"
  const targetDateOnly = targetDate.split(' ')[0].trim();

  for (let i = 0; i < dateRange.length; i++) {
    const rawA = dateRange[i][0];
    if (!rawA) continue;

    let aStr = '';
    if (rawA instanceof Date) {
      aStr = Utilities.formatDate(rawA, tz, 'M/d');
    } else {
      // 字串格式：可能是 "4/6" 或 "4/6 週一" 或 "2026/4/6"
      const s = rawA.toString().trim();
      // 嘗試解析 YYYY/M/D 格式
      const ymMatch = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
      if (ymMatch) {
        aStr = parseInt(ymMatch[2]) + '/' + parseInt(ymMatch[3]);
      } else {
        aStr = s.split(' ')[0].trim(); // 取日期部分
      }
    }

    const bStr = dateRange[i][1] ? dateRange[i][1].toString().trim() : '';
    const sheetDateFull = (aStr + (bStr ? ' ' + bStr : '')).trim();
    const sheetDateOnly = aStr;

    // 先嘗試完整比對（含星期），再嘗試只比對日期
    if (sheetDateFull === targetDate || sheetDateOnly === targetDateOnly) {
      return i + 2;
    }
  }
  return null;
}

function logShiftChange(sheetName, date, oldShift, newShift, shiftType, empId, remark) {
  // BUG 6 修正：自動補齊星期，確保日誌格式一致（getScheduleChanges regex 需要 週X）
  let dateWithDay = date ? date.toString().trim() : '';
  if (dateWithDay && !/週./.test(dateWithDay)) {
    const weekDays = ['日','一','二','三','四','五','六'];
    const parts = dateWithDay.match(/(\d+)\/(\d+)/);
    if (parts) {
      try {
        const d = new Date(new Date().getFullYear(), parseInt(parts[1])-1, parseInt(parts[2]));
        dateWithDay = dateWithDay + ' 週' + weekDays[d.getDay()];
      } catch(e) {}
    }
  }
  const logSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
  const logRangeObj = logSheet.getRange('A40:A460');
  const logRange = logRangeObj.getValues();
  let logRow = 40;
  let found = false;
  for (let i = 0; i < logRange.length; i++) {
    if (logRange[i][0] === "") { logRow = i + 40; found = true; break; }
  }
  if (!found) {
    // bug #23 / #25: 原本 419 次個別 setValue/getValue 太慢易超時，且只迴圈到 i<459 → A460 那筆未被覆蓋會殘留。
    // 改用 getValues/setValues 一次讀寫 + 整段往上 shift 一格，最後一格 A460 由 setValue 寫入新紀錄。
    const shifted = [];
    for (let i = 1; i < logRange.length; i++) shifted.push([logRange[i][0]]);
    shifted.push(['']); // 末位先清空，等下方 setValue 寫入新紀錄
    logRangeObj.setValues(shifted);
    logRow = 460;
  }
  const empIdStr = empId ? `(${empId}) ` : '';
  const remarkStr = remark ? ' 備註:' + remark.toString().trim().split('\n').join(' ') : '';
  logSheet.getRange(logRow, 1).setValue(
    `${dateWithDay} ${shiftType} 原${oldShift}→新${newShift} ${empIdStr}更換時間: ${new Date().toLocaleString()}${remarkStr}`
  );
}

function sendEmailNotification(oldShift, newShift, targetDate, shiftType) {
  const emailSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
  const namesRange = emailSheet.getRange('I1:I11').getValues().flat();
  const emailsRange = emailSheet.getRange('L1:L11').getValues().flat();
  let emails = [];
  newShift.split(', ').forEach(shift => {
    const idx = namesRange.indexOf(shift);
    if (idx !== -1) emails.push(emailsRange[idx]);
  });
  const oldIdx = namesRange.indexOf(oldShift);
  if (oldIdx !== -1) emails.push(emailsRange[oldIdx]);
  emails = [...new Set(emails)];
  if (emails.length > 0) {
    MailApp.sendEmail({
      to: emails.join(','),
      subject: "班表更換通知",
      body: `班表更換完成:\n\n日期: ${targetDate}\n班別種類: ${shiftType}\n原人員: ${oldShift}\n新人員: ${newShift}\n更換時間: ${new Date().toLocaleString()}`,
      name: '班表管理系統'
    });
  }
}

// ── 驗證員工編號是否對應特定姓名（個人統計用）──────────────────
// 用 H欄 empId 對應 I欄 姓名 (H=empId, I=name，與班表設定一致)
function verifyEmpIdForPerson(empId, personName) {
  if (!empId || !personName) return false;
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    const names  = sheet.getRange('I1:I11').getValues().flat().map(s => s.toString().trim());
    // H欄：與 I欄平行的員工編號（也接受 M欄作為後備）
    const hIds   = sheet.getRange('H1:H11').getValues().flat().map(s => s.toString().trim());
    const mIds   = sheet.getRange('M1:M11').getValues().flat().map(s => s.toString().trim());

    const normalizeId = id => {
      if (!id) return '';
      id = id.trim();
      return (id.length > 0 && isNaN(parseInt(id.charAt(0))))
        ? id.charAt(0).toUpperCase() + id.substring(1) : id;
    };
    const normInput = normalizeId(empId);

    for (let i = 0; i < names.length; i++) {
      if (names[i] === personName) {
        if (normalizeId(hIds[i]) === normInput) return true;
        if (normalizeId(mIds[i]) === normInput) return true;
      }
    }
    return false;
  } catch(e) { return false; }
}

// ── 驗證員工編號（換班用，只驗證編號存在）───────────────────────
function verifyEmpId(empId) {
  if (!empId) return false;
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    const empIds = sheet.getRange('M1:M11').getValues().flat().map(String);
    let norm = empId;
    if (empId.length > 0 && isNaN(parseInt(empId.charAt(0))))
      norm = empId.charAt(0).toUpperCase() + empId.substring(1);
    const normIds = empIds.map(id =>
      (id.length > 0 && isNaN(parseInt(id.charAt(0)))) ? id.charAt(0).toUpperCase() + id.substring(1) : id
    );
    return normIds.includes(norm);
  } catch(e) { return false; }
}

// ── 解析換班 Log，回傳 {dateShiftKey → [{old, new, time, empId}]} ──
// log 格式：「3/15 週六 值班 原王小明→新鄭兆鑫 (a00460) 更換時間: ...」
function getScheduleChanges(sheetName, preloadedHeaders) {
  try {
    const spreadsheet  = getSpreadsheet();
    const settingSheet = spreadsheet.getSheetByName(EMAIL_SHEET_NAME);
    const logs = settingSheet.getRange('A40:A460').getValues()
      .flat().filter(r => r && r.toString().trim() !== '');

    // 使用傳入的 headers，否則才讀一次
    let sheetHdrs = preloadedHeaders || [];
    if (!sheetHdrs.length) {
      try {
        const sh = spreadsheet.getSheetByName(sheetName);
        if (sh) sheetHdrs = sh.getRange('C1:M1').getValues()[0].map(h => h.toString().trim());
      } catch(e) {}
    } else {
      sheetHdrs = sheetHdrs.map(h => h ? h.toString().trim() : '');
    }
    const hdrSet = new Set(sheetHdrs.map(h => h ? h.toString().trim() : '').filter(h => h));

    // 讀取員工名單，供 empId → name 轉換
    const staffNames = settingSheet.getRange('I1:I11').getValues().flat().map(String).filter(n=>n.trim());
    const staffEmpIds = settingSheet.getRange('M1:M11').getValues().flat().map(String);
    function empIdToName(eid) {
      if (!eid) return '';
      const norm = id => (id.length>0 && isNaN(parseInt(id.charAt(0))))
        ? id.charAt(0).toUpperCase()+id.substring(1) : id;
      const normEid = norm(eid);
      for (let i=0; i<staffEmpIds.length; i++) {
        if (norm(staffEmpIds[i]) === normEid && staffNames[i]) return staffNames[i];
      }
      return eid; // fallback: show raw empId
    }

    // 標頭正規化 map（舊日誌「值班人員」→「值班」等）
    const hdrNormMap = {};
    sheetHdrs.forEach(h => { if (h) hdrNormMap[h] = h; });
    // 常見舊名稱對應
    const LEGACY_HDR = { '值班人員':'值班', '登革熱二線':'停班2線', '協助掛號':'支援' };
    Object.keys(LEGACY_HDR).forEach(old => {
      if (!hdrNormMap[old]) hdrNormMap[old] = LEGACY_HDR[old];
    });
    // 若 sheetHdrs 本身有舊名，直接對應到 sheetHdrs 中的名稱
    sheetHdrs.forEach(h => { if (h && LEGACY_HDR[h]) hdrNormMap[LEGACY_HDR[h]] = h; });

    const changes = {};
    logs.forEach(log => {
      const s = log.toString().trim();
      // 支援可選備註：「...更換時間: XXX 備註:YYY」
      const m = s.match(/^(\d+\/\d+)\s+週.\s+(.+?)\s+原(.+?)→新(.+?)\s+(?:\(([^)]+)\)\s+)?更換時間:\s+([^備]+?)(?:\s+備註:(.+))?$/);
      if (!m) return;
      const [, date, shiftType, oldVal, newVal, empId, time, remark] = m;
      const stRaw = shiftType.trim();
      // 標頭正規化：舊名稱轉換為當前標頭
      const st = hdrNormMap[stRaw] || stRaw;
      // hdrSet 為空時不過濾（容錯），不為空才比對
      if (hdrSet.size > 0 && !hdrSet.has(st)) return;
      const key = date + '|' + st;
      if (!changes[key]) changes[key] = [];
      changes[key].push({
        oldVal, newVal,
        empId:  empId || '',
        name:   empIdToName(empId || ''),  // ★ 員工姓名
        time:   time.trim(),
        remark: remark ? remark.trim() : '',  // ★ 備註
        isDragLog: false
      });
    });

    // ★ 解析 排班-XXX / 審核-XXX 拖曳紀錄
    logs.forEach(log => {
      const s2 = log.toString().trim();
      const m2 = s2.match(/^(排班|審核)-(.+?)\s+(\d+\/\d+)(?:\s+週.)?\s+(.+?)\s+原(.+?)→新(.+?)\s+更換時間:\s+([^備]+?)(?:\s+備註:(.+))?$/);
      if (!m2) return;
      const [, role, arrangerName2, date2, shiftType2, oldVal2, newVal2, time2, remark2] = m2;
      const norm2 = hdrNormMap[shiftType2.trim()] || shiftType2.trim();
      if (hdrSet.size > 0 && !hdrSet.has(norm2)) return;
      const key2 = date2 + '|' + norm2;
      if (!changes[key2]) changes[key2] = [];
      changes[key2].push({
        oldVal: oldVal2, newVal: newVal2,
        empId: '', name: role + '-' + arrangerName2,
        time: time2.trim(),
        remark: remark2 ? remark2.trim() : '',
        isDragLog: true,
        isAuditLog: role === '審核'
      });
    });

    return changes;
  } catch(e) { return {}; }
}

function escapeHtml(text) {
  if (typeof text !== 'string') return text;
  return text.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
             .replace(/"/g,'&quot;').replace(/'/g,'&#039;');
}

function logError(msg) {
  try {
    const logSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    logSheet.getRange('B40').setValue(`[${new Date().toLocaleString()}] ${msg}`);
  } catch(e) {}
}

// ── 操作紀錄（M40:M460）────────────────────────────────────────────
function writeOpLog(action, detail) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    const raw   = sheet.getRange(GLOBAL_CONFIG.OP_LOG_RANGE).getValues();
    let row = 40;
    for (let i = 0; i < raw.length; i++) {
      if (!raw[i][0]) { row = i + 40; break; }
      if (i === raw.length - 1) row = 40; // 滿了就覆蓋最舊的
    }
    const tz  = SpreadsheetApp.openById(SHEET_ID).getSpreadsheetTimeZone();
    const ts  = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm:ss');
    sheet.getRange(row, 13).setValue(`[${ts}] [${action}] ${detail}`);
  } catch(e) {}
}

function getOperationLog() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    const rows  = sheet.getRange(GLOBAL_CONFIG.OP_LOG_RANGE).getValues();
    const logs  = rows.map(r => r[0] ? r[0].toString() : '').filter(v => v);
    return { success: true, logs: logs.reverse() }; // 最新在前
  } catch(e) { return { success: false, logs: [], message: e.message }; }
}

// ── 通知名單（Z欄 TRUE/FALSE）────────────────────────────────────────
function getNotifySettings() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    const names   = sheet.getRange('I1:I11').getValues().flat();
    const emails  = sheet.getRange('L1:L11').getValues().flat();
    const notify  = sheet.getRange(GLOBAL_CONFIG.NOTIFY_RANGE).getValues().flat();
    const result  = [];
    for (let i = 0; i < 11; i++) {
      if (!names[i]) continue;
      result.push({
        _row:   i,
        name:   names[i].toString().trim(),
        email:  emails[i] ? emails[i].toString().trim() : '',
        notify: notify[i] === true || notify[i] === 'TRUE' || notify[i] === 1
      });
    }
    return { success: true, staff: result };
  } catch(e) { return { success: false, message: e.message, staff: [] }; }
}

function saveNotifySettings(adminPassword, settings) {
  if (!verifyAdminPassword(adminPassword)) return { success: false, message: '管理員密碼錯誤' };
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    (settings || []).forEach(function(s) {
      const r = Number(s._row);
      if (r < 0 || r > 10) return;
      sheet.getRange('Z' + (r + 1)).setValue(s.notify ? true : false);
    });
    writeOpLog('通知設定', '更新通知名單');
    return { success: true, message: '通知設定已儲存' };
  } catch(e) { return { success: false, message: e.message }; }
}



// =============================================
// 自動排班模組
// =============================================

function getStaffList() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    const cfg = GLOBAL_CONFIG.AUTO_SCHEDULE;
    const names      = sheet.getRange(cfg.STAFF_RANGE).getValues().flat();
    const empIds     = sheet.getRange(cfg.EMPID_RANGE).getValues().flat().map(String);
    const emails     = sheet.getRange(cfg.EMAIL_RANGE).getValues().flat();
    const joinDates  = sheet.getRange(GLOBAL_CONFIG.JOIN_DATE_RANGE).getValues().flat();
    const leaveDates = sheet.getRange(GLOBAL_CONFIG.LEAVE_DATE_RANGE).getValues().flat();
    let points = [];
    try { points = sheet.getRange(cfg.POINTS_RANGE).getValues().flat(); } catch(e) {}
    const staff = [];
    for (let i = 0; i < names.length; i++) {
      if (names[i]) {
        staff.push({
          name:      names[i],
          empId:     empIds[i] || '',
          email:     emails[i] || '',
          points:    Number(points[i]) || 0,
          joinDate:  joinDates[i]  ? joinDates[i].toString().trim()  : '',
          leaveDate: leaveDates[i] ? leaveDates[i].toString().trim() : ''
        });
      }
    }
    return staff;
  } catch(e) {
    logError('getStaffList: ' + e.message);
    return [];
  }
}

function getPointsDashboard() {
  try {
    const staff = getStaffList();
    let totalDebt = 0, totalCredit = 0;
    staff.forEach(s => {
      if (s.points < 0) totalDebt += Math.abs(s.points);
      else totalCredit += s.points;
    });
    const maxCompensation = Math.ceil(totalDebt / Math.max(staff.length, 1));
    return {
      staff: staff.map(s => ({
        name: s.name, empId: s.empId, points: s.points,
        status: s.points > 0 ? 'credit' : s.points < 0 ? 'debt' : 'balanced'
      })),
      summary: { totalDebt, totalCredit, maxCompensation }
    };
  } catch(e) { return { staff: [], summary: { totalDebt:0, totalCredit:0, maxCompensation:0 } }; }
}

// ── 全年門診系列次數統計（護理師均勻看板）──────────────────────
// 統計 E~K 欄（門診、流注1/2、預登1/2、預注1/2）全年各月累積
function getYearlyClinicStats() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const now      = new Date();
    const year     = now.getFullYear();
    const rocYear  = year - 1911;
    const prefix   = rocNumToStr(rocYear) + '年';
    const allNames = spreadsheet.getSheets().map(s => s.getName());

    const monthOrder = ['一月','二月','三月','四月','五月','六月',
                        '七月','八月','九月','十月','十一月','十二月'];
    const yearSheets = allNames
      .filter(n => n !== EMAIL_SHEET_NAME && n.includes(prefix))
      .sort((a, b) => {
        const ai = monthOrder.findIndex(m => a.includes(m));
        const bi = monthOrder.findIndex(m => b.includes(m));
        return ai - bi;
      });

    // 護理師名單（K欄，門診系列候選）
    const settingSheet = spreadsheet.getSheetByName(EMAIL_SHEET_NAME);
    const nurseNames   = settingSheet.getRange('K1:K8').getValues().flat()
      .filter(n => n && n.toString().trim() !== '')
      .map(n => n.toString().trim());

    // 動態讀取職務名稱（E~L欄 = C1:M1 的 index 2~9），與班表試算表標題同步
    const DEFAULT_CLINIC_HEADERS = ['門診','流注1','流注2','預登1','預登2','預注1','預注2','卡介苗'];
    let CLINIC_HEADERS = DEFAULT_CLINIC_HEADERS.slice();
    if (yearSheets.length > 0) {
      try {
        const firstSh = spreadsheet.getSheetByName(yearSheets[0]);
        if (firstSh) {
          const shHdrs = firstSh.getRange('C1:M1').getValues()[0];
          const dynHdrs = shHdrs.slice(2, 10).map((v, i) =>
            v ? v.toString().trim() : DEFAULT_CLINIC_HEADERS[i]);
          if (dynHdrs.some(v => v)) CLINIC_HEADERS = dynHdrs;
        }
      } catch(e) { /* 讀不到就用預設 */ }
    }
    // E=門診, F=流注1, G=流注2, H=預登1, I=預登2, J=預注1, K=預注2, L=卡介苗
    // 在試算表 C2:M32 中：E=col5=idx2, ..., K=col11=idx8, L=col12=idx9

    // counts[name][ci] = 全年累積次數（ci 0-7 對應上述8個班別，含卡介苗）
    const counts = {};
    nurseNames.forEach(n => { counts[n] = new Array(8).fill(0); });

    // 月份明細
    const monthly = []; // [{label, sheetName, counts:{name:[]}}]

    // ── 週別分欄設定 ─────────────────────────────────────────────────
    // 原始 E2:L 8欄順序：門診(0)/掛號(1)/前台(2)/預登1(3)/預登2注(4)/注射1(5)/注射2(6)/卡介苗(7)
    // 混合日欄（週二+週四）：idx 0,3,4,5 → 拆成 _二 / _四 兩子欄
    // 週四專屬欄：idx 1,2（掛號/前台）→ 不拆
    // 週二專屬欄：idx 6（注射2）→ 不拆
    // 卡介苗：idx 7 → 不拆
    // 輸出 12 欄：[門診(二),門診(四),掛號,前台,預登1(二),預登1(四),預登2注(二),預登2注(四),注射1(二),注射1(四),注射2,卡介苗]
    const MIX_IDX = new Set([0, 3, 4, 5]); // 需要拆分的原始索引
    const wkMap = {'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'日':0};

    // 建立 12 欄的 counts 結構
    const wdCounts = {};
    nurseNames.forEach(n => { wdCounts[n] = new Array(12).fill(0); });

    // 原始欄 idx → 輸出欄 idx 的映射（不分日）
    // 混合欄 0→[0,1]  單欄 1→2  2→3  混合欄 3→[4,5]  4→[6,7]  5→[8,9]  6→10  7→11
    function getOutIdx(srcIdx, dow) {
      if (srcIdx === 0) return dow === 2 ? 0 : (dow === 4 ? 1 : 0); // 門診
      if (srcIdx === 1) return 2;   // 掛號（週四）
      if (srcIdx === 2) return 3;   // 前台（週四）
      if (srcIdx === 3) return dow === 2 ? 4 : (dow === 4 ? 5 : 4); // 預登1
      if (srcIdx === 4) return dow === 2 ? 6 : (dow === 4 ? 7 : 6); // 預登2注
      if (srcIdx === 5) return dow === 2 ? 8 : (dow === 4 ? 9 : 8); // 注射1
      if (srcIdx === 6) return 10;  // 注射2（週二）
      if (srcIdx === 7) return 11;  // 卡介苗
      return 0;
    }

    yearSheets.forEach(sName => {
      const sh = spreadsheet.getSheetByName(sName);
      if (!sh) return;
      const lr = sh.getLastRow();
      if (lr < 2) return;
      const maxRow = Math.min(lr, 32);
      // E到L欄（含卡介苗）= 門診~卡介苗；同時讀 B 欄星期字串
      const data    = sh.getRange('E2:L' + maxRow).getValues();
      const wkData  = sh.getRange('B2:B' + maxRow).getValues();

      const mCounts = {};
      nurseNames.forEach(n => { mCounts[n] = new Array(8).fill(0); });
      const mWdCounts = {};
      nurseNames.forEach(n => { mWdCounts[n] = new Array(12).fill(0); });

      data.forEach((row, ri) => {
        // 解析星期
        const wb  = wkData[ri][0] ? wkData[ri][0].toString() : '';
        const wm  = wb.match(/週([一二三四五六日])/);
        const dow = wm ? (wkMap[wm[1]] !== undefined ? wkMap[wm[1]] : -1) : -1;

        for (let ci = 0; ci < 8; ci++) {
          const val = row[ci] ? row[ci].toString().trim() : '';
          if (!val || !counts.hasOwnProperty(val)) continue;
          counts[val][ci]++;
          mCounts[val][ci]++;
          const oi = getOutIdx(ci, dow);
          wdCounts[val][oi]++;
          mWdCounts[val][oi]++;
        }
      });

      const label = sName.replace(/一百[一二三四五六七八九十]+年/,'').replace('班表','').trim();
      monthly.push({ label, sheetName: sName, mCounts, mWdCounts });
    });

    // ── 讀取排班排除清單（請長假等），判斷目前仍可排班的人 ──────────
    const tz = spreadsheet.getSpreadsheetTimeZone();
    const exclRaw = settingSheet.getRange(GLOBAL_CONFIG.EXCLUSION_RANGE).getValues();
    function fmtD2(v) {
      if (!v) return '';
      if (v instanceof Date) return Utilities.formatDate(v, tz, 'M/d');
      return v.toString().trim();
    }
    const exclusions = exclRaw
      .filter(r => r[0] && r[0].toString().trim())
      .map(r => ({
        name:     r[0].toString().trim(),
        startMD:  fmtD2(r[1]),
        endMD:    fmtD2(r[2]),
        weekdays: r[3] ? r[3].toString().trim() : ''
      }));

    // 判斷某人「目前」是否在請假排除期間（不指定特定日期，只看日期範圍是否覆蓋今日）
    const today = new Date();
    const todayM = today.getMonth() + 1;
    const todayD = today.getDate();
    function isCurrentlyExcluded(name) {
      return exclusions.some(ex => {
        if (ex.name !== name) return false;
        let startM = 0, startD = 0;
        if (ex.startMD) { const m = ex.startMD.match(/(\d+)\/(\d+)/); if(m){startM=parseInt(m[1]);startD=parseInt(m[2]);} }
        let endM = 99, endD = 99;
        if (ex.endMD) { const m = ex.endMD.match(/(\d+)\/(\d+)/); if(m){endM=parseInt(m[1]);endD=parseInt(m[2]);} }
        const cur = todayM * 100 + todayD;
        const start = startM * 100 + startD;
        const end   = endM * 100 + endD;
        if (cur < start || cur > end) return false;
        // 有 weekdays 代表是條件式排除（如只排週末），不算「完全暫停」
        if (ex.weekdays && ex.weekdays.trim() !== '') return false;
        return true;
      });
    }

    // ── 計算「已排班月份數」作為分母（而非整年12個月）──────────────
    const scheduledMonthCount = yearSheets.length || 1;

    // 讀取到職/離職日（X/Y欄）
    const joinDatesRaw  = settingSheet.getRange(GLOBAL_CONFIG.JOIN_DATE_RANGE).getValues().flat();
    const leaveDatesRaw = settingSheet.getRange(GLOBAL_CONFIG.LEAVE_DATE_RANGE).getValues().flat();
    const staffNames    = settingSheet.getRange('I1:I11').getValues().flat().map(n=>n?n.toString().trim():'');

    function parseMD(v) {
      if (!v) return null;
      if (v instanceof Date) return { m: v.getMonth()+1, d: v.getDate() };
      const m = v.toString().match(/(\d+)\/(\d+)/);
      return m ? { m: parseInt(m[1]), d: parseInt(m[2]) } : null;
    }

    // 取得已排班月份的月份號碼列表（e.g. [1,2,3,4]）
    const scheduledMonthNums = yearSheets.map(sn => {
      const p = parseYearMonthFromSheetName(sn);
      return p.valid ? p.month : 0;
    }).filter(m => m > 0);

    // 各人在已排班月份中的在職月數
    const activeRatio = {};
    const activeMonths = {}; // 實際在職的月份數
    const inferredJoinMonth = {}; // 自動推算的到職月

    // 先掃描所有已排班月份，找出每人最早出現的月份
    const firstAppearance = {}; // name → 最早出現的月份號
    yearSheets.forEach(sName => {
      const sh = spreadsheet.getSheetByName(sName);
      if (!sh) return;
      const p = parseYearMonthFromSheetName(sName);
      if (!p.valid) return;
      const data = sh.getRange('C2:M32').getValues();
      data.forEach(row => {
        row.forEach(cell => {
          const name = cell ? cell.toString().trim() : '';
          if (name && nurseNames.includes(name)) {
            if (!firstAppearance[name] || p.month < firstAppearance[name]) {
              firstAppearance[name] = p.month;
            }
          }
        });
      });
    });

    nurseNames.forEach(name => {
      const idx = staffNames.indexOf(name);
      const jd = idx !== -1 ? parseMD(joinDatesRaw[idx])  : null;
      const ld = idx !== -1 ? parseMD(leaveDatesRaw[idx]) : null;

      // 到職月：優先用 X 欄明確設定，否則用首次出現月，再否則用 1 月
      const joinM  = jd ? jd.m : (firstAppearance[name] || 1);
      const leaveM = ld ? ld.m : 12;
      inferredJoinMonth[name] = joinM;

      // 在已排班月份中，有多少月份此人在職
      const inMonths = scheduledMonthNums.filter(m => m >= joinM && m <= leaveM).length;
      activeMonths[name] = inMonths;
      activeRatio[name] = scheduledMonthCount > 0
        ? Math.max(0, Math.min(1, inMonths / scheduledMonthCount))
        : 1.0;
    });

    // 離職後排除：若 leaveMonth < 最後已排班月，不加入優先排序比較
    const lastScheduledMonth = scheduledMonthNums.length > 0 ? Math.max(...scheduledMonthNums) : 12;

    // 只比較「目前仍在職」的人（leaveM >= lastScheduledMonth）
    const activeNurses = nurseNames.filter(name => {
      const idx = staffNames.indexOf(name);
      const ld = idx !== -1 ? parseMD(leaveDatesRaw[idx]) : null;
      const leaveM = ld ? ld.m : 12;
      return leaveM >= lastScheduledMonth;
    });

    // 每月平均次數（用全年在職人員的總次數 / 已排班月數，再除以人數）
    // = 每人每月平均排幾次門診系列
    const fullYearNurses = activeNurses.filter(n => activeRatio[n] >= 1.0);
    const refNurses = fullYearNurses.length > 0 ? fullYearNurses : activeNurses;
    const avgPerMonth = (scheduledMonthCount > 0 && refNurses.length > 0)
      ? refNurses.reduce((s,n) => s + counts[n].reduce((a,b)=>a+b,0), 0)
        / refNurses.length / scheduledMonthCount
      : 0;

    const rows = nurseNames.map(name => {
      const total    = counts[name].reduce((a, b) => a + b, 0);
      const ratio    = activeRatio[name];
      const inMonths = activeMonths[name];
      const joinM    = inferredJoinMonth[name] || 1;
      const expected = Math.round(avgPerMonth * inMonths);
      const idx = staffNames.indexOf(name);
      const jd = idx !== -1 ? parseMD(joinDatesRaw[idx]) : null;
      const ld = idx !== -1 ? parseMD(leaveDatesRaw[idx]) : null;
      const leaveM = ld ? ld.m : 12;
      // isResigned = 已設定離職日且離職月 < 最後排班月（本年度已結算，不計入排名）
      const isResigned = !!(ld && leaveM < lastScheduledMonth);
      const isActive   = !isResigned;
      const onLeave    = isCurrentlyExcluded(name);
      const joinDateSet = !!jd;
      return { name, counts: counts[name], total, expected, ratio, inMonths, isActive, isResigned, onLeave, joinM, joinDateSet };
    }).sort((a, b) => {
      // 離職人員排到最後，不參與優先比較
      if (a.isActive !== b.isActive) return a.isActive ? -1 : 1;
      return a.total - b.total; // 在職者少者優先
    });

    const passedMonths = scheduledMonthCount; // 已排班月份數

    // 12 欄標頭（週別分欄後的名稱）
    const h = CLINIC_HEADERS;
    const WD_HEADERS = [
      (h[0]||'門診')+'(二)',  (h[0]||'門診')+'(四)',
      h[1]||'掛號',            h[2]||'前台',
      (h[3]||'預登1')+'(二)', (h[3]||'預登1')+'(四)',
      (h[4]||'預登2注')+'(二)',(h[4]||'預登2注')+'(四)',
      (h[5]||'注射1')+'(二)', (h[5]||'注射1')+'(四)',
      h[6]||'注射2',           h[7]||'卡介苗'
    ];

    // 將 wdCounts 轉為 rows 相同格式供前端使用
    const wdRows = rows.map(row => ({
      ...row,
      wdCounts: wdCounts[row.name] || new Array(12).fill(0),
      wdTotal:  (wdCounts[row.name] || []).reduce((s,v)=>s+v,0)
    }));

    return {
      success:  true,
      headers:  CLINIC_HEADERS,
      wdHeaders: WD_HEADERS,
      rows,
      wdRows,
      monthly,
      nurses:   nurseNames,
      sheets:   yearSheets,
      passedMonths
    };
  } catch(e) {
    return { success: false, message: e.message, headers:[], rows:[], monthly:[], nurses:[], sheets:[] };
  }
}

// ── 一鍵排班：建立工作表 + 自動排班（單月）──────────────────────
// mode: 'preview' | 'execute'
// month: sheetName（工作表名稱）
// targetYear: 西元年（如 2026），預設用當年
function quickSchedule(adminPassword, mode, scope, month, targetYear) {
  if (!verifyAdminPassword(adminPassword)) return { success: false, message: '管理員密碼錯誤。' };
  // BUG 9 修正：整年排班已移除，scope 僅接受 'month'
  if (scope && scope !== 'month') {
    return { success: false, message: '整年排班功能已停用，請逐月排班。' };
  }
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    {
      // 單月：若工作表不存在先建立
      if (!spreadsheet.getSheetByName(month)) {
        const p = parseYearMonthFromSheetName(month);
        if (p.valid) createScheduleSheet(adminPassword, p.year, p.month, month);
      }
      if (mode === 'preview') {
        const r = previewAutoSchedule(month, adminPassword);
        // 判斷此月是否有排班資料
        const sh = spreadsheet.getSheetByName(month);
        let hasData = false;
        if (sh) {
          const vals = sh.getRange('C2:M10').getValues().flat();
          hasData = vals.some(function(v){ return v && v.toString().trim(); });
        }
        r.hasData = hasData;
        return r;
      } else {
        return autoSchedule(month, adminPassword, { overwrite: true, sendNotify: false });
      }
    }
  } catch(e) {
    return { success: false, message: '一鍵排班失敗：' + e.message };
  }
}

// 取得所有115年班表名稱（含已存在 + 未來可建立的），供單月選單用
function getAllPossibleSheets(targetYear) {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const useYear  = targetYear ? Number(targetYear) : new Date().getFullYear();
  const rocY     = useYear - 1911;
  const rocPfx   = rocNumToStr(rocY);
  const rocMonthNames = ['','一','二','三','四','五','六','七','八','九','十','十一','十二'];
  const result = [];
  for (let m = 1; m <= 12; m++) {
    const name   = rocPfx + '年' + rocMonthNames[m] + '月班表';
    const sh     = spreadsheet.getSheetByName(name);
    const exists = !!sh;
    let reviewStatus = '', schedTime = '', writeCount = 0;
    if (sh) {
      try {
        const note = sh.getRange('N1').getNote() || '';
        note.split('\n').forEach(line => {
          if (line.startsWith('審核狀態:'))  reviewStatus = line.replace('審核狀態:', '').trim();
          if (line.startsWith('排定時間:'))  schedTime    = line.replace('排定時間:', '').trim();
          if (line.startsWith('writeCount:')) writeCount  = parseInt(line.replace('writeCount:', '').trim()) || 0;
        });
      } catch(e) {}
    }
    const wcLetter = writeCount > 0 ? wcToLetter(writeCount) : '';
    result.push({ name, exists, month: m, year: useYear, reviewStatus, schedTime, wcLetter });
  }
  return result;
}
function getSystemSettings() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    const tz     = SpreadsheetApp.openById(SHEET_ID).getSpreadsheetTimeZone();
    const names  = sheet.getRange('I1:I11').getValues().flat();
    const empIds = sheet.getRange('M1:M11').getValues().flat();
    const emails = sheet.getRange('L1:L11').getValues().flat();
    const joinDates  = sheet.getRange(GLOBAL_CONFIG.JOIN_DATE_RANGE).getValues().flat();
    const leaveDates = sheet.getRange(GLOBAL_CONFIG.LEAVE_DATE_RANGE).getValues().flat();
    const pointsRaw  = sheet.getRange('P1:P11').getValues().flat();
    const kkNames    = sheet.getRange('J1:J3').getValues().flat();
    const clinicNames= sheet.getRange('K1:K8').getValues().flat();
    const bcgNames   = sheet.getRange('Q1:Q2').getValues().flat();
    // 值班順序（E欄）、登革熱順序（B欄）
    const dutyNames   = sheet.getRange('E1:E11').getValues().flat();
    const dengueNames = sheet.getRange('B1:B11').getValues().flat();

    function fmtD(v){ if(!v) return ''; if(v instanceof Date) return Utilities.formatDate(v,tz,'M/d'); return v.toString().trim(); }

    const staff = [];
    for (let i = 0; i < 11; i++) {
      if (names[i]) {
        staff.push({
          _row:      i,
          name:      names[i].toString().trim(),
          empId:     empIds[i] ? empIds[i].toString().trim() : '',
          email:     emails[i] ? emails[i].toString().trim() : '',
          joinDate:  fmtD(joinDates[i]),
          leaveDate: fmtD(leaveDates[i]),
          points:    Number(pointsRaw[i]) || 0
        });
      }
    }
    return {
      success: true, staff,
      kkNames:     kkNames.filter(n=>n).map(n=>n.toString().trim()),
      clinicNames: clinicNames.filter(n=>n).map(n=>n.toString().trim()),
      bcgNames:    bcgNames.filter(n=>n).map(n=>n.toString().trim()),
      dutyNames:   dutyNames.filter(n=>n).map(n=>n.toString().trim()),
      dengueNames: dengueNames.filter(n=>n).map(n=>n.toString().trim())
    };
  } catch(e) { return { success: false, message: e.message }; }
}

function saveSystemSettings(adminPassword, settings) {
  if (!verifyAdminPassword(adminPassword)) return { success: false, message: '管理員密碼錯誤。' };
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    if (settings.staff) {
      // 用 _row 精準定位，避免陣列索引因空格而錯位
      settings.staff.forEach(function(s) {
        const r = (s._row !== undefined ? Number(s._row) : -1);
        if (r < 0 || r > 10) return;
        const row = r + 1;
        sheet.getRange('I'+row).setValue(s.name      || '');
        sheet.getRange('M'+row).setValue(s.empId     || '');
        sheet.getRange('L'+row).setValue(s.email     || '');
        sheet.getRange('X'+row).setValue(s.joinDate  || '');
        sheet.getRange('Y'+row).setValue(s.leaveDate || '');
        if (s.points !== undefined && s.points !== '') {
          sheet.getRange('P'+row).setValue(Number(s.points) || 0);
        }
      });
    }
    if (settings.kkNames)     for (let i=0;i<3;i++)  sheet.getRange('J'+(i+1)).setValue(settings.kkNames[i]||'');
    if (settings.clinicNames)  for (let i=0;i<8;i++)  sheet.getRange('K'+(i+1)).setValue(settings.clinicNames[i]||'');
    if (settings.bcgNames)    for (let i=0;i<2;i++)  sheet.getRange('Q'+(i+1)).setValue(settings.bcgNames[i]||'');
    if (settings.dutyNames)   for (let i=0;i<11;i++) sheet.getRange('E'+(i+1)).setValue(settings.dutyNames[i]||'');
    if (settings.dengueNames) for (let i=0;i<11;i++) sheet.getRange('B'+(i+1)).setValue(settings.dengueNames[i]||'');
    if (settings.newAdminPwd && settings.newAdminPwd.trim()) sheet.getRange('N2').setValue(settings.newAdminPwd.trim());
    writeOpLog('系統設定', '儲存系統設定');
    return { success: true, message: '系統設定已儲存！' };
  } catch(e) { return { success: false, message: '儲存失敗：'+e.message }; }
}

// ── 接班人設定（一鍵換人，自動同步所有名單）────────────────────────
function setSuccessor(adminPassword, rowIndex, newStaff) {
  if (!verifyAdminPassword(adminPassword)) return { success: false, message: '管理員密碼錯誤' };
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(EMAIL_SHEET_NAME);
    const tz    = ss.getSpreadsheetTimeZone();
    if (rowIndex < 0 || rowIndex > 10) return { success: false, message: '無效的崗位列號' };
    const row     = rowIndex + 1;
    const oldName = (sheet.getRange('I'+row).getValue() || '').toString().trim();

    // ── 計算在職人員點數平均，作為新人初始點數 ──────────────────
    let finalPoints = 0;
    if (newStaff.points !== undefined && newStaff.points !== '') {
      finalPoints = Number(newStaff.points) || 0;
    } else {
      const allNames   = sheet.getRange('I1:I11').getValues().flat();
      const allPoints  = sheet.getRange('P1:P11').getValues().flat();
      const leaveDates = sheet.getRange('Y1:Y11').getValues().flat();
      const now        = new Date();
      const pool = [];
      for (let i = 0; i < 11; i++) {
        if (i === rowIndex || !allNames[i]) continue;
        const ld = leaveDates[i];
        if (ld) {
          const d = ld instanceof Date ? ld : new Date(ld);
          if (!isNaN(d) && d < now) continue;
        }
        pool.push(Number(allPoints[i]) || 0);
      }
      finalPoints = pool.length
        ? Math.round(pool.reduce(function(a,b){return a+b;},0) / pool.length)
        : 0;
    }

    // ── 寫入新人主要欄位 ─────────────────────────────────────────
    sheet.getRange('I'+row).setValue(newStaff.name     || '');
    sheet.getRange('M'+row).setValue(newStaff.empId    || '');
    sheet.getRange('L'+row).setValue(newStaff.email    || '');
    sheet.getRange('X'+row).setValue(newStaff.joinDate || '');
    sheet.getRange('Y'+row).setValue('');            // 清離職日
    sheet.getRange('P'+row).setValue(finalPoints);   // 初始點數

    // ── 各候選名單自動將舊名→新名 ────────────────────────────────
    if (oldName && newStaff.name && oldName !== newStaff.name) {
      ['K1:K8','J1:J3','E1:E11','Q1:Q2','B1:B11'].forEach(function(r) {
        replaceInRange_(sheet, r, oldName, newStaff.name);
      });
    }

    // ── 清除排除清單中舊人的所有條目 ─────────────────────────────
    if (oldName) {
      const excl = sheet.getRange(GLOBAL_CONFIG.EXCLUSION_RANGE).getValues();
      excl.forEach(function(r, i) {
        if (r[0] && r[0].toString().trim() === oldName) {
          sheet.getRange('T'+(i+1)+':W'+(i+1)).clearContent();
        }
      });
    }

    const msg = oldName
      ? oldName + ' → ' + newStaff.name + ' 接班完成，初始點數 ' + finalPoints
      : newStaff.name + ' 新增完成，初始點數 ' + finalPoints;
    writeOpLog('接班設定', msg);
    return { success: true, message: msg, avgPoints: finalPoints };
  } catch(e) {
    return { success: false, message: '接班設定失敗：' + e.message };
  }
}

function replaceInRange_(sheet, rangeA1, oldName, newName) {
  try {
    const rng  = sheet.getRange(rangeA1);
    const data = rng.getValues();
    let changed = false;
    data.forEach(function(row) {
      for (let j = 0; j < row.length; j++) {
        if (row[j] && row[j].toString().trim() === oldName) {
          row[j] = newName;
          changed = true;
        }
      }
    });
    if (changed) rng.setValues(data);
  } catch(e) { /* 欄位不存在時略過 */ }
}


function resetStaffPoints(name, adminPassword) {
  if (!verifyAdminPassword(adminPassword)) return '管理員密碼錯誤，無法執行清算。';
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    const names = sheet.getRange(GLOBAL_CONFIG.AUTO_SCHEDULE.STAFF_RANGE).getValues().flat();
    const idx = names.indexOf(name);
    if (idx === -1) return `找不到人員：${name}`;
    const oldPoints = Number(sheet.getRange(`P${idx+1}`).getValue()) || 0;
    sheet.getRange(`P${idx+1}`).setValue(0);
    const logRange = sheet.getRange('A40:A460').getValues();
    let logRow = 40;
    for (let i = 0; i < logRange.length; i++) {
      if (logRange[i][0] === "") { logRow = i + 40; break; }
    }
    sheet.getRange(logRow, 1).setValue(
      `[離職清算] ${name} 點數歸零（原為 ${oldPoints}） 清算時間: ${new Date().toLocaleString()}`
    );
    return `${name} 點數清算完成（原為 ${oldPoints}，已歸零）。`;
  } catch(e) { return '清算失敗：' + e.message; }
}

// =============================================
// 核心排班規則
// 班別欄位（colIdx 0-10，對應 C-M）：
//   0=值班(C)      → 每日（含假日），積分作為 tiebreaker，辦公日/假日分開輪
//   1=協助掛號(D)  → 只有週四工作日，純次數輪，不用積分
//   2=門診(E)      → 週二＋週四工作日，週二/週四分開輪序，每人每月同欄≤2次
//   3=流注1(F)     → 只有週四工作日，純次數輪，每人每月同欄≤2次
//   4=流注2(G)     → 只有週四工作日，純次數輪，每人每月同欄≤2次
//   5=預登1(H)     → 週二＋週四工作日，週二/週四分開輪序，每人每月同欄≤2次
//   6=預登2(I)     → 週二＋週四工作日，週二/週四分開輪序，每人每月同欄≤2次
//   7=預注1(J)     → 週二＋週四工作日，週二/週四分開輪序，每人每月同欄≤2次
//   8=預注2(K)     → 只有週二工作日，純次數輪，每人每月同欄≤2次
//   9=卡介苗(L)    → 只有每月第一個週二工作日，純次數輪
//  10=登革熱二線(M)→ 只有假日（週六日＋補假），純次數輪
// =============================================

// ★ 執行時讀取動態規則（每次 runAutoSchedule 開始時重設）
let _shiftDayRulesCache = null;

function shouldAssignShift(date, colIdx, year, month) {
  const holiday = isHoliday(date);
  const dow = date.getDay(); // 0=日,1=一,2=二,3=三,4=四,5=五,6=六

  // 值班/停班2線固定
  if (colIdx === 0 || colIdx === 10) return true;
  // 卡介苗固定（每月第一個週二工作日）
  if (colIdx === 9) {
    if (holiday || dow !== 2) return false;
    return date.getDate() === getFirstTuesdayWorkday(year, month);
  }
  if (holiday) return false;

  // ★ 優先使用動態規則快取
  if (_shiftDayRulesCache && _shiftDayRulesCache[colIdx]) {
    return _shiftDayRulesCache[colIdx].includes(dow);
  }

  // Fallback 硬編碼預設
  switch(colIdx) {
    case 1: return dow === 4;                   // 支援：週四
    case 2: return dow === 2 || dow === 4;      // 門診：週二＋週四
    case 3: case 4: return dow === 4;           // 掛號/前台：週四
    case 5: case 6: case 7: return dow === 2 || dow === 4; // 預登/注射1：週二＋週四
    case 8: return dow === 2;                   // 注射2：週二
    default: return false;
  }
}

function parseYearMonthFromSheetName(sheetName) {
  // BUG20 修正：改為通用解析，不再硬編碼 114/115 年
  // 支援所有民國 100~199 年（西元 2011~2110）
  const rocMonthMap = {
    '一':1,'二':2,'三':3,'四':4,'五':5,'六':6,
    '七':7,'八':8,'九':9,'十':10,'十一':11,'十二':12
  };

  // 動態匹配所有民國百位年份：「一百X十Y年」
  // 例如「一百一十五年四月班表」→ 民國 115 → 西元 2026
  const rocMatch = sheetName.match(/(一百[一二三四五六七八九]?十?[一二三四五六七八九]?)年([一二三四五六七八九十]+)月/);
  if (rocMatch) {
    const rocYearStr = rocMatch[1];
    const mStr = rocMatch[2];
    const month = rocMonthMap[mStr];
    if (month) {
      // 反向解析民國漢字年份為數字
      const rocYear = rocStrToNum(rocYearStr);
      if (rocYear > 0) {
        return { year: rocYear + 1911, month, valid: true };
      }
    }
  }

  // Fallback：數字格式（2026年4月）
  const numMatch = sheetName.match(/(\d{4})年(\d{1,2})月/);
  if (numMatch) return { year: parseInt(numMatch[1]), month: parseInt(numMatch[2]), valid: true };
  return { year: new Date().getFullYear(), month: 1, valid: false };
}

// 民國漢字 → 數字（一百一十五 → 115）
function rocStrToNum(str) {
  if (!str) return 0;
  const unitMap = {'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9};
  // 格式：一百 + [X十] + [Y]
  if (!str.startsWith('一百')) return 0;
  const rem = str.substring(2); // 去掉「一百」
  if (!rem) return 100;
  let tens = 0, units = 0;
  const tenMatch = rem.match(/([一二三四五六七八九])十/);
  if (tenMatch) {
    tens = unitMap[tenMatch[1]] || 0;
  } else if (rem.includes('十')) {
    tens = 1; // 「十」=10（不帶前綴）
  }
  // 個位：十後面的字，或沒有十時直接的字
  const afterTen = rem.split('十');
  const lastPart = afterTen.length > 1 ? afterTen[afterTen.length - 1] : (tens === 0 ? rem : '');
  if (lastPart && unitMap[lastPart]) units = unitMap[lastPart];
  return 100 + tens * 10 + units;
}

function getShiftStaffMap() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
  const result = {};
  const cols = ['C','D','E','F','G','H','I','J','K','L','M'];
  cols.forEach(col => {
    const rangeKey = GLOBAL_CONFIG.SHIFT_OPTIONS[col];
    if (rangeKey) {
      result[col] = sheet.getRange(rangeKey).getValues().flat().filter(n => n);
    } else {
      result[col] = [];
    }
  });
  return result;
}

// =============================================
// 自動排班核心 ver4.4
//
// ★ 優先階層（最高，三者可互相重疊，不互相干預）：
//   值班(0)  協助掛號(1)  卡介苗(9)
//
// ★ 排除名單：直接向後遞補，不留空
//
// ★ 當月卡介苗人員：整月不排門診相關職務（colIdx 2-8），值班/協助掛號不受限
//
// ★ 門診、流注、預登、預注(colIdx 2-8)：整月均勻（次數少者優先）
//
// ★ 登革熱二線：衝突假日值班 → 與下一位 swap；再衝突再 swap，以此類推
// =============================================
function runAutoSchedule(sheetName, adminPassword, options) {
  if (!verifyAdminPassword(adminPassword)) return { success: false, message: '管理員密碼錯誤。' };

  // ★ 載入動態排班日規則快取（每次排班重新讀取）
  try {
    const rulesRes = getShiftDayRules();
    _shiftDayRulesCache = (rulesRes.rules && typeof rulesRes.rules === 'object') ? rulesRes.rules : null;
    if (_shiftDayRulesCache) {
      // 轉換 keys 為數字（JSON 序列化後 key 為字串）
      const numericCache = {};
      for (const k in _shiftDayRulesCache) { numericCache[parseInt(k)] = _shiftDayRulesCache[k]; }
      _shiftDayRulesCache = numericCache;
    }
  } catch(e) { _shiftDayRulesCache = null; }

  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const sheet       = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: '找不到工作表：' + sheetName };

    const { year, month } = parseYearMonthFromSheetName(sheetName);
    const staffAll = getStaffList();
    // ★ 只保留本月在職人員（過濾已離職 & 尚未到職）
    const staff = staffAll.filter(s => isStaffActiveForMonth(s, year, month));
    if (staff.length === 0) return { success: false, message: '無法取得人員清單。' };
    staff.forEach((s, i) => { s.staffIdx = i; });

    const exclusions    = getExclusionList();
    // ★ 讀取班別標題 C1:M1（colIdx 0-10），供注射欄判斷使用
    const shiftHdrs = sheet.getRange('C1:M1').getValues()[0].map(v => v ? v.toString().trim() : '');
    const shiftStaffMap = getShiftStaffMap();
    const cols = ['C','D','E','F','G','H','I','J','K','L','M'];

    const datesRange   = sheet.getRange('A2:B32').getValues();
    const existingData = sheet.getRange('C2:M32').getValues();
    const headers      = sheet.getRange('C1:M1').getValues()[0];
    const tz           = spreadsheet.getSpreadsheetTimeZone();

    // ── 預先解析日期 ────────────────────────────────────────────────
    const dateObjByRow  = [];
    const combinedDates = [];
    for (let r = 0; r < datesRange.length; r++) {
      const raw  = datesRange[r][0];
      const aVal = raw instanceof Date
        ? Utilities.formatDate(raw, tz, 'M/d')
        : (raw ? raw.toString().trim() : '');
      const bVal = datesRange[r][1] ? datesRange[r][1].toString().trim() : '';
      combinedDates.push((aVal + ' ' + bVal).trim());
      let d = null;
      if (raw instanceof Date) {
        d = raw;
      } else {
        const m = aVal.match(/(\d+)\/(\d+)/);
        if (m) { try { d = new Date(year, parseInt(m[1])-1, parseInt(m[2])); } catch(e){} }
      }
      dateObjByRow.push(d);
    }

    // ── 通用：從指針位置找下一個未被排除的人（遞補邏輯）────────────
    // 傳入 candidateStaff（已過濾的陣列）、ptr、dateObj
    // 回傳 { name, nextPtr }；若全部排除回傳 { name:'', nextPtr: ptr+1 }
    function pickFromPtr(candidateStaff, ptr, d) {
      const n = candidateStaff.length;
      if (n === 0) return { name: '', nextPtr: 0 };
      for (let attempt = 0; attempt < n; attempt++) {
        const idx    = (ptr + attempt) % n;
        const person = candidateStaff[idx];
        if (!isPersonExcluded(person.name, d, exclusions, year)) {
          return { name: person.name, nextPtr: (idx + 1) % n };
        }
      }
      // 全部排除 → 留空，ptr 前進一格
      return { name: '', nextPtr: (ptr + 1) % n };
    }

    // ════════════════════════════════════════════════════════════════
    // 值班：指針輪轉，平日/假日各自獨立指針，排除直接遞補
    // ★ 跨月接續：先從上個月的末尾找最後一位，再從本月已有資料還原
    // ★ 使用「原始排班人」（換班 log 反推），不用換班後的人
    // ════════════════════════════════════════════════════════════════
    const allDutyNames  = shiftStaffMap['C'] || [];
    // ★ 值班用 E 欄順位（dutyIdx），不用 I 欄的 staffIdx
    const allDutyStaff  = allDutyNames
      .filter(n => n)
      .map((name, dutyIdx) => {
        const s = staff.find(x => x.name === name);
        return s ? { ...s, dutyIdx } : null;
      })
      .filter(Boolean);
    let workdayPtr = 0;
    let holidayPtr = 0;

    // ── 讀取上一個月換班紀錄，建立「換班後→原始」還原 map ──────────
    function buildOriginalMap(prevSheetName) {
      // 回傳 { 'date|colIdx': originalName }
      const origMap = {};
      try {
        const settingSheet = spreadsheet.getSheetByName(EMAIL_SHEET_NAME);
        const logs = settingSheet.getRange('A40:A460').getValues()
          .flat().filter(r => r && r.toString().trim() !== '');
        const sh = spreadsheet.getSheetByName(prevSheetName);
        if (!sh) return origMap;
        const hdrs = sh.getRange('C1:M1').getValues()[0].map(h => h.toString().trim());

        logs.forEach(log => {
          const s = log.toString().trim();

          // ── 格式1：一般換班日誌 ──
          const m1 = s.match(/^(\d+\/\d+)\s+週.\s+(.+?)\s+原(.+?)→新(.+?)\s+(?:\([^)]+\)\s+)?更換時間:/);
          if (m1) {
            const [, date, shiftType, oldVal] = m1;
            const ci = hdrs.indexOf(shiftType);
            if (ci !== -1) {
              const key = date + '|' + ci;
              if (!origMap[key]) origMap[key] = oldVal;
            }
            return;
          }

          // ── 格式2：拖曳日誌（排班-/審核-）── 同 BUG 15 修正
          const m2 = s.match(/^(?:排班|審核)-.+?\s+(\d+\/\d+)\s+(.+?)\s+原(.+?)→新(.+?)\s+更換時間:/);
          if (m2) {
            const [, date, shiftType, oldVal] = m2;
            const ci = hdrs.indexOf(shiftType);
            if (ci !== -1) {
              const key = date + '|' + ci;
              if (!origMap[key]) origMap[key] = oldVal;
            }
          }
        });
      } catch(e) {}
      return origMap;
    }

    // ── 從上個月最後一格找接續位置（使用原始排班人）─────────────────
    function getPrevMonthLastPtr(colIdx, staffArr) {
      if (staffArr.length === 0) return { workday: 0, holiday: 0 };
      const prevMonth = month === 1 ? 12 : month - 1;
      const prevYear  = month === 1 ? year - 1 : year;
      const rocMonths = ['','一','二','三','四','五','六','七','八','九','十','十一','十二'];
      const prevSheetName = rocNumToStr(prevYear - 1911) + '年' + rocMonths[prevMonth] + '月班表';
      try {
        const prevSheet = spreadsheet.getSheetByName(prevSheetName);
        if (!prevSheet) return { workday: 0, holiday: 0 };

        const origMap = buildOriginalMap(prevSheetName);
        const prevData = prevSheet.getRange('A2:M32').getValues();
        let lastWorkdayName = '', lastHolidayName = '';

        for (let r = 0; r < prevData.length; r++) {
          const raw = prevData[r][0];
          if (!raw) continue;
          let pd = null;
          if (raw instanceof Date) {
            pd = raw;
          } else {
            const aStr = raw.toString().trim();
            const mm = aStr.match(/(\d+)\/(\d+)/);
            if (mm) { try { pd = new Date(prevYear, parseInt(mm[1])-1, parseInt(mm[2])); } catch(e){} }
          }
          if (!pd) continue;

          // 取當日日期字串（M/D 格式），用於查 origMap
          const tz = spreadsheet.getSpreadsheetTimeZone();
          const dateStr = Utilities.formatDate(pd, tz, 'M/d');
          const mapKey  = dateStr + '|' + colIdx;

          // 優先用原始排班人，若無換班記錄則用當前格子的值
          // ★ A2:M32 中，A=0(日期), B=1(星期), C=2(值班)...故需 +2 對齊班別欄
          const cellVal = prevData[r][colIdx + 2] ? prevData[r][colIdx + 2].toString().trim() : '';
          const name    = origMap[mapKey] || cellVal;
          if (!name) continue;

          if (isHoliday(pd)) lastHolidayName = name;
          else               lastWorkdayName = name;
        }

        const wIdx = staffArr.findIndex(s => s.name === lastWorkdayName);
        const hIdx = staffArr.findIndex(s => s.name === lastHolidayName);
        return {
          workday: wIdx !== -1 ? (wIdx + 1) % staffArr.length : 0,
          holiday: hIdx !== -1 ? (hIdx + 1) % staffArr.length : 0
        };
      } catch(e) { return { workday: 0, holiday: 0 }; }
    }

    // 先從上個月接續
    const dutyCarry = getPrevMonthLastPtr(0, allDutyStaff);
    workdayPtr = dutyCarry.workday || 0;
    holidayPtr = dutyCarry.holiday || 0;

    // 再用本月已有資料覆蓋（不覆蓋模式）
    if (!options.overwrite) {
      for (let r = 0; r < datesRange.length; r++) {
        const d = dateObjByRow[r];
        if (!d) continue;
        const ex = existingData[r][0] ? existingData[r][0].toString().trim() : '';
        if (!ex) continue;
        const idx = allDutyStaff.findIndex(s => s.name === ex);
        if (idx === -1) continue;
        if (isHoliday(d)) holidayPtr = (idx + 1) % allDutyStaff.length;
        else              workdayPtr = (idx + 1) % allDutyStaff.length;
      }
    }

    // ════════════════════════════════════════════════════════════════
    // 協助掛號：指針輪轉（週四工作日），排除直接遞補
    // ★ 跨月接續（同樣用原始排班人）
    // ════════════════════════════════════════════════════════════════
    const allKkNames  = shiftStaffMap['D'] || [];
    // ★ 協助掛號用 J 欄順位（kkIdx）
    const allKkStaff  = allKkNames
      .filter(n => n)
      .map((name, kkIdx) => {
        const s = staff.find(x => x.name === name);
        return s ? { ...s, kkIdx } : null;
      })
      .filter(Boolean);
    let kkPtr = 0;

    function getPrevMonthLastPtrSingle(colIdx, staffArr) {
      if (staffArr.length === 0) return 0;
      const prevMonth = month === 1 ? 12 : month - 1;
      const prevYear  = month === 1 ? year - 1 : year;
      const rocMonths = ['','一','二','三','四','五','六','七','八','九','十','十一','十二'];
      const prevSheetName = rocNumToStr(prevYear - 1911) + '年' + rocMonths[prevMonth] + '月班表';
      try {
        const prevSheet = spreadsheet.getSheetByName(prevSheetName);
        if (!prevSheet) return 0;
        const origMap = buildOriginalMap(prevSheetName);
        const prevData = prevSheet.getRange('A2:M32').getValues();
        const tz = spreadsheet.getSpreadsheetTimeZone();
        let lastName = '';
        for (let r = 0; r < prevData.length; r++) {
          const raw = prevData[r][0];
          if (!raw) continue;
          let pd = null;
          if (raw instanceof Date) pd = raw;
          else {
            const mm = raw.toString().match(/(\d+)\/(\d+)/);
            if (mm) { try { pd = new Date(prevYear, parseInt(mm[1])-1, parseInt(mm[2])); } catch(e){} }
          }
          const dateStr = pd ? Utilities.formatDate(pd, tz, 'M/d') : '';
          const mapKey  = dateStr + '|' + colIdx;
          // ★ A2:M32 中 A=0,B=1,C=2...故需 +2 對齊班別欄（colIdx 0=C值班,1=D協助掛號...）
          const cellVal = prevData[r][colIdx + 2] ? prevData[r][colIdx + 2].toString().trim() : '';
          const name    = origMap[mapKey] || cellVal;
          if (name) lastName = name;
        }
        const idx = staffArr.findIndex(s => s.name === lastName);
        return idx !== -1 ? (idx + 1) % staffArr.length : 0;
      } catch(e) { return 0; }
    }

    kkPtr = getPrevMonthLastPtrSingle(1, allKkStaff);

    if (!options.overwrite) {
      for (let r = 0; r < datesRange.length; r++) {
        const d = dateObjByRow[r];
        if (!d) continue;
        const ex = existingData[r][1] ? existingData[r][1].toString().trim() : '';
        if (!ex) continue;
        const idx = allKkStaff.findIndex(s => s.name === ex);
        if (idx !== -1) kkPtr = (idx + 1) % allKkStaff.length;
      }
    }

    // ════════════════════════════════════════════════════════════════
    // 計數桶：colIdx 2-10（門診系列+卡介苗+登革熱）
    // ★ 門診系列(colIdx 2-8)：使用全年累積次數作為基準
    //   - 掃描本年度所有已排好的月份工作表（不含本月）
    //   - 新到職者：按全年剩餘月份比例折算基準積分，補至平均水準
    //   - 離職後歸零：人員不在名單就不統計（已從候選池移除）
    // ════════════════════════════════════════════════════════════════
    const assignCount = {};
    cols.forEach((col, idx) => {
      assignCount[idx] = {};
      staff.forEach(s => { assignCount[idx][s.name] = 0; });
    });

    // ── 全年累積計數（門診系列 colIdx 2-8）──────────────────────────
    // 「真正的新到職者」先宣告（carry=0），try block 內才賦值，確保後續所有邏輯可存取
    const trulyNewStaff = new Set();
    // ★ 週二/週四分開計數器（宣告在 try 外，確保全函式可用）
    const TUE_THU_CIS = new Set([2, 5, 6, 7]);
    const assignCountTue = {}; // 會在 try 內初始化
    const assignCountThu = {};

    try {
      const allSheets = spreadsheet.getSheets().map(s => s.getName());
      const rocYear   = year - 1911;
      const prefix    = rocNumToStr(rocYear) + '年';
      // ★ 修正 BUG 7：只累積「本月之前」已排月份的門診次數。
      //   原本掃「所有其他月份」（含未來月），導致第2次排班時1月能看到2-12月資料，
      //   打亂優先序 → 整年排班 vs 單月連排結果不一致，且多次重排結果不同。
      //   修正後：以 year/month 為界，只看已過去的月份，算法穩定可重複。
      const yearSheets = allSheets.filter(n => {
        if (n === EMAIL_SHEET_NAME || !n.includes(prefix) || n === sheetName) return false;
        const pm = parseYearMonthFromSheetName(n);
        // 只取同年、月份嚴格小於當前排班月的工作表
        return pm.valid && pm.year === year && pm.month < month;
      });
      const cliHeaders = [null, null, 'E','F','G','H','I','J','K', null, null]; // colIdx 2-8

      // ★ 週二/週四分開計數器初始化（TUE_THU_CIS 宣告於 try 外）
      TUE_THU_CIS.forEach(ci => {
        assignCountTue[ci] = {}; assignCountThu[ci] = {};
        staff.forEach(s => { assignCountTue[ci][s.name] = 0; assignCountThu[ci][s.name] = 0; });
      });

      // ★ 效能優化：批次取得所有工作表物件，避免重複 getSheetByName
      const allSheetsMap = {};
      spreadsheet.getSheets().forEach(sh => { allSheetsMap[sh.getName()] = sh; });

      yearSheets.forEach(sn => {
        const s = allSheetsMap[sn];
        if (!s) return;
        const lr = s.getLastRow();
        if (lr < 2) return;
        const maxRow = Math.min(lr, 32);
        // 一次讀 B2:K(maxRow)：B=星期字串, C-D跳過, E-K=門診資料(偏移調整)
        // 改為分兩欄：B欄(星期) + E:K(門診)
        const weekdayCol = s.getRange('B2:B' + maxRow).getValues(); // 星期字串
        const data       = s.getRange('E2:K' + maxRow).getValues(); // 門診欄
        const wkMap2 = {'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'日':0};
        data.forEach((row, ri) => {
          // 從 B 欄星期字串直接取 dow（快，不 new Date）
          let dow = -1;
          const wb = weekdayCol[ri][0] ? weekdayCol[ri][0].toString() : '';
          const wm = wb.match(/週([一二三四五六日])/);
          if (wm && wkMap2[wm[1]] !== undefined) dow = wkMap2[wm[1]];

          for (let ci = 2; ci <= 8; ci++) {
            const val = row[ci - 2] ? row[ci - 2].toString().trim() : '';
            if (!val || !assignCount[ci].hasOwnProperty(val)) continue;
            assignCount[ci][val]++;
            if (TUE_THU_CIS.has(ci)) {
              if (dow === 2) assignCountTue[ci][val] = (assignCountTue[ci][val]||0) + 1;
              if (dow === 4) assignCountThu[ci][val] = (assignCountThu[ci][val]||0) + 1;
            }
          }
        });
      });

      // ── 記錄「真正的新到職者」（carry 全為 0 的人，補基準前偵測）──
      {
        const clinicCiList2 = [2,3,4,5,6,7,8];
        Object.keys(assignCount[2] || {}).forEach(n => {
          // BUG 10 修正：運算子優先序 || 低於 +，必須加括號確保正確累加
          const total0 = clinicCiList2.reduce((s,ci2)=>((assignCount[ci2]||{})[n]||0)+s, 0);
          if (total0 === 0) trulyNewStaff.add(n);
        });
      }

      // ── 新到職者補齊積分（全局總量基準，平均分配各欄）─────────────
      // 用「全局門診總次數」的平均作為基準，再均分到各欄
      // 避免 per-column 為 0 但 total 卻偏低，導致新人月月補過頭
      const totalMonths  = 12;
      const passedMonths = Math.max(month - 1, 0); // 排本月時，已過去的月數
      // ★ 門診自動代償已移除：改為在前端統計面板顯示差異備註，不影響排班計數
      // ── 建立全局總次數 totalClinicCount（所有門診系列 ci=2-8 合計）──────
      //    用來在次數不均時以「總量」優先均衡，防止per-column偏差疊加
      // 注意：這僅用於 sort 的第一層，不改變 assignCount 的各欄統計
    } catch(e) {
      Logger.log('[runAutoSchedule] 全年累積計數失敗: ' + e.message);
    }

    // ════════════════════════════════════════════════════════════════
    // ★ 預先計算當月卡介苗日期 + 人員（pre-pass）
    //   排門診相關(colIdx 2-8)時，只在卡介苗「當日」排除該人員
    //   （不影響當月其餘日期的門診排班）
    // ════════════════════════════════════════════════════════════════
    let bcgPersonThisMonth = '';  // 保留供 successMsg 顯示
    let bcgDateKey = '';           // ★ 卡介苗當日的日期 key（yyyy-M-d）
    const qCandNames = shiftStaffMap['L'] || [];
    for (let r = 0; r < dateObjByRow.length; r++) {
      const d = dateObjByRow[r];
      if (!d || !shouldAssignShift(d, 9, year, month)) continue;
      // 找到卡介苗日
      bcgDateKey = d.getFullYear()+'-'+(d.getMonth()+1)+'-'+d.getDate();
      const qPool = staff.filter(s =>
        qCandNames.includes(s.name) &&
        !isPersonExcluded(s.name, d, exclusions, year)
      );
      if (qPool.length === 0) break;

      // 不覆蓋模式且已有資料 → 用現有資料
      if (!options.overwrite) {
        const ex9 = existingData[r][9] ? existingData[r][9].toString().trim() : '';
        if (ex9) { bcgPersonThisMonth = ex9; break; }
      }

      qPool.sort((a, b) => {
        const diff = (assignCount[9][a.name]||0) - (assignCount[9][b.name]||0);
        if (diff !== 0) return diff;
        const aIdx = qCandNames.indexOf(a.name);
        const bIdx = qCandNames.indexOf(b.name);
        return month % 2 === 1 ? bIdx - aIdx : aIdx - bIdx;
      });
      bcgPersonThisMonth = qPool[0].name;
      break;
    }

    // ── 計算每人全年門診總次數（ci 2-8 合計，含跨月carry）────────────
    // 用於排班優先排序：少者優先，避免手動排班造成的 per-column 不均拖累 total 均衡
    const totalClinicCount = {};
    staff.forEach(s => {
      let sum = 0;
      for (let ci = 2; ci <= 8; ci++) sum += (assignCount[ci][s.name] || 0);
      totalClinicCount[s.name] = sum;
    });

    // ── 預估本月門診總 slot 數 → 計算每人預期月次數（用於月內硬上限）──────
    // 掃一遍本月所有日期，計算實際會排的門診欄總數
    let totalClinicSlots = 0;
    for (let r2 = 0; r2 < dateObjByRow.length; r2++) {
      const d2 = dateObjByRow[r2];
      if (!d2) continue;
      for (let ci2 = 2; ci2 <= 8; ci2++) {
        if (shouldAssignShift(d2, ci2, year, month)) totalClinicSlots++;
      }
    }
    const clinicNurseCount = Object.keys(totalClinicCount).length || 1;
    // 預期每人月次數 = 總 slot / 護理師人數（至少 1）
    const expectedMonthlyClinic = totalClinicSlots / clinicNurseCount;

    // ── 本月門診次數計數器（用於月內上限）────────────────────────────
    const monthlyClinicCount = {};
    staff.forEach(s => { monthlyClinicCount[s.name] = 0; });
    // ★ 本月每欄次數計數器（每人每欄每月不可超過 2 次）
    const monthlyCountPerCi = {};
    for (let ci3 = 2; ci3 <= 8; ci3++) {
      monthlyCountPerCi[ci3] = {};
      staff.forEach(s => { monthlyCountPerCi[ci3][s.name] = 0; });
    }
    // ★ 混合日欄（週二+週四）分開月計，確保每人週二和週四各輪到一次
    const monthlyCountPerCiTue = {};
    const monthlyCountPerCiThu = {};
    for (const ci3 of [2, 5, 6, 7]) {  // 門診/預登1/預登2注/注射1
      monthlyCountPerCiTue[ci3] = {};
      monthlyCountPerCiThu[ci3] = {};
      staff.forEach(s => {
        monthlyCountPerCiTue[ci3][s.name] = 0;
        monthlyCountPerCiThu[ci3][s.name] = 0;
      });
    }

    // 恢復不覆蓋模式計數（colIdx 2-10，排除0跟1）
    const dutyByDate = {};
    const bcgByDate  = {};
    if (!options.overwrite) {
      for (let r = 0; r < datesRange.length; r++) {
        const d = dateObjByRow[r];
        if (!d) continue;
        const dk = d.getFullYear()+'-'+(d.getMonth()+1)+'-'+d.getDate();
        const v0 = existingData[r][0] ? existingData[r][0].toString().trim() : '';
        if (v0) dutyByDate[dk] = v0;
        for (let ci = 2; ci <= 10; ci++) {
          const val = existingData[r][ci] ? existingData[r][ci].toString().trim() : '';
          if (!val) continue;
          if (ci === 9) bcgByDate[dk] = val;
          if (assignCount[ci] && assignCount[ci].hasOwnProperty(val)) {
            assignCount[ci][val]++;
            // ★ 不覆蓋模式下，恢復週二/週四分開計數
            if (TUE_THU_CIS && TUE_THU_CIS.has(ci)) {
              const d_ex = dateObjByRow[r];
              const dow_ex = d_ex ? d_ex.getDay() : -1;
              if (dow_ex === 2 && assignCountTue[ci]) assignCountTue[ci][val] = (assignCountTue[ci][val]||0)+1;
              if (dow_ex === 4 && assignCountThu[ci]) assignCountThu[ci][val] = (assignCountThu[ci][val]||0)+1;
            }
            if (monthlyCountPerCi[ci]) monthlyCountPerCi[ci][val] = (monthlyCountPerCi[ci][val]||0)+1;
            // 恢復混合日欄分開月計
            if (TUE_THU_CIS && TUE_THU_CIS.has(ci) && monthlyCountPerCiTue[ci] && monthlyCountPerCiThu[ci]) {
              if (dow_ex === 2) monthlyCountPerCiTue[ci][val] = (monthlyCountPerCiTue[ci][val]||0)+1;
              if (dow_ex === 4) monthlyCountPerCiThu[ci][val] = (monthlyCountPerCiThu[ci][val]||0)+1;
            }
          }
        }
      }
    }

    // ── 預算各欄在每個 rowIdx 之後還剩幾天（含當天），混合日欄分週別 ──
    const colFutureDays    = {};  // 合計（用於非混合日欄）
    const colFutureDaysTue = {};  // 僅週二（用於混合日欄）
    const colFutureDaysThu = {};  // 僅週四（用於混合日欄）
    for (let ci2 = 2; ci2 <= 8; ci2++) {
      const days    = new Array(dateObjByRow.length).fill(0);
      const daysTue = new Array(dateObjByRow.length).fill(0);
      const daysThu = new Array(dateObjByRow.length).fill(0);
      let cnt=0, cntT=0, cntH=0;
      for (let r2 = dateObjByRow.length - 1; r2 >= 0; r2--) {
        const d2 = dateObjByRow[r2];
        if (d2 && shouldAssignShift(d2, ci2, year, month)) {
          cnt++;
          const dw2 = d2.getDay();
          if (dw2 === 2) cntT++;
          if (dw2 === 4) cntH++;
        }
        days[r2]    = cnt;
        daysTue[r2] = cntT;
        daysThu[r2] = cntH;
      }
      colFutureDays[ci2]    = days;
      colFutureDaysTue[ci2] = daysTue;
      colFutureDaysThu[ci2] = daysThu;
    }

    // ════════════════════════════════════════════════════════════════
    // ★ 第一回合：colIdx 0,1,9,2-8（不含登革熱10）
    //   處理順序：值班(0) → 協助掛號(1) → 卡介苗(9) → 門診系列(2-8)
    //   三最高優先可同日同人，不互相干預
    //   ★ 門診系列(2-9)各欄之間：同日不可重複出現同一人
    // ════════════════════════════════════════════════════════════════
    const result = [];

    for (let rowIdx = 0; rowIdx < datesRange.length; rowIdx++) {
      const d      = dateObjByRow[rowIdx];
      const rowRes = new Array(cols.length).fill('');
      const dk     = d ? d.getFullYear()+'-'+(d.getMonth()+1)+'-'+d.getDate() : 'nd-'+rowIdx;
      const hol    = d ? isHoliday(d) : false;

      // ★ 每日門診系列已排人員集合（colIdx 2-9，確保同日不重複）
      const clinicAssigned = new Set();

      // ★ 門診系列（ci=2-8）的處理順序依 rowIdx 輪轉
      //   確保補償人員（總次數最低者）不會每天都被分到同一個職務欄
      //   值班(0)、支援(1)、卡介苗(9) 維持固定順序
      const clinicCols = [2, 3, 4, 5, 6, 7, 8];
      const clinicStart = rowIdx % clinicCols.length;
      const rotatedClinic = [
        ...clinicCols.slice(clinicStart),
        ...clinicCols.slice(0, clinicStart)
      ];
      // ★ 今日門診欄排序：0次護理師較多的欄優先處理
      //   確保稀缺資源（0次護理師）先分配給最需要的欄，避免被其他欄「搶走」
      const sortedClinic = rotatedClinic.slice().sort((a, b) => {
        const colA = cols[a];
        const colB = cols[b];
        const namesA = shiftStaffMap[colA] || [];
        const namesB = shiftStaffMap[colB] || [];
        const zeroA = namesA.filter(n => n && !isPersonExcluded(n, d, exclusions, year) &&
          ((monthlyCountPerCi[a] && monthlyCountPerCi[a][n]) || 0) === 0).length;
        const zeroB = namesB.filter(n => n && !isPersonExcluded(n, d, exclusions, year) &&
          ((monthlyCountPerCi[b] && monthlyCountPerCi[b][n]) || 0) === 0).length;
        return zeroB - zeroA; // 0次護理師多的欄先處理
      });
      const processOrder = [0, 1, 9, ...sortedClinic];

      for (const ci of processOrder) {

        // 不覆蓋：保留現有
        if (!options.overwrite) {
          const ex = existingData[rowIdx][ci] ? existingData[rowIdx][ci].toString().trim() : '';
          if (ex !== '') {
            rowRes[ci] = ex;
            if (ci === 0) dutyByDate[dk] = ex;
            if (ci === 9) { bcgByDate[dk] = ex; clinicAssigned.add(ex); }
            if (ci >= 2 && ci <= 8) clinicAssigned.add(ex);
            continue;
          }
        }

        if (!d || !shouldAssignShift(d, ci, year, month)) { rowRes[ci] = ''; continue; }

        const col    = cols[ci];
        const cNames = shiftStaffMap[col] || [];
        if (cNames.length === 0) { rowRes[ci] = ''; continue; }

        // ────────────────────────────────────────────────────────────
        // 值班(0)：指針輪轉，排除遞補
        // ────────────────────────────────────────────────────────────
        if (ci === 0) {
          const ptr = hol ? holidayPtr : workdayPtr;
          const { name, nextPtr } = pickFromPtr(allDutyStaff, ptr, d);
          rowRes[0] = name;
          if (name) dutyByDate[dk] = name;
          if (hol) holidayPtr = nextPtr; else workdayPtr = nextPtr;

        // ────────────────────────────────────────────────────────────
        // 協助掛號(1)：指針輪轉，排除遞補，可與值班/卡介苗重疊
        // ────────────────────────────────────────────────────────────
        } else if (ci === 1) {
          const { name, nextPtr } = pickFromPtr(allKkStaff, kkPtr, d);
          rowRes[1] = name;
          kkPtr = nextPtr;

        // ────────────────────────────────────────────────────────────
        // 卡介苗(9)：動態 Q欄，奇偶月優先，排除遞補
        // ────────────────────────────────────────────────────────────
        } else if (ci === 9) {
          const qPool = staff.filter(s =>
            cNames.includes(s.name) &&
            !isPersonExcluded(s.name, d, exclusions, year)
          );
          if (qPool.length === 0) { rowRes[9] = ''; continue; }
          qPool.sort((a, b) => {
            const diff = (assignCount[9][a.name]||0) - (assignCount[9][b.name]||0);
            if (diff !== 0) return diff;
            const aIdx = cNames.indexOf(a.name);
            const bIdx = cNames.indexOf(b.name);
            return month % 2 === 1 ? bIdx - aIdx : aIdx - bIdx;
          });
          const chosen = qPool[0].name;
          rowRes[9] = chosen;
          assignCount[9][chosen] = (assignCount[9][chosen]||0) + 1;
          bcgByDate[dk] = chosen;
          clinicAssigned.add(chosen); // ★ 記入門診不重複集合

        // ────────────────────────────────────────────────────────────
        // 門診系列(2-8)：次數均勻，排除遞補
        // ★ 卡介苗人員只在卡介苗「當日」不排此類職務（其餘日正常排）
        // ★ 同日已排門診系列的人不可再排（clinicAssigned）
        // ────────────────────────────────────────────────────────────
        } else {
          // ★ 門診系列：次數最少者優先
          //   ★ 週二/週四分開輪序：週二日排週二次數少者優先，週四日排週四次數少者優先
          //   ★ 每人每欄每月不可超過 2 次（per-ci monthly cap）
          const dow = d ? d.getDay() : -1;
          const isMixedDayCol = TUE_THU_CIS && TUE_THU_CIS.has(ci);

          const pool = cNames
            .filter(n => n)
            .map((name, clinicIdx) => {
              const s = staff.find(x => x.name === name);
              return s ? { ...s, clinicIdx } : null;
            })
            .filter(s => s &&
              !(s.name === bcgPersonThisMonth && dk === bcgDateKey) &&
              !isPersonExcluded(s.name, d, exclusions, year) &&
              !clinicAssigned.has(s.name)
            );
          if (pool.length === 0) { rowRes[ci] = ''; continue; }

          const dayPhase = (rowIdx * (ci + 1)) % Math.max(cNames.length, 1);

          // ★ 月內限制：
          //   (A) 每欄每人每月不可超過 2 次
          //   (B) 總月次數不超過 expectedMonthlyClinic + 2（一般）或 ceil（新人）
          const CLINIC_CAP = 2;
          const hardCap   = Math.ceil(expectedMonthlyClinic);
          const normalCap = hardCap + CLINIC_CAP;
          // ═══════════════════════════════════════════════════════════
          // ★ 分配邏輯（月總次數優先，確保 max-min ≤ 1）：
          //
          //   數學分析：50格 / 7人 = 7.14 → 1人得8格，6人得7格
          //   只需一個規則：月總最少者優先（自然均衡）
          //   per-col hard cap = 2（防止極端集中），但不做兩階段阻擋
          //
          //   排序鍵：
          //   ① 月內總次數（最少優先）← 核心均衡鍵
          //   ② 此欄此日（週二/週四）月計（0優先1）← 欄內分散
          //   ③ 跨月歷史總次數（少優先）← 跨月均衡
          //   ④ 週二/週四分欄輪序
          //   ⑤ Round-robin
          // ═══════════════════════════════════════════════════════════

          // 混合日欄依 dow 使用日別月計（確保週二/週四各自公平）
          const isMixedDayCi = TUE_THU_CIS && TUE_THU_CIS.has(ci);
          const getPerCiDay = (name) => {
            if (isMixedDayCi) {
              if (dow === 2) return (monthlyCountPerCiTue[ci] && monthlyCountPerCiTue[ci][name]) || 0;
              if (dow === 4) return (monthlyCountPerCiThu[ci] && monthlyCountPerCiThu[ci][name]) || 0;
            }
            return (monthlyCountPerCi[ci] && monthlyCountPerCi[ci][name]) || 0;
          };

          // 此欄此日的 hard cap（≤2），防止極端集中
          const HARD_CAP_PER_SLOT = 2;
          const finalPool = pool.filter(p => getPerCiDay(p.name) < HARD_CAP_PER_SLOT);
          const effectivePool = finalPool.length > 0 ? finalPool : pool;

          effectivePool.sort((a, b) => {
            // ① 此欄此日=0優先（主要：確保所有人輪1次後才有2次）
            const mca = getPerCiDay(a.name);
            const mcb = getPerCiDay(b.name);
            if (mca !== mcb) return mca - mcb;

            // ② 月內總次數少者優先（次要：月內均衡，max-min≤1）
            const mthA = monthlyClinicCount[a.name] || 0;
            const mthB = monthlyClinicCount[b.name] || 0;
            if (mthA !== mthB) return mthA - mthB;

            // ③ 跨月歷史總次數少者優先（補償）
            const ta = totalClinicCount[a.name] || 0;
            const tb = totalClinicCount[b.name] || 0;
            if (ta !== tb) return ta - tb;

            // ④ 週二/週四分開輪序
            let ca, cb;
            if (isMixedDayCol && dow === 2) {
              ca = (assignCountTue[ci] && assignCountTue[ci][a.name]) || 0;
              cb = (assignCountTue[ci] && assignCountTue[ci][b.name]) || 0;
            } else if (isMixedDayCol && dow === 4) {
              ca = (assignCountThu[ci] && assignCountThu[ci][a.name]) || 0;
              cb = (assignCountThu[ci] && assignCountThu[ci][b.name]) || 0;
            } else {
              ca = assignCount[ci][a.name] || 0;
              cb = assignCount[ci][b.name] || 0;
            }
            if (ca !== cb) return ca - cb;

            // ⑤ Round-robin
            const plen = Math.max(effectivePool.length, 1);
            const ra = (a.clinicIdx - dayPhase % plen + plen) % plen;
            const rb = (b.clinicIdx - dayPhase % plen + plen) % plen;
            return ra - rb;
          });
          const finalPoolRef = effectivePool;
          const chosen = finalPoolRef[0].name;
          rowRes[ci] = chosen;
          assignCount[ci][chosen] = (assignCount[ci][chosen] || 0) + 1;
          // ★ 同步更新週二/週四分開計數器
          if (isMixedDayCol) {
            if (dow === 2) assignCountTue[ci][chosen] = (assignCountTue[ci][chosen]||0) + 1;
            if (dow === 4) assignCountThu[ci][chosen] = (assignCountThu[ci][chosen]||0) + 1;
          }
          totalClinicCount[chosen]   = (totalClinicCount[chosen]   || 0) + 1;
          monthlyClinicCount[chosen] = (monthlyClinicCount[chosen] || 0) + 1;
          if (monthlyCountPerCi[ci]) monthlyCountPerCi[ci][chosen] = (monthlyCountPerCi[ci][chosen]||0) + 1;
          // ★ 混合日欄：同步更新日別月計
          if (isMixedDayCi) {
            if (dow === 2 && monthlyCountPerCiTue[ci]) monthlyCountPerCiTue[ci][chosen] = (monthlyCountPerCiTue[ci][chosen]||0) + 1;
            if (dow === 4 && monthlyCountPerCiThu[ci]) monthlyCountPerCiThu[ci][chosen] = (monthlyCountPerCiThu[ci][chosen]||0) + 1;
          }
          clinicAssigned.add(chosen);
        }
      }

      result.push(rowRes);
    }

    // ════════════════════════════════════════════════════════════════
    // ════════════════════════════════════════════════════════════════
    // ★ 第二回合：停班2線(colIdx=10)
    //   假日、平日各自用B欄指針輪轉，各自內部swap衝突解決
    //   假日只和假日swap，平日只和平日swap，不可混換
    // ════════════════════════════════════════════════════════════════
    const dengCandNames = shiftStaffMap['M'] || [];
    const dengStaff = dengCandNames
      .filter(n => n)
      .map((name, dengueIdx) => {
        const s = staff.find(x => x.name === name);
        return s ? { ...s, dengueIdx } : null;
      })
      .filter(Boolean);

    // ── 讀取累積換班次數（N1 備註中 swapCount:name1=3,name2=1...）───
    let swapCounts = {};
    try {
      const n1Note = sheet.getRange('N1').getNote() || '';
      n1Note.split('\n').forEach(line => {
        if (line.startsWith('swapCount:')) {
          line.replace('swapCount:', '').split(',').forEach(pair => {
            const [nm, cnt] = pair.split('=');
            if (nm && cnt) swapCounts[nm.trim()] = parseInt(cnt) || 0;
          });
        }
      });
    } catch(e) {}

    // ── 共用工具：B欄指針輪轉分配一組 slot，衝突時優先找換最少次的人 swap ────
    const tz2 = spreadsheet.getSpreadsheetTimeZone();
    function assignSlotsWithPointer(slots, startPtr, isHolGroup) {
      const n = dengStaff.length;
      let ptr = startPtr;
      const assign = [];
      // Step1: 指針輪轉分配
      for (let i = 0; i < slots.length; i++) {
        const { d } = slots[i];
        const avail = dengStaff.filter(s => !isPersonExcluded(s.name, d, exclusions, year));
        let chosen = '', nextPtr = ptr;
        for (let attempt = 0; attempt < n; attempt++) {
          const idx  = (ptr + attempt) % n;
          const name = dengStaff[idx] ? dengStaff[idx].name : '';
          if (name && avail.find(s => s.name === name)) {
            chosen = name; nextPtr = (idx + 1) % n; break;
          }
        }
        assign.push(chosen);
        ptr = nextPtr;
      }
      // Step2: 衝突 swap（多輪迴圈，直到無衝突為止）
      // BUG 14 修正：原先只做單輪 j > i 掃描，無法處理：
      //   (1) i 之前的位置才能提供互換對象；(2) swap 後新衝突未被偵測
      // 現改為：全範圍掃描 + 多輪迴圈，最多執行 assign.length 輪
      const swappedSlots = new Set();
      const swapInfo = {};
      let maxPass = assign.length + 1;
      let hasConflict = true;
      while (hasConflict && maxPass-- > 0) {
        hasConflict = false;
        for (let i = 0; i < assign.length; i++) {
          if (!assign[i]) continue;
          const dutyI = result[slots[i].rowIdx][0] || dutyByDate[slots[i].dk] || '';
          if (assign[i] !== dutyI) continue;
          // 衝突：assign[i] 與當日值班者相同
          hasConflict = true;
          // 找同組內「換後不衝突」的最佳對象（全範圍搜尋，不限 j > i）
          let bestJ = -1, bestCount = Infinity;
          for (let j = 0; j < assign.length; j++) {
            if (j === i || !assign[j]) continue;
            if (assign[j] === dutyI) continue; // j 本身也衝突，換了沒用
            // 確認把 assign[j] 放到 slot[i] 後不會再衝突
            const dutyJ = result[slots[j].rowIdx][0] || dutyByDate[slots[j].dk] || '';
            if (assign[i] === dutyJ) continue; // assign[i] 換去 j 位置也會衝突
            let cnt = swapCounts[assign[j]] || 0;
            // ★ BUG 26 修正：避免把 slot 0（跨月接續首位）當交換對象，
            //   否則第一天的人被換走，跨月輪序自檢會失敗
            if (j === 0) cnt += 9999;
            if (cnt < bestCount) { bestCount = cnt; bestJ = j; }
          }
          if (bestJ !== -1) {
            swapCounts[assign[i]]     = (swapCounts[assign[i]]     || 0) + 1;
            swapCounts[assign[bestJ]] = (swapCounts[assign[bestJ]] || 0) + 1;
            const tmp = assign[i]; assign[i] = assign[bestJ]; assign[bestJ] = tmp;
            swappedSlots.add(i); swappedSlots.add(bestJ);
            swapInfo[i]     = { originalSlot: bestJ };
            swapInfo[bestJ] = { originalSlot: i };
          }
          // 找不到可換對象：保留原始分配（衝突由 Check 3 偵測）
        }
      }
      return { assign, swappedSlots, swapInfo, finalPtr: ptr };
    }

    // ── 從上個月接續 hSlot（假日）及 wdSlot（平日）指針 ─────────────
    function getPrevLastPtr(prevIsHol) {
      const prevM2 = month === 1 ? 12 : month - 1;
      const prevY2 = month === 1 ? year - 1 : year;
      const rocM2  = ['','一','二','三','四','五','六','七','八','九','十','十一','十二'];
      const prevSN = rocNumToStr(prevY2 - 1911) + '年' + rocM2[prevM2] + '月班表';
      try {
        const prevSh = spreadsheet.getSheetByName(prevSN);
        if (!prevSh) return 0;
        const origMap2 = buildOriginalMap(prevSN);
        const prevD2   = prevSh.getRange('A2:M32').getValues();
        // ver4.6：讀上月 M 欄 swap note，跳過系統挪移的列（避免下月接續找錯起點）
        const prevMNotes = prevSh.getRange('M2:M32').getNotes();
        let lastName = '';
        for (let r = 0; r < prevD2.length; r++) {
          const rawA = prevD2[r][0];
          let pd = null;
          if (rawA instanceof Date) pd = rawA;
          else { const mm = rawA ? rawA.toString().match(/(\d+)\/(\d+)/) : null; if(mm) try { pd = new Date(prevY2, parseInt(mm[1])-1, parseInt(mm[2])); } catch(e){} }
          if (!pd) continue;
          if (prevIsHol !== isHoliday(pd)) continue; // 只看對應類型
          // 系統挪移列（'swap:M/d' note）→ 跳過，找真正輪序的最後一位
          const mNote = (prevMNotes[r] && prevMNotes[r][0]) ? prevMNotes[r][0].toString() : '';
          if (mNote.indexOf('swap:') === 0) continue;
          const ds2 = Utilities.formatDate(pd, tz2, 'M/d');
          // ★ M欄（停班2線）在 A2:M32 中是 index 12，不是 10
          const cv2 = prevD2[r][12] ? prevD2[r][12].toString().trim() : '';
          const nm2 = (origMap2[ds2 + '|10']) || cv2;
          if (nm2) lastName = nm2;
        }
        const wi = dengStaff.findIndex(s => s.name === lastName);
        return wi !== -1 ? (wi + 1) % dengStaff.length : 0;
      } catch(e) { return 0; }
    }

    // ── 分開假日槽和平日槽 ──────────────────────────────────────────
    const hSlots = [];
    const wdSlots = [];
    for (let r = 0; r < dateObjByRow.length; r++) {
      const d = dateObjByRow[r];
      if (!d) continue;
      const dk = d.getFullYear()+'-'+(d.getMonth()+1)+'-'+d.getDate();
      const hol = isHoliday(d);
      if (!options.overwrite) {
        const ex10 = existingData[r][10] ? existingData[r][10].toString().trim() : '';
        if (ex10 !== '') {
          result[r][10] = ex10;
          if (assignCount[10].hasOwnProperty(ex10)) assignCount[10][ex10]++;
          continue;
        }
      }
      if (hol) hSlots.push({ rowIdx: r, d, dk });
      else     wdSlots.push({ rowIdx: r, d, dk });
    }

    // ── 假日：B欄指針輪轉 + 同組swap ────────────────────────────────
    const hStartPtr = getPrevLastPtr(true);
    const hResult   = assignSlotsWithPointer(hSlots, hStartPtr, true);

    // ── 平日：B欄指針輪轉 + 同組swap ────────────────────────────────
    const wdStartPtr = getPrevLastPtr(false);
    const wdResult   = assignSlotsWithPointer(wdSlots, wdStartPtr, false);

    // ── 寫入結果 + 記錄挪移 ─────────────────────────────────────────
    const dengSwappedRows = new Set();
    const dengSwapDateMap = {};
    function writeSlotResult(slots, res) {
      const { assign, swappedSlots, swapInfo } = res;
      for (let i = 0; i < slots.length; i++) {
        result[slots[i].rowIdx][10] = assign[i] || '';
        if (swappedSlots.has(i)) {
          dengSwappedRows.add(slots[i].rowIdx);
          if (swapInfo[i]) {
            const origSlot = swapInfo[i].originalSlot;
            const origD = slots[origSlot] ? slots[origSlot].d : null;
            if (origD) dengSwapDateMap[slots[i].rowIdx] = Utilities.formatDate(origD, tz2, 'M/d');
          }
        }
      }
    }
    writeSlotResult(hSlots,  hResult);
    writeSlotResult(wdSlots, wdResult);

    // ── 寫入試算表 ──────────────────────────────────────────────────
    if (!options.dryRun) {
      // ★ 強制 A2:A32 為純文字，確保日期不被轉成 Date 物件
      sheet.getRange('A2:A32').setNumberFormat('@');
      sheet.getRange('C2:M32').setValues(result);

      // ★ 記錄系統排定時間到 N1 備註（保留原有審核狀態 + 累加版次）
      const schedTimeStr = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'yyyy/MM/dd HH:mm');
      let existingReview = '';
      let writeCount = 0;
      try {
        const oldNote = sheet.getRange('N1').getNote() || '';
        oldNote.split('\n').forEach(line => {
          if (line.startsWith('審核狀態:'))   existingReview = line.trim();
          if (line.startsWith('writeCount:')) writeCount = parseInt(line.replace('writeCount:', '').trim()) || 0;
        });
      } catch(e) {}
      writeCount += 1; // 每次寫入 +1（1=A, 2=B, ...）
      // ★ 自動排班後一律設為審核中，待管理員核准後才開放換班
      // 同時保存累積換班次數
      const swapCountStr = Object.keys(swapCounts).length > 0
        ? '\nswapCount:' + Object.keys(swapCounts).map(n => n+'='+swapCounts[n]).join(',')
        : '';
      // ver4.6：記錄停班2線系統挪移列（rowIdx → 原日期），讓自檢/前端識別非 bug
      const dengSwapStr = Object.keys(dengSwapDateMap).length > 0
        ? '\ndengSwapRows:' + Object.keys(dengSwapDateMap).map(rIdx => rIdx+'='+dengSwapDateMap[rIdx]).join(',')
        : '';
      const newNote = '排定時間:' + schedTimeStr + '\n審核狀態:pending\nwriteCount:' + writeCount + swapCountStr + dengSwapStr;
      sheet.getRange('N1').setNote(newNote);

      // ★ 將挪移資訊寫入儲存格備註（M欄 = 登革熱二線，欄序13）
      // 先清空 M2:M32 所有備註，再設定有挪移的格子
      const mCol = sheet.getRange('M2:M32');
      mCol.clearNote();
      Object.keys(dengSwapDateMap).forEach(rowIdxStr => {
        const rowIdx = parseInt(rowIdxStr);
        const origDate = dengSwapDateMap[rowIdx];
        if (origDate && result[rowIdx][10]) {
          // row 2 + rowIdx（result 從 0 開始對應第2列）
          sheet.getRange(rowIdx + 2, 13).setNote('swap:' + origDate);
        }
      });

      if (options.sendNotify) {
        try { sendAutoScheduleNotification(sheetName, staff); } catch(e) {}
      }
    }

    const successMsg = `${sheetName} 排班完成！（${year}年${month}月，共 ${result.length} 天，卡介苗：${bcgPersonThisMonth||'未定'}）`;
    writeOpLog('自動排班', successMsg);
    // ★ 回傳哪些列是假日，讓前端正確標色
    const holidayRows = dateObjByRow.map(d => d ? isHoliday(d) : false);
    // ★ 門診候選護理師清單（E..K各欄合集，供前端「新增」功能使用）
    const clinicStaffSet = new Set();
    for (let ci = 2; ci <= 8; ci++) {
      const col2 = cols[ci];
      (shiftStaffMap[col2] || []).forEach(n => { if(n) clinicStaffSet.add(n.toString().trim()); });
    }
    return {
      success: true,
      message: successMsg,
      dates:   combinedDates.map(d => [d]),
      headers,
      preview: result,
      holidayRows,
      dengSwapped:     Array.from(dengSwappedRows),
      dengSwapDateMap: dengSwapDateMap,
      clinicStaff:     Array.from(clinicStaffSet)
    };

  } catch(e) {
    logError('runAutoSchedule: ' + e.message);
    return { success: false, message: '排班失敗：' + e.message };
  }
}


function previewAutoSchedule(sheetName, adminPassword) {
  return runAutoSchedule(sheetName, adminPassword, { overwrite: true, sendNotify: false, dryRun: true });
}

function autoSchedule(sheetName, adminPassword, options) {
  return runAutoSchedule(sheetName, adminPassword, {
    overwrite: options.overwrite || false,
    sendNotify: options.sendNotify || false,
    dryRun: false
  });
}

function sendAutoScheduleNotification(sheetName, staff) {
  const subject = `班表通知：${sheetName} 自動排班已完成`;
  const body = `${sheetName} 已完成自動排班，請登入系統確認班表內容。\n\n排班時間：${new Date().toLocaleString()}\n\n此為系統自動發送的通知信件。`;
  const emails = staff.map(s => s.email).filter(e => e && e.includes('@'));
  if (emails.length > 0) {
    MailApp.sendEmail({ to: emails.slice(0,50).join(','), subject, body, name: '班表管理系統' });
  }
}

function sendWeeklyNotification(sheetName, adminPassword) {
  if (!verifyAdminPassword(adminPassword)) return '管理員密碼錯誤。';
  try {
    sendAutoScheduleNotification(sheetName, getStaffList());
    return '本週通知已成功寄送！';
  } catch(e) { return '寄送失敗：' + e.message; }
}

// ── 取得某月可選日期（for 手動補寄 UI）────────────────────────────
function getScheduleDates(sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const timezone    = spreadsheet.getSpreadsheetTimeZone();
    const sched = spreadsheet.getSheetByName(sheetName);
    if (!sched) return { success: false, dates: [] };
    const dateCol  = sched.getRange('A2:A32').getValues();
    const headers  = sched.getRange('C1:M1').getValues()[0];
    const data     = sched.getRange('C2:M32').getValues();
    const result   = [];
    dateCol.forEach((r, i) => {
      const v = r[0];
      if (!v) return;
      const s = v instanceof Date ? Utilities.formatDate(v, timezone, 'M/d') : v.toString().trim();
      const row = data[i] || [];
      const hasDuty = row.some(c => c && c.toString().trim());
      if (s && hasDuty) result.push(s);
    });
    return { success: true, dates: result, headers: headers.filter(Boolean) };
  } catch(e) { return { success: false, dates: [], message: e.message }; }
}

// =============================================
// 一鍵排整年：115年4月～12月
// =============================================

// getYearlyScheduleSheets / autoScheduleFullYear / previewFullYear 已移除
// （整年排班功能已移除，統一使用單月排班）

// =============================================
// 職務統計看板
// 讀取指定工作表（或全部班表），統計每人各班別出現次數
// =============================================
function getScheduleStats(sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const headers = GLOBAL_CONFIG.SHIFT_HEADERS;
    const staff = getStaffList();
    const staffNames = staff.map(s => s.name);

    // 決定要統計的工作表清單（只限115年/2026年班表）
    let sheets = [];
    if (sheetName === '__ALL__') {
      const allNames = spreadsheet.getSheets().map(s => s.getName());
      sheets = allNames.filter(n => {
        if (n === EMAIL_SHEET_NAME) return false;
        if (!n.includes('班表')) return false;
        const ym = parseYearMonthFromSheetName(n);
        return ym.valid && ym.year === new Date().getFullYear();
      });
    } else {
      sheets = [sheetName];
    }

    // 初始化統計矩陣：stats[name][colIdx] = count（只統計名單內人員）
    const stats = {};
    staffNames.forEach(n => {
      stats[n] = new Array(headers.length).fill(0);
    });
    const staffSet = new Set(staffNames);

    let totalDays = 0;

    sheets.forEach(sName => {
      const sh = spreadsheet.getSheetByName(sName);
      if (!sh) return;
      const data = sh.getRange('C2:M32').getValues();
      data.forEach(row => {
        let hasData = false;
        row.forEach((cell, ci) => {
          const val = cell ? cell.toString().trim() : '';
          if (!val) return;
          hasData = true;
          val.split(',').forEach(v => {
            const name = v.trim();
            if (staffSet.has(name)) stats[name][ci]++;  // 只計名單內人員
          });
        });
        if (hasData) totalDays++;
      });
    });

    // 轉成陣列回傳
    const rows = Object.keys(stats).map(name => ({
      name,
      counts: stats[name],
      total: stats[name].reduce((a, b) => a + b, 0)
    }));
    rows.sort((a, b) => b.total - a.total);

    return { success: true, headers, rows, sheets, totalDays };
  } catch(e) {
    return { success: false, message: e.message, headers: [], rows: [], sheets: [], totalDays: 0 };
  }
}

// =============================================
// 排班排除清單（請長假 / 離職 / 到職日設定）
// 「班表設定」工作表：
//   T欄 = 姓名
//   U欄 = 開始排除日期（M/D，空白=從最早開始）
//   V欄 = 結束排除日期（M/D，空白=持續排除到底）
// =============================================

/**
 * 讀取排除清單
 * @returns {Array} [{name, startMD, endMD}]
 */
function getExclusionList() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    const tz    = SpreadsheetApp.openById(SHEET_ID).getSpreadsheetTimeZone();
    const raw   = sheet.getRange(GLOBAL_CONFIG.EXCLUSION_RANGE).getValues();

    // 日期格式化：若儲存格是 Date 物件則轉成 M/d，否則直接取字串
    function fmtDate(v) {
      if (!v) return '';
      if (v instanceof Date) return Utilities.formatDate(v, tz, 'M/d');
      return v.toString().trim();
    }

    return raw
      .filter(r => r[0] && r[0].toString().trim() !== '')
      .map(r => ({
        name:     r[0].toString().trim(),
        startMD:  fmtDate(r[1]),
        endMD:    fmtDate(r[2]),
        weekdays: r[3] ? r[3].toString().trim() : ''
      }));
  } catch(e) { return []; }
}

function saveExclusionList(adminPassword, list) {
  if (!verifyAdminPassword(adminPassword)) return { success: false, message: '管理員密碼錯誤。' };
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    sheet.getRange(GLOBAL_CONFIG.EXCLUSION_RANGE).clearContent();
    if (list && list.length > 0) {
      const values = list.slice(0, 30).map(item => [
        item.name || '', item.startMD || '', item.endMD || '', item.weekdays || ''
      ]);
      sheet.getRange(1, 20, values.length, 4).setValues(values);
    }
    writeOpLog('排除設定', '儲存排除清單，共 ' + (list||[]).length + ' 筆');
    return { success: true, message: '排除清單已儲存！' };
  } catch(e) {
    return { success: false, message: '儲存失敗：' + e.message };
  }
}

function isPersonExcluded(name, dateObj, exclusions, year) {
  if (!exclusions || exclusions.length === 0) return false;
  const month = dateObj.getMonth() + 1;
  const day   = dateObj.getDate();
  const dow   = dateObj.getDay();
  const DOW_MAP = {'日':0,'一':1,'二':2,'三':3,'四':4,'五':5,'六':6};

  return exclusions.some(ex => {
    if (ex.name !== name) return false;
    let startM=0,startD=0;
    if (ex.startMD) { const m=ex.startMD.match(/(\d+)\/(\d+)/); if(m){startM=parseInt(m[1]);startD=parseInt(m[2]);} }
    let endM=99,endD=99;
    if (ex.endMD) { const m=ex.endMD.match(/(\d+)\/(\d+)/); if(m){endM=parseInt(m[1]);endD=parseInt(m[2]);} }
    const cur=month*100+day, start=startM*100+startD, end=endM*100+endD;
    if (cur<start||cur>end) return false;
    if (ex.weekdays && ex.weekdays!=='') {
      const wds=ex.weekdays.split(/[,、，\s]+/).map(s=>s.trim());
      if (wds.includes('假日') && isHoliday(dateObj)) return true;
      for (const wd of wds) { if (wd in DOW_MAP && DOW_MAP[wd]===dow) return true; }
      return false;
    }
    return true;
  });
}

// ── 判斷某人員在指定年月是否仍在職 ──────────────────────────────────
// 有離職日且 leaveDate < 排班月第一天 → 已離職，不應排入
// 有到職日且 joinDate > 排班月最後一天 → 尚未到職，不應排入
function isStaffActiveForMonth(staffObj, schedYear, schedMonth) {
  // BUG 11 修正：改用完整日期比較（年+月+日），避免跨年誤判
  // bug #8: X/Y 欄只存「月/日」沒有年份。原本一律代 schedYear 解讀，
  //   • 2024/10 離職者，排 2026/1 → leaveFullDate=2026/10/15，不小於 2026/1/1 → 誤判仍在職
  //   • 2026/10 才到職者，排 2026/1 → joinFullDate=2026/10/15 > 2026/1/31 → 正確（恰好）
  //   • 但 2025/2 到職、排 2026/1 → joinFullDate=2026/2/15 > 2026/1/31 → 誤判尚未到職
  // 解法：支援「YYYY/M/D」完整年月日（首選）；若只有 M/D，採「最近過去」推斷年份：
  //   - 離職日：取 ≤ schedFirstDay 最近的 M/D 那年（即 schedYear 或 schedYear-1）
  //   - 到職日：取 ≤ schedLastDay 最近的 M/D 那年
  const parseDateLocal = (str) => {
    if (!str) return null;
    const s = str.toString().trim();
    // 完整 YYYY/M/D 或 YYY/M/D（民國年）
    const mFull = s.match(/^(\d{3,4})\/(\d{1,2})\/(\d{1,2})$/);
    if (mFull) {
      let yr = parseInt(mFull[1]);
      if (yr < 1900) yr += 1911; // 民國 → 西元
      return { year: yr, month: parseInt(mFull[2]), day: parseInt(mFull[3]), hasYear: true };
    }
    const mMD = s.match(/(\d+)\/(\d+)/);
    return mMD ? { year: null, month: parseInt(mMD[1]), day: parseInt(mMD[2]), hasYear: false } : null;
  };

  const schedFirstDay = new Date(schedYear, schedMonth - 1, 1);
  const schedLastDay  = new Date(schedYear, schedMonth, 0);

  // 離職日檢查：以「最近過去」原則推斷年份
  if (staffObj.leaveDate) {
    const ld = parseDateLocal(staffObj.leaveDate);
    if (ld) {
      let yr = ld.hasYear ? ld.year : schedYear;
      let leaveFullDate = new Date(yr, ld.month - 1, ld.day);
      // 若 M/D 解讀為當年比排班月最後一天還晚 → 表示「實際發生在去年」
      if (!ld.hasYear && leaveFullDate > schedLastDay) {
        yr = schedYear - 1;
        leaveFullDate = new Date(yr, ld.month - 1, ld.day);
      }
      if (leaveFullDate < schedFirstDay) return false;
    }
  }
  // 到職日檢查：相同邏輯
  if (staffObj.joinDate) {
    const jd = parseDateLocal(staffObj.joinDate);
    if (jd) {
      let yr = jd.hasYear ? jd.year : schedYear;
      let joinFullDate = new Date(yr, jd.month - 1, jd.day);
      if (!jd.hasYear && joinFullDate > schedLastDay) {
        yr = schedYear - 1;
        joinFullDate = new Date(yr, jd.month - 1, jd.day);
      }
      if (joinFullDate > schedLastDay) return false;
    }
  }
  return true;
}

/**
 * 取得個人統計
 * @param {string} personName 人員姓名
 * @param {string} sheetName  '__ALL__' 或特定月份
 */
function getPersonalStats(personName, sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const headers = GLOBAL_CONFIG.SHIFT_HEADERS;

    // 決定工作表清單
    let sheets = [];
    if (sheetName === '__ALL__') {
      const allNames = spreadsheet.getSheets().map(s => s.getName());
      const rocYear  = new Date().getFullYear() - 1911;
      const prefix   = rocNumToStr(rocYear) + '年';
      sheets = allNames.filter(n =>
        n !== EMAIL_SHEET_NAME && n.includes(prefix)
      );
      // 依月份排序
      const monthOrder = ['一月','二月','三月','四月','五月','六月',
                          '七月','八月','九月','十月','十一月','十二月'];
      sheets.sort((a, b) =>
        monthOrder.findIndex(m => a.includes(m)) -
        monthOrder.findIndex(m => b.includes(m))
      );
    } else {
      sheets = [sheetName];
    }

    // 各班別計數
    const counts = new Array(headers.length).fill(0);
    // 月份明細
    const monthlyDetail = [];
    // 找出該人擔任的所有日期（供行事曆顯示）
    const myDates = [];

    sheets.forEach(sName => {
      const sh = spreadsheet.getSheetByName(sName);
      if (!sh) return;
      const tz       = spreadsheet.getSpreadsheetTimeZone();
      const lastRow  = sh.getLastRow();
      if (lastRow < 2) return;
      const datesRng = sh.getRange('A2:B' + lastRow).getValues();
      const dataRng  = sh.getRange('C2:M' + Math.min(lastRow, 32)).getValues();

      const monthCounts = new Array(headers.length).fill(0);

      datesRng.forEach((dateRow, i) => {
        const row = dataRng[i];
        if (!row) return;
        let dateStr = dateRow[0] instanceof Date
          ? Utilities.formatDate(dateRow[0], tz, 'M/d')
          : (dateRow[0] ? dateRow[0].toString().trim() : '');
        const weekStr = dateRow[1] ? dateRow[1].toString().trim() : '';
        if (!dateStr) return;

        const myShifts = [];
        row.forEach((cell, ci) => {
          if (!cell) return;
          const val = cell.toString();
          if (val.split(',').map(v => v.trim()).includes(personName)) {
            counts[ci]++;
            monthCounts[ci]++;
            myShifts.push(headers[ci]);
          }
        });
        if (myShifts.length > 0) {
          myDates.push({ date: dateStr, week: weekStr, shifts: myShifts, sheet: sName });
        }
      });

      monthlyDetail.push({
        sheetName: sName,
        label: sName.replace('班表','').replace(/一百[一二三四五六七八九十]+年/, (m) => { const r = rocStrToNum(m.replace('年','')); return r > 0 ? r + '年' : m; }),
        counts: [...monthCounts],
        total: monthCounts.reduce((a,b) => a+b, 0)
      });
    });

    return {
      success:       true,
      personName,
      headers,
      counts,
      total:         counts.reduce((a,b) => a+b, 0),
      monthlyDetail: monthlyDetail.filter(m => m.total > 0),
      myDates,
      sheets
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * 取得所有人員姓名（供前端選人用）
 */
function getStaffNames() {
  return getStaffList().map(s => s.name);
}
// 規則：每次平日值班 = 1.5小時
//       每月結算：整數時數累計保留，小數（0.5）捨棄歸零
//       例：3次 = 4.5h → 保留4h，0.5h捨棄
//           4次 = 6.0h → 全部保留6h
// =============================================
function getOvertimeStats(sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const staff = getStaffList();
    const staffNames = staff.map(s => s.name);

    // 決定要統計的工作表清單（依月份排序）
    let sheetList = [];
    if (sheetName === '__ALL__') {
      const allNames = spreadsheet.getSheets().map(s => s.getName());
      sheetList = allNames
        .filter(n => {
          if (n === EMAIL_SHEET_NAME) return false;
          if (!n.includes('班表')) return false;
          const ym = parseYearMonthFromSheetName(n);
          return ym.valid && ym.year === new Date().getFullYear();
        })
        .map(n => ({ name: n, ...parseYearMonthFromSheetName(n) }))
        .sort((a, b) => (a.year * 100 + a.month) - (b.year * 100 + b.month));
    } else {
      const ym = parseYearMonthFromSheetName(sheetName);
      sheetList = [{ name: sheetName, ...ym }];
    }

    // 累計整數加班時數（跨月保留）
    const accumulated = {};  // { name: integer hours }
    staffNames.forEach(n => { accumulated[n] = 0; });

    // 每月明細
    const monthlyDetail = [];

    sheetList.forEach(({ name: sName, year, month }) => {
      const sh = spreadsheet.getSheetByName(sName);
      if (!sh) return;

      const datesRange = sh.getRange('A2:B32').getValues();
      const dutyCol = sh.getRange('C2:C32').getValues();
      const tz = spreadsheet.getSpreadsheetTimeZone();

      // 本月每人平日值班次數
      const monthCount = {};
      staffNames.forEach(n => { monthCount[n] = 0; });

      for (let i = 0; i < datesRange.length; i++) {
        const rawA = datesRange[i][0];
        let dateObj = null;
        if (rawA instanceof Date) {
          dateObj = rawA;
        } else {
          const aVal = rawA ? rawA.toString().trim() : '';
          const mMatch = aVal.match(/(\d+)\/(\d+)/);
          if (mMatch) {
            try { dateObj = new Date(year, parseInt(mMatch[1]) - 1, parseInt(mMatch[2])); } catch(e) {}
          }
        }
        if (!dateObj) continue;

        // 只算平日（非假日）
        if (isHoliday(dateObj)) continue;

        const person = dutyCol[i][0] ? dutyCol[i][0].toString().trim() : '';
        if (!person) continue;
        if (!monthCount.hasOwnProperty(person)) monthCount[person] = 0;
        monthCount[person]++;
      }

      // 計算本月原始加班時數 & 結算
      const monthRow = { label: `${year}/${month}`, sheetName: sName, persons: {} };
      staffNames.forEach(n => {
        const cnt = monthCount[n] || 0;
        const rawHours = cnt * 1.5;
        const intHours = Math.floor(rawHours);
        const forfeited = rawHours - intHours; // 0 or 0.5
        accumulated[n] = (accumulated[n] || 0) + intHours;
        monthRow.persons[n] = {
          count: cnt,
          rawHours: rawHours,
          intHours: intHours,
          forfeited: forfeited,
          accumulated: accumulated[n]
        };
      });
      monthlyDetail.push(monthRow);
    });

    // 整理回傳：最終累計 + 各月明細
    const summary = staffNames.map(n => ({
      name: n,
      accumulated: accumulated[n] || 0
    })).sort((a, b) => b.accumulated - a.accumulated);

    return {
      success: true,
      summary,
      monthlyDetail,
      staffNames
    };
  } catch(e) {
    return { success: false, message: e.message, summary: [], monthlyDetail: [], staffNames: [] };
  }
}

// =============================================
// 國定假日週出勤統計
// 統計每人在假日（週六日 + 政府放假日）值班(C欄)的次數
// 同時細分：週六、週日、補假日（政府公告非週末的補假）
// =============================================
function getHolidayAttendanceStats(sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const staff = getStaffList();
    const staffNames = staff.map(s => s.name);

    let sheetList = [];
    if (sheetName === '__ALL__') {
      const allNames = spreadsheet.getSheets().map(s => s.getName());
      sheetList = allNames
        .filter(n => {
          if (n === EMAIL_SHEET_NAME) return false;
          if (!n.includes('班表')) return false;
          const ym = parseYearMonthFromSheetName(n);
          return ym.valid && ym.year === new Date().getFullYear();
        })
        .map(n => ({ name: n, ...parseYearMonthFromSheetName(n) }))
        .sort((a, b) => (a.year * 100 + a.month) - (b.year * 100 + b.month));
    } else {
      const ym = parseYearMonthFromSheetName(sheetName);
      sheetList = [{ name: sheetName, ...ym }];
    }

    // 累計統計
    const total     = {};  // 全假日
    const bySat     = {};  // 週六
    const bySun     = {};  // 週日
    const bySubHol  = {};  // 補假（非週末的政府放假日）
    staffNames.forEach(n => { total[n]=0; bySat[n]=0; bySun[n]=0; bySubHol[n]=0; });

    // 每月明細
    const monthlyDetail = [];

    sheetList.forEach(({ name: sName, year, month }) => {
      const sh = spreadsheet.getSheetByName(sName);
      if (!sh) return;

      const datesRange = sh.getRange('A2:B32').getValues();
      const dutyCol    = sh.getRange('C2:C32').getValues();  // 值班欄
      const tz = spreadsheet.getSpreadsheetTimeZone();

      const monthCount    = {};
      const monthSat      = {};
      const monthSun      = {};
      const monthSubHol   = {};
      staffNames.forEach(n => { monthCount[n]=0; monthSat[n]=0; monthSun[n]=0; monthSubHol[n]=0; });

      for (let i = 0; i < datesRange.length; i++) {
        const rawA = datesRange[i][0];
        let dateObj = null;
        if (rawA instanceof Date) {
          dateObj = rawA;
        } else {
          const aVal = rawA ? rawA.toString().trim() : '';
          const mMatch = aVal.match(/(\d+)\/(\d+)/);
          if (mMatch) {
            try { dateObj = new Date(year, parseInt(mMatch[1])-1, parseInt(mMatch[2])); } catch(e) {}
          }
        }
        if (!dateObj) continue;
        if (!isHoliday(dateObj)) continue;  // 只算假日

        const person = dutyCol[i][0] ? dutyCol[i][0].toString().trim() : '';
        if (!person) continue;
        if (!monthCount.hasOwnProperty(person)) {
          monthCount[person]=0; monthSat[person]=0; monthSun[person]=0; monthSubHol[person]=0;
        }

        const dow = dateObj.getDay();
        const y = dateObj.getFullYear();
        const m2 = String(dateObj.getMonth()+1).padStart(2,'0');
        const d2 = String(dateObj.getDate()).padStart(2,'0');
        const dateKey = `${y}-${m2}-${d2}`;
        const isWeekend = (dow === 6 || dow === 0);
        const govYearSet = GOV_HOLIDAYS[y];
        const isGovHol  = govYearSet ? govYearSet.has(dateKey) : false;

        monthCount[person]++;
        if (dow === 6) monthSat[person]++;
        else if (dow === 0) monthSun[person]++;
        else if (isGovHol) monthSubHol[person]++;  // 補假（工作日但政府公告放假）
      }

      // 累計
      Object.keys(monthCount).forEach(n => {
        total[n]    = (total[n]||0)    + monthCount[n];
        bySat[n]    = (bySat[n]||0)    + monthSat[n];
        bySun[n]    = (bySun[n]||0)    + monthSun[n];
        bySubHol[n] = (bySubHol[n]||0) + monthSubHol[n];
      });

      monthlyDetail.push({
        label: `${year}/${month}`,
        sheetName: sName,
        count:    { ...monthCount },
        sat:      { ...monthSat },
        sun:      { ...monthSun },
        subHol:   { ...monthSubHol }
      });
    });

    const summary = staffNames.map(n => ({
      name:   n,
      total:  total[n]    || 0,
      sat:    bySat[n]    || 0,
      sun:    bySun[n]    || 0,
      subHol: bySubHol[n] || 0
    })).sort((a, b) => b.total - a.total);

    return { success: true, summary, monthlyDetail, staffNames };
  } catch(e) {
    return { success: false, message: e.message, summary: [], monthlyDetail: [], staffNames: [] };
  }
}

// ╔══════════════════════════════════════════════════════════════╗
// ║          Line Bot 班表查詢整合模組  ver4.1                    ║
// ║  「班表設定」工作表 R / S 欄設定：                            ║
// ║    R1  = Line Channel Access Token                           ║
// ║    R2  = 啟動搜尋關鍵字（前綴；空白=回應所有訊息）            ║
// ║    R3  = 模糊搜尋 TRUE/FALSE（空白預設 TRUE）                 ║
// ║    R4:R15 = 指定搜尋工作表（空白=自動全年115年班表）          ║
// ║    S1:S13 = 搜尋欄位 A~M（空白=搜尋全部欄位）                ║
// ╚══════════════════════════════════════════════════════════════╝

// ─── 讀取完整 Line Bot 設定 ─────────────────────────────────────
function getLineBotConfig() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);

  // R1：Token
  const token = sheet.getRange(GLOBAL_CONFIG.LINE_TOKEN_RANGE).getValue().toString().trim();

  // R2：啟動搜尋關鍵字（前綴）
  const keyword = sheet.getRange(GLOBAL_CONFIG.LINE_KEYWORD_RANGE).getValue().toString().trim();

  // R3：模糊搜尋（空白 → 預設 TRUE）
  const fuzzyRaw = sheet.getRange(GLOBAL_CONFIG.LINE_FUZZY_RANGE).getValue();
  const fuzzy = (fuzzyRaw === true || fuzzyRaw === 'TRUE' || fuzzyRaw === '' || fuzzyRaw === null);

  // R4:R15：搜尋工作表清單（空白 → 自動取全年115年班表）
  const sheetsRaw = sheet.getRange(GLOBAL_CONFIG.LINE_SHEETS_RANGE)
    .getValues().flat()
    .filter(s => s && s.toString().trim() !== '');
  const targetSheets = sheetsRaw.length > 0
    ? sheetsRaw.map(s => s.toString().trim())
    : getAllYear115Sheets();

  // S1:S13：搜尋欄位（空白 → 搜尋全部 A~M，索引 0~12）
  const colsRaw = sheet.getRange(GLOBAL_CONFIG.LINE_COLS_RANGE)
    .getValues().flat()
    .filter(s => s && s.toString().trim() !== '');
  const searchColIndices = colsRaw.length > 0
    ? colsRaw.map(c => c.toString().toUpperCase().charCodeAt(0) - 65).filter(n => n >= 0 && n <= 12)
    : null; // null = 搜尋全部

  return { token, keyword, fuzzy, targetSheets, searchColIndices };
}

// ─── 供前端顯示目前設定 ─────────────────────────────────────────
function getLineBotSettings() {
  const cfg = getLineBotConfig();
  const masked = cfg.token.length > 8
    ? cfg.token.substring(0, 6) + '****' + cfg.token.slice(-4)
    : (cfg.token ? '(已設定)' : '(未設定)');
  return {
    tokenMasked:      masked,
    hasToken:         cfg.token.length > 0,
    keyword:          cfg.keyword,
    fuzzy:            cfg.fuzzy,
    targetSheets:     cfg.targetSheets,
    totalSheets:      cfg.targetSheets.length,
    searchColIndices: cfg.searchColIndices,
    searchColsDisplay: cfg.searchColIndices
      ? cfg.searchColIndices.map(i => String.fromCharCode(65 + i)).join(', ')
      : '全部（A~M）'
  };
}

// ─── 供前端儲存設定 ─────────────────────────────────────────────
function updateLineBotSettings(adminPassword, token, keyword, fuzzy) {
  if (!verifyAdminPassword(adminPassword)) return { success: false, message: '管理員密碼錯誤。' };
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    if (token && token.trim() !== '') {
      sheet.getRange(GLOBAL_CONFIG.LINE_TOKEN_RANGE).setValue(token.trim());
    }
    sheet.getRange(GLOBAL_CONFIG.LINE_KEYWORD_RANGE).setValue(keyword.trim());
    sheet.getRange(GLOBAL_CONFIG.LINE_FUZZY_RANGE).setValue(
      fuzzy === true || fuzzy === 'true' ? true : false
    );
    return { success: true, message: '設定已儲存！' };
  } catch(e) {
    return { success: false, message: '儲存失敗：' + e.message };
  }
}

// ─── doPost：LINE Webhook ─────────────────────────────────────
function doPost(e) {
  var output = ContentService.createTextOutput('ok');
  try {
    var body = (e && e.postData) ? e.postData.contents : '';
    if (!body) return output;

    var userData;
    try {
      userData = JSON.parse(body);
    } catch (err) {
      console.error(err);
      Logger.log('[doPost] invalid json: ' + err.message);
      return ContentService
        .createTextOutput('{"ok":false,"error":"invalid json"}')
        .setMimeType(ContentService.MimeType.JSON);
    }
    var event    = userData.events && userData.events[0];
    if (!event || !event.message || event.message.type !== 'text') return output;

    var cfg = getLineBotConfig();
    if (!cfg.token) return output;

    var kw = event.message.text.trim();
    Logger.log('[doPost] 收到訊息: ' + kw + ' | 前綴: ' + cfg.keyword);

    // 前綴過濾
    if (cfg.keyword !== '' && kw.indexOf(cfg.keyword) !== 0) return output;
    kw = cfg.keyword ? kw.substring(cfg.keyword.length).trim() : kw;

    if (!kw) {
      sendLineReply(cfg.token, event.replyToken, [makeHelpMessage(cfg.keyword)]);
      return output;
    }

    // ── AQI 空氣品質查詢 ──────────────────────────────────────────
    var kwLower = kw.toLowerCase().trim();
    Logger.log('[doPost] kwLower=' + kwLower);
    if (kwLower === 'aqi' || kwLower === '空氣' || kwLower === '空氣品質' || kwLower === '善化' || kw === 'AQI') {
      sendLineReply(cfg.token, event.replyToken, [makeAqiMessage()]);
      return output;
    }

    // ── AI 問答（關鍵字：問 XXX）────────────────────────────────
    if (kw.indexOf('問 ') === 0 || kw.indexOf('問　') === 0) {
      var question = kw.substring(2).trim();
      if (question) {
        var aiReply = askGemini(question);
        sendLineReply(cfg.token, event.replyToken, [{ type: 'text', text: aiReply }]);
        return output;
      }
    }

    // ── 日期時間查詢 ──────────────────────────────────────────────
    if (kwLower === '時間' || kwLower === '日期' || kwLower === '現在' || kwLower === '幾點' || kwLower === 'time') {
      sendLineReply(cfg.token, event.replyToken, [makeDateTimeMessage()]);
      return output;
    }

    // 解析月份後綴
    var parsed      = parseKwWithMonth(kw);
    var baseKw      = parsed.baseKw;
    var monthFilter = parsed.monthFilter;
    Logger.log('[doPost] baseKw=' + baseKw + ' monthFilter=' + monthFilter);

    var targetSheets = monthFilter
      ? cfg.targetSheets.filter(function(s){ return s.indexOf(monthFilter) !== -1; })
      : cfg.targetSheets;

    var results = lineSearchSchedule(baseKw, targetSheets, cfg.fuzzy, cfg.searchColIndices);
    Logger.log('[doPost] 搜尋結果: ' + results.length + ' 筆');

    var msgs;
    if (results.length === 0) {
      msgs = [makeNoResultMessage(baseKw, cfg.keyword)];
    } else if (!monthFilter && results.length > 7) {
      var byMonth = {};
      for (var i = 0; i < results.length; i++) {
        var sn = results[i].sheetName;
        byMonth[sn] = (byMonth[sn] || 0) + 1;
      }
      msgs = [makeMonthSelectMessage(baseKw, cfg.keyword, byMonth)];
    } else {
      var bubbles = results.map(function(r){ return makeScheduleFlexBubble(r, baseKw); });
      msgs = buildLineCarouselMessages(bubbles, baseKw);
    }
    sendLineReply(cfg.token, event.replyToken, msgs);

  } catch(err) {
    Logger.log('[doPost ERROR] ' + err.message + ' | ' + err.stack);
  }
  return output;
}

/**
 * 由時間觸發器每 1 分鐘呼叫一次，處理 LINE 訊息
 * 設定方式：GAS 編輯器 → 觸發器 → 新增觸發器
 *   函式：processLinePending
 *   事件來源：時間驅動
 *   類型：分鐘計時器
 *   間隔：每 1 分鐘
 */
function processLinePending() {
  const cache = CacheService.getScriptCache();
  const keysRaw = cache.get('line_pending_keys');
  if (!keysRaw) return;

  let keys;
  try { keys = JSON.parse(keysRaw); } catch(e) { return; }
  if (!keys || keys.length === 0) return;

  // 清除 key 清單（避免重複處理）
  cache.remove('line_pending_keys');

  keys.forEach(ts => {
    const body = cache.get('line_pending_' + ts);
    cache.remove('line_pending_' + ts);
    if (!body) return;
    try {
      processLineEvent(body);
    } catch(err) {
      Logger.log('[processLinePending] ' + err.message);
    }
  });
}

/**
 * 實際處理 LINE 事件
 */
function processLineEvent(bodyStr) {
  var cfg = getLineBotConfig();
  if (!cfg.token) return;

  var userData;
  try {
    userData = JSON.parse(bodyStr);
  } catch (err) {
    console.error(err);
    Logger.log('[processLineEvent] invalid json: ' + err.message);
    return;
  }
  var event    = userData.events && userData.events[0];
  if (!event) return;

  var replyToken = event.replyToken;
  if (!replyToken) return;
  if (!event.message || event.message.type !== 'text') return;

  var searchContent = event.message.text.trim();
  if (!searchContent) return;

  // 關鍵字前綴過濾
  if (cfg.keyword !== '') {
    if (searchContent.indexOf(cfg.keyword) !== 0) return;
    searchContent = searchContent.substring(cfg.keyword.length).trim();
    if (!searchContent) {
      sendLineReply(cfg.token, replyToken, [makeHelpMessage(cfg.keyword)]);
      return;
    }
  }

  // 搜尋班表（支援月份後綴篩選）
  var parsed      = parseKwWithMonth(searchContent);
  var baseKw      = parsed.baseKw;
  var monthFilter = parsed.monthFilter;

  var targetSheets = monthFilter
    ? cfg.targetSheets.filter(function(s){ return s.indexOf(monthFilter) !== -1; })
    : cfg.targetSheets;

  var results = lineSearchSchedule(baseKw, targetSheets, cfg.fuzzy, cfg.searchColIndices);

  if (results.length === 0) {
    sendLineReply(cfg.token, replyToken, [makeNoResultMessage(baseKw, cfg.keyword)]);
    return;
  }

  if (!monthFilter && results.length > 7) {
    var byMonth = {};
    for (var i = 0; i < results.length; i++) {
      var sn = results[i].sheetName;
      byMonth[sn] = (byMonth[sn] || 0) + 1;
    }
    sendLineReply(cfg.token, replyToken, [makeMonthSelectMessage(baseKw, cfg.keyword, byMonth)]);
    return;
  }

  var bubbles  = results.map(function(r){ return makeScheduleFlexBubble(r, baseKw); });
  var messages = buildLineCarouselMessages(bubbles, baseKw);
  sendLineReply(cfg.token, replyToken, messages);
}

// ─── 班表搜尋核心 ───────────────────────────────────────────────
function lineSearchSchedule(keyword, sheetNames, fuzzy, searchColIndices) {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const tz = spreadsheet.getSpreadsheetTimeZone();
  const kw = keyword.toLowerCase();
  const results = [];

  sheetNames.forEach(sName => {
    const sheet = spreadsheet.getSheetByName(sName);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const headers  = sheet.getRange('A1:M1').getValues()[0];
    const bodyData = sheet.getRange(2, 1, lastRow - 1, 13).getValues();

    bodyData.forEach(row => {
      let dateStr = row[0] instanceof Date
        ? Utilities.formatDate(row[0], tz, 'M/d')
        : (row[0] ? row[0].toString().trim() : '');
      if (!dateStr) return;

      const weekStr  = row[1] ? row[1].toString().trim() : '';
      const fullDate = `${dateStr} ${weekStr}`.trim();

      // ── 依 searchColIndices 決定搜尋範圍 ──
      let searchText;
      if (!searchColIndices) {
        searchText = [fullDate, ...row.slice(2).map(c => c ? c.toString() : '')].join(' ');
      } else {
        searchText = searchColIndices.map(ci => {
          if (ci === 0) return dateStr;
          if (ci === 1) return weekStr;
          return row[ci] ? row[ci].toString() : '';
        }).join(' ');
      }

      // ── 模糊 vs 精確 ──
      const matched = fuzzy
        ? searchText.toLowerCase().indexOf(kw) !== -1
        : searchText.toLowerCase().split(/[\s,]+/).some(t => t === kw);

      if (!matched) return;

      const shifts = [];
      for (let ci = 2; ci < 13; ci++) {
        const val = row[ci] ? row[ci].toString().trim() : '';
        if (val) shifts.push({ header: headers[ci], value: val });
      }
      results.push({ sheetName: sName, dateStr, weekStr, fullDate, shifts });
    });
  });

  return results;
}

// ─── 班別圖示 / 顏色 ────────────────────────────────────────────
// ─── 班別樣式對應（圖示 + 顏色，與推播系統一致）────────────────
const LINE_SHIFT_STYLE = {
  '值班':      { icon: '👤', color: '#E74C3C' },
  '協助掛號':  { icon: '📋', color: '#3498DB' },
  '門診':      { icon: '🏥', color: '#9B59B6' },
  '流注1':     { icon: '💉', color: '#E67E22' },
  '流注2':     { icon: '💉', color: '#E67E22' },
  '預登1':     { icon: '📝', color: '#27AE60' },
  '預登2':     { icon: '📝', color: '#27AE60' },
  '預注1':     { icon: '💉', color: '#F39C12' },
  '預注2':     { icon: '💉', color: '#F39C12' },
  '卡介苗':    { icon: '🩹', color: '#E74C3C' },
  '登革熱二線':{ icon: '🦟', color: '#8E44AD' }
};

function getLineShiftStyle(name) {
  return LINE_SHIFT_STYLE[name] || { icon: '📌', color: '#2C3E50' };
}

// ─── 灰色分隔線（節與節之間）─────────────────────────────────────
function lineSep(marginVal) {
  return { type: 'separator', margin: marginVal || 'md', color: '#E0E0E0' };
}

// ─── 白色圓角卡片容器 ─────────────────────────────────────────────
function lineCard(contents) {
  return {
    type: 'box', layout: 'vertical',
    backgroundColor: '#FFFFFF',
    cornerRadius: '8px',
    paddingAll: '15px',
    contents: contents
  };
}

// ─── 卡片標題列（圖示 + 粗體標題）───────────────────────────────
function lineCardTitle(icon, title, color) {
  return {
    type: 'box', layout: 'horizontal',
    contents: [
      { type: 'text', text: icon,  size: 'lg', color: color || '#1DB446', flex: 0 },
      { type: 'text', text: title, size: 'lg', weight: 'bold', color: '#2C3E50', flex: 1, margin: 'sm' }
    ]
  };
}

// ─── 建立單日班表 Flex Bubble（仿推播卡片樣式）───────────────────
function makeScheduleFlexBubble(r, searchKeyword) {
  const isWknd    = r.weekStr && (r.weekStr.includes('六') || r.weekStr.includes('日'));
  const headerBg  = isWknd ? '#C0392B' : '#1DB446';
  const now       = new Date();
  const ts        = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')} ${String(now.getHours()).padStart(2,'0')}:${String(now.getMinutes()).padStart(2,'0')}`;

  // ── 判斷關鍵字是否為人名（含在某個班別人名中）─────────────────
  // 若搜尋詞與某人名完全吻合（或包含其中），則在該人名後加 ← 標記
  const kw = (searchKeyword || '').trim();
  const markName = (nameVal) => {
    if (!kw) return nameVal;
    // 逗號分隔多人的情況
    return nameVal.split(',').map(n => {
      const nm = n.trim();
      return nm.includes(kw) || kw.includes(nm)
        ? nm + ' ←'
        : nm;
    }).join(', ');
  };

  const dutyShift   = r.shifts.find(s => s.header === '值班');
  const otherShifts = r.shifts.filter(s => s.header !== '值班');

  const dutyCardContents = [
    lineCardTitle('👥', '勤務資訊', '#1DB446'),
    lineSep('md'),
    {
      type: 'box', layout: 'horizontal', margin: 'md',
      contents: [
        { type: 'text', text: '📅', size: 'md', flex: 0 },
        { type: 'text', text: `${r.dateStr} ${r.weekStr}`, weight: 'bold', size: 'md', color: '#3498DB', flex: 1, margin: 'sm' }
      ]
    }
  ];

  // 值班（突出顯示）
  if (dutyShift) {
    const markedVal = markName(dutyShift.value);
    const isMarked  = markedVal !== dutyShift.value;
    dutyCardContents.push({
      type: 'box', layout: 'horizontal', margin: 'sm',
      contents: [
        { type: 'text', text: '👤', size: 'md', flex: 0 },
        {
          type: 'text',
          text: `值班：${markedVal}`,
          weight: 'bold', size: 'md',
          color: isMarked ? '#E67E22' : '#E74C3C',
          flex: 1, margin: 'sm'
        }
      ]
    });
  }

  // 其他班別
  otherShifts.forEach(s => {
    const style     = getLineShiftStyle(s.header);
    const isDengue  = (s.header === '登革熱二線');
    const markedVal = markName(s.value);
    const isMarked  = markedVal !== s.value;
    dutyCardContents.push({
      type: 'box', layout: 'horizontal', margin: 'sm',
      contents: [
        { type: 'text', text: style.icon, size: isDengue ? 'xs' : 'sm', flex: 0 },
        {
          type: 'text',
          text: `${s.header}：${markedVal}`,
          size: isDengue ? 'xs' : 'sm',
          color: isMarked ? '#E67E22' : style.color,
          weight: isMarked ? 'bold' : 'regular',
          flex: 1, margin: 'sm', wrap: true
        }
      ]
    });
  });

  if (r.shifts.length === 0) {
    dutyCardContents.push({
      type: 'text', text: '本日無排班資料',
      size: 'sm', color: '#AAAAAA', align: 'center', margin: 'md'
    });
  }

  const bodyContents = [
    lineCard(dutyCardContents),
    lineSep('md'),
    {
      type: 'box', layout: 'horizontal',
      contents: [
        { type: 'text', text: '📂', size: 'xs', flex: 0 },
        { type: 'text', text: r.sheetName, size: 'xs', color: '#7F8C8D', flex: 1, margin: 'sm' }
      ]
    }
  ];

  return {
    type: 'bubble',
    size: 'mega',
    // ── Header ─────────────────────────────────────────────────
    header: {
      type: 'box', layout: 'vertical',
      backgroundColor: headerBg,
      paddingAll: '20px',
      contents: [
        {
          type: 'text',
          text: '📋 班表查詢結果',
          weight: 'bold', color: '#FFFFFF', size: 'xl', align: 'center'
        }
      ]
    },
    // ── Body ───────────────────────────────────────────────────
    body: {
      type: 'box', layout: 'vertical',
      paddingAll: '0px', spacing: 'none',
      contents: bodyContents
    },
    // ── Footer ─────────────────────────────────────────────────
    footer: {
      type: 'box', layout: 'vertical',
      paddingAll: '10px',
      contents: [
        {
          type: 'text',
          text: `📅 查詢時間：${ts}`,
          size: 'xs', color: '#7F8C8D', align: 'center'
        }
      ]
    }
  };
}

// ─── 查無結果（卡片樣式）──────────────────────────────────────────
function makeNoResultMessage(keyword, prefix) {
  const examples = prefix
    ? [
        { q: `${prefix}王小明`, hint: '查詢個人所有班表' },
        { q: `${prefix}3/20`,   hint: '查詢當日班表' },
        { q: `${prefix}卡介苗`, hint: '查詢班別排班' }
      ]
    : [
        { q: '王小明', hint: '查詢個人所有班表' },
        { q: '3/20',   hint: '查詢當日班表' },
        { q: '卡介苗', hint: '查詢班別排班' }
      ];

  const rows = examples.map(e => ({
    type: 'box', layout: 'horizontal', margin: 'sm',
    contents: [
      { type: 'text', text: '✅', size: 'sm', flex: 0 },
      { type: 'text', text: e.q,    size: 'sm', color: '#27AE60', weight: 'bold', flex: 2, margin: 'sm' },
      { type: 'text', text: e.hint, size: 'xs', color: '#7F8C8D', flex: 3, margin: 'sm', wrap: true }
    ]
  }));

  return {
    type: 'flex',
    altText: `查無「${keyword}」的班表資料`,
    contents: {
      type: 'bubble',
      size: 'mega',
      header: {
        type: 'box', layout: 'vertical',
        backgroundColor: '#E74C3C', paddingAll: '20px',
        contents: [
          { type: 'text', text: '🔍 查無班表資料', weight: 'bold', color: '#FFFFFF', size: 'xl', align: 'center' }
        ]
      },
      body: {
        type: 'box', layout: 'vertical', paddingAll: '0px',
        contents: [
          lineCard([
            lineCardTitle('❓', `查無「${keyword}」的結果`, '#E74C3C'),
            lineSep('md'),
            { type: 'text', text: '💡 搜尋技巧', weight: 'bold', size: 'sm', color: '#2C3E50', margin: 'md' },
            ...rows
          ])
        ]
      }
    }
  };
}

// ─── 使用說明（卡片樣式）──────────────────────────────────────────
function makeHelpMessage(prefix) {
  var p = prefix || '';
  var examples = [
    { q: p + '王小明',       hint: '查詢個人所有班表' },
    { q: p + '王小明 三月',  hint: '查詢個人三月班表' },
    { q: p + '3/20',         hint: '查詢當日班表' },
    { q: p + '空氣',         hint: '善化站 AQI 空氣品質' },
    { q: p + '問 臺南市佳里區推薦午餐店名地址電話', hint: 'AI 問答' },
    { q: p + '時間',         hint: '查詢目前日期時間' }
  ];

  var rows = examples.map(function(e) {
    return {
      type: 'box', layout: 'horizontal', margin: 'sm',
      contents: [
        { type: 'text', text: '✅', size: 'sm', flex: 0 },
        { type: 'text', text: e.q,    size: 'sm', color: '#27AE60', weight: 'bold', flex: 3, margin: 'sm' },
        { type: 'text', text: e.hint, size: 'xs', color: '#7F8C8D', flex: 4, margin: 'sm', wrap: true }
      ]
    };
  });

  return {
    type: 'flex',
    altText: '班表查詢說明',
    contents: {
      type: 'bubble',
      size: 'mega',
      header: {
        type: 'box', layout: 'vertical',
        backgroundColor: '#1DB446', paddingAll: '20px',
        contents: [
          { type: 'text', text: '🏥 班表查詢系統', weight: 'bold', color: '#FFFFFF', size: 'xl', align: 'center' }
        ]
      },
      body: {
        type: 'box', layout: 'vertical', paddingAll: '0px',
        contents: [
          lineCard([
            lineCardTitle('🔍', '搜尋說明', '#3498DB'),
            lineSep('md'),
            {
              type: 'box', layout: 'horizontal', margin: 'md',
              contents: [
                { type: 'text', text: '🔑', size: 'sm', flex: 0 },
                {
                  type: 'text',
                  text: prefix ? ('前綴關鍵字：「' + prefix + '」') : '無需前綴，直接輸入關鍵字',
                  size: 'sm', weight: 'bold', color: '#1DB446', flex: 1, margin: 'sm'
                }
              ]
            },
            lineSep('md'),
            { type: 'text', text: '📝 使用範例', weight: 'bold', size: 'sm', color: '#2C3E50', margin: 'md' }
          ].concat(rows)),
          lineSep('md'),
          lineCard([
            lineCardTitle('📂', '搜尋範圍', '#9B59B6'),
            lineSep('md'),
            { type: 'text', text: (new Date().getFullYear()-1911)+'年全年班表（1月～12月）', size: 'sm', color: '#7F8C8D', margin: 'sm' }
          ])
        ]
      }
    }
  };
}

// ─── 組裝輪播訊息 ────────────────────────────────────────────────
// ─── 月份縮寫對照表（供關鍵字篩選用）───────────────────────────
var MONTH_NAMES = ['一月','二月','三月','四月','五月','六月',
                   '七月','八月','九月','十月','十一月','十二月'];

/**
 * 從 results 分組產生「月份選擇」Flex 卡片
 */
function makeMonthSelectMessage(keyword, prefix, resultsByMonth) {
  var totalCount = 0;
  var keys = Object.keys(resultsByMonth);
  keys.forEach(function(k){ totalCount += resultsByMonth[k]; });

  // 依月份順序排列
  keys.sort(function(a, b) {
    var ai = -1, bi = -1;
    for (var n = 0; n < MONTH_NAMES.length; n++) {
      if (ai === -1 && a.indexOf(MONTH_NAMES[n]) !== -1) ai = n;
      if (bi === -1 && b.indexOf(MONTH_NAMES[n]) !== -1) bi = n;
    }
    return ai - bi;
  });

  var p = prefix || '';
  var lines = ['🔍 共找到 ' + totalCount + ' 筆，請加上月份再查詢：', ''];
  keys.forEach(function(sName) {
    var label = sName.replace(/一百[一二三四五六七八九十]+年/,'').replace('班表','').trim();
    lines.push('👉 ' + p + keyword + ' ' + label);
  });
  lines.push('');
  lines.push('⚠️ 注意：名字與月份之間要有空格');

  return { type: 'text', text: lines.join('\n') };
}

/**
 * 解析關鍵字中的月份後綴
 * 例："鄭兆鑫 三月" → {baseKw:"鄭兆鑫", monthFilter:"三月"}
 */
function parseKwWithMonth(keyword) {
  for (var i = 0; i < MONTH_NAMES.length; i++) {
    var m = MONTH_NAMES[i];
    // 以空格+月份 或 直接結尾月份
    if (keyword.length > m.length) {
      if (keyword.slice(-m.length) === m) {
        var baseKw = keyword.slice(0, keyword.length - m.length).trim();
        if (baseKw.length > 0) return { baseKw: baseKw, monthFilter: m };
      }
    }
  }
  return { baseKw: keyword, monthFilter: null };
}

function buildLineCarouselMessages(bubbles, keyword) {
  const MAX_PER = 10, MAX_MSG = 5;
  const messages = [];
  for (let i = 0; i < bubbles.length && messages.length < MAX_MSG; i += MAX_PER) {
    messages.push({
      type: 'flex',
      altText: `🔍「${keyword}」班表查詢結果`,
      contents: { type: 'carousel', contents: bubbles.slice(i, i + MAX_PER) }
    });
  }
  const maxShow = MAX_PER * MAX_MSG;
  if (bubbles.length > maxShow) {
    messages.push({
      type: 'text',
      text: `⚠️ 共找到 ${bubbles.length} 筆結果，僅顯示前 ${maxShow} 筆。\n請輸入更精確的關鍵字縮小範圍。`
    });
  }
  return messages;
}

// ─── 日期時間查詢 ────────────────────────────────────────────────
function makeDateTimeMessage() {
  var now  = new Date();
  var tz   = 'Asia/Taipei';
  var DAYS = ['日','一','二','三','四','五','六'];

  var dateStr = Utilities.formatDate(now, tz, 'yyyy/MM/dd');
  var timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');
  var dow     = DAYS[now.getDay()];
  var rocYear = now.getFullYear() - 1911;

  // 農曆節氣 / 假日提示（簡易版）
  var hour = parseInt(Utilities.formatDate(now, tz, 'HH'));
  var greeting;
  if      (hour >= 5  && hour < 12) greeting = '☀️ 早安';
  else if (hour >= 12 && hour < 14) greeting = '🌤 午安';
  else if (hour >= 14 && hour < 18) greeting = '🌤 午後';
  else if (hour >= 18 && hour < 22) greeting = '🌙 晚安';
  else                              greeting = '🌃 深夜';

  return {
    type: 'flex',
    altText: '現在時間 ' + dateStr + ' ' + timeStr,
    contents: {
      type: 'bubble',
      size: 'kilo',
      body: {
        type: 'box',
        layout: 'vertical',
        paddingAll: '18px',
        spacing: 'sm',
        contents: [
          {
            type: 'text',
            text: greeting,
            size: 'sm',
            color: '#888888'
          },
          {
            type: 'text',
            text: timeStr,
            size: 'xxl',
            weight: 'bold',
            color: '#2C3E50'
          },
          { type: 'separator', margin: 'sm' },
          {
            type: 'text',
            text: '民國 ' + rocYear + ' 年　' + dateStr + '　週' + dow,
            size: 'sm',
            color: '#555555',
            wrap: true
          }
        ]
      }
    }
  };
}

// ─── AQI 空氣品質查詢（善化站）────────────────────────────────
function getAqiData() {
  try {
    var url = 'https://data.moenv.gov.tw/api/v2/aqx_p_432?api_key=846e44e1-8cc5-4893-ad87-c79d2d383706&limit=1000&sort=ImportDate%20desc&format=json';
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    var data = JSON.parse(resp.getContentText());
    var records = Array.isArray(data) ? data : (data.records || []);
    for (var i = 0; i < records.length; i++) {
      if (records[i].sitename === '善化') return records[i];
    }
    return null;
  } catch(e) {
    Logger.log('[getAqiData] ' + e.message);
    return null;
  }
}

function makeAqiMessage() {
  var rec = getAqiData();
  if (!rec) {
    return { type: 'text', text: '⚠️ 無法取得善化站 AQI 資料，請稍後再試。' };
  }

  var aqi    = rec.aqi    ? rec.aqi.toString().trim()    : '--';
  var status = rec.status ? rec.status.toString().trim() : '--';
  var pm25   = rec.pm2_5  || rec['pm2.5'] || '--';
  var pm10   = rec.pm10   || '--';
  var time   = rec.datacreationdate || rec.ImportDate || '';

  var aqiNum = parseInt(aqi);
  var color;
  if      (isNaN(aqiNum))    color = '#888888';
  else if (aqiNum <= 50)     color = '#27AE60';
  else if (aqiNum <= 100)    color = '#E6AC00';
  else if (aqiNum <= 150)    color = '#E67E22';
  else if (aqiNum <= 200)    color = '#E74C3C';
  else if (aqiNum <= 300)    color = '#9B59B6';
  else                       color = '#7D3C98';

  var now  = new Date();
  var queryTime = (now.getMonth()+1) + '/' + now.getDate()
    + ' ' + now.getHours().toString().padStart(2,'0')
    + ':' + now.getMinutes().toString().padStart(2,'0');

  return {
    type: 'flex',
    altText: '善化站 AQI ' + aqi + ' ' + status,
    contents: {
      type: 'bubble',
      size: 'kilo',
      body: {
        type: 'box',
        layout: 'vertical',
        spacing: 'sm',
        paddingAll: '16px',
        contents: [
          {
            type: 'text',
            text: '🌬️ 善化站空氣品質',
            size: 'sm',
            color: '#555555',
            weight: 'bold'
          },
          {
            type: 'text',
            text: 'AQI　' + aqi + '　' + status,
            size: 'xl',
            weight: 'bold',
            color: color
          },
          { type: 'separator', margin: 'sm' },
          {
            type: 'text',
            text: 'PM2.5　' + pm25 + ' μg/m³　PM10　' + pm10 + ' μg/m³',
            size: 'xs',
            color: '#888888',
            wrap: true
          },
          {
            type: 'text',
            text: '🕐 查詢時間　' + queryTime,
            size: 'xs',
            color: '#AAAAAA'
          }
        ]
      }
    }
  };
}

// ─── Google AI Studio（Gemini）問答 ──────────────────────────
// API Key 存於「班表設定」R18 儲存格
function askGemini(question) {
  try {
    var sheet  = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    var apiKey = sheet.getRange(GLOBAL_CONFIG.GEMINI_API_KEY_RANGE).getValue().toString().trim();
    if (!apiKey) return '⚠️ 尚未設定 Gemini API Key（班表設定 R18）';

    // 依序嘗試不同模型（遇到 429/404 自動換下一個）
    var models = [
      'gemini-2.0-flash',
      'gemini-2.5-flash',
      'gemini-2.0-flash-lite'
    ];

    var body = {
      contents: [{ parts: [{ text: question }] }],
      system_instruction: {
        parts: [{ text: '請用繁體中文回答。回答必須簡潔，200字以內。不要使用Markdown格式（不要用**粗體**、*斜體*、#標題、-列表符號等），直接用純文字回答。' }]
      },
      generationConfig: { maxOutputTokens: 800, temperature: 0.7 }
    };

    for (var i = 0; i < models.length; i++) {
      var url  = 'https://generativelanguage.googleapis.com/v1beta/models/'
               + models[i] + ':generateContent?key=' + apiKey;
      var resp = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(body),
        muteHttpExceptions: true
      });
      var code = resp.getResponseCode();
      Logger.log('[askGemini] 模型=' + models[i] + ' HTTP=' + code);

      if (code === 429 || code === 404) continue;  // 超限或模型不存在 → 試下一個

      if (code !== 200) {
        return '⚠️ AI 查詢失敗（HTTP ' + code + '）';
      }

      var data = JSON.parse(resp.getContentText());
      var text = data.candidates
        && data.candidates[0]
        && data.candidates[0].content
        && data.candidates[0].content.parts
        && data.candidates[0].content.parts[0]
        && data.candidates[0].content.parts[0].text;

      return text ? ('🤖 ' + text.trim()) : '⚠️ AI 未回傳內容';
    }

    // 三個模型都 429 / 404 或 continue 路徑掉出迴圈 → 保底回傳
    return '⚠️ AI 服務暫時忙碌，請稍後再試';

  } catch(e) {
    Logger.log('[askGemini] ' + e.message);
    return '⚠️ AI 發生錯誤：' + e.message;
  }
}

// ─── 發送 Line 回覆 ──────────────────────────────────────────────
function sendLineReply(token, replyToken, messages) {
  var resp = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + token
    },
    payload: JSON.stringify({ replyToken: replyToken, messages: messages }),
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  Logger.log('[sendLineReply] HTTP ' + code + ' | ' + resp.getContentText().substring(0, 200));
  if (code < 200 || code >= 300) {
    console.error('LINE API failed', code, resp.getContentText());
    return false;
  }
  return true;
}

// =============================================
// LINE Bot 診斷工具
// 在 GAS 編輯器選取 diagLineBot → 執行
// 查看「執行記錄」即可看到所有問題
// =============================================
function diagLineBot() {
  const results = [];
  const pass = '✅';
  const fail = '❌';
  const warn = '⚠️';

  // 1. 讀取設定
  let cfg;
  try {
    cfg = getLineBotConfig();
    results.push(pass + ' getLineBotConfig() 執行成功');
  } catch(e) {
    results.push(fail + ' getLineBotConfig() 失敗：' + e.message);
    Logger.log(results.join('\n'));
    return;
  }

  // 2. Token 檢查
  if (!cfg.token) {
    results.push(fail + ' R1（Token）未設定 → LINE Bot 無法運作');
  } else if (cfg.token.length < 100) {
    results.push(warn + ' R1（Token）長度偏短（' + cfg.token.length + ' 字元），請確認是否完整');
  } else {
    results.push(pass + ' R1（Token）已設定（長度：' + cfg.token.length + ' 字元）');
  }

  // 3. 關鍵字前綴
  results.push((cfg.keyword !== '' ? pass : warn)
    + ' R2（關鍵字前綴）：' + (cfg.keyword || '（空白，回應所有訊息）'));

  // 4. 模糊搜尋
  results.push(pass + ' R3（模糊搜尋）：' + (cfg.fuzzy ? '開啟' : '關閉'));

  // 5. 搜尋工作表
  if (!cfg.targetSheets || cfg.targetSheets.length === 0) {
    results.push(fail + ' 找不到任何可搜尋的班表工作表！');
  } else {
    results.push(pass + ' 搜尋工作表共 ' + cfg.targetSheets.length + ' 個：'
      + cfg.targetSheets.slice(0, 3).join('、')
      + (cfg.targetSheets.length > 3 ? '…' : ''));
  }

  // 6. 搜尋欄位
  results.push(pass + ' 搜尋欄位：'
    + (cfg.searchColIndices ? cfg.searchColIndices.map(i => String.fromCharCode(65+i)).join(',') : '全部 A~M'));

  // 7. 驗證 Token 是否有效（呼叫 LINE profile API）
  if (cfg.token) {
    try {
      const resp = UrlFetchApp.fetch('https://api.line.me/v2/bot/info', {
        method: 'get',
        headers: { 'Authorization': 'Bearer ' + cfg.token },
        muteHttpExceptions: true
      });
      const code = resp.getResponseCode();
      if (code === 200) {
        const info = JSON.parse(resp.getContentText());
        results.push(pass + ' Token 驗證成功！Bot 名稱：' + (info.displayName || '未知'));
      } else {
        const body = resp.getContentText();
        results.push(fail + ' Token 驗證失敗（HTTP ' + code + '）：' + body);
        results.push('  → 請至 LINE Developers 重新產生 Channel Access Token 並填入 R1');
      }
    } catch(e) {
      results.push(warn + ' Token 驗證網路錯誤：' + e.message);
    }
  }

  // 8. 測試搜尋功能（搜尋第一個工作表是否能讀到資料）
  if (cfg.targetSheets && cfg.targetSheets.length > 0) {
    try {
      const testSheet = cfg.targetSheets[0];
      const testResults = lineSearchSchedule('值班', [testSheet], true, null);
      results.push(pass + ' 搜尋測試（在「' + testSheet + '」搜尋「值班」）：找到 '
        + testResults.length + ' 筆');
    } catch(e) {
      results.push(fail + ' 搜尋測試失敗：' + e.message);
    }
  }

  // 9. 部署提示
  results.push('');
  results.push('── 部署設定確認清單 ──');
  results.push('□ 部署類型：網頁應用程式');
  results.push('□ 執行身分：我（' + Session.getActiveUser().getEmail() + '）');
  results.push('□ 存取權限：任何人（含匿名）');
  results.push('□ 已點「建立新版本」後再部署（不是用舊版本）');
  results.push('□ LINE Developers Webhook URL 填的是 /exec 結尾（不是 /dev）');
  results.push('□ LINE Developers 已開啟「Use webhook」');
  results.push('□ LINE Developers 已點「Verify」且顯示 Success');

  const output = results.join('\n');
  Logger.log(output);
  return output;
}

// ── 班表欄位設定讀取與儲存（N欄：職務名稱、候選人員來源）──────────
// 試算表 N1:N11 存自訂職務名稱（空白=用預設 SHIFT_HEADERS）
// 格式：每格一行，C欄~M欄共11個職務
function getShiftColumnConfig() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    // C1:M1 直接讀班表第一列的欄位標題（最權威的來源）
    const allSheets = SpreadsheetApp.openById(SHEET_ID).getSheets()
      .filter(s => s.getName() !== EMAIL_SHEET_NAME && s.getName().includes('年'));
    let headers = GLOBAL_CONFIG.SHIFT_HEADERS.slice();
    if (allSheets.length > 0) {
      const h = allSheets[0].getRange('C1:M1').getValues()[0];
      if (h.some(v => v)) headers = h.map((v,i) => v ? v.toString().trim() : GLOBAL_CONFIG.SHIFT_HEADERS[i] || '');
    }
    return { success: true, headers };
  } catch(e) { return { success: false, headers: GLOBAL_CONFIG.SHIFT_HEADERS.slice() }; }
}
// 在 GAS 編輯器手動執行一次即可修復所有已建立的班表
function fixAllSheetDateFormats() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const tz = spreadsheet.getSpreadsheetTimeZone();
  const sheets = spreadsheet.getSheets();
  let fixed = 0;
  sheets.forEach(sh => {
    const name = sh.getName();
    if (!name.includes('年') || name === EMAIL_SHEET_NAME) return;
    const range = sh.getRange('A2:A32');
    range.setNumberFormat('@'); // 強制純文字
    const vals = range.getValues();
    const newVals = vals.map(row => {
      const v = row[0];
      if (!v) return [''];
      if (v instanceof Date) {
        return [Utilities.formatDate(v, tz, 'M/d')];
      }
      const m = v.toString().match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
      if (m) return [parseInt(m[2]) + '/' + parseInt(m[3])];
      return [v.toString().trim()];
    });
    range.setValues(newVals);
    fixed++;
    Logger.log('修復：' + name);
  });
  const msg = '已修復 ' + fixed + ' 個工作表的日期格式。';
  Logger.log(msg);
  return msg;
}

// ── 審核驗算：通過或退回 ────────────────────────────────────────────────
// action: 'approve' | 'reject'
// empId: 審核者員工編號（通過時需驗證）
// sheetName: 班表工作表名稱
// errorNote: 退回時的錯誤說明（reject 才用）
// checkedItems: 已勾選的驗算項目（array of strings）
function auditSchedule(empId, sheetName, action, errorNote, checkedItems) {
  // bug #7: reject 會 deleteSheet 整月班表，原本完全沒驗證可被任意呼叫造成資料毀滅。
  // 改為 reject 也須通過 verifyEmpId（與 approve 對稱）。
  if ((action === 'approve' || action === 'reject') && !verifyEmpId(empId)) {
    return { success: false, message: '員工編號驗證失敗，請重新輸入。' };
  }
  try {
    const spreadsheet = getSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: '找不到工作表：' + sheetName };

    // 取得審核者姓名
    const settingSheet = spreadsheet.getSheetByName(EMAIL_SHEET_NAME);
    const empIds = settingSheet.getRange('M1:M11').getValues().flat().map(String);
    const names  = settingSheet.getRange('I1:I11').getValues().flat().map(String);
    let auditorName = empId;
    empIds.forEach((id, i) => { if (id.trim() === empId.trim()) auditorName = names[i] || empId; });

    if (action === 'approve') {
      // 核准：設為 approved
      const oldNote = sheet.getRange('N1').getNote() || '';
      const lines = oldNote.split('\n').filter(l => l && !l.startsWith('審核狀態:'));
      lines.push('審核狀態:approved');
      sheet.getRange('N1').setNote(lines.join('\n'));
      const itemStr = (checkedItems && checkedItems.length) ? checkedItems.join('、') : '（未指定）';
      writeOpLog('審核通過', `${sheetName} 由 ${auditorName}(${empId}) 通過，驗算項目：${itemStr}`);
      return { success: true, message: `「${sheetName}」已核准，開放換班。` };

    } else if (action === 'reject') {
      // 退回：刪除工作表 + 寄送錯誤回饋 email
      const itemStr = (checkedItems && checkedItems.length) ? checkedItems.join('、') : '（未指定）';
      const now = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'yyyy/MM/dd HH:mm');

      // 寄送 email
      const subject = `【班表審核退回】${sheetName} — 審核者：${auditorName}`;
      const html = `
<div style="font-family:sans-serif;max-width:600px">
  <h2 style="color:#c62828;border-bottom:2px solid #ef9a9a;padding-bottom:8px">⚠️ 班表審核退回通知</h2>
  <table style="width:100%;border-collapse:collapse;margin-top:12px">
    <tr><td style="padding:8px;background:#fafafa;border:1px solid #eee;font-weight:700;width:120px">班表名稱</td>
        <td style="padding:8px;border:1px solid #eee">${sheetName}</td></tr>
    <tr><td style="padding:8px;background:#fafafa;border:1px solid #eee;font-weight:700">審核者</td>
        <td style="padding:8px;border:1px solid #eee">${auditorName}（${empId}）</td></tr>
    <tr><td style="padding:8px;background:#fafafa;border:1px solid #eee;font-weight:700">退回時間</td>
        <td style="padding:8px;border:1px solid #eee">${now}</td></tr>
    <tr><td style="padding:8px;background:#fafafa;border:1px solid #eee;font-weight:700">已驗算項目</td>
        <td style="padding:8px;border:1px solid #eee">${itemStr}</td></tr>
    <tr><td style="padding:8px;background:#fff3f3;border:1px solid #eee;font-weight:700;color:#c62828">錯誤說明</td>
        <td style="padding:8px;background:#fff3f3;border:1px solid #eee;color:#c62828;font-weight:700">${errorNote || '（未提供說明）'}</td></tr>
  </table>
  <p style="color:#666;font-size:.85rem;margin-top:16px">此班表已從系統刪除，請重新執行自動排班。</p>
</div>`;
      try {
        MailApp.sendEmail({ to: 'za869765@gmail.com', subject, htmlBody: html, name: '佳里衛生所班表系統' });
      } catch(e) {
        writeOpLog('審核退回Email失敗', e.message);
      }

      // 刪除工作表
      spreadsheet.deleteSheet(sheet);
      writeOpLog('審核退回', `${sheetName} 由 ${auditorName}(${empId}) 退回，原因：${errorNote}`);
      return { success: true, message: `「${sheetName}」已退回並刪除，已發送通知 email。` };
    }
    return { success: false, message: '未知操作。' };
  } catch(e) {
    return { success: false, message: '執行失敗：' + e.message };
  }
}

// ── 驗算輔助：取得上月末位 / 本月首位 值班資訊 ────────────────────────
function getAuditHints(sheetName) {
  try {
    const spreadsheet = getSpreadsheet();
    const parsed = parseYearMonthFromSheetName(sheetName);
    if (!parsed.valid) return { success: false, message: '無法解析班表月份' };
    const { year, month } = parsed;
    const rocMonths = ['','一','二','三','四','五','六','七','八','九','十','十一','十二'];
    const prevMonth = month === 1 ? 12 : month - 1;
    const prevYear  = month === 1 ? year - 1 : year;
    const prevSheetName = rocNumToStr(prevYear - 1911) + '年' + rocMonths[prevMonth] + '月班表';
    const prevSheet = spreadsheet.getSheetByName(prevSheetName);
    let prevWdLast = '', prevWdLastDate = '', prevHolLast = '', prevHolLastDate = '';
    if (prevSheet) {
      const prevData = prevSheet.getRange('A2:M32').getValues();
      const tz = spreadsheet.getSpreadsheetTimeZone();
      for (let r = 0; r < prevData.length; r++) {
        const raw = prevData[r][0];
        if (!raw) continue;
        let pd = null;
        if (raw instanceof Date) { pd = raw; }
        else { const mm = raw.toString().trim().match(/(\d+)\/(\d+)/); if (mm) { try { pd = new Date(prevYear, parseInt(mm[1])-1, parseInt(mm[2])); } catch(e){} } }
        if (!pd) continue;
        const person = (prevData[r][2] || '').toString().trim();
        if (!person) continue;
        const isHol = isHoliday(pd);
        const ds = Utilities.formatDate(pd, tz, 'M/d');
        if (isHol) { prevHolLast = person; prevHolLastDate = ds; }
        else        { prevWdLast  = person; prevWdLastDate  = ds; }
      }
    }
    const curSheet = spreadsheet.getSheetByName(sheetName);
    if (!curSheet) return { success: false, message: '找不到工作表：' + sheetName };
    const curData = curSheet.getRange('A2:M32').getValues();
    let curWdFirst = '', curWdFirstDate = '', curHolFirst = '', curHolFirstDate = '';
    const tz2 = spreadsheet.getSpreadsheetTimeZone();
    for (let r = 0; r < curData.length; r++) {
      const raw = curData[r][0];
      if (!raw) continue;
      let cd = null;
      if (raw instanceof Date) { cd = raw; }
      else { const mm = raw.toString().trim().match(/(\d+)\/(\d+)/); if (mm) { try { cd = new Date(year, parseInt(mm[1])-1, parseInt(mm[2])); } catch(e){} } }
      if (!cd) continue;
      const person = (curData[r][2] || '').toString().trim();
      if (!person) continue;
      const isHol = isHoliday(cd);
      const ds = Utilities.formatDate(cd, tz2, 'M/d');
      if (!isHol && !curWdFirst)  { curWdFirst  = person; curWdFirstDate  = ds; }
      if (isHol  && !curHolFirst) { curHolFirst = person; curHolFirstDate = ds; }
      if (curWdFirst && curHolFirst) break;
    }
    const settingSheet = spreadsheet.getSheetByName(EMAIL_SHEET_NAME);
    const exclusions = settingSheet.getRange('T1:T30').getValues().flat().filter(n => n).map(String);
    return {
      success: true, prevSheetName, prevExists: !!prevSheet,
      prevWdLast, prevWdLastDate, prevHolLast, prevHolLastDate,
      curWdFirst, curWdFirstDate, curHolFirst, curHolFirstDate,
      exclusions: exclusions.join('、') || '（無設定）', month, year
    };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── 從換班日誌重建「原始排班人」map（供 autoValidateSchedule Check 1 使用）──
// 與 runAutoSchedule 內部的 buildOriginalMap 邏輯完全一致，
// 差別在此為頂層函式，使用 getSpreadsheet()。
function buildOriginalScheduleMap(prevSheetName) {
  const origMap = {};
  try {
    const spreadsheet = getSpreadsheet();
    const settingSheet = spreadsheet.getSheetByName(EMAIL_SHEET_NAME);
    const logs = settingSheet.getRange('A40:A460').getValues()
      .flat().filter(r => r && r.toString().trim() !== '');
    const sh = spreadsheet.getSheetByName(prevSheetName);
    if (!sh) return origMap;
    const hdrs = sh.getRange('C1:M1').getValues()[0].map(h => h.toString().trim());

    // bug #22: 與 getScheduleChanges 對齊 — 舊日誌可能寫「值班人員/登革熱二線/協助掛號」，
    // 新 header 為「值班/停班2線/支援」。原本只用 hdrs.indexOf 直接比對拿不到，
    // origMap 缺值 → autoValidateSchedule Check 1 跨月輪序誤判。
    const LEGACY_HDR = { '值班人員':'值班', '登革熱二線':'停班2線', '協助掛號':'支援' };
    const resolveHdrIdx = (name) => {
      let ci = hdrs.indexOf(name);
      if (ci !== -1) return ci;
      const mapped = LEGACY_HDR[name];
      if (mapped) {
        ci = hdrs.indexOf(mapped);
        if (ci !== -1) return ci;
      }
      // 反向：log 寫新名、sheet header 還是舊名
      for (const oldName in LEGACY_HDR) {
        if (LEGACY_HDR[oldName] === name) {
          ci = hdrs.indexOf(oldName);
          if (ci !== -1) return ci;
        }
      }
      return -1;
    };

    logs.forEach(log => {
      const s = log.toString().trim();

      // ── 格式1：一般換班日誌（含星期）──
      // 格式：M/D 週X 班別 原A→新B (empId) 更換時間: ...
      const m1 = s.match(/^(\d+\/\d+)\s+週.\s+(.+?)\s+原(.+?)→新(.+?)\s+(?:\([^)]+\)\s+)?更換時間:/);
      if (m1) {
        const [, date, shiftType, oldVal] = m1;
        const ci = resolveHdrIdx(shiftType);
        if (ci !== -1) {
          const key = date + '|' + ci;
          if (!origMap[key]) origMap[key] = oldVal;
        }
        return;
      }

      // ── 格式2：拖曳日誌（排班者 / 審核者）──
      // BUG 15 修正：補充解析 排班-姓名 / 審核-姓名 格式，避免原始排班還原不完整
      // 格式：排班-姓名 M/D 班別 原A→新B 更換時間: ...
      //       審核-姓名 M/D 班別 原A→新B 更換時間: ...
      const m2 = s.match(/^(?:排班|審核)-.+?\s+(\d+\/\d+)\s+(.+?)\s+原(.+?)→新(.+?)\s+更換時間:/);
      if (m2) {
        const [, date, shiftType, oldVal] = m2;
        const ci = resolveHdrIdx(shiftType);
        if (ci !== -1) {
          const key = date + '|' + ci;
          if (!origMap[key]) origMap[key] = oldVal;
        }
      }
    });
  } catch(e) {}
  return origMap;
}

// ── 系統自動驗算（背景執行，三項檢查）────────────────────────────────
function autoValidateSchedule(sheetName) {
  try {
    const spreadsheet = getSpreadsheet();
    const parsed = parseYearMonthFromSheetName(sheetName);
    if (!parsed.valid) return { success: false, checks: [], error: '無法解析班表月份' };
    const { year, month } = parsed;
    const tz = spreadsheet.getSpreadsheetTimeZone();
    const rocMonths = ['','一','二','三','四','五','六','七','八','九','十','十一','十二'];
    const SHIFT_NAMES = ['值班','支援','門診','掛號','前台','預登1','預登2注','注射1','注射2','卡介苗','停班2線'];

    // ── 讀取本月班表 A2:M32 ──────────────────────────────────────────
    const curSheet = spreadsheet.getSheetByName(sheetName);
    if (!curSheet) return { success: false, checks: [], error: '找不到工作表' };
    const curData = curSheet.getRange('A2:M32').getValues();

    // 解析日期
    function parseDate(raw) {
      if (!raw) return null;
      if (raw instanceof Date) return raw;
      const mm = raw.toString().trim().match(/(\d+)\/(\d+)/);
      if (mm) { try { return new Date(year, parseInt(mm[1])-1, parseInt(mm[2])); } catch(e){} }
      return null;
    }

    // 收集本月每天資料 { date, isHol, duty(C=idx2), deng(M=idx12), kk(D=idx3) }
    const curRows = curData.map(r => {
      const d = parseDate(r[0]);
      if (!d) return null;
      return {
        date: d,
        ds: Utilities.formatDate(d, tz, 'M/d'),
        isHol: isHoliday(d),
        duty: (r[2] || '').toString().trim(),   // C欄 值班
        kk:   (r[3] || '').toString().trim(),   // D欄 協助掛號
        deng: (r[12]|| '').toString().trim(),   // M欄 停班2線
      };
    }).filter(Boolean);

    const errors = [];

    // ════════════════════════════════════════════════════════════════
    // 檢查 1：上月接續（值班、停班2線、協助掛號）
    // ════════════════════════════════════════════════════════════════
    const prevM = month === 1 ? 12 : month - 1;
    const prevY = month === 1 ? year - 1 : year;
    const prevSN = rocNumToStr(prevY - 1911) + '年' + rocMonths[prevM] + '月班表';
    const prevSheet = spreadsheet.getSheetByName(prevSN);

    // 取值班/停班2線名單輪序
    const settingSheet = spreadsheet.getSheetByName(EMAIL_SHEET_NAME);
    const dutyOrder = settingSheet.getRange('E1:E11').getValues().flat().map(String).filter(n => n.trim());
    const dengOrder = settingSheet.getRange('B1:B11').getValues().flat().map(String).filter(n => n.trim());
    const kkOrder   = settingSheet.getRange('J1:J3').getValues().flat().map(String).filter(n => n.trim());

    const checkContinuity = function(label, order, prevLastName, curFirstName) {
      if (!order.length || !prevLastName || !curFirstName) return;
      const prevIdx = order.indexOf(prevLastName);
      const curIdx  = order.indexOf(curFirstName);
      if (prevIdx === -1 || curIdx === -1) return; // 有人不在名單，略過
      const expected = (prevIdx + 1) % order.length;
      if (curIdx !== expected) {
        errors.push({
          check: 1,
          msg: `【${label}】上月末位「${prevLastName}」→本月應接「${order[expected]}」，但實際為「${curFirstName}」（跳位或重回順位1）`
        });
      }
    }

    if (prevSheet) {
      const prevData = prevSheet.getRange('A2:M32').getValues();
      // ★ 修正 BUG 1：使用換班日誌還原「原始排班人」，與排班引擎的指針邏輯一致。
      //   若只讀儲存格現值（含換班後的人），會與排班引擎的起始指針不一致，導致誤報。
      const prevOrigMap = buildOriginalScheduleMap(prevSN);
      let prevLastWdDuty='', prevLastHolDuty='';
      let prevLastWdDeng='', prevLastHolDeng='';
      let prevLastWdKk='',   prevLastHolKk='';

      prevData.forEach(r => {
        const d = parseDate(r[0]);
        if (!d) return;
        const isHol = isHoliday(d);
        const ds = Utilities.formatDate(d, tz, 'M/d');
        // 優先取原始排班人（換班前），無記錄則用儲存格現值
        const duty = prevOrigMap[ds + '|0']  || (r[2]||'').toString().trim();
        const kk   = prevOrigMap[ds + '|1']  || (r[3]||'').toString().trim();
        const deng = prevOrigMap[ds + '|10'] || (r[12]||'').toString().trim();
        if (isHol) {
          if (duty) prevLastHolDuty = duty;
          if (deng) prevLastHolDeng = deng;
          if (kk)   prevLastHolKk   = kk;
        } else {
          if (duty) prevLastWdDuty = duty;
          if (deng) prevLastWdDeng = deng;
          if (kk)   prevLastWdKk   = kk;
        }
      });

      // 本月首位
      const curFirstWdDuty = (curRows.find(r => !r.isHol && r.duty) || {}).duty || '';
      const curFirstHolDuty = (curRows.find(r => r.isHol && r.duty) || {}).duty || '';
      const curFirstWdDeng = (curRows.find(r => !r.isHol && r.deng) || {}).deng || '';
      const curFirstHolDeng = (curRows.find(r => r.isHol && r.deng) || {}).deng || '';
      const curFirstWdKk   = (curRows.find(r => !r.isHol && r.kk) || {}).kk || '';

      checkContinuity('平日值班', dutyOrder, prevLastWdDuty, curFirstWdDuty);
      checkContinuity('假日值班', dutyOrder, prevLastHolDuty, curFirstHolDuty);
      checkContinuity('平日停班2線', dengOrder, prevLastWdDeng, curFirstWdDeng);
      checkContinuity('假日停班2線', dengOrder, prevLastHolDeng, curFirstHolDeng);
      if (kkOrder.length) checkContinuity('平日協助掛號', kkOrder, prevLastWdKk, curFirstWdKk);
    } else {
      errors.push({ check: 1, msg: `無法讀取上月「${prevSN}」，跳過接續檢查` });
    }

    // ════════════════════════════════════════════════════════════════
    // 檢查 2：連續同人問題（平日/假日各自分開）
    // ════════════════════════════════════════════════════════════════
    const checkConsecutive = function(label, rows, field) {
      let prev = '';
      rows.forEach((r, i) => {
        const val = r[field];
        if (val && val === prev) {
          errors.push({ check: 2, msg: `【${label}連續】${r.ds} 與前一日（${rows[i-1].ds}）均為「${val}」` });
        }
        // ★ 修正 BUG 3：空白（全員排除日）時重置 prev，避免跨空格誤判連排
        prev = val || '';
      });
    }
    const wdRows  = curRows.filter(r => !r.isHol);
    const holRows = curRows.filter(r => r.isHol);
    checkConsecutive('平日值班', wdRows, 'duty');
    checkConsecutive('假日值班', holRows, 'duty');
    checkConsecutive('平日停班2線', wdRows, 'deng');
    checkConsecutive('假日停班2線', holRows, 'deng');

    // ════════════════════════════════════════════════════════════════
    // 檢查 3：同日 值班 = 停班2線 衝突（應已挪移）
    // ════════════════════════════════════════════════════════════════
    curRows.forEach(r => {
      if (r.duty && r.deng && r.duty === r.deng) {
        errors.push({ check: 3, msg: `【衝突】${r.ds}（${r.isHol?'假日':'平日'}）值班「${r.duty}」與停班2線「${r.deng}」為同一人（應已挪移）` });
      }
    });

    // ════════════════════════════════════════════════════════════════
    // 檢查 4：各性質班別公平度（同性質班別最高-最低不得差 4 次以上 即 ±2）
    //   群組1：門診合計（E~K 欄，colIdx 2-8，排除支援D欄 及 卡介苗L欄）→ 護理師 K欄名單
    //   群組2：值班平日次數   → 全員 E欄輪序名單
    //   群組3：值班假日次數   → 全員 E欄輪序名單
    //   群組4：停班2線平日    → 全員 B欄輪序名單
    //   群組5：停班2線假日    → 全員 B欄輪序名單
    //   （支援 D欄、卡介苗 L欄 不檢查）
    // ════════════════════════════════════════════════════════════════
    const check4Sheet = spreadsheet.getSheetByName(EMAIL_SHEET_NAME);

    // 名單
    const nurseNames  = check4Sheet.getRange('K1:K8').getValues().flat().map(String).filter(n=>n.trim());
    const dutyNames   = check4Sheet.getRange('E1:E11').getValues().flat().map(String).filter(n=>n.trim());
    const dengNames   = check4Sheet.getRange('B1:B11').getValues().flat().map(String).filter(n=>n.trim());

    // 計算各性質群組次數
    const clinicTotal = {}; // 門診合計
    const dutyWd  = {}, dutyHol  = {}; // 值班平日/假日
    const dengWd  = {}, dengHol  = {}; // 停班2線平日/假日

    nurseNames.forEach(n => { clinicTotal[n] = 0; });
    dutyNames.forEach(n  => { dutyWd[n]=0;  dutyHol[n]=0; });
    dengNames.forEach(n  => { dengWd[n]=0;  dengHol[n]=0; });

    // ★ 修正 BUG 6：curData = A2:M32，index 0=A,1=B,2=C(值班)...4=E(門診)...10=K(注射2)
    //   門診系列 E~K = array indices 4~10（原錯誤用 2~8 包含值班/協助掛號並漏掉注射1/2）
    curData.forEach(row => {
      for (let idx = 4; idx <= 10; idx++) {
        const nm = (row[idx]||'').toString().trim();
        if (nm && clinicTotal[nm] !== undefined) clinicTotal[nm]++;
      }
    });

    // 值班 / 停班2線 計數（從 curRows 直接取，來源已正確）
    curRows.forEach(r => {
      if (r.duty && dutyWd[r.duty] !== undefined) {
        if (r.isHol) dutyHol[r.duty]++; else dutyWd[r.duty]++;
      }
      if (r.deng && dengWd[r.deng] !== undefined) {
        if (r.isHol) dengHol[r.deng]++; else dengWd[r.deng]++;
      }
    });

    // 各群組允許差距上限（超過此值即報錯）
    // 門診合計 ±2 → max-min > 4（即 ≥ 5）
    // 值班平日/假日 ±1 → max-min > 2（即 ≥ 3）
    // 停班2線平日/假日 ±2 → max-min > 4（即 ≥ 5）
    const fairCheckT = function(label, countMap, maxAllowed) {
      const active = Object.entries(countMap);
      if (active.length < 2) return;
      const vals = active.map(([,v])=>v);
      const maxV = Math.max(...vals), minV = Math.min(...vals);
      if (maxV - minV > maxAllowed * 2) {
        const maxN = active.find(([,v])=>v===maxV)[0];
        const minN = active.find(([,v])=>v===minV)[0];
        errors.push({ check: 4,
          msg: `【${label}】最多「${maxN}」${maxV}次 vs 最少「${minN}」${minV}次，差距 ${maxV-minV} 次（超過允許的 ±${maxAllowed}）`
        });
      }
    }

    fairCheckT('門診合計',   clinicTotal, 2);
    fairCheckT('值班平日',   dutyWd,      1);
    fairCheckT('值班假日',   dutyHol,     1);
    fairCheckT('停班2線平日', dengWd,     2);
    fairCheckT('停班2線假日', dengHol,    2);

    // ── 整理結果 ─────────────────────────────────────────────────────
    const check1Errors = errors.filter(e => e.check === 1);
    const check2Errors = errors.filter(e => e.check === 2);
    const check3Errors = errors.filter(e => e.check === 3);
    const check4Errors = errors.filter(e => e.check === 4);

    return {
      success: true,
      sheetName,
      checks: [
        { id: 1, label: '跨月輪序接續',   pass: check1Errors.length === 0, errors: check1Errors.map(e => e.msg) },
        { id: 2, label: '連續同人班次',   pass: check2Errors.length === 0, errors: check2Errors.map(e => e.msg) },
        { id: 3, label: '值班/停班2線衝突', pass: check3Errors.length === 0, errors: check3Errors.map(e => e.msg) },
        { id: 4, label: '門診公平度(±2)',  pass: check4Errors.length === 0, errors: check4Errors.map(e => e.msg) },
      ],
      allPass: errors.length === 0,
      totalErrors: errors.length
    };
  } catch(e) {
    return { success: false, checks: [], error: e.message };
  }
}
// =============================================
// 全年模擬驗證 — simulateFullYearValidation
// 逐月執行 dryRun 排班，接著對結果跑 4 項自檢邏輯，
// 回傳每月是否通過及詳細錯誤清單。
// 對已有資料的月份直接讀 sheet；未排班月份執行 dryRun。
// =============================================
function simulateFullYearValidation(adminPassword) {
  const spreadsheet = getSpreadsheet();
  const tz = spreadsheet.getSpreadsheetTimeZone();
  const rocMonths = ['','一','二','三','四','五','六','七','八','九','十','十一','十二'];
  const year = new Date().getFullYear();

  // 讀設定表名單
  const settingSheet = spreadsheet.getSheetByName(EMAIL_SHEET_NAME);
  const dutyOrder  = settingSheet.getRange('E1:E11').getValues().flat().map(String).filter(n=>n.trim());
  const dengOrder  = settingSheet.getRange('B1:B11').getValues().flat().map(String).filter(n=>n.trim());
  const kkOrder    = settingSheet.getRange('J1:J3').getValues().flat().map(String).filter(n=>n.trim());
  const nurseNames = settingSheet.getRange('K1:K8').getValues().flat().map(String).filter(n=>n.trim());
  const dutyNames  = settingSheet.getRange('E1:E11').getValues().flat().map(String).filter(n=>n.trim());
  const dengNames  = settingSheet.getRange('B1:B11').getValues().flat().map(String).filter(n=>n.trim());

  function parseRawDate(raw, yr, mn) {
    if (!raw) return null;
    if (raw instanceof Date) return raw;
    const mm = raw.toString().trim().match(/(\d+)\/(\d+)/);
    if (mm) { try { return new Date(yr, parseInt(mm[1])-1, parseInt(mm[2])); } catch(e){} }
    return null;
  }

  const monthResults = [];
  // prevSummary 記錄上月末位（模擬月使用，已排月份從 sheet 讀）
  let prevSummary = null;

  for (let month = 1; month <= 12; month++) {
    const sheetName = rocNumToStr(year - 1911) + '年' + rocMonths[month] + '月班表';
    const sheet = spreadsheet.getSheetByName(sheetName);
    const mres = { month, sheetName, source: '', pass: false, errors: [] };

    if (!sheet) {
      mres.source = 'no_sheet';
      mres.pass = null;
      mres.errors.push('工作表不存在，跳過');
      monthResults.push(mres);
      prevSummary = null;
      continue;
    }

    // 判斷是否已有排班資料
    const rawSheet = sheet.getRange('A2:M32').getValues();
    const hasData  = rawSheet.some(r => (r[2]||'').toString().trim() || (r[12]||'').toString().trim());

    // 建立 curRows（統一格式 {date,ds,isHol,duty,kk,deng,rawIdx}）
    let curRows = [];
    // clinic raw data（用於 Check 4）：array of rows, 每 row index 對應欄位
    // existing → rawSheet[r][4..10] = E..K；dryRun → preview[r][2..8] = colIdx 2..8
    let clinicSource = null;
    let clinicIsRaw  = true; // true=rawSheet(index需+0), false=preview(colIdx直接用)

    if (hasData) {
      mres.source = 'existing';
      rawSheet.forEach((r, ri) => {
        const d = parseRawDate(r[0], year, month);
        if (!d) return;
        curRows.push({
          date: d,
          ds:   Utilities.formatDate(d, tz, 'M/d'),
          isHol: isHoliday(d),
          duty: (r[2]||'').toString().trim(),
          kk:   (r[3]||'').toString().trim(),
          deng: (r[12]||'').toString().trim(),
          rawRow: r
        });
      });
      clinicSource = rawSheet;
      clinicIsRaw  = true;

    } else if (adminPassword) {
      mres.source = 'dry_run';
      const dr = runAutoSchedule(sheetName, adminPassword, { overwrite: true, sendNotify: false, dryRun: true });
      if (!dr.success) {
        mres.errors.push('dryRun失敗: ' + dr.message);
        mres.pass = false;
        monthResults.push(mres);
        prevSummary = null;
        continue;
      }
      // dr.dates = array of [dateStr] (from combinedDates.map(d=>[d]))
      // dr.preview = array of rowRes[0..10]
      for (let r = 0; r < dr.preview.length; r++) {
        const dateStr = Array.isArray(dr.dates[r]) ? dr.dates[r][0] : (dr.dates[r]||'');
        if (!dateStr || !dateStr.trim()) continue;
        const mm = dateStr.match(/(\d+)\/(\d+)/);
        if (!mm) continue;
        let d; try { d = new Date(year, parseInt(mm[1])-1, parseInt(mm[2])); } catch(e){ continue; }
        const row = dr.preview[r];
        curRows.push({
          date: d,
          ds:   mm[1]+'/'+parseInt(mm[2]),
          isHol: isHoliday(d),
          duty: (row[0]||'').toString().trim(),   // colIdx 0
          kk:   (row[1]||'').toString().trim(),   // colIdx 1
          deng: (row[10]||'').toString().trim(),  // colIdx 10
          rawRow: row
        });
      }
      clinicSource = dr.preview;
      clinicIsRaw  = false; // preview 索引 0-10 = colIdx

    } else {
      mres.source = 'no_data';
      mres.pass = null;
      mres.errors.push('無排班資料且未提供密碼，跳過');
      monthResults.push(mres);
      prevSummary = null;
      continue;
    }

    const errors = [];

    // ── Check 1：跨月輪序接續 ──────────────────────────────────────────
    // 上月末位：若 prevSummary 存在（上月為 dryRun）用 prevSummary；否則讀 sheet
    let prevLastWdDuty='', prevLastHolDuty='', prevLastWdDeng='', prevLastHolDeng='', prevLastWdKk='';

    if (month > 1 && prevSummary) {
      prevLastWdDuty  = prevSummary.lastWdDuty;
      prevLastHolDuty = prevSummary.lastHolDuty;
      prevLastWdDeng  = prevSummary.lastWdDeng;
      prevLastHolDeng = prevSummary.lastHolDeng;
      prevLastWdKk    = prevSummary.lastWdKk;
    } else if (month > 1) {
      const prevM = month - 1;
      const prevSN = rocNumToStr(year - 1911) + '年' + rocMonths[prevM] + '月班表';
      const prevSh = spreadsheet.getSheetByName(prevSN);
      if (prevSh) {
        prevSh.getRange('A2:M32').getValues().forEach(r => {
          const d = parseRawDate(r[0], year, prevM);
          if (!d) return;
          const isHol = isHoliday(d);
          const duty = (r[2]||'').toString().trim();
          const kk   = (r[3]||'').toString().trim();
          const deng = (r[12]||'').toString().trim();
          if (isHol) { if(duty) prevLastHolDuty=duty; if(deng) prevLastHolDeng=deng; }
          else       { if(duty) prevLastWdDuty=duty;  if(deng) prevLastWdDeng=deng; if(kk) prevLastWdKk=kk; }
        });
      }
    }

    const chkCont = (label, order, prev, cur) => {
      if (!order.length || !prev || !cur) return;
      const pi = order.indexOf(prev), ci2 = order.indexOf(cur);
      if (pi === -1 || ci2 === -1) return;
      const exp = (pi + 1) % order.length;
      if (ci2 !== exp) errors.push(`[C1] 【${label}】上月末「${prev}」→應接「${order[exp]}」，實際「${cur}」`);
    };
    const curFirstWdDuty  = (curRows.find(r=>!r.isHol&&r.duty)||{}).duty||'';
    const curFirstHolDuty = (curRows.find(r=> r.isHol&&r.duty)||{}).duty||'';
    const curFirstWdDeng  = (curRows.find(r=>!r.isHol&&r.deng)||{}).deng||'';
    const curFirstHolDeng = (curRows.find(r=> r.isHol&&r.deng)||{}).deng||'';
    const curFirstWdKk    = (curRows.find(r=>!r.isHol&&r.kk)||{}).kk||'';
    chkCont('平日值班',    dutyOrder, prevLastWdDuty,  curFirstWdDuty);
    chkCont('假日值班',    dutyOrder, prevLastHolDuty, curFirstHolDuty);
    chkCont('平日停班2線', dengOrder, prevLastWdDeng,  curFirstWdDeng);
    chkCont('假日停班2線', dengOrder, prevLastHolDeng, curFirstHolDeng);
    if (kkOrder.length) chkCont('平日協助掛號', kkOrder, prevLastWdKk, curFirstWdKk);

    // ── Check 2：連續同人 ──────────────────────────────────────────────
    const wdRows  = curRows.filter(r=>!r.isHol);
    const holRows = curRows.filter(r=> r.isHol);
    const chkConsec = (label, rows, field) => {
      // BUG 16 修正：空值時必須重置 pv，否則 A,空,A 會誤報為連續相同
      // 與 autoValidateSchedule Check 2 邏輯保持一致
      let pv = '';
      rows.forEach((r,i) => {
        const val = (r[field] || '').toString().trim();
        if (val && val === pv) errors.push(`[C2] 【${label}連續】${r.ds} 與前日（${rows[i-1].ds}）均為「${val}」`);
        pv = val; // 空值自動重置為 ''，不保留前次值
      });
    };
    chkConsec('平日值班',    wdRows,  'duty');
    chkConsec('假日值班',    holRows, 'duty');
    chkConsec('平日停班2線', wdRows,  'deng');
    chkConsec('假日停班2線', holRows, 'deng');

    // ── Check 3：值班 = 停班2線 衝突 ──────────────────────────────────
    curRows.forEach(r => {
      if (r.duty && r.deng && r.duty === r.deng)
        errors.push(`[C3] 【衝突】${r.ds}（${r.isHol?'假':'平'}）值班=停班2線=「${r.duty}」`);
    });

    // ── Check 4（修正版）：公平度，門診合計用正確欄位 E~K ─────────────
    const clinicTotal = {}; nurseNames.forEach(n=>{ clinicTotal[n]=0; });
    const dutyWd={}, dutyHol={}; dutyNames.forEach(n=>{ dutyWd[n]=0; dutyHol[n]=0; });
    const dengWd={}, dengHol={}; dengNames.forEach(n=>{ dengWd[n]=0; dengHol[n]=0; });

    curRows.forEach(r => {
      if (r.duty) { if(r.isHol){ if(dutyHol[r.duty]!==undefined) dutyHol[r.duty]++; }
                    else        { if(dutyWd[r.duty]!==undefined)  dutyWd[r.duty]++; } }
      if (r.deng) { if(r.isHol){ if(dengHol[r.deng]!==undefined) dengHol[r.deng]++; }
                    else        { if(dengWd[r.deng]!==undefined)  dengWd[r.deng]++; } }
    });

    // 門診合計：根據資料來源決定正確 index
    if (clinicSource) {
      clinicSource.forEach(row => {
        if (!row) return;
        if (clinicIsRaw) {
          // rawSheet: A2:M32 → index 4(E)..10(K) = colIdx 2..8
          for (let idx=4; idx<=10; idx++) {
            const nm = (row[idx]||'').toString().trim();
            if (nm && clinicTotal[nm]!==undefined) clinicTotal[nm]++;
          }
        } else {
          // dryRun preview: index 0..10 = colIdx 0..10 → colIdx 2..8 = 門診系列
          for (let ci=2; ci<=8; ci++) {
            const nm = (row[ci]||'').toString().trim();
            if (nm && clinicTotal[nm]!==undefined) clinicTotal[nm]++;
          }
        }
      });
    }

    const fairT = (label, map, maxAllowed) => {
      const active = Object.entries(map);
      if (active.length < 2) return;
      const vals = active.map(([,v])=>v);
      const mx = Math.max(...vals), mn2 = Math.min(...vals);
      if (mx - mn2 > maxAllowed * 2) {
        const mxN = active.find(([,v])=>v===mx)[0];
        const mnN = active.find(([,v])=>v===mn2)[0];
        errors.push(`[C4] 【${label}】${mxN}:${mx}次 vs ${mnN}:${mn2}次，差${mx-mn2}（允許±${maxAllowed}）`);
      }
    };
    fairT('門診合計',   clinicTotal, 2);
    fairT('值班平日',   dutyWd,      1);
    fairT('值班假日',   dutyHol,     1);
    fairT('停班2線平日', dengWd,     2);
    fairT('停班2線假日', dengHol,    2);

    // ── 額外偵測：停班2線空白（BUG 2 symptom）─────────────────────────
    curRows.forEach(r => {
      if (!r.deng) errors.push(`[B2] 停班2線空白: ${r.ds}（${r.isHol?'假':'平'}）`);
      if (!r.duty) errors.push(`[DUTY] 值班空白: ${r.ds}`);
    });

    mres.errors = errors;
    mres.pass = errors.length === 0;

    // 記錄本月末位供下月 Check 1 使用
    let lwDuty='', lhDuty='', lwDeng='', lhDeng='', lwKk='';
    curRows.forEach(r => {
      if (r.isHol) { if(r.duty) lhDuty=r.duty; if(r.deng) lhDeng=r.deng; }
      else         { if(r.duty) lwDuty=r.duty;  if(r.deng) lwDeng=r.deng; if(r.kk) lwKk=r.kk; }
    });
    prevSummary = { lastWdDuty:lwDuty, lastHolDuty:lhDuty, lastWdDeng:lwDeng, lastHolDeng:lhDeng, lastWdKk:lwKk };
    monthResults.push(mres);
  }

  // 整理摘要
  const passCount = monthResults.filter(r=>r.pass===true).length;
  const failCount = monthResults.filter(r=>r.pass===false && r.source !== 'no_sheet' && r.source !== 'no_data').length;
  const lines = [];
  const rocYearDisp = new Date().getFullYear() - 1911;
  lines.push(`=== ${rocYearDisp}年全年排班自檢模擬結果 ===`);
  lines.push(`通過: ${passCount} 月 | 不通過: ${failCount} 月`);
  lines.push('');
  monthResults.forEach(r => {
    const tag = r.pass===true ? '✅' : r.pass===false ? '❌' : '⚠️';
    lines.push(`${tag} ${r.month}月（${r.sheetName}）[${r.source}]`);
    if (r.errors.length > 0) {
      r.errors.forEach(e => lines.push('   ' + e));
    }
  });

  Logger.log(lines.join('\n'));
  return { success: true, summary: monthResults, report: lines.join('\n') };
}

// =============================================
// 排班一致性 + 自檢完整測試（單月連排版）
// runScheduleConsistencyTest(adminPassword)
//
// 測試項目：
//   Phase 1 — 單月連排第 1 次 → autoValidateSchedule 全 12 月
//   Phase 2 — 單月連排第 2 次 → 驗算 → 與 Phase 1 快照比對（重複性）
//
// 注意：函式實際寫入試算表，完成後班表為 Phase 2 的結果。
// =============================================
function runScheduleConsistencyTest(adminPassword) {
  if (!verifyAdminPassword(adminPassword))
    return { success: false, message: '管理員密碼錯誤' };

  const spreadsheet = getSpreadsheet();
  const year  = 2026;
  const rocY  = year - 1911;
  const rocMonths = ['','一','二','三','四','五','六','七','八','九','十','十一','十二'];

  function sn(m)  { return rocNumToStr(rocY) + '年' + rocMonths[m] + '月班表'; }
  function ts()   { return new Date().toLocaleTimeString(); }

  // 擷取 C2:M32 排班結果（只比較 C~M 欄，不含日期 A/B）
  function captureSchedule(m) {
    const sh = spreadsheet.getSheetByName(sn(m));
    if (!sh) return null;
    return sh.getRange('C2:M32').getValues().map(r =>
      r.map(v => v ? v.toString().trim() : '')
    );
  }

  // 比對兩份排班快照，回傳差異清單
  function diffSnaps(a, b, label) {
    if (!a || !b) return [`${label}: 缺少快照`];
    const diffs = [];
    const colLabels = ['C(值班)','D(協掛)','E(門診)','F(流注1)','G(流注2)',
                       'H(預登1)','I(預登2)','J(注射1)','K(注射2)','L(卡介)','M(停班2)'];
    for (let r = 0; r < 31; r++) {
      for (let c = 0; c < 11; c++) {
        const va = (a[r]||[])[c] || '', vb = (b[r]||[])[c] || '';
        if (va !== vb) diffs.push(`R${r+2} ${colLabels[c]}: 「${va}」→「${vb}」`);
      }
    }
    return diffs;
  }

  // 執行單月連排（month 1→12，overwrite）
  function runSequential() {
    Logger.log(`[${ts()}] 單月連排 start`);
    for (let m = 1; m <= 12; m++) {
      const name = sn(m);
      const sh   = spreadsheet.getSheetByName(name);
      if (!sh) { Logger.log(`  ${m}月 no_sheet skip`); continue; }
      const res = runAutoSchedule(name, adminPassword, { overwrite: true, sendNotify: false, dryRun: false });
      Logger.log(`  ${m}月 ${res.success ? 'ok' : 'FAIL: ' + res.message}`);
    }
    Logger.log(`[${ts()}] 單月連排 end`);
  }

  // 對全部月份跑 autoValidateSchedule，回傳摘要
  function validateAll(tag) {
    const rows = [];
    for (let m = 1; m <= 12; m++) {
      const name = sn(m);
      const sh   = spreadsheet.getSheetByName(name);
      if (!sh) { rows.push({ m, tag, status: 'no_sheet', errors: [] }); continue; }
      const hasData = sh.getRange('C2:C5').getValues().flat().some(v => v && v.toString().trim());
      if (!hasData) { rows.push({ m, tag, status: 'no_data', errors: [] }); continue; }
      const v = autoValidateSchedule(name);
      const errs = v.checks ? v.checks.filter(c => !c.pass).flatMap(c => c.errors) : [v.error || ''];
      rows.push({ m, tag, status: v.allPass ? 'PASS' : 'FAIL', checks: v.checks, errors: errs });
    }
    return rows;
  }

  // 擷取全部月份快照
  function captureAll() {
    const s = {};
    for (let m = 1; m <= 12; m++) s[m] = captureSchedule(m);
    return s;
  }

  // ═══════════════════════════════════════
  // Phase 1：單月連排第 1 次
  // ═══════════════════════════════════════
  runSequential();
  const snap1 = captureAll();
  const val1  = validateAll('單月-1次');

  // ═══════════════════════════════════════
  // Phase 2：單月連排第 2 次（重複性驗證）
  // ═══════════════════════════════════════
  runSequential();
  const snap2 = captureAll();
  const val2  = validateAll('單月-2次');

  // ── 比對快照 ──────────────────────────────────────────────────────
  const compare = (snapA, snapB, label) => {
    const result = { label, consistent: true, monthDiffs: {} };
    for (let m = 1; m <= 12; m++) {
      const d = diffSnaps(snapA[m], snapB[m], `${m}月`);
      if (d.length > 0) { result.consistent = false; result.monthDiffs[m] = d; }
    }
    return result;
  };

  const cmp_1vs2 = compare(snap1, snap2, '單月連排 第1次 vs 第2次（重複性）');

  // ── 產生文字報告 ──────────────────────────────────────────────────
  const L = [];
  L.push('════════════════════════════════════════');
  L.push('  排班一致性 + 自檢完整測試報告');
  L.push('════════════════════════════════════════');

  [[val1,'Phase 1 單月-1次'], [val2,'Phase 2 單月-2次']].forEach(([vals, label]) => {
    const pass = vals.filter(v => v.status === 'PASS').length;
    const fail = vals.filter(v => v.status === 'FAIL').length;
    const skip = vals.filter(v => v.status === 'no_sheet' || v.status === 'no_data').length;
    L.push('');
    L.push(`【${label}】自檢結果：✅${pass} / ❌${fail} / ⚠️${skip}`);
    vals.forEach(v => {
      const icon = v.status === 'PASS' ? '✅' : v.status === 'FAIL' ? '❌' : '⚠️';
      L.push(`  ${icon} ${v.m}月（${v.status}）`);
      (v.errors || []).forEach(e => L.push(`       └ ${e}`));
    });
  });

  L.push('');
  L.push('────────────────────────────────────────');
  L.push('  快照比對（重複排班結果一致性）');
  L.push('────────────────────────────────────────');
  L.push('');
  L.push(`${cmp_1vs2.consistent ? '✅' : '❌'} ${cmp_1vs2.label}`);
  if (!cmp_1vs2.consistent) {
    Object.entries(cmp_1vs2.monthDiffs).forEach(([m, diffs]) => {
      L.push(`  ${m}月（${diffs.length}處差異）：`);
      diffs.slice(0, 8).forEach(d => L.push(`    ${d}`));
      if (diffs.length > 8) L.push(`    ...（共 ${diffs.length} 處）`);
    });
  }

  L.push('');
  L.push('════════════════════════════════════════');
  const allValPass = [...val1,...val2].every(v => v.status === 'PASS' || v.status === 'no_sheet' || v.status === 'no_data');
  const allCmpOk   = cmp_1vs2.consistent;
  L.push(`總結：自檢 ${allValPass ? '全部通過 ✅' : '有不合格項目 ❌'}`);
  L.push(`      一致性 ${allCmpOk   ? '完全一致 ✅' : '存在差異 ❌'}`);
  L.push('════════════════════════════════════════');

  const report = L.join('\n');
  Logger.log(report);
  return {
    success: true,
    report,
    allValidationPass: allValPass,
    allConsistent:     allCmpOk,
    details: { val1, val2, cmp_1vs2 }
  };
}

// ── 排班星期規則設定 ────────────────────────────────────────────────────
// 儲存在「班表設定」工作表 N3:N13（每格一個 colIdx 1-10 的規則字串）
// 格式：colIdx|day1,day2  例如 "8|2" 表示 注射2 只排週二
// day 數字：1=週一 2=週二 3=週三 4=週四 5=週五 6=週六 7=週日 0=假日含平日
// 留空 = 使用預設值

// 預設規則（與 shouldAssignShift 的硬編碼一致）
const SHIFT_DAY_DEFAULTS = {
  1: [4],           // 支援：週四
  2: [2, 4],        // 門診：週二＋週四
  3: [4],           // 掛號：週四
  4: [4],           // 前台：週四
  5: [2, 4],        // 預登1：週二＋週四
  6: [2, 4],        // 預登2注：週二＋週四
  7: [2, 4],        // 注射1：週二＋週四
  8: [2],           // 注射2：週二（已修正）
  9: 'bcg',         // 卡介苗：每月第一個週二（特殊規則）
  10: 'all',        // 停班2線：每日
};


// ── 解析員工編號→姓名（供前端確認排班者身份）──────────────────────
function resolveEmpIdToName(empId) {
  if (!empId) return '';
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    const names  = sheet.getRange('I1:I11').getValues().flat().map(s => s.toString().trim());
    const mIds   = sheet.getRange('M1:M11').getValues().flat().map(s => s.toString().trim());
    const norm   = id => (id.length>0 && isNaN(parseInt(id.charAt(0))))
      ? id.charAt(0).toUpperCase()+id.substring(1) : id;
    const ne = norm(empId);
    for (let i=0; i<mIds.length; i++) {
      if (norm(mIds[i]) === ne && names[i]) return names[i];
    }
    return '';
  } catch(e) { return ''; }
}

// ── 寫入排班者拖曳異動紀錄到 A欄日誌 ──────────────────────────────
// logs = [{date, header, oldName, newName, reason}]
function writeDragShiftLog(sheetName, arrangerEmpId, logs) {
  try {
    const arrangerName = resolveEmpIdToName(arrangerEmpId) || arrangerEmpId;
    const logSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    const logRange = logSheet.getRange('A40:A460').getValues();
    let logRow = 40;
    for (let i = 0; i < logRange.length; i++) {
      if (!logRange[i][0]) { logRow = i + 40; break; }
      if (i === logRange.length - 1) logRow = 460;
    }

    // ★ 建立去重集合：從現有日誌提取「排班者/審核者+日期+欄位+原→新」作為 key
    // 格式：「排班-XXX D/D 欄位 原A→新B」或「審核-XXX D/D 欄位 原A→新B」（忽略時間和備註）
    // bug #21: 原本 startsWith('排班-') filter 把所有「審核-」前綴的紀錄排除在 dedup 之外，
    // 審核者反覆寫入相同變更會無限重複堆疊，擠掉舊換班日誌。
    const existingKeys = new Set();
    logRange.forEach(row => {
      const s = (row[0] || '').toString().trim();
      if (!s.startsWith('排班-') && !s.startsWith('審核-')) return;
      // 擷取「排班-XXX D/D 欄位 原A→新B」這段（更換時間之前）
      const keyMatch = s.match(/^((排班|審核)-.+?\s+\d+\/\d+\s+.+?\s+原.+?→新.+?)\s+更換時間/);
      if (keyMatch) existingKeys.add(keyMatch[1]);
    });

    const tz = SpreadsheetApp.openById(SHEET_ID).getSpreadsheetTimeZone();
    const ts = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm');
    let written = 0;
    (logs || []).forEach(log => {
      const remarkStr = log.reason ? ' 備註:' + log.reason.trim() : '';
      const prefix = log.isAudit ? `審核-${arrangerName}` : `排班-${arrangerName}`;
      const d1 = (log.date||'').split(' ')[0];
      if(log.oldName !== log.newName){
        const key1 = `${prefix} ${d1} ${log.header} 原${log.oldName||'–'}→新${log.newName||'–'}`;
        if(!existingKeys.has(key1)){
          logSheet.getRange(logRow, 1).setValue(`${key1} 更換時間: ${ts}${remarkStr}`);
          existingKeys.add(key1);
          logRow = Math.min(logRow + 1, 460);
          written++;
        }
      }
      const d2 = (log.date2||'').split(' ')[0];
      if(log.date2 && log.header2 && log.oldName2 !== log.newName2 &&
         !(d1===d2 && log.header===log.header2)){
        const key2 = `${prefix} ${d2} ${log.header2} 原${log.oldName2||'–'}→新${log.newName2||'–'}`;
        if(!existingKeys.has(key2)){
          logSheet.getRange(logRow, 1).setValue(`${key2} 更換時間: ${ts}${remarkStr}`);
          existingKeys.add(key2);
          logRow = Math.min(logRow + 1, 460);
          written++;
        }
      }
    });
    if(written > 0) writeOpLog('排班拖曳', `${sheetName} ${logs[0]?.isAudit?'審核者':'排班者'}${arrangerName} 異動 ${written} 筆（去重後）`);
  } catch(e) { Logger.log('writeDragShiftLog error: ' + e.message); }
}

// ── 寫入已拖曳調整的預覽資料（直接覆蓋排班欄 C:M）──────────────────
// previewRows: [[col0..col10], ...] 同 runAutoSchedule 的 result 格式
function writeDraggedPreview(sheetName, adminPassword, previewRows) {
  if (!verifyAdminPassword(adminPassword)) return { success: false, message: '管理員密碼錯誤。' };
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: '找不到工作表：' + sheetName };
    if (!Array.isArray(previewRows) || !previewRows.length)
      return { success: false, message: '預覽資料為空。' };

    // Sanitize: ensure each row has exactly 11 cols (C:M), replace null/undefined with ''
    const rows = previewRows.map(function(row) {
      var r = Array.isArray(row) ? row.slice(0, 11) : [];
      while (r.length < 11) r.push('');
      return r.map(function(v){ return v == null ? '' : v; });
    });

    sheet.getRange('A2:A32').setNumberFormat('@');
    sheet.getRange(2, 3, rows.length, 11).setValues(rows); // C2:M(2+n)

    // Update N1 note
    // bug #24: 原本只 parse「審核狀態:」「writeCount:」，遇到 runAutoSchedule 累積寫入的
    //   「swapCount:」「公平性指標」等其他行就被丟棄。重寫 N1 後，跨月挪移計數歸零，
    //   下次自動排班 assignSlotsWithPointer 公平性會從 0 重來，長期挪移集中到少數人。
    //   保留所有不認識的行原樣寫回。
    const tz = spreadsheet.getSpreadsheetTimeZone();
    const ts = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm');
    let wc = 0;
    const passthroughLines = [];
    try {
      const old = sheet.getRange('N1').getNote() || '';
      old.split('\n').forEach(function(l){
        const trimmed = l.trim();
        if (!trimmed) return;
        if (trimmed.startsWith('writeCount:')) {
          wc = parseInt(trimmed.replace('writeCount:','').trim())||0;
          return;
        }
        // 排定時間 / 審核狀態 由本函式重寫，不沿用
        if (trimmed.startsWith('排定時間:') || trimmed.startsWith('審核狀態:')) return;
        // 其他行（含 swapCount、公平性、自訂備註）一律保留
        passthroughLines.push(trimmed);
      });
    } catch(e2){}
    const letter = wcToLetter(wc + 1);
    const baseLines = [`排定時間: ${ts}`, `審核狀態:pending`, `writeCount: ${wc+1}`];
    const newNote = baseLines.concat(passthroughLines).join('\n');
    sheet.getRange('N1').setNote(newNote);
    sheet.getRange('N1').setValue(sheetName + '　' + letter);

    writeOpLog('排班寫入（拖曳調整）', `${sheetName} 寫入 ${rows.length} 列`);
    return {
      success: true,
      message: `${sheetName} 排班完成（含拖曳調整）！（共 ${rows.length} 天）`,
      writeLetter: letter
    };
  } catch(e) {
    return { success: false, message: '寫入失敗：' + e.message };
  }
}

function getShiftDayRules() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    const stored = sheet.getRange('N3:N13').getValues().flat();
    const result = {};
    // 讀取已儲存的規則
    stored.forEach(raw => {
      const s = (raw || '').toString().trim();
      if (!s) return;
      const parts = s.split('|');
      if (parts.length < 2) return;
      const ci = parseInt(parts[0]);
      const days = parts[1].split(',').map(Number).filter(n => !isNaN(n));
      if (!isNaN(ci)) result[ci] = days;
    });
    // 合併預設值（有儲存的優先）
    const merged = {};
    for (const ci in SHIFT_DAY_DEFAULTS) {
      merged[ci] = result[ci] !== undefined ? result[ci] : SHIFT_DAY_DEFAULTS[ci];
    }
    return { success: true, rules: merged, defaults: SHIFT_DAY_DEFAULTS };
  } catch(e) {
    return { success: false, rules: SHIFT_DAY_DEFAULTS, defaults: SHIFT_DAY_DEFAULTS, error: e.message };
  }
}

function saveShiftDayRules(adminPassword, rules) {
  if (!verifyAdminPassword(adminPassword)) return { success: false, message: '管理員密碼錯誤' };
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_SHEET_NAME);
    // Clear N3:N13
    sheet.getRange('N3:N13').clearContent();
    const rows = [];
    let r = 0;
    for (const ciStr in rules) {
      const ci = parseInt(ciStr);
      const days = rules[ciStr];
      if (!Array.isArray(days)) continue; // skip 'bcg'/'all' special
      rows.push([ci + '|' + days.join(',')]);
      r++;
      if (r >= 11) break;
    }
    if (rows.length) sheet.getRange(3, 14, rows.length, 1).setValues(rows);
    writeOpLog('排班日設定', '已更新各職務排班星期規則');
    return { success: true, message: '排班規則已儲存' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}