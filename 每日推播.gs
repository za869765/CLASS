// ===== 完整整合版 Google Apps Script 代碼 =====
// 版本: 2026-08-28 v6.4
//
// === v6.4 修正 (2026-08-28) ===
// 天氣抓取強化(2026-08-27 06:19 自動推播整批「無法連線」——Open-Meteo 對
// Google 共用出口 IP 偶發限流/失敗,單次失敗即整篇無資料)
//   • 重試:非 200 或例外時間隔 2 秒重試,最多打 3 次,並記錄 HTTP 狀態碼
//   • 持久備援:成功即存精簡版資料到 ScriptProperties(跨執行存活);
//     3 次全失敗時改用 24 小時內的上次成功資料(forecast 為兩天份,舊數小時仍可用)
//   • 精簡欄位:只保留 daily 全部 + hourly.precipitation_probability/uv_index,
//     避免 ScriptProperties 9KB 上限風險(hourly.time 等未使用欄位不存)
//
// === v6.3 修正 (2026-08-25) ===
// 天氣來源由 wttr.in 更換為 Open-Meteo(https://open-meteo.com)
//   • 原因:wttr.in 對雲端 IP(GAS 共用 Google 出口 IP)限流,隨機回 500,
//     2026-08-25 起整批天氣/UV/日出日落顯示「無法連線」
//   • Open-Meteo 免金鑰、對雲端友善、經緯度直接鎖佳里(23.165, 120.177)
//   • 快取鍵改為 weatherData2_ 前綴,舊 wttr 格式快取自動失效,不需手動清
//   • 顯示字串格式維持原樣(降雨X%/A-B°C),getWeatherAdvice 判讀不受影響
//   • 失敗一律不寫入快取(維持原設計)
//
// === v6.2 新增 (2026-05-12) ===
// 「65 歲以上新冠疫苗第２劑開放提醒」推播區塊
//   • 觸發:今日為週一~週五才顯示,週六/週日跳過
//   • 目標週四:今天=週一~週四 → 本週四;今天=週五 → 下週四
//   • 公告日 = 目標週四 - 180 天 (民國年 YYY-MM-DD)
//   • 數字採用 xxl 字體放大顯示
//
// === v6.1 ===
// 加入驗證 Log,確認讀取的試算表正確
//
// === v6.0 ===
// 完全由程式碼主導,不依賴任何工作表讀取資料
// 1. 直接從來源試算表讀取班表(含欄位名稱動態讀取)
// 2. 直接呼叫天氣/UV/AQI API,不需要任何工作表
// 3. 圖示和顏色靠欄位順序(index)對應,改名不影響樣式
// 4. 新增/刪除欄位、人名即時生效(無快取)
// 5. 自動合併所有月份班表,永久解決跨月/跨年問題
// 6. 奇數日才執行推播

// ===== 來源試算表設定 =====
const SOURCE_SPREADSHEET_ID = '1NMiyJr0p6Vq6J2ubZy8xr3UArJhO-Vp3s4UXLOeqOUQ';

// ===== 天氣 API 設定(v6.3:Open-Meteo,佳里區座標)=====
const WEATHER_API_URL =
  'https://api.open-meteo.com/v1/forecast' +
  '?latitude=23.165&longitude=120.177' +
  '&daily=weather_code,temperature_2m_max,temperature_2m_min,sunrise,sunset' +
  '&hourly=precipitation_probability,uv_index' +
  '&forecast_days=2&timezone=Asia%2FTaipei';

// ===== 快取機制(僅天氣資料有快取,班表資料完全無快取)=====
var globalWeatherCache     = null;
var globalWeatherCacheTime = null;

function getWeatherDataCached() {
  var now = new Date().getTime();
  if (globalWeatherCache != null && globalWeatherCacheTime != null) {
    if (now - globalWeatherCacheTime < 300000) {
      Logger.log("使用全域變數快取");
      return globalWeatherCache;
    }
  }
  var cache    = CacheService.getScriptCache();
  // v6.3:鍵名改 weatherData2_,與舊 wttr 格式快取自然區隔
  var cacheKey = 'weatherData2_' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMddHH');
  var cached   = cache.get(cacheKey);
  if (cached != null) {
    Logger.log("使用 Script Cache");
    var data = JSON.parse(cached);
    globalWeatherCache     = data;
    globalWeatherCacheTime = now;
    return data;
  }
  // v6.4:重試最多 3 次(間隔 2 秒),並記錄每次失敗的 HTTP 狀態碼
  var data = null;
  for (var attempt = 1; attempt <= 3; attempt++) {
    try {
      Logger.log("從 API 取得天氣資料 (Open-Meteo)，第 " + attempt + " 次");
      var response = UrlFetchApp.fetch(WEATHER_API_URL, { muteHttpExceptions: true });
      var code = response.getResponseCode();
      if (code !== 200) {
        Logger.log("天氣 API 回應錯誤: HTTP " + code);
      } else {
        var parsed = JSON.parse(response.getContentText());
        if (!parsed.daily || !parsed.hourly) {
          Logger.log("天氣 API 資料格式異常");
        } else {
          data = slimWeatherData_(parsed);
          break;
        }
      }
    } catch (e) {
      Logger.log("天氣 API 錯誤: " + e.toString());
    }
    if (attempt < 3) Utilities.sleep(2000);
  }

  if (data) {
    cache.put(cacheKey, JSON.stringify(data), 3600);
    globalWeatherCache     = data;
    globalWeatherCacheTime = now;
    // v6.4:持久備援(跨執行存活;失敗日可用前次成功資料)
    try {
      PropertiesService.getScriptProperties()
        .setProperty('lastWeatherOk', JSON.stringify({ ts: now, data: data }));
    } catch (e) { Logger.log("備援寫入失敗: " + e.toString()); }
    return data;
  }

  // v6.4:3 次全失敗 → 讀 24 小時內的上次成功資料當備援
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('lastWeatherOk');
    if (raw) {
      var backup = JSON.parse(raw);
      if (backup && backup.data && (now - backup.ts) < 24 * 3600 * 1000) {
        Logger.log("⚠️ API 全數失敗，改用 " + Math.round((now - backup.ts) / 3600000) + " 小時前的備援天氣資料");
        globalWeatherCache     = backup.data;
        globalWeatherCacheTime = now;
        return backup.data;
      }
      Logger.log("備援資料已超過 24 小時，不使用");
    }
  } catch (e) { Logger.log("備援讀取失敗: " + e.toString()); }
  return null;
}

// v6.4:只留推播實際用到的欄位(daily 全部 + 降雨機率/UV)，
// 縮小 ScriptProperties 用量(單一屬性上限 9KB，hourly.time 等不存)
function slimWeatherData_(data) {
  return {
    daily: data.daily,
    hourly: {
      precipitation_probability: data.hourly.precipitation_probability,
      uv_index:                  data.hourly.uv_index
    }
  };
}

// ===== 設定區域 =====
const CONFIG = {
  LOG_SHEET_NAME: '系統日誌',

  LINE: {
    // ⚠️ repo 為公開備份,Token 不入庫;部署時貼回 GAS 請填入實際 LINE Channel Access Token
    ACCESS_TOKEN: '【LINE_CHANNEL_ACCESS_TOKEN】',
    GROUP_ID:     'C15ba3b411cf7bd2e1ec694ed1642cc64',
    API_URL:      'https://api.line.me/v2/bot/message/push'
  },

  COLORS: {
    PRIMARY:        '#1DB446',
    SECONDARY:      '#FF6B35',
    WARNING:        '#F39C12',
    DANGER:         '#E74C3C',
    INFO:           '#3498DB',
    SUCCESS:        '#27AE60',
    TEXT_PRIMARY:   '#2C3E50',
    TEXT_SECONDARY: '#7F8C8D'
  },

  AIR_QUALITY_COLORS: {
    '良好': '#27AE60', '普通': '#F1C40F',
    '橘色': '#E67E22', '紅色': '#E74C3C',
    '紫色': '#9B59B6', '危害': '#8E44AD',
    '綠':   '#27AE60', '黃':   '#F1C40F',
    '橘':   '#E67E22', '紅':   '#E74C3C',
    '紫':   '#9B59B6', '褐紅': '#8E44AD'
  }
};

// ===== 圖示和顏色設定(靠欄位 index 對應)=====
const DUTY_STYLES = [
  { icon: '📋', color: '#3498DB' },  // index 0:D欄(協助掛號/門診支援)
  { icon: '🏥', color: '#9B59B6' },  // index 1:E欄(門診)
  { icon: '💉', color: '#E67E22' },  // index 2:F欄(流注1/流感注射)
  { icon: '💉', color: '#E67E22' },  // index 3:G欄(流注2/流感注射)
  { icon: '📝', color: '#27AE60' },  // index 4:H欄(預登1/預注登記)
  { icon: '📝', color: '#27AE60' },  // index 5:I欄(預登2/預注登記)
  { icon: '💉', color: '#F39C12' },  // index 6:J欄(預注1/預防注射)
  { icon: '💉', color: '#F39C12' },  // index 7:K欄(預注2/預防注射)
  { icon: '💉', color: '#E74C3C' },  // index 8:L欄(卡介苗)
  { icon: '🦟', color: '#8E44AD' },  // index 9:M欄(登革熱二線)
];

// ===== 核心輔助函數 =====

function rocNumberToChinese(num) {
  const digits   = ['', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
  const hundreds = Math.floor(num / 100);
  const tens     = Math.floor((num % 100) / 10);
  const ones     = num % 10;
  let result = '';
  if (hundreds > 0) result += digits[hundreds] + '百';
  if (tens > 0)     result += digits[tens] + '十';
  if (ones > 0)     result += digits[ones];
  return result;
}

function toDateKey(date) {
  return `${date.getMonth() + 1}/${date.getDate()}`;
}

function formatDateWithWeekday(date) {
  const weekdays = ['週日', '週一', '週二', '週三', '週四', '週五', '週六'];
  return `${toDateKey(date)}${weekdays[date.getDay()]}`;
}

// ===== 核心:掃描來源試算表 =====

function getAllDutyData() {
  const dutyMap    = {};
  let   dutyLabels = null;

  try {
    const sourceSS  = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
    const allSheets = sourceSS.getSheets();

    Logger.log('========== 來源試算表驗證 ==========');
    Logger.log(`試算表名稱: ${sourceSS.getName()}`);
    Logger.log(`試算表 URL: ${sourceSS.getUrl()}`);
    Logger.log(`試算表 ID:  ${SOURCE_SPREADSHEET_ID}`);
    Logger.log(`共有工作表: ${allSheets.length} 個`);
    Logger.log('所有工作表清單:');
    allSheets.forEach(s => Logger.log(`  - ${s.getName()}`));
    Logger.log('=====================================');

    const rocYear   = new Date().getFullYear() - 1911;
    const rocYearCN = rocNumberToChinese(rocYear) + '年';
    Logger.log(`當前民國年: ${rocYear} → ${rocYearCN}`);

    for (const sheet of allSheets) {
      const name = sheet.getName();

      if (!name.includes('月班表') || !name.includes(rocYearCN)) {
        Logger.log(`跳過: ${name}`);
        continue;
      }

      Logger.log(`讀取: ${name}`);
      const values = sheet.getDataRange().getValues();
      if (values.length === 0) continue;

      if (!dutyLabels) {
        const headerRow = values[0];
        dutyLabels = headerRow
          .slice(3)
          .map(h => h ? h.toString().trim() : '')
          .filter(h => h !== '');
        Logger.log(`讀取到欄位標籤(${dutyLabels.length}個): ${JSON.stringify(dutyLabels)}`);
      }

      let rowCount = 0;
      for (const row of values) {
        const dateStr = row[0] ? row[0].toString().trim() : '';
        if (!dateStr.match(/^\d{1,2}\/\d{1,2}$/)) continue;
        dutyMap[dateStr] = row;
        rowCount++;
      }
      Logger.log(`  → 讀取 ${rowCount} 筆資料`);
    }

    Logger.log(`共合併 ${Object.keys(dutyMap).length} 天的班表資料`);

    return {
      dutyMap:    dutyMap,
      dutyLabels: dutyLabels || []
    };

  } catch (e) {
    Logger.log(`讀取來源試算表失敗: ${e.toString()}`);
    throw new Error(`無法讀取來源試算表,請確認 ID 及授權: ${e.toString()}`);
  }
}

// ===== 主要資料讀取函數 =====

function getSheetData() {
  try {
    const { dutyMap, dutyLabels } = getAllDutyData();

    const today    = new Date();
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);

    const todayKey    = toDateKey(today);
    const tomorrowKey = toDateKey(tomorrow);

    Logger.log(`查詢今日: ${todayKey}`);
    Logger.log(`查詢明日: ${tomorrowKey}`);

    const todayRow    = dutyMap[todayKey]    || null;
    const tomorrowRow = dutyMap[tomorrowKey] || null;

    if (!todayRow)    Logger.log(`⚠️ 找不到今日 ${todayKey} 的班表資料`);
    if (!tomorrowRow) Logger.log(`⚠️ 找不到明日 ${tomorrowKey} 的班表資料`);

    const buildExtraDuties = (row) =>
      dutyLabels.map((_, i) => [row ? (row[3 + i] || '') : '']);

    const data = {};
    data.todayDate        = formatDateWithWeekday(today);
    data.todayDuty        = todayRow ? (todayRow[2] || '(無)') : '(無)';
    data.todayDutyLabels  = dutyLabels.map(l => [l]);
    data.todayExtraDuties = buildExtraDuties(todayRow);

    Logger.log(`今日值班: ${data.todayDuty}`);
    Logger.log(`今日勤務明細:`);
    dutyLabels.forEach((l, i) => {
      const p = data.todayExtraDuties[i][0];
      if (p) Logger.log(`  ${l}: ${p}`);
    });

    data.tomorrowDate        = formatDateWithWeekday(tomorrow);
    data.tomorrowDuty        = tomorrowRow ? (tomorrowRow[2] || '(無)') : '(無)';
    data.tomorrowDutyLabels  = dutyLabels.map(l => [l]);
    data.tomorrowExtraDuties = buildExtraDuties(tomorrowRow);

    Logger.log(`明日值班: ${data.tomorrowDuty}`);

    Logger.log('開始取得天氣/AQI 資料...');
    data.todayWeather    = getTodayWeather();
    data.todayUV         = getTodayUV();
    data.todaySunrise    = getTodaySunrise();
    data.todaySunset     = getTodaySunset();
    data.todayAQI        = getTodayAQI();
    data.tomorrowWeather = getTomorrowWeather();
    data.tomorrowUV      = getTomorrowUV();

    Logger.log(`今日天氣: ${data.todayWeather}`);
    Logger.log(`今日AQI:  ${data.todayAQI}`);

    // ===== v6.2 新增:疫苗提醒 =====
    data.vaccineNotice = getVaccineNotice();
    Logger.log('疫苗提醒: ' + (data.vaccineNotice ? JSON.stringify(data.vaccineNotice) : '不顯示(週六/日)'));

    return data;

  } catch (error) {
    throw new Error(`讀取資料失敗: ${error.toString()}`);
  }
}

// ===== v6.2 新增:疫苗提醒功能 =====

/**
 * 西元 Date 轉民國年字串 YYY-MM-DD
 * 例:2025-11-15 → 114-11-15
 */
function toRocDate(date) {
  var rocYear = date.getFullYear() - 1911;
  var m       = (date.getMonth() + 1).toString().padStart(2, '0');
  var d       = date.getDate().toString().padStart(2, '0');
  return rocYear.toString().padStart(3, '0') + '-' + m + '-' + d;
}

/**
 * 距離天數轉自然語句:0→今天、1→明天、≥2→X 天後
 */
function formatDaysLater(days) {
  if (days === 0) return '今天';
  if (days === 1) return '明天';
  return days + ' 天後';
}

/**
 * 計算疫苗提醒資料
 * - 今天為週六/日 → 回傳 null(不顯示)
 * - 今天為週一~週四 → 目標 = 本週的週四
 * - 今天為週五      → 目標 = 下週四
 * - 公告日 = 目標週四 - 180 天(民國年格式)
 */
function getVaccineNotice() {
  var today = new Date();
  var dow   = today.getDay(); // 0=日 ~ 6=六

  if (dow === 0 || dow === 6) return null;

  // 週一(1)→+3, 週二(2)→+2, 週三(3)→+1, 週四(4)→0, 週五(5)→+6
  var addDays = (dow <= 4) ? (4 - dow) : (4 - dow + 7);

  var thursday = new Date(today.getFullYear(), today.getMonth(), today.getDate() + addDays);
  var eligible = new Date(thursday.getTime() - 180 * 24 * 60 * 60 * 1000);

  return {
    displayDate:   (thursday.getMonth() + 1) + '/' + thursday.getDate(),
    daysLaterText: formatDaysLater(addDays),
    eligibleDate:  toRocDate(eligible)
  };
}

/**
 * 疫苗提醒卡片
 * - 標題:65 歲以上新冠疫苗第２劑提醒
 * - 開放日:黑字單行,5/14 略大(lg)
 * - 接種日:紅字,114-11-15 最大(xxl),可跨兩行,無綠勾
 */
function createVaccineSection(notice) {
  return {
    type: 'box', layout: 'vertical',
    contents: [
      // 標題(單行)
      { type: 'box', layout: 'baseline', contents: [
        { type: 'text', text: '💉', size: 'lg', color: CONFIG.COLORS.DANGER, flex: 0 },
        { type: 'text', text: '65歲以上新冠疫苗第２劑提醒',
          weight: 'bold', color: CONFIG.COLORS.TEXT_PRIMARY, size: 'md',
          flex: 1, margin: 'sm' }
      ]},
      { type: 'separator', margin: 'sm' },

      // 開放施打日(黑字單行,5/14 略大)
      { type: 'box', layout: 'baseline', margin: 'md', contents: [
        { type: 'text', text: '📅', size: 'md', flex: 0 },
        { type: 'text', text: '於', size: 'sm', color: CONFIG.COLORS.TEXT_PRIMARY, flex: 0, margin: 'sm' },
        { type: 'text', text: notice.displayDate,
          size: 'lg', weight: 'bold', color: CONFIG.COLORS.TEXT_PRIMARY, flex: 0, margin: 'xs' },
        { type: 'text', text: '(週四,' + notice.daysLaterText + ')當天',
          size: 'sm', color: CONFIG.COLORS.TEXT_PRIMARY, flex: 1, margin: 'xs' }
      ]},

      // 接種日提醒(紅字,可跨兩行,無綠勾)
      { type: 'box', layout: 'vertical', margin: 'md', contents: [
        // 第一行:114-11-15(xxl)+ (含)以前
        { type: 'box', layout: 'baseline', contents: [
          { type: 'text', text: notice.eligibleDate,
            size: 'xxl', weight: 'bold', color: CONFIG.COLORS.DANGER, flex: 0 },
          { type: 'text', text: '(含)以前',
            size: 'sm', color: CONFIG.COLORS.DANGER, weight: 'bold', flex: 1, margin: 'sm' }
        ]},
        // 第二行:已施打第一劑者,可打第２劑。
        { type: 'text', text: '已施打第一劑者,可打第２劑。',
          size: 'sm', color: CONFIG.COLORS.DANGER, weight: 'bold',
          wrap: true, margin: 'xs' }
      ]}
    ],
    backgroundColor: '#FFFFFF', cornerRadius: '8px', paddingAll: '12px'
  };
}

// ===== 主要觸發函數 =====

/**
 * 每日自動執行(觸發器:每天 06:19)
 * 奇數日推播,偶數日跳過
 */
function dailyTriggerCheck() {
  const today = new Date();
  const date  = today.getDate();
  let logMsg  = formatTimestamp(today) + " - ";

  if (date % 2 === 1) {
    logMsg += `今天是 ${date} 日(奇數日),執行推播`;
    Logger.log(logMsg);
    logToSheet(logMsg);
    try {
      const result = sendDailyBroadcast();
      const msg = result.success
        ? `推播成功: ${result.message}`
        : `推播失敗: ${result.message}`;
      Logger.log(msg);
      logToSheet(msg);
    } catch (error) {
      const msg = `推播發生錯誤: ${error.toString()}`;
      Logger.log(msg);
      logToSheet(msg);
    }
  } else {
    logMsg += `今天是 ${date} 日(偶數日),不執行推播`;
    Logger.log(logMsg);
    logToSheet(logMsg);
  }
}

// ===== 推播函數 =====

function sendDailyBroadcast() {
  try {
    Logger.log(`${formatTimestamp(new Date())} - 開始執行推播`);
    const data    = getSheetData();
    Logger.log('所有資料讀取成功');
    const message = createEnhancedMessage(data);
    Logger.log('訊息建立成功');
    const result  = sendLineMessage(message);
    const logMsg  = result.success
      ? `成功: ${result.message}`
      : `失敗: ${result.message}`;
    logToSheet(`${formatTimestamp(new Date())} - 推播結果: ${logMsg}`);
    Logger.log(`推播結果: ${logMsg}`);
    return result;
  } catch (error) {
    const errMsg = `推播執行失敗: ${error.toString()}`;
    logToSheet(`${formatTimestamp(new Date())} - ${errMsg}`);
    Logger.log(errMsg);
    throw error;
  }
}

function sendLineMessage(message) {
  try {
    const options = {
      method: 'post',
      headers: {
        'Content-Type':  'application/json',
        'Authorization': `Bearer ${CONFIG.LINE.ACCESS_TOKEN}`
      },
      payload: JSON.stringify({ to: CONFIG.LINE.GROUP_ID, messages: [message] }),
      muteHttpExceptions: true
    };
    const response     = UrlFetchApp.fetch(CONFIG.LINE.API_URL, options);
    const responseCode = response.getResponseCode();
    if (responseCode === 200) {
      return { success: true,  message: `推播訊息已發送 (回應碼: ${responseCode})` };
    } else {
      return { success: false, message: `LINE API 回應錯誤 (${responseCode}): ${response.getContentText()}` };
    }
  } catch (error) {
    return { success: false, message: `發送失敗: ${error.toString()}` };
  }
}

// ===== LINE Flex Message =====

function createEnhancedMessage(data) {
  return {
    type: 'flex',
    altText: `值班機器人 - ${data.todayDate}`,
    contents: {
      type: 'bubble', size: 'mega',
      header: createEnhancedHeader(),
      body:   createEnhancedBody(data),
      footer: createEnhancedFooter()
    }
  };
}

function createEnhancedHeader() {
  return {
    type: 'box', layout: 'vertical',
    contents: [{
      type: 'text', text: '👥 值班人員資訊',
      weight: 'bold', color: '#FFFFFF', size: 'xl', align: 'center'
    }],
    backgroundColor: CONFIG.COLORS.PRIMARY,
    paddingAll: '20px'
  };
}

function createEnhancedBody(data) {
  const contents = [];

  const todayDuty = createEnhancedDutySection(data, 'today');
  if (todayDuty) { contents.push(todayDuty); contents.push(createSeparator()); }

  // ★ v6.2 新增:疫苗提醒(今日為週一~週五才顯示)
  if (data.vaccineNotice) {
    contents.push(createVaccineSection(data.vaccineNotice));
    contents.push(createSeparator());
  }

  const todayEnv = createEnhancedEnvironmentSection(
    data.todayWeather, data.todayUV, data.todayAQI, '今日'
  );
  if (todayEnv) { contents.push(todayEnv); contents.push(createSeparator()); }

  contents.push(createEnhancedSunTimeSection(data.todaySunrise, data.todaySunset));
  contents.push(createSeparator());

  const tomorrowDuty = createEnhancedDutySection(data, 'tomorrow');
  if (tomorrowDuty) { contents.push(tomorrowDuty); contents.push(createSeparator()); }

  const tomorrowEnv = createEnhancedEnvironmentSection(
    data.tomorrowWeather, data.tomorrowUV, null, '明日'
  );
  if (tomorrowEnv) { contents.push(tomorrowEnv); }

  return {
    type: 'box', layout: 'vertical',
    contents: contents, paddingAll: '0px', spacing: 'none'
  };
}

function createEnhancedDutySection(data, type) {
  const isToday     = (type === 'today');
  const date        = isToday ? data.todayDate        : data.tomorrowDate;
  const duty        = isToday ? data.todayDuty        : data.tomorrowDuty;
  const extraDuties = isToday ? data.todayExtraDuties : data.tomorrowExtraDuties;
  const dutyLabels  = isToday ? data.todayDutyLabels  : data.tomorrowDutyLabels;
  const label       = isToday ? '今日' : '明日';

  const contents = [
    { type: 'box', layout: 'horizontal', contents: [
      { type: 'text', text: '👥', size: 'lg', color: CONFIG.COLORS.PRIMARY, flex: 0 },
      { type: 'text', text: '勤務資訊', weight: 'bold', color: CONFIG.COLORS.TEXT_PRIMARY, size: 'lg', flex: 1, margin: 'sm' }
    ]},
    { type: 'separator', margin: 'md' },
    { type: 'box', layout: 'horizontal', margin: 'md', contents: [
      { type: 'text', text: '📅', size: 'md', flex: 0 },
      { type: 'text', text: `${label}: ${date}`, weight: 'bold', size: 'md', color: CONFIG.COLORS.INFO, flex: 1, margin: 'sm' }
    ]},
    { type: 'box', layout: 'horizontal', margin: 'sm', contents: [
      { type: 'text', text: '👤', size: 'md', flex: 0 },
      { type: 'text', text: `值班: ${duty}`, weight: 'bold', size: 'md', color: CONFIG.COLORS.DANGER, flex: 1, margin: 'sm' }
    ]}
  ];

  for (let i = 0; i < extraDuties.length; i++) {
    const lbl    = dutyLabels[i][0];
    const person = extraDuties[i][0];
    if (person && person.toString().trim() !== '') {
      const style = DUTY_STYLES[i] || { icon: '📌', color: CONFIG.COLORS.TEXT_PRIMARY };
      contents.push({
        type: 'box', layout: 'horizontal', margin: 'sm', contents: [
          { type: 'text', text: style.icon, size: 'md', flex: 0 },
          { type: 'text', text: `${lbl}: ${person}`, size: 'sm', color: style.color, flex: 1, margin: 'sm' }
        ]
      });
    }
  }

  return {
    type: 'box', layout: 'vertical', contents: contents,
    backgroundColor: '#FFFFFF', cornerRadius: '8px', paddingAll: '15px'
  };
}

function createEnhancedEnvironmentSection(weather, uv, aqi, label) {
  const contents = [
    { type: 'box', layout: 'horizontal', contents: [
      { type: 'text', text: '🌤️', size: 'lg', color: CONFIG.COLORS.SECONDARY, flex: 0 },
      { type: 'text', text: `${label}環境資訊`, weight: 'bold', color: CONFIG.COLORS.TEXT_PRIMARY, size: 'lg', flex: 1, margin: 'sm' }
    ]},
    { type: 'separator', margin: 'md' }
  ];

  if (weather) contents.push({
    type: 'text', text: `☁️ 天氣: ${weather.toString()}`,
    size: 'md', color: CONFIG.COLORS.TEXT_PRIMARY, wrap: true, margin: 'sm'
  });

  if (uv) {
    const uvVal = parseInt((uv.toString().match(/(\d+)/) || [0, 0])[1]);
    contents.push({
      type: 'text', text: `☀️ 紫外線指數: ${uv.toString()}`,
      size: 'md', color: getUVColor(uvVal), weight: 'bold', margin: 'sm'
    });
  }

  if (aqi) contents.push({
    type: 'text', text: `🌬️ 空氣品質: ${aqi.toString()}`,
    size: 'md', color: getAirQualityColor(aqi.toString()), weight: 'bold', wrap: true, margin: 'sm'
  });

  if (weather && uv) {
    const advice = getWeatherAdvice(
      weather.toString(), uv.toString(), (aqi || '無 AQI 資料').toString()
    );
    contents.push({ type: 'separator', margin: 'md' });
    contents.push({
      type: 'text', text: `💡 ${advice}`,
      size: 'sm', color: CONFIG.COLORS.INFO, weight: 'bold', wrap: true, margin: 'sm'
    });
  }

  return {
    type: 'box', layout: 'vertical', contents: contents,
    backgroundColor: '#FFFFFF', cornerRadius: '8px', paddingAll: '15px'
  };
}

function createEnhancedSunTimeSection(sunriseTime, sunsetTime) {
  return {
    type: 'box', layout: 'vertical',
    contents: [
      { type: 'box', layout: 'horizontal', contents: [
        { type: 'text', text: '🌅', size: 'lg', color: CONFIG.COLORS.WARNING, flex: 0 },
        { type: 'text', text: '日出日落', weight: 'bold', color: CONFIG.COLORS.TEXT_PRIMARY, size: 'lg', flex: 1, margin: 'sm' }
      ]},
      { type: 'separator', margin: 'md' },
      { type: 'box', layout: 'horizontal', contents: [
        { type: 'box', layout: 'vertical', flex: 1, contents: [
          { type: 'text', text: '🌅 日出', size: 'sm', color: CONFIG.COLORS.TEXT_SECONDARY },
          { type: 'text', text: sunriseTime || '--:--', weight: 'bold', size: 'lg', color: CONFIG.COLORS.WARNING }
        ]},
        { type: 'box', layout: 'vertical', flex: 1, contents: [
          { type: 'text', text: '🌇 日落', size: 'sm', color: CONFIG.COLORS.TEXT_SECONDARY },
          { type: 'text', text: sunsetTime || '--:--', weight: 'bold', size: 'lg', color: CONFIG.COLORS.DANGER }
        ]}
      ]}
    ],
    backgroundColor: '#FFFFFF', cornerRadius: '8px', paddingAll: '15px'
  };
}

function createSeparator() {
  return { type: 'separator', margin: 'md', color: '#E0E0E0' };
}

function createEnhancedFooter() {
  return {
    type: 'box', layout: 'vertical',
    contents: [{
      type: 'text', text: `📅 更新時間: ${formatTimestamp(new Date())}`,
      size: 'xs', color: CONFIG.COLORS.TEXT_SECONDARY, align: 'center'
    }],
    paddingAll: '10px'
  };
}

// ===== 輔助函數 =====

function getAirQualityColor(description) {
  for (const [key, color] of Object.entries(CONFIG.AIR_QUALITY_COLORS)) {
    if (description.includes(key)) return color;
  }
  return CONFIG.COLORS.TEXT_PRIMARY;
}

function formatTimestamp(date) {
  return `${date.getFullYear()}-` +
    `${(date.getMonth()+1).toString().padStart(2,'0')}-` +
    `${date.getDate().toString().padStart(2,'0')} ` +
    `${date.getHours().toString().padStart(2,'0')}:` +
    `${date.getMinutes().toString().padStart(2,'0')}`;
}

function logToSheet(message) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
      logSheet.appendRow(['時間', '訊息']);
    }
    logSheet.appendRow([formatTimestamp(new Date()), message]);
  } catch (error) {
    Logger.log(`記錄日誌失敗: ${error.toString()}`);
  }
}

// ===== 天氣 API 函數(v6.3:Open-Meteo 資料格式)=====

// 組出「晴  降雨30%(下午2點) 24-29°C」格式字串(今日 day=0 / 明日 day=1)
// 字串格式與 v6.2 完全相同,getWeatherAdvice 的判讀正則不受影響
function buildWeatherText(data, day) {
  var desc = wmoToChinese(data.daily.weather_code[day]);
  var probs = data.hourly.precipitation_probability.slice(day * 24, day * 24 + 24);
  var maxRain = 0, maxRainTime = "";
  for (var i = 0; i < probs.length; i++) {
    var r = parseInt(probs[i]);
    if (!isNaN(r) && r > maxRain) {
      maxRain = r;
      maxRainTime = i < 6 ? "凌晨" + i + "點" : i < 12 ? "上午" + i + "點" : i === 12 ? "中午12點" : "下午" + (i - 12) + "點";
    }
  }
  var minT = Math.round(data.daily.temperature_2m_min[day]);
  var maxT = Math.round(data.daily.temperature_2m_max[day]);
  return desc +
    (maxRain > 0 ? `  降雨${maxRain}%(${maxRainTime})` : `  降雨${maxRain}%`) +
    ` ${minT}-${maxT}°C`;
}

function getTodayWeather() {
  try {
    var data = getWeatherDataCached();
    if (!data) return "無法連線到天氣服務";
    return buildWeatherText(data, 0);
  } catch (e) { return "今日天氣讀取失敗"; }
}

function getTomorrowWeather() {
  try {
    var data = getWeatherDataCached();
    if (!data) return "無法連線到天氣服務";
    return buildWeatherText(data, 1);
  } catch (e) { return "明日天氣讀取失敗"; }
}

// UV 取白天時段(06~18時)的最小-最大(夜間恆為 0,計入會失真)
function buildUVText(data, day) {
  var uvs = data.hourly.uv_index
    .slice(day * 24 + 6, day * 24 + 19)
    .map(function(v) { return Math.round(v); })
    .filter(function(v) { return !isNaN(v); });
  if (!uvs.length) return "無 UV 資料";
  var min = Math.min.apply(null, uvs), max = Math.max.apply(null, uvs);
  return (min === max ? min : min + "-" + max) + " " + getUVLevel(max);
}

function getTodayUV() {
  try {
    var data = getWeatherDataCached();
    if (!data) return "無法連線";
    return buildUVText(data, 0);
  } catch (e) { return "紫外線資料讀取失敗"; }
}

function getTomorrowUV() {
  try {
    var data = getWeatherDataCached();
    if (!data) return "無法連線";
    return buildUVText(data, 1);
  } catch (e) { return "紫外線資料讀取失敗"; }
}

function getTodaySunrise() {
  try {
    var data = getWeatherDataCached();
    if (!data) return "無法連線";
    var v = data.daily.sunrise[0]; // ISO 格式 "2026-08-25T05:39"
    return v ? v.slice(11, 16) : "無日出資料";
  } catch (e) { return "日出時間讀取失敗"; }
}

function getTodaySunset() {
  try {
    var data = getWeatherDataCached();
    if (!data) return "無法連線";
    var v = data.daily.sunset[0];
    return v ? v.slice(11, 16) : "無日落資料";
  } catch (e) { return "日落時間讀取失敗"; }
}

function getTodayAQI() {
  try {
    var url  = "https://data.moenv.gov.tw/api/v2/aqx_p_432" +
               "?api_key=846e44e1-8cc5-4893-ad87-c79d2d383706" +
               "&limit=1000&sort=ImportDate%20desc&format=json";
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return "無法連線";
    var data    = JSON.parse(resp.getContentText());
    var records = Array.isArray(data) ? data : (data && data.records ? data.records : null);
    if (!records) return "資料格式錯誤";
    for (var i = 0; i < records.length; i++) {
      if (records[i].sitename === "善化") {
        var aqi = records[i].aqi;
        if (!aqi || aqi === "" || aqi === "-") return "善化站資料維護中";
        return `${aqi} ${records[i].status} (${getAQIColor(parseInt(aqi))})`;
      }
    }
    return "善化站無資料";
  } catch (e) { return "AQI 讀取失敗"; }
}

// ===== 計算輔助函數 =====

function getAQIColor(aqi) {
  if (isNaN(aqi)) return "無";
  if (aqi <= 50)  return "綠";
  if (aqi <= 100) return "黃";
  if (aqi <= 150) return "橘";
  if (aqi <= 200) return "紅";
  if (aqi <= 300) return "紫";
  return "褐紅";
}

function getUVLevel(uv) {
  if (uv <= 2)  return "低量級";
  if (uv <= 5)  return "中量級";
  if (uv <= 7)  return "高量級";
  if (uv <= 10) return "過量級";
  return "危險級";
}

function getUVColor(uv) {
  if (uv <= 5) return CONFIG.COLORS.SUCCESS;
  if (uv <= 7) return CONFIG.COLORS.WARNING;
  return CONFIG.COLORS.DANGER;
}

// v6.3:WMO 天氣代碼 → 中文(Open-Meteo weather_code)
function wmoToChinese(code) {
  var t = {
    0:"晴", 1:"大致晴朗", 2:"多雲", 3:"陰",
    45:"霧", 48:"霧淞",
    51:"毛毛雨", 53:"毛毛雨", 55:"毛毛雨",
    56:"凍雨", 57:"凍雨",
    61:"小雨", 63:"中雨", 65:"大雨",
    66:"凍雨", 67:"凍雨",
    71:"小雪", 73:"中雪", 75:"大雪", 77:"雪粒",
    80:"陣雨", 81:"陣雨", 82:"強陣雨",
    85:"陣雪", 86:"陣雪",
    95:"雷雨", 96:"雷陣雨", 99:"強雷雨"
  };
  return t[code] !== undefined ? t[code] : "多雲";
}

function getWeatherAdvice(weather, uvInfo, aqiInfo) {
  var advice   = [];
  var rainProb = parseInt((weather.match(/降雨(\d+)%/) || [0,0])[1]) || 0;
  var tMatch   = weather.match(/(\d+)-(\d+)°C/);
  var minTemp  = tMatch ? parseInt(tMatch[1]) : 20;
  var maxTemp  = tMatch ? parseInt(tMatch[2]) : 25;
  var tempDiff = maxTemp - minTemp;
  var uvMatch  = uvInfo.match(/(\d+)-(\d+)/);
  var uvMin, uvMax;
  if (uvMatch) { uvMin=parseInt(uvMatch[1]); uvMax=parseInt(uvMatch[2]); }
  else { var s=uvInfo.match(/(\d+)/); uvMin=uvMax=s?parseInt(s[1]):0; }
  var aqi = parseInt((aqiInfo.match(/(\d+)/) || [0,50])[1]) || 50;

  if      (rainProb >= 70) advice.push("高機率降雨請帶傘");
  else if (rainProb >= 50) advice.push("記得帶傘");
  else if (rainProb >= 30) advice.push("建議帶傘備用");

  if      (tempDiff >= 8) advice.push("早晚溫差大請帶外套保暖");
  else if (tempDiff >= 6) advice.push("注意溫差建議帶外套");

  if      (maxTemp >= 35) advice.push(minTemp>=35?"天氣酷熱請多補充水分":"中午天氣酷熱請多補充水分");
  else if (maxTemp >= 30) advice.push(minTemp>=30?"天氣焰熱多補充水分":"中午天氣焰熱多補充水分");
  else if (minTemp <= 15) advice.push("天氣寒冷請保暖");
  else if (minTemp <= 21) advice.push("天涼請穿外套");

  if      (uvMax >= 11) advice.push(uvMin>=11?"建議不要外出":"中午建議不要外出");
  else if (uvMax >= 8)  advice.push(uvMin>=8?"請減少外出並做好防曬":"中午請減少外出並做好防曬");
  else if (uvMax >= 6)  advice.push(uvMin>=6?"外出請注意防曬":"中午外出請注意防曬");

  if      (aqi >= 200) advice.push("建議不要外出,門窗關好");
  else if (aqi >= 150) advice.push("減少外出並戴口罩");
  else if (aqi >  100) advice.push("外出請戴口罩");
  else if (aqi >   50) advice.push("敏感族群外出請戴口罩");

  if (!advice.length) return "天氣良好,適合戶外活動";
  return advice.slice(0,3).join(",");
}

// ===== 測試函數 =====

/**
 * 測試1:驗證試算表連線及所有資料(先執行這個)
 * 執行後查看 Log,確認「試算表名稱」正確
 */
function testGetSheetData() {
  try {
    const data = getSheetData();
    Logger.log('===== 今日勤務 =====');
    Logger.log('日期: ' + data.todayDate);
    Logger.log('值班: ' + data.todayDuty);
    data.todayDutyLabels.forEach((l, i) => {
      const p = data.todayExtraDuties[i][0];
      if (p) Logger.log(`  ${l[0]}: ${p}`);
    });
    Logger.log('===== 明日勤務 =====');
    Logger.log('日期: ' + data.tomorrowDate);
    Logger.log('值班: ' + data.tomorrowDuty);
    data.tomorrowDutyLabels.forEach((l, i) => {
      const p = data.tomorrowExtraDuties[i][0];
      if (p) Logger.log(`  ${l[0]}: ${p}`);
    });
    Logger.log('===== 環境資料 =====');
    Logger.log('今日天氣: '   + data.todayWeather);
    Logger.log('今日UV: '     + data.todayUV);
    Logger.log('今日日出: '   + data.todaySunrise);
    Logger.log('今日日落: '   + data.todaySunset);
    Logger.log('今日AQI: '    + data.todayAQI);
    Logger.log('明日天氣: '   + data.tomorrowWeather);
    Logger.log('明日UV: '     + data.tomorrowUV);
    Logger.log('===== v6.2 疫苗提醒 =====');
    Logger.log('疫苗資料: ' + (data.vaccineNotice ? JSON.stringify(data.vaccineNotice) : '不顯示(週六/日)'));
  } catch (e) {
    Logger.log('測試失敗: ' + e.toString());
  }
}

/**
 * 測試2:清除天氣快取後重新推播
 */
function clearCacheAndTest() {
  globalWeatherCache     = null;
  globalWeatherCacheTime = null;
  var cache    = CacheService.getScriptCache();
  var cacheKey = 'weatherData2_' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMddHH');
  cache.remove(cacheKey);
  Logger.log('天氣快取已清除');
  sendDailyBroadcast();
}

/**
 * 測試3:v6.2 新增 — 單獨測試疫苗提醒邏輯
 * 印出今天/明天/各種星期的計算結果
 */
function testVaccineNotice() {
  Logger.log('===== 疫苗提醒測試 =====');
  var notice = getVaccineNotice();
  if (notice) {
    Logger.log('開放施打日: ' + notice.displayDate);
    Logger.log('距今: '       + notice.daysLaterText);
    Logger.log('接種日(含)以前: ' + notice.eligibleDate);
  } else {
    Logger.log('今天為週六/日,不顯示疫苗區塊');
  }
}
