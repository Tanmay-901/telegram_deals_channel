/**
 * Telegram Multi-Channel Deals Pipeline
 *
 * SCRIPT PROPERTIES (Project Settings):
 *   BOT_TOKEN        - Telegram Bot Token from @BotFather
 *   CUELINKS_API_KEY - Cuelinks API Key
 *
 * TIME-DRIVEN TRIGGERS (Triggers icon):
 *   sweepPendingRows    -> Minutes timer (every 2-3 min)
 *   runAllIntakes       -> Minutes timer (every 30-60 min)
 *   archiveRecentDeals  -> Hour timer    (every 4 hours)
 *   runDailyMaintenance -> Day timer     (4:00 AM - 5:00 AM)
 */

// ======================================================================
// SHEET NAMES AND COLUMN INDICES
// ======================================================================

const CHANNELS_SHEET      = 'Channels';
const LIVE_QUEUE_SHEET    = 'Live_Queue';
const RECENT_POSTED_SHEET = 'Recent_Posted';
const CONFIG_SHEET        = 'Configuration';
const ARCHIVE_SHEET       = 'Historical_Archive';
const LOGS_SHEET          = 'Logs';
const HEADER_ROW          = 1;
const ROW_WIDTH           = 8; // All deal sheets are A-H

// ======================================================================
// Execution-time monitoring (daily total runtime)
// ======================================================================
// Best-effort: if the “Configurations” sheet isn’t present, the script will
// still track counters in PropertiesService, and will never break the
// main pipeline.
const CONFIG_MON_SHEET_NAME = 'Configurations';
const CONFIG_DAILY_RUNTIME_CELL_A1 = 'B2';

// Per-handler last runtime visualization.
// (If you change handler function names, update this mapping.)
const CONFIG_LAST_RUNTIME_CELL_BY_HANDLER = {
  runAllIntakes: 'B3',
  onEditInstallable: 'B4',
  sweepPendingRows: 'B5'
};

// Throttle writes to the sheet to reduce overhead.
const CONFIG_WRITE_MIN_INTERVAL_MS = 30 * 1000;

// Where we store counters/state in PropertiesService
const PROP_EXEC_DAILY_DATE = 'DAILY_EXEC_DATE';
const PROP_EXEC_DAILY_TOTAL_MS = 'DAILY_EXEC_TOTAL_MS';
const PROP_EXEC_LAST_WRITE_AT = 'CONFIG_LAST_WRITE_AT_EPOCH_MS';

// For PropertiesService-only snapshots (optional; no dependency on sheet)
const PROP_EXEC_LAST_RUNTIME_MS_PREFIX = 'LAST_RUNTIME_MS_';

// Channels tab (1-based)
const CH = {
  NAME:1, CHAT_ID:2, KEYWORDS:3, NEG_KEYWORDS:4,
  EMOJI:5, ENABLED:6, INTERVAL:7, MAX_DEALS:8
};

// Live_Queue / Recent_Posted / Archive (1-based)
// Column E = Target_Channel (replaced Category from v1)
const COL = {
  NAME:1, LINK:2, ORIG:3, DEAL:4,
  TARGET:5, STATUS:6, COUPON:7, PCT_OFF:8
};

// ======================================================================
// SHEET ACCESSORS
// ======================================================================

function ss_() { return SpreadsheetApp.getActive(); }

function msToHuman_(ms) {
  ms = Number(ms) || 0;
  const totalSeconds = Math.floor(ms / 1000);
  const s = totalSeconds % 60;
  const totalMinutes = Math.floor(totalSeconds / 60);
  const m = totalMinutes % 60;
  const h = Math.floor(totalMinutes / 60);
  if (h > 0) return `${h}h ${m}m ${s}s`;
  if (m > 0) return `${m}m ${s}s`;
  return `${s}s`;
}

function epochMsToHuman_(epochMs) {
  const n = Number(epochMs);
  if (!isFinite(n) || n <= 0) return '';
  return new Date(n).toLocaleString();
}

function getTodayKey_() {
  const d = new Date();
  // YYYY-MM-DD (local time)
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function recordExecutionRuntime_(handlerName, startEpochMs_) {
  // Never allow monitoring to break the main deal pipeline.
  try {
    const start = Number(startEpochMs_);
    if (!isFinite(start) || start <= 0) return;

    const elapsedMs = Date.now() - start;
    if (!isFinite(elapsedMs) || elapsedMs < 0) return;

    const props = PropertiesService.getScriptProperties();
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return;

    try {
      const todayKey = getTodayKey_();

      // Reset on first run after date changes
      const lastResetDate = props.getProperty(PROP_EXEC_DAILY_DATE);
      if (lastResetDate !== todayKey) {
        props.setProperty(PROP_EXEC_DAILY_DATE, todayKey);
        props.setProperty(PROP_EXEC_DAILY_TOTAL_MS, '0');
      }

      const prevTotalMs = Number(props.getProperty(PROP_EXEC_DAILY_TOTAL_MS) || '0');
      const newTotalMs = prevTotalMs + elapsedMs;
      props.setProperty(PROP_EXEC_DAILY_TOTAL_MS, String(newTotalMs));

      // Best-effort store per-handler snapshot in properties (no required UI dependency)
      props.setProperty(PROP_EXEC_LAST_RUNTIME_MS_PREFIX + handlerName, String(elapsedMs));

      // Throttle sheet writes
      const now = Date.now();
      const lastWrite = Number(props.getProperty(PROP_EXEC_LAST_WRITE_AT) || '0');
      if (now - lastWrite < CONFIG_WRITE_MIN_INTERVAL_MS) return;
      props.setProperty(PROP_EXEC_LAST_WRITE_AT, String(now));

      // Best-effort: update the Configurations tab cells (if present)
      const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG_MON_SHEET_NAME);
      if (!sheet) return;

      // Daily total
      sheet.getRange(CONFIG_DAILY_RUNTIME_CELL_A1).setValue(msToHuman_(newTotalMs));

      // Per-handler last runtime
      const cellA1 = (CONFIG_LAST_RUNTIME_CELL_BY_HANDLER && CONFIG_LAST_RUNTIME_CELL_BY_HANDLER[handlerName])
        ? CONFIG_LAST_RUNTIME_CELL_BY_HANDLER[handlerName]
        : null;

      if (cellA1) {
        const lastRunAtEpochMs = Date.now();
        const prettyTime = epochMsToHuman_(lastRunAtEpochMs);
        sheet.getRange(cellA1).setValue(`${msToHuman_(elapsedMs)} | ${prettyTime}`);
      }

    } finally {
      lock.releaseLock();
    }
  } catch (e) {
    // swallow
  }
}

function getOrCreateSheet_(name, headers) {
  let sheet = ss_().getSheetByName(name);
  if (!sheet) {
    sheet = ss_().insertSheet(name);
    if (headers && headers.length) sheet.appendRow(headers);
  }
  return sheet;
}

function getChannelsSheet_() {
  const sheet = getOrCreateSheet_(CHANNELS_SHEET, [
    'Channel_Name','Chat_ID','Keywords','Negative_Keywords',
    'Emoji','Enabled','Min_Interval_Min','Max_Deals_Per_Run','Notes'
  ]);
  if (sheet.getLastRow() === 1) {
    sheet.getRange(2, 1, 4, 9).setValues([
      ['Gym & Fitness', '@YourGymChannel', 'whey, protein, creatine, dumbbell, shaker, gym', 'tshirt, apparel, socks', '💪', false, 5, 2, 'Example fitness channel'],
      ['Pet Care', '@YourPetChannel', 'dog, cat, pedigree, supertails, kibble, leash, pet', 'plush toy, stuffed', '🐾', false, 10, 2, 'Example pet care channel'],
      ['Automotive', '@YourVehicleChannel', 'car, bike, dashcam, tyre, helmet, lubricant', 'toy car, hot wheels, rc car', '🚗', false, 10, 2, 'Example auto channel'],
      ['Mega Loot (All Deals)', '@YourAllDealsChannel', '*', 'expired, sample', '🔥', false, 2, 5, 'Wildcard channel - matches all deals']
    ]);
  }
  return sheet;
}
function getLiveQueueSheet_() {
  return getOrCreateSheet_(LIVE_QUEUE_SHEET, [
    'Product_Name','Raw_URL','Original_Price','Deal_Price',
    'Target_Channel','Status','Coupon_Code','Percent_Off'
  ]);
}
function getConfigSheet_() {
  const sheet = getOrCreateSheet_(CONFIG_SHEET, ['Key', 'Value']);
  if (sheet.getLastRow() === 1) {
    sheet.getRange(2, 1, 5, 2).setValues([
      ['IS_SALE_MODE_ACTIVE', false],
      ['DEDUP_WINDOW_HOURS', 4],
      ['LOG_RETENTION_HOURS', 24],
      ['LOG_MAX_ROWS', 1000],
      ['ARCHIVE_RETENTION_DAYS', 30]
    ]);
  }
  return sheet;
}
function getRecentPostedSheet_() {
  return getOrCreateSheet_(RECENT_POSTED_SHEET, [
    'Product_Name','Raw_URL','Original_Price','Deal_Price',
    'Target_Channel','Sent_At','Coupon_Code','Percent_Off'
  ]);
}
function getArchiveSheet_() {
  return getOrCreateSheet_(ARCHIVE_SHEET, [
    'Product_Name','Raw_URL','Original_Price','Deal_Price',
    'Target_Channel','Sent_At','Coupon_Code','Percent_Off'
  ]);
}

// ======================================================================
// CONFIGURATION DASHBOARD
// ======================================================================

const CONFIG_DEFAULTS = {
  IS_SALE_MODE_ACTIVE:    false,
  DEDUP_WINDOW_HOURS:     4,
  LOG_RETENTION_HOURS:    24,
  LOG_MAX_ROWS:           1000,
  ARCHIVE_RETENTION_DAYS: 30
};

const SALE_MODE_BURST_CAP = 20;

function readConfig_() {
  const sheet = getConfigSheet_();
  if (!sheet) return {};
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return {};
  const values = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, 2).getValues();
  const map = {};
  values.forEach(row => { const k = String(row[0]).trim(); if (k) map[k] = row[1]; });
  return map;
}

function getConfigNumber_(config, key, fallback) {
  const v = config[key]; const n = Number(v);
  return (v !== undefined && v !== '' && !isNaN(n)) ? n : fallback;
}
function getConfigBool_(config, key, fallback) {
  const v = config[key];
  if (v === undefined || v === '') return fallback;
  if (typeof v === 'boolean') return v;
  return String(v).trim().toUpperCase() === 'TRUE';
}
function getConfigDate_(config, key) {
  const v = config[key]; if (!v) return null;
  const d = (v instanceof Date) ? v : new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function setConfigValue_(key, value) {
  const sheet = getConfigSheet_(); if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow > HEADER_ROW) {
    const keys = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).trim() === key) {
        sheet.getRange(HEADER_ROW + 1 + i, 2).setValue(value);
        return;
      }
    }
  }
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, 2).setValues([[key, value]]);
}

// ======================================================================
// CHANNELS TAB - read and compile channel configs
// ======================================================================

/**
 * Reads all enabled rows from the Channels tab and returns compiled objects:
 * { name, chatId, emoji, isWildcard, posRegexes, negRegexes,
 *   minInterval (minutes), maxDeals }
 * Regex arrays are compiled once here and reused throughout each run.
 */
function readChannels_() {
  const sheet = getChannelsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return [];

  const values = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, CH.MAX_DEALS).getValues();
  const channels = [];

  values.forEach(row => {
    const enabledRaw = row[CH.ENABLED - 1];
    const enabled = (typeof enabledRaw === 'boolean')
      ? enabledRaw : String(enabledRaw).trim().toUpperCase() === 'TRUE';
    if (!enabled) return;

    const name   = String(row[CH.NAME   - 1]).trim();
    const chatId = String(row[CH.CHAT_ID - 1]).trim();
    if (!name || !chatId) return;

    const emoji    = String(row[CH.EMOJI    - 1]).trim() || '';
    const interval = Math.max(1, Number(row[CH.INTERVAL  - 1]) || 5);
    const maxDeals = Math.max(1, Number(row[CH.MAX_DEALS - 1]) || 2);

    const kwStr    = String(row[CH.KEYWORDS - 1] || '').trim();
    const isWildcard = (kwStr === '*' || kwStr.toUpperCase() === 'ALL' || kwStr === '');

    const keywords    = isWildcard ? [] : parseKeywords_(row[CH.KEYWORDS     - 1], []);
    const negKeywords =                   parseKeywords_(row[CH.NEG_KEYWORDS - 1], []);

    const posRegexes = keywords.map(kw =>
      new RegExp('\\b' + escapeRegex_(kw) + 's?\\b', 'i')
    );
    const negRegexes = negKeywords.map(kw =>
      new RegExp('\\b' + escapeRegex_(kw) + 's?\\b', 'i')
    );

    channels.push({ name, chatId, emoji, isWildcard,
                    posRegexes, negRegexes,
                    minInterval: interval, maxDeals });
  });

  return channels;
}

function escapeRegex_(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Returns the names of channels that a given haystack matches.
 * Negative keywords are evaluated first; any match disqualifies that channel.
 * Then positive keywords (or wildcard) are checked.
 */
function matchChannels_(haystack, channels) {
  const matched = [];
  for (const ch of channels) {
    if (ch.negRegexes.some(re => re.test(haystack))) continue;
    if (ch.isWildcard || ch.posRegexes.some(re => re.test(haystack))) matched.push(ch.name);
  }
  return matched;
}

// ======================================================================
// KEYWORD PARSER (CSV and JSON-array format)
// ======================================================================

function parseKeywords_(rawVal, fallback) {
  if (!rawVal && rawVal !== 0) return fallback;
  if (Array.isArray(rawVal)) return rawVal.map(String).map(s => s.trim()).filter(Boolean);
  let str = String(rawVal).trim();
  if (!str) return fallback;
  if (str.startsWith('[') && str.endsWith(']')) {
    try {
      const parsed = JSON.parse(str.replace(/'/g, '"'));
      if (Array.isArray(parsed) && parsed.length > 0)
        return parsed.map(k => String(k).trim()).filter(Boolean);
    } catch (e) { str = str.slice(1, -1); }
  }
  const list = str.split(',').map(k => k.replace(/['"[\]]/g, '').trim()).filter(Boolean);
  return list.length > 0 ? list : fallback;
}

// ======================================================================
// DEDUPLICATION - scoped per (URL :: Channel_Name)
// ======================================================================

/**
 * Builds a Set of "URL::ChannelName" composite keys from Live_Queue and
 * Recent_Posted. Historical_Archive is excluded so deals older than the
 * dedup window can be re-posted to the same channel.
 */
function buildDedupIndex_() {
  const set = new Set();
  [getLiveQueueSheet_(), getRecentPostedSheet_()].forEach(sheet => {
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow <= HEADER_ROW) return;
    // Columns B(LINK=2) through E(TARGET=5) -> width = 4
    const rows = sheet.getRange(
      HEADER_ROW + 1, COL.LINK,
      lastRow - HEADER_ROW, COL.TARGET - COL.LINK + 1
    ).getValues();
    rows.forEach(r => {
      const link   = String(r[0]).trim();
      const target = String(r[COL.TARGET - COL.LINK]).trim();
      if (link && target) set.add(link + '::' + target);
    });
  });
  return set;
}

function dedupKey_(link, channelName) {
  return String(link).trim() + '::' + String(channelName).trim();
}
function isDuplicate_(link, channelName, dedupSet) {
  return dedupSet.has(dedupKey_(link, channelName));
}

// ======================================================================
// PRICE EXTRACTION HELPERS
// ======================================================================

function extractPriceFromText_(text) {
  const pctMatch = text.match(/(\d{1,3})\s*%\s*(?:off|discount|extra)?/i);
  if (!pctMatch) return null;
  const pct = Number(pctMatch[1]);
  if (isNaN(pct) || pct <= 0 || pct >= 100) return null;
  const priceMatch = text.match(
    /(?:only pay|starting at|starts at|just)\s*(?:rs\.?|inr|Rs)\s*([\d,]+(?:\.\d+)?)/i
  );
  if (!priceMatch) return null;
  const deal = Number(priceMatch[1].replace(/,/g, ''));
  if (isNaN(deal) || deal <= 0) return null;
  const original = deal / (1 - pct / 100);
  if (!isFinite(original) || original <= deal) return null;
  return {
    original: Math.round(original * 100) / 100,
    deal:     Math.round(deal     * 100) / 100
  };
}

function extractPercentOff_(text) {
  const m = text.match(/(\d{1,3})\s*%\s*(?:off|discount|extra)?/i);
  if (!m) return null;
  const pct = Number(m[1]);
  return (isNaN(pct) || pct <= 0 || pct >= 100) ? null : pct;
}

// ======================================================================
// CUELINKS INTAKE
// ======================================================================

const CUELINKS_API_BASE  = 'https://developers.cuelinks.com/pub_api/v3';
const CUELINKS_PAGE_SIZE = 100;
const CUELINKS_MAX_PAGES = 10;

/**
 * Master intake: fetches Cuelinks offers, routes each deal to all matching
 * channels, and appends one Live_Queue row per (deal, channel) pair.
 * Updates LAST_INTAKE_RESULT and manages consecutive-zero-intake alerting.
 */
function runAllIntakes() {
  const __trackStart = Date.now();
  try {
    setConfigValue_('LAST_INTAKE_RUN_TIMESTAMP', new Date());

    const config   = readConfig_();
    const channels = readChannels_();

    if (channels.length === 0) {
      setConfigValue_('LAST_INTAKE_RESULT', 'No enabled channels in Channels tab');
      return;
    }

    const dedupSet   = buildDedupIndex_();
    let   totalAdded = 0;

    try {
      const r = fetchCuelinksOffers_(channels, config, dedupSet);
      totalAdded += r.total;
    } catch (err) { logError_('runAllIntakes', err); }

    let resultMsg;
    if (totalAdded > 0) {
      resultMsg = 'Added ' + totalAdded + ' queue row(s) across ' + channels.length + ' channel(s)';
      setConfigValue_('CONSECUTIVE_ZERO_INTAKE_COUNT', 0);
    } else {
      resultMsg = '0 deals added (all duped, quota-full, or no keyword match)';
      const prev = getConfigNumber_(config, 'CONSECUTIVE_ZERO_INTAKE_COUNT', 0);
      const next  = prev + 1;
      setConfigValue_('CONSECUTIVE_ZERO_INTAKE_COUNT', next);
      if (next === 3) {
        logError_('runAllIntakes',
          new Error('3 consecutive zero-intake runs. Check API key, keywords, or channel quotas.'));
      }
    }
    setConfigValue_('LAST_INTAKE_RESULT', resultMsg);
    Logger.log('runAllIntakes: ' + resultMsg);
  } finally {
    recordExecutionRuntime_('runAllIntakes', __trackStart);
  }
}

function fetchCuelinksOffers_(channels, config, dedupSet) {
  const sheet  = getLiveQueueSheet_();
  const apiKey = PropertiesService.getScriptProperties().getProperty('CUELINKS_API_KEY');
  if (!apiKey) {
    logError_('fetchCuelinksOffers_', new Error('CUELINKS_API_KEY not set in Script Properties'));
    return { total: 0 };
  }

  const isSaleMode   = getConfigBool_(config, 'IS_SALE_MODE_ACTIVE', CONFIG_DEFAULTS.IS_SALE_MODE_ACTIVE);
  const quotaUsed    = {};
  channels.forEach(ch => { quotaUsed[ch.name] = 0; });
  const rowsToAppend = [];

  let page = 1, totalPages = 1;

  pageLoop:
  while (page <= totalPages && page <= CUELINKS_MAX_PAGES) {
    const allFull = channels.every(ch =>
      (quotaUsed[ch.name] || 0) >= (isSaleMode ? ch.maxDeals * 7 : ch.maxDeals)
    );
    if (allFull) break;

    const url = CUELINKS_API_BASE + '/offers?page=' + page +
      '&per_page=' + CUELINKS_PAGE_SIZE + '&sort=created_at&order=desc';
    let json;
    try {
      const resp = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: { 'Authorization': 'Token ' + apiKey },
        muteHttpExceptions: true
      });
      const code = resp.getResponseCode();
      if (code !== 200)
        throw new Error('Cuelinks API HTTP ' + code + ': ' + resp.getContentText().slice(0, 200));
      json = JSON.parse(resp.getContentText());
    } catch (err) { logError_('fetchCuelinksOffers_ page ' + page, err); break; }

    totalPages = (json.meta && json.meta.total_pages) || 1;

    for (const offer of (json.data || [])) {
      const allFullInner = channels.every(ch =>
        (quotaUsed[ch.name] || 0) >= (isSaleMode ? ch.maxDeals * 7 : ch.maxDeals)
      );
      if (allFullInner) break pageLoop;

      try {
        let title         = String(offer.title || '').trim();
        const link        = offer.tracking_url;
        const description = offer.description || '';
        let couponCode    = String(offer.coupon_code || '').trim();
        if (!title || !link) continue;

        // Prevent Google Sheets formula injection (e.g. leading =, +, -, @)
        if (/^[=+\-@]/.test(title)) {
          title = title.replace(/^[=+\-@\s]+/, '');
        }
        if (!title || title.toUpperCase() === '#ERROR!') continue;

        // Ignore placeholder coupons case-insensitively (e.g. DEAL ACTIVATED, NO CODE REQUIRED)
        if (/^(deal\s*activated|no\s*code\s*(required)?|not\s*required|none|na|n\/a)$/i.test(couponCode)) {
          couponCode = '';
        }

        // 3-tier price extraction
        let orig = offer.original_price, deal = offer.discount_price;
        let percentOff = '', isEstimated = false;

        const hasRealPrices =
          orig != null && deal != null &&
          !isNaN(Number(orig)) && !isNaN(Number(deal)) && Number(deal) < Number(orig);

        if (!hasRealPrices) {
          const extracted = extractPriceFromText_(title + ' ' + description);
          if (extracted) {
            orig = extracted.original; deal = extracted.deal; isEstimated = true;
          } else {
            orig = ''; deal = '';
            const sp = offer.percent_off;
            percentOff = (sp != null && !isNaN(Number(sp)))
              ? Number(sp) : extractPercentOff_(title + ' ' + description);
          }
        }

        // Skip if nothing postable
        if (!hasRealPrices && !isEstimated && percentOff == null && !couponCode) continue;

        const haystack     = title + ' ' + description;
        const matchedNames = matchChannels_(haystack, channels);

        for (const chName of matchedNames) {
          const ch    = channels.find(c => c.name === chName);
          const quota = isSaleMode ? ch.maxDeals * 7 : ch.maxDeals;
          if (!ch || (quotaUsed[ch.name] || 0) >= quota) continue;
          if (isDuplicate_(link, ch.name, dedupSet)) continue;

          const displayName = isEstimated ? title.trim() + ' (Est.)' : title.trim();
          rowsToAppend.push([
            displayName, link.trim(),
            orig === '' ? '' : Number(orig),
            deal === '' ? '' : Number(deal),
            ch.name, 'Pending', couponCode,
            percentOff === '' || percentOff == null ? '' : Number(percentOff)
          ]);
          quotaUsed[ch.name] = (quotaUsed[ch.name] || 0) + 1;
          dedupSet.add(dedupKey_(link, ch.name));
        }
      } catch (err) { logError_('fetchCuelinksOffers_ page ' + page + ' item', err); }
    }
    page++;
  }

  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, ROW_WIDTH).setValues(rowsToAppend);
  }
  Logger.log('fetchCuelinksOffers_: appended ' + rowsToAppend.length + ' rows.');
  return { total: rowsToAppend.length };
}

// ======================================================================
// POSTING ENGINE - per-channel pacing with LockService
// ======================================================================

function processQueue_(triggerLabel) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(12000)) {
    Logger.log('processQueue_ (' + triggerLabel + '): lock unavailable, skipping.');
    return;
  }

  try {
    setConfigValue_('LAST_' + triggerLabel + '_RUN_TIMESTAMP', new Date());

    const config     = readConfig_();
    const isSaleMode = getConfigBool_(config, 'IS_SALE_MODE_ACTIVE', CONFIG_DEFAULTS.IS_SALE_MODE_ACTIVE);
    const channels   = readChannels_();
    if (channels.length === 0) return;

    const sheet   = getLiveQueueSheet_(); if (!sheet) return;
    const lastRow = sheet.getLastRow();   if (lastRow <= HEADER_ROW) return;

    // Single batch read of entire queue
    const allRows = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, ROW_WIDTH).getValues();

    // Group pending rows by target channel (preserve original sheet row numbers)
    const pendingByChannel = {};
    allRows.forEach((row, idx) => {
      if (String(row[COL.STATUS - 1]) === 'Pending') {
        const chName = String(row[COL.TARGET - 1]).trim();
        if (!pendingByChannel[chName]) pendingByChannel[chName] = [];
        pendingByChannel[chName].push({ rowNum: HEADER_ROW + 1 + idx, rowValues: row });
      }
    });

    const botToken = PropertiesService.getScriptProperties().getProperty('BOT_TOKEN');
    if (!botToken) { logError_('processQueue_', new Error('BOT_TOKEN not set')); return; }

    let totalPosted = 0;

    for (const ch of channels) {
      const pending = pendingByChannel[ch.name];
      if (!pending || pending.length === 0) continue;

      if (!isSaleMode) {
        // Normal mode: one post per channel per sweep, with per-channel cooldown
        const lastPost = getConfigDate_(config, 'LAST_POST_' + ch.name);
        if (lastPost && Date.now() - lastPost.getTime() < ch.minInterval * 60 * 1000) continue;

        const first = pending[0];
        const ok = postQueueRow_(sheet, first.rowNum, first.rowValues, ch, botToken);
        if (ok) { setConfigValue_('LAST_POST_' + ch.name, new Date()); totalPosted++; }

      } else {
        // Sale mode burst: post up to SALE_MODE_BURST_CAP rows per channel
        // Reverse so row deletions don't shift the indices of later rows
        const toProcess = pending.slice(0, SALE_MODE_BURST_CAP).reverse();
        let anyPosted = false;
        for (const item of toProcess) {
          if (String(item.rowValues[COL.STATUS - 1]) !== 'Pending') continue;
          const ok = postQueueRow_(sheet, item.rowNum, item.rowValues, ch, botToken);
          if (ok) { anyPosted = true; totalPosted++; }
          Utilities.sleep(1500); // Telegram rate-limit gap between posts
        }
        if (anyPosted) setConfigValue_('LAST_POST_' + ch.name, new Date());
      }
    }

    Logger.log('processQueue_ (' + triggerLabel + '): posted ' + totalPosted + ' deal(s).');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Sends ONE Live_Queue row to Telegram.
 * Success: appends to Recent_Posted, deletes from Live_Queue. Returns true.
 * Failure: marks Status cell as "Error: <reason>". Returns false.
 */
function postQueueRow_(sheet, rowNum, rowValues, ch, botToken) {
  const name       = rowValues[COL.NAME    - 1];
  const link       = rowValues[COL.LINK    - 1];
  const origPrice  = rowValues[COL.ORIG    - 1];
  const dealPrice  = rowValues[COL.DEAL    - 1];
  const couponCode = rowValues[COL.COUPON  - 1];
  const pctOffRaw  = rowValues[COL.PCT_OFF - 1];

  try {
    if (!name || String(name).trim() === '#ERROR!') throw new Error('Missing or invalid Product_Name');
    if (!link)      throw new Error('Missing Raw_URL');
    if (!ch.chatId) throw new Error('No Chat_ID for channel "' + ch.name + '"');

    const cleanCoupon = (couponCode && !/^(deal\s*activated|no\s*code\s*(required)?|not\s*required|none|na|n\/a)$/i.test(String(couponCode).trim()))
      ? String(couponCode).trim()
      : null;

    const hasOrig   = origPrice  !== '' && origPrice  != null;
    const hasDeal   = dealPrice  !== '' && dealPrice  != null;
    const hasCoupon = cleanCoupon != null;
    const hasPct    = pctOffRaw  !== '' && pctOffRaw  != null && !isNaN(Number(pctOffRaw));

    let message;

    if (hasOrig || hasDeal) {
      if (!hasOrig || !hasDeal)
        throw new Error('Original_Price and Deal_Price must both be set or both blank');
      const orig = Number(origPrice), deal = Number(dealPrice);
      if (isNaN(orig) || orig <= 0) throw new Error('Invalid Original_Price');
      if (isNaN(deal) || deal <= 0) throw new Error('Invalid Deal_Price');
      if (deal >= orig)             throw new Error('Deal_Price must be less than Original_Price');
      const pct = Math.round(((orig - deal) / orig) * 100);
      message = buildDealMessage_(ch, name, orig, deal, pct, link, hasCoupon ? cleanCoupon : null);

    } else if (hasPct || hasCoupon) {
      const pct = hasPct ? Math.round(Number(pctOffRaw)) : null;
      message = buildCouponMessage_(ch, name, hasCoupon ? cleanCoupon : null, pct, link);

    } else {
      throw new Error('No price, percent-off, or coupon code - nothing to post');
    }

    const result = sendToTelegram_(botToken, ch.chatId, message, link);
    if (!result.ok) throw new Error(result.error);

    prependToRecentPosted_([
      name, link,
      hasOrig  ? Number(origPrice) : '',
      hasDeal  ? Number(dealPrice) : '',
      ch.name, new Date(),
      hasCoupon ? couponCode        : '',
      hasPct    ? Number(pctOffRaw) : ''
    ]);
    sheet.deleteRow(rowNum);
    return true;

  } catch (err) {
    sheet.getRange(rowNum, COL.STATUS).setValue('Error: ' + err.message);
    logError_('postQueueRow_ "' + ch.name + '" row ' + rowNum, err);
    return false;
  }
}

function prependToRecentPosted_(rowValues) {
  const sheet = getRecentPostedSheet_();
  sheet.insertRowAfter(HEADER_ROW);
  sheet.getRange(HEADER_ROW + 1, 1, 1, rowValues.length).setValues([rowValues]);
}

// ======================================================================
// MESSAGE BUILDERS
// ======================================================================

function buildDealMessage_(ch, name, orig, deal, savingsPct, link, couponCode) {
  const safeName = escapeMarkdownV2_(name);
  const origStr  = escapeMarkdownV2_('Rs.' + orig.toLocaleString('en-IN'));
  const dealStr  = escapeMarkdownV2_('Rs.' + deal.toLocaleString('en-IN'));
  const pctStr   = escapeMarkdownV2_(savingsPct + '% OFF');
  const safeLink = String(link).replace(/([)\\])/g, '\\$1');

  const lines = [
    ch.emoji + ' *' + escapeMarkdownV2_(ch.name) + '* ' + ch.emoji,
    '',
    '*' + safeName + '*',
    '',
    '~' + origStr + '~  \\-\\>  *' + dealStr + '*',
    '*' + pctStr + '*'
  ];
  if (couponCode) lines.push('Code: `' + escapeMarkdownV2Code_(couponCode) + '`');
  lines.push('', '[Buy Now](' + safeLink + ')');
  return lines.join('\n');
}

function buildCouponMessage_(ch, name, couponCode, percentOff, link) {
  const safeName = escapeMarkdownV2_(name);
  const safeLink = String(link).replace(/([)\\])/g, '\\$1');

  const lines = [
    ch.emoji + ' *' + escapeMarkdownV2_(ch.name) + '* ' + ch.emoji,
    '',
    '*' + safeName + '*'
  ];
  if (percentOff) lines.push('', '*' + escapeMarkdownV2_(percentOff + '% OFF') + '*');
  if (couponCode) lines.push('Code: `' + escapeMarkdownV2Code_(couponCode) + '`');
  lines.push('', '[Grab This Deal](' + safeLink + ')');
  return lines.join('\n');
}

function escapeMarkdownV2_(text) {
  return String(text).replace(/[_*[\]()~`>#+\-=|{}.!]/g, '\\$&');
}
function escapeMarkdownV2Code_(text) {
  return String(text).replace(/([`\\])/g, '\\$1');
}

// ======================================================================
// TELEGRAM - HTTP send with retry and rate-limit back-off
// ======================================================================

function sendToTelegram_(botToken, chatId, text, link) {
  const url = 'https://api.telegram.org/bot' + botToken + '/sendMessage';
  const payload = {
    chat_id: chatId,
    text: text,
    parse_mode: 'MarkdownV2',
    disable_web_page_preview: false,
    reply_markup: JSON.stringify({
      inline_keyboard: [[{ text: 'Buy Now', url: link }]]
    })
  };
  const options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: payload,
    muteHttpExceptions: true
  };
  let lastError = '';
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      const resp = UrlFetchApp.fetch(url, options);
      const code = resp.getResponseCode();
      const body = JSON.parse(resp.getContentText());
      if (code === 200 && body.ok) return { ok: true };
      lastError = 'HTTP ' + code + ': ' + (body.description || 'unknown');
      if (code === 429 && body.parameters && body.parameters.retry_after) {
        Utilities.sleep(body.parameters.retry_after * 1000); continue;
      }
      if (code >= 500) { Utilities.sleep(attempt * 2000); continue; }
      break;
    } catch (err) { lastError = err.message; Utilities.sleep(attempt * 1500); }
  }
  return { ok: false, error: lastError };
}

// ======================================================================
// ARCHIVING - move expired Recent_Posted rows to Historical_Archive
// ======================================================================

function archiveRecentDeals() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) { Logger.log('archiveRecentDeals: lock unavailable.'); return; }
  try {
    setConfigValue_('LAST_ARCHIVE_RUN_TIMESTAMP', new Date());
    const config      = readConfig_();
    const windowHours = getConfigNumber_(config, 'DEDUP_WINDOW_HOURS', CONFIG_DEFAULTS.DEDUP_WINDOW_HOURS);
    const cutoff      = Date.now() - windowHours * 3600000;

    const recentSheet = getRecentPostedSheet_();
    const lastRow     = recentSheet.getLastRow();
    if (lastRow <= HEADER_ROW) return;

    const values    = recentSheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, ROW_WIDTH).getValues();
    const toArchive = [], stillHere = [];

    values.forEach(row => {
      const sentAt = row[5]; // column F = Sent_At (zero-indexed: 5)
      const d = (sentAt instanceof Date) ? sentAt : new Date(sentAt);
      (!isNaN(d.getTime()) && d.getTime() < cutoff ? toArchive : stillHere).push(row);
    });

    if (toArchive.length === 0) {
      Logger.log('archiveRecentDeals: nothing to archive (window=' + windowHours + 'h).');
      return;
    }

    const archSheet = getArchiveSheet_();
    archSheet.getRange(archSheet.getLastRow() + 1, 1, toArchive.length, ROW_WIDTH).setValues(toArchive);

    recentSheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, ROW_WIDTH).clearContent();
    if (stillHere.length > 0)
      recentSheet.getRange(HEADER_ROW + 1, 1, stillHere.length, ROW_WIDTH).setValues(stillHere);

    Logger.log('archiveRecentDeals: moved ' + toArchive.length + ', kept ' + stillHere.length + '.');
  } catch (err) { logError_('archiveRecentDeals', err); }
  finally { lock.releaseLock(); }
}

// ======================================================================
// DAILY MAINTENANCE - trim Logs and purge old archive rows
// ======================================================================

function runDailyMaintenance() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) { Logger.log('runDailyMaintenance: lock unavailable.'); return; }
  try {
    setConfigValue_('LAST_MAINTENANCE_RUN_TIMESTAMP', new Date());
    const config = readConfig_();
    cleanLogs_(config);
    cleanHistoricalArchive_(config);
    Logger.log('runDailyMaintenance: complete.');
  } catch (err) { logError_('runDailyMaintenance', err); }
  finally { lock.releaseLock(); }
}

function cleanLogs_(config) {
  const logSheet = ss_().getSheetByName(LOGS_SHEET);
  if (!logSheet) return;
  const lastRow = logSheet.getLastRow();
  if (lastRow <= HEADER_ROW) return;

  const retHours = getConfigNumber_(config, 'LOG_RETENTION_HOURS', CONFIG_DEFAULTS.LOG_RETENTION_HOURS);
  const maxRows  = getConfigNumber_(config, 'LOG_MAX_ROWS',         CONFIG_DEFAULTS.LOG_MAX_ROWS);
  const cutoff   = Date.now() - retHours * 3600000;

  let rows = logSheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, 3).getValues();
  rows = rows.filter(r => {
    const d = (r[0] instanceof Date) ? r[0] : new Date(r[0]);
    return !isNaN(d.getTime()) && d.getTime() >= cutoff;
  });
  if (rows.length > maxRows) rows = rows.slice(rows.length - maxRows);

  logSheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, 3).clearContent();
  if (rows.length > 0) logSheet.getRange(HEADER_ROW + 1, 1, rows.length, 3).setValues(rows);
}

function cleanHistoricalArchive_(config) {
  const sheet = getArchiveSheet_(); if (!sheet) return;
  const lastRow = sheet.getLastRow(); if (lastRow <= HEADER_ROW) return;

  const retDays = getConfigNumber_(config, 'ARCHIVE_RETENTION_DAYS', CONFIG_DEFAULTS.ARCHIVE_RETENTION_DAYS);
  const cutoff  = Date.now() - retDays * 86400000;

  const rows = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, ROW_WIDTH).getValues();
  const keep = rows.filter(r => {
    const d = (r[5] instanceof Date) ? r[5] : new Date(r[5]);
    return isNaN(d.getTime()) || d.getTime() >= cutoff;
  });

  if (keep.length !== rows.length) {
    sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, ROW_WIDTH).clearContent();
    if (keep.length > 0)
      sheet.getRange(HEADER_ROW + 1, 1, keep.length, ROW_WIDTH).setValues(keep);
    Logger.log('cleanHistoricalArchive_: purged ' + (rows.length - keep.length) +
               ' rows older than ' + retDays + ' days.');
  }
}

// ======================================================================
// TRIGGERS AND PUBLIC ENTRY POINTS
// ======================================================================

function onEditInstallable(e) {
  const __trackStart = Date.now();
  try {
    if (e.range.getSheet().getName() !== LIVE_QUEUE_SHEET) return;
    if (e.range.getRow() <= HEADER_ROW) return;
    processQueue_('ONEDIT');
  } catch (err) { logError_('onEditInstallable', err); }
  finally {
    recordExecutionRuntime_('onEditInstallable', __trackStart);
  }
}

function sweepPendingRows() {
  const __trackStart = Date.now();
  try { processQueue_('SWEEP'); }
  catch (err) { logError_('sweepPendingRows', err); }
  finally {
    recordExecutionRuntime_('sweepPendingRows', __trackStart);
  }
}

/** Run ONCE from the Apps Script editor to install the onEdit installable trigger. */
function createTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onEditInstallable') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditInstallable').forSpreadsheet(ss).onEdit().create();
  Logger.log('===== SETUP COMPLETE =====');
  Logger.log('onEdit trigger installed. Now add 4 time-driven triggers manually:');
  Logger.log('  1. sweepPendingRows    -> Minutes timer, every 1-5 min');
  Logger.log('  2. runAllIntakes       -> Minutes/Hour timer, every 30-120 min');
  Logger.log('  3. archiveRecentDeals  -> Hour timer, every 4 hours');
  Logger.log('  4. runDailyMaintenance -> Day timer, 4:00 AM - 5:00 AM');
}

// ======================================================================
// ERROR LOGGING
// ======================================================================

function logError_(context, err) {
  Logger.log('[ERROR][' + context + '] ' + err.message);
  try {
    let logSheet = ss_().getSheetByName(LOGS_SHEET);
    if (!logSheet) {
      logSheet = ss_().insertSheet(LOGS_SHEET);
      logSheet.appendRow(['Timestamp', 'Context', 'Error']);
    }
    logSheet.appendRow([new Date(), context, err.message]);
  } catch (e) { /* never let logging break the main flow */ }
}
