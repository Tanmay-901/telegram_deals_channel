/**
 * Sheets -> Telegram MULTI-CHANNEL pipeline
 * Light-Queue / Fast-Dedup / Configuration-Dashboard / Historical-Archive architecture
 * -------------------------------------------------------------------------
 * REQUIRED TABS:
 *
 *   1. Live_Queue        — active, lightweight. Columns A-H:
 *        A: Product_Name | B: Raw_URL | C: Original_Price | D: Deal_Price
 *        E: Category (Gym / Pets / Vehicle) | F: Status (Pending / Error: ...)
 *        G: Coupon_Code (optional) | H: Percent_Off (optional)
 *      C/D (prices) OR H (percent) should be populated — NOT both required.
 *      NOTE: successfully-posted rowas are DELETED from here and moved to
 *      Recent_Posted (active dedup window) — Live_Queue never accumulates "Sent" rows.
 *
 *   2. Recent_Posted     — dedicated 4-hour deduplication sheet. Columns A-H:
 *        A: Product_Name | B: Raw_URL | C: Original_Price | D: Deal_Price
 *        E: Category | F: Sent_At | G: Coupon_Code | H: Percent_Off
 *      Every posted deal lands here first. The dedup index ONLY checks Live_Queue +
 *      Recent_Posted, keeping dedup lookups instant (<50ms). Rows older than
 *      DEDUP_WINDOW_HOURS (default 4h) are moved to Historical_Archive by archiveRecentDeals.
 *      (Auto-created by script if not present).
 *
 *   3. Configuration     — key/value dashboard. Column A = Key, B = Value.
 *      Row 1: headers (Key, Value)
 *      Configurable Keys:
 *        POSTING_INTERVAL_MINUTES      e.g. 5
 *        MAX_DEALS_PER_INTAKE_RUN      e.g. 2
 *        IS_SALE_MODE_ACTIVE           TRUE or FALSE
 *        GYM_KEYWORDS                  e.g. ['whey', 'protein', 'gym'] or whey, protein, gym
 *        PETS_KEYWORDS                 e.g. ['dog', 'cat', 'pet'] or dog, cat, pet
 *        VEHICLE_KEYWORDS              e.g. ['car', 'bike', 'dashcam'] or car, bike, dashcam
 *        DEDUP_WINDOW_HOURS            e.g. 4
 *        LOG_RETENTION_HOURS           e.g. 24
 *        LOG_MAX_ROWS                  e.g. 1000
 *        ARCHIVE_RETENTION_DAYS        e.g. 30
 *      Auto-updated Diagnostic Keys:
 *        LAST_INTAKE_RESULT            (e.g. "Added 3 deal(s) (Gym: 2, Pets: 1, Vehicle: 0)")
 *        CONSECUTIVE_ZERO_INTAKE_COUNT (increments on 0-deal runs; alerts at 3; resets on deal)
 *        LAST_POST_TIMESTAMP
 *        LAST_INTAKE_RUN_TIMESTAMP
 *        LAST_SWEEP_RUN_TIMESTAMP
 *        LAST_ARCHIVE_RUN_TIMESTAMP
 *        LAST_MAINTENANCE_RUN_TIMESTAMP
 *
 *   4. Historical_Archive — permanent log of posted deals retained for ARCHIVE_RETENTION_DAYS.
 *      Columns A-H: Product_Name | Raw_URL | Original_Price | Deal_Price
 *      | Category | Sent_At | Coupon_Code | Percent_Off
 *
 *   5. Logs              — records errors and 3-consecutive-zero intake warnings.
 *      Columns A-C: Timestamp | Context | Error
 *      Auto-trimmed by runDailyMaintenance (>24 hours old or >1,000 rows).
 *
 * SETUP & TRIGGERS:
 * 1. Create tabs with headers (or let the script auto-create Recent_Posted and Logs).
 * 2. Extensions > Apps Script, paste this file in.
 * 3. Project Settings > Script Properties, add:
 *      BOT_TOKEN         = <your bot token from @BotFather>
 *      CHANNEL_GYM       = @YourGymChannel   (or numeric -100... id)
 *      CHANNEL_PETS      = @YourPetChannel
 *      CHANNEL_VEHICLE   = @YourVehicleChannel
 *      CUELINKS_API_KEY  = <your key from https://www.cuelinks.com/api-key>
 * 4. Run `createTrigger` once from the editor (installs the onEdit trigger).
 * 5. Set up 4 TIME-DRIVEN TRIGGERS manually (Triggers icon > Add Trigger):
 *      - sweepPendingRows    -> Every 1 to 5 minutes
 *      - runAllIntakes       -> Every 30 to 120 minutes
 *      - archiveRecentDeals  -> Every 4 hours
 *      - runDailyMaintenance -> Daily (between 4:00 AM - 5:00 AM)
 */

// =====================================================================
// SHEET NAMES & LAYOUT
// =====================================================================

const LIVE_QUEUE_SHEET = 'Live_Queue';
const RECENT_POSTED_SHEET = 'Recent_Posted';
const CONFIG_SHEET = 'Configuration';
const ARCHIVE_SHEET = 'Historical_Archive';
const LOGS_SHEET = 'Logs';
const HEADER_ROW = 1;

const COL = { NAME: 1, LINK: 2, ORIG: 3, DEAL: 4, CATEGORY: 5, STATUS: 6, COUPON_CODE: 7, PERCENT_OFF: 8 };
const ARCHIVE_COL = { NAME: 1, LINK: 2, ORIG: 3, DEAL: 4, CATEGORY: 5, SENT_AT: 6, COUPON_CODE: 7, PERCENT_OFF: 8 };
const ROW_WIDTH = 8; // A-H, kept as one constant so every read/write range stays in sync

function getLiveQueueSheet_() {
  return SpreadsheetApp.getActive().getSheetByName(LIVE_QUEUE_SHEET);
}
function getConfigSheet_() {
  return SpreadsheetApp.getActive().getSheetByName(CONFIG_SHEET);
}
function getArchiveSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(ARCHIVE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(ARCHIVE_SHEET);
    sheet.appendRow(['Product_Name', 'Raw_URL', 'Original_Price', 'Deal_Price', 'Category', 'Sent_At', 'Coupon_Code', 'Percent_Off']);
  }
  return sheet;
}
function getRecentPostedSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(RECENT_POSTED_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(RECENT_POSTED_SHEET);
    sheet.appendRow(['Product_Name', 'Raw_URL', 'Original_Price', 'Deal_Price', 'Category', 'Sent_At', 'Coupon_Code', 'Percent_Off']);
  }
  return sheet;
}

// =====================================================================
// CONFIGURATION DASHBOARD — read/write helpers
// =====================================================================

// Fallbacks used only if a key is missing/blank in the Configuration tab —
// the tab itself is always the source of truth when a value is present.
const CONFIG_DEFAULTS = {
  POSTING_INTERVAL_MINUTES: 5,
  MAX_DEALS_PER_INTAKE_RUN: 2,
  IS_SALE_MODE_ACTIVE: false,
  DEDUP_WINDOW_HOURS: 4,
  LOG_RETENTION_HOURS: 24,
  LOG_MAX_ROWS: 1000,
  ARCHIVE_RETENTION_DAYS: 30
};

// Sale mode widens intake volume. Multiplies base quota unless overridden.
const SALE_MODE_INTAKE_MULTIPLIER = 7;

/** Reads the whole Configuration tab into a plain {key: value} object. */
function readConfig_() {
  const sheet = getConfigSheet_();
  if (!sheet) return {};
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow <= HEADER_ROW) return map;
  const values = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, 2).getValues();
  values.forEach(row => {
    const key = String(row[0]).trim();
    if (key) map[key] = row[1];
  });
  return map;
}

function getConfigNumber_(config, key, fallback) {
  const v = config[key];
  const n = Number(v);
  return (v !== undefined && v !== '' && !isNaN(n)) ? n : fallback;
}
function getConfigBool_(config, key, fallback) {
  const v = config[key];
  if (v === undefined || v === '') return fallback;
  if (typeof v === 'boolean') return v;
  return String(v).trim().toUpperCase() === 'TRUE';
}
function getConfigDate_(config, key) {
  const v = config[key];
  if (!v) return null;
  const d = (v instanceof Date) ? v : new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Writes a single key's value. Updates the existing row if the key is
 * already present; otherwise appends a new row automatically.
 */
function setConfigValue_(key, value) {
  const sheet = getConfigSheet_();
  if (!sheet) return;
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

/**
 * Per-category intake cap for this run, sourced from Configuration.
 */
function getIntakeQuota_(config) {
  const isSaleMode = getConfigBool_(config, 'IS_SALE_MODE_ACTIVE', CONFIG_DEFAULTS.IS_SALE_MODE_ACTIVE);
  const baseQuota = getConfigNumber_(config, 'MAX_DEALS_PER_INTAKE_RUN', CONFIG_DEFAULTS.MAX_DEALS_PER_INTAKE_RUN);
  if (!isSaleMode) return baseQuota;

  const saleOverride = Number(config['MAX_DEALS_PER_INTAKE_RUN_SALE']);
  if (config['MAX_DEALS_PER_INTAKE_RUN_SALE'] !== undefined && !isNaN(saleOverride)) {
    return saleOverride;
  }
  return baseQuota * SALE_MODE_INTAKE_MULTIPLIER;
}

// =====================================================================
// DEDUPLICATION INDEX — Live_Queue + Recent_Posted (fast O(1) lookups)
// =====================================================================

/**
 * Builds a single Set of every Raw_URL currently in Live_Queue OR Recent_Posted.
 * Deals older than DEDUP_WINDOW_HOURS have moved to Historical_Archive and are
 * deliberately excluded, allowing products to be posted again after the window expires.
 */
function buildDedupIndex_() {
  const set = new Set();
  [getLiveQueueSheet_(), getRecentPostedSheet_()].forEach(sheet => {
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow <= HEADER_ROW) return;
    const links = sheet.getRange(HEADER_ROW + 1, COL.LINK, lastRow - HEADER_ROW, 1).getValues();
    links.forEach(row => {
      const val = String(row[0]).trim();
      if (val) set.add(val);
    });
  });
  return set;
}

function isDuplicateLink_(link, dedupSet) {
  return dedupSet.has(String(link).trim());
}

// =====================================================================
// CATEGORY ROUTING & KEYWORDS
// =====================================================================

const CATEGORY_MAP = {
  'Gym':     { propKey: 'CHANNEL_GYM',     emoji: '💪', label: 'Gym Deal' },
  'Pets':    { propKey: 'CHANNEL_PETS',    emoji: '🐾', label: 'Pet Deal' },
  'Vehicle': { propKey: 'CHANNEL_VEHICLE', emoji: '🚗', label: 'Vehicle Deal' }
};

const CATEGORY_KEYWORDS = {
  'Gym':     ['whey', 'protein', 'creatine', 'dumbbell', 'shaker', 'bcaa', 'gym', 'workout', 'fitness'],
  'Pets':    ['dog', 'cat', 'pedigree', 'kibble', 'leash', 'pet', 'puppy', 'aquarium', 'litter'],
  'Vehicle': ['car', 'bike', 'dashcam', 'tyre', 'helmet', 'lubricant', 'seat cover', 'automotive']
};

function buildCategoryRegexes_() {
  const map = {};
  for (const category in CATEGORY_KEYWORDS) {
    map[category] = CATEGORY_KEYWORDS[category].map(
      kw => new RegExp('\\b' + kw.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + 's?\\b', 'i')
    );
  }
  return map;
}
const CATEGORY_KEYWORD_REGEXES = buildCategoryRegexes_();

function classifyCategory_(text) {
  for (const category in CATEGORY_KEYWORD_REGEXES) {
    for (const re of CATEGORY_KEYWORD_REGEXES[category]) {
      if (re.test(text)) return category;
    }
  }
  return null;
}

function extractPriceFromText_(text) {
  const percentMatch = text.match(/(\d{1,3})\s*%\s*(?:off|discount|extra)?/i);
  if (!percentMatch) return null;

  const pct = Number(percentMatch[1]);
  if (isNaN(pct) || pct <= 0 || pct >= 100) return null;

  const priceMatch = text.match(
    /(?:only pay|starting at|starts at|prices? (?:from|start(?:ing)? (?:at|from))|just)\s*(?:rs\.?|inr|₹)\s*([\d,]+(?:\.\d+)?)/i
  );
  if (!priceMatch) return null;

  const deal = Number(priceMatch[1].replace(/,/g, ''));
  if (isNaN(deal) || deal <= 0) return null;

  const original = deal / (1 - pct / 100);
  if (!isFinite(original) || original <= deal) return null;

  return {
    original: Math.round(original * 100) / 100,
    deal: Math.round(deal * 100) / 100
  };
}

function extractPercentOff_(text) {
  const m = text.match(/(\d{1,3})\s*%\s*(?:off|discount|extra)?/i);
  if (!m) return null;
  const pct = Number(m[1]);
  if (isNaN(pct) || pct <= 0 || pct >= 100) return null;
  return pct;
}

// =====================================================================
// XML/RSS GENERIC INTAKE ENGINE
// =====================================================================

const DEAL_SOURCES = [
  {
    name: 'EarnKaro',
    enabled: false,
    feedUrl: 'PASTE_EARNKARO_FEED_URL_HERE_IF_CONFIRMED',
    tags: { itemNode: 'item', titleTag: 'title', linkTag: 'link', descriptionTag: 'description' }
  },
  {
    name: 'INRDeals',
    enabled: false,
    feedUrl: 'PASTE_INRDEALS_FEED_URL_HERE_IF_CONFIRMED',
    tags: { itemNode: 'item', titleTag: 'title', linkTag: 'link', descriptionTag: 'description' }
  }
];

function fetchDealsFromSource_(source, config, dedupSet) {
  const sheet = getLiveQueueSheet_();
  dedupSet = dedupSet || buildDedupIndex_();

  if (!source.feedUrl || source.feedUrl.indexOf('PASTE_') === 0) {
    logError_(`fetchDealsFromSource_ (${source.name})`, new Error('feedUrl not configured'));
    return { total: 0, byCategory: { 'Gym': 0, 'Pets': 0, 'Vehicle': 0 } };
  }

  let xmlText;
  try {
    const response = UrlFetchApp.fetch(source.feedUrl, { muteHttpExceptions: true });
    const code = response.getResponseCode();
    if (code !== 200) throw new Error(`Feed fetch failed with HTTP ${code}`);
    xmlText = response.getContentText();
  } catch (err) {
    logError_(`fetchDealsFromSource_ (${source.name}) fetch`, err);
    return { total: 0, byCategory: { 'Gym': 0, 'Pets': 0, 'Vehicle': 0 } };
  }

  let items;
  try {
    const document = XmlService.parse(xmlText);
    items = findAllDescendants_(document.getRootElement(), source.tags.itemNode);
  } catch (err) {
    logError_(`fetchDealsFromSource_ (${source.name}) parse`, err);
    return { total: 0, byCategory: { 'Gym': 0, 'Pets': 0, 'Vehicle': 0 } };
  }

  if (!items || items.length === 0) {
    logError_(`fetchDealsFromSource_ (${source.name})`, new Error('Feed parsed but returned 0 items'));
    return { total: 0, byCategory: { 'Gym': 0, 'Pets': 0, 'Vehicle': 0 } };
  }

  const quotaLimit = getIntakeQuota_(config);
  const quotaUsed = { 'Gym': 0, 'Pets': 0, 'Vehicle': 0 };
  const categoryRegexes = buildCategoryRegexes_(config);
  const rowsToAppend = [];

  for (const item of items) {
    const allFull = Object.keys(quotaUsed).every(cat => quotaUsed[cat] >= quotaLimit);
    if (allFull) break;

    try {
      const title = getChildText_(item, source.tags.titleTag);
      const link = getChildText_(item, source.tags.linkTag);
      const description = getChildText_(item, source.tags.descriptionTag);

      if (!title || !link) continue;
      if (isDuplicateLink_(link, dedupSet)) continue;

      const haystack = `${title} ${description || ''}`.toLowerCase();
      const category = classifyCategory_(haystack, categoryRegexes);
      if (!category) continue;
      if (quotaUsed[category] >= quotaLimit) continue;

      rowsToAppend.push([title.trim(), link.trim(), '', '', category, 'Pending', '', '']);
      quotaUsed[category]++;
      dedupSet.add(link.trim());

    } catch (err) {
      logError_(`fetchDealsFromSource_ (${source.name}) item`, err);
    }
  }

  if (rowsToAppend.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, ROW_WIDTH).setValues(rowsToAppend);
  }

  Logger.log(`fetchDealsFromSource_ (${source.name}): appended ${rowsToAppend.length} rows.`);
  return { total: rowsToAppend.length, byCategory: quotaUsed };
}

function findAllDescendants_(element, tagName) {
  let results = [];
  const children = element.getChildren();
  for (const child of children) {
    if (child.getName() === tagName) results.push(child);
    results = results.concat(findAllDescendants_(child, tagName));
  }
  return results;
}

function getChildText_(itemElement, tagName) {
  const direct = itemElement.getChild(tagName);
  if (direct) return direct.getText();
  const nested = findAllDescendants_(itemElement, tagName);
  return nested.length > 0 ? nested[0].getText() : '';
}

// =====================================================================
// CUELINKS — real JSON REST intake
// =====================================================================

const CUELINKS_API_BASE = 'https://developers.cuelinks.com/pub_api/v3';
const CUELINKS_PAGE_SIZE = 100;
const CUELINKS_MAX_PAGES = 10;

/**
 * Runs Cuelinks + every enabled XML source.
 * Tracks total deals added, updates LAST_INTAKE_RESULT in Configuration,
 * and logs an alert if 3 consecutive runs return 0 deals.
 */
function runAllIntakes() {
  setConfigValue_('LAST_INTAKE_RUN_TIMESTAMP', new Date());
  const config = readConfig_();
  const dedupSet = buildDedupIndex_();

  let totalAdded = 0;
  const categoryAdded = { 'Gym': 0, 'Pets': 0, 'Vehicle': 0 };

  try {
    const cuelinksStats = fetchCuelinksOffers(config, dedupSet);
    if (cuelinksStats) {
      totalAdded += cuelinksStats.total;
      for (const cat in cuelinksStats.byCategory) {
        categoryAdded[cat] = (categoryAdded[cat] || 0) + cuelinksStats.byCategory[cat];
      }
    }
  } catch (err) {
    logError_('runAllIntakes (Cuelinks)', err);
  }

  DEAL_SOURCES.forEach(source => {
    if (!source.enabled) return;
    try {
      const sourceStats = fetchDealsFromSource_(source, config, dedupSet);
      if (sourceStats) {
        totalAdded += sourceStats.total;
        for (const cat in sourceStats.byCategory) {
          categoryAdded[cat] = (categoryAdded[cat] || 0) + sourceStats.byCategory[cat];
        }
      }
    } catch (err) {
      logError_(`runAllIntakes (${source.name})`, err);
    }
  });

  // Update diagnostic result in Configuration
  let resultMsg = '';
  if (totalAdded > 0) {
    const breakdown = Object.keys(categoryAdded).map(c => `${c}: ${categoryAdded[c]}`).join(', ');
    resultMsg = `Added ${totalAdded} deal(s) (${breakdown})`;
    setConfigValue_('CONSECUTIVE_ZERO_INTAKE_COUNT', 0);
  } else {
    resultMsg = '0 deals added (quotas full or no matching keyword/unseen deals)';
    const prevZeroCount = getConfigNumber_(config, 'CONSECUTIVE_ZERO_INTAKE_COUNT', 0);
    const newZeroCount = prevZeroCount + 1;
    setConfigValue_('CONSECUTIVE_ZERO_INTAKE_COUNT', newZeroCount);
    if (newZeroCount === 3) {
      logError_('runAllIntakes', new Error('3 consecutive intake runs returned 0 deals. Check API quota, token, or category filters.'));
    }
  }
  setConfigValue_('LAST_INTAKE_RESULT', resultMsg);
  Logger.log(`runAllIntakes: ${resultMsg}`);
}

function fetchCuelinksOffers(config, dedupSet) {
  const sheet = getLiveQueueSheet_();
  dedupSet = dedupSet || buildDedupIndex_();

  const apiKey = PropertiesService.getScriptProperties().getProperty('CUELINKS_API_KEY');
  if (!apiKey) {
    logError_('fetchCuelinksOffers', new Error('CUELINKS_API_KEY not set in Script Properties'));
    return { total: 0, byCategory: { 'Gym': 0, 'Pets': 0, 'Vehicle': 0 } };
  }

  const quotaLimit = getIntakeQuota_(config);
  const quotaUsed = { 'Gym': 0, 'Pets': 0, 'Vehicle': 0 };
  const categoryRegexes = buildCategoryRegexes_(config);
  const rowsToAppend = [];
  let page = 1;
  let totalPages = 1;

  pageLoop:
  while (page <= totalPages && page <= CUELINKS_MAX_PAGES) {
    const allFull = Object.keys(quotaUsed).every(cat => quotaUsed[cat] >= quotaLimit);
    if (allFull) break;

    const url = `${CUELINKS_API_BASE}/offers?page=${page}&per_page=${CUELINKS_PAGE_SIZE}&sort=created_at&order=desc`;

    let json;
    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: { 'Authorization': 'Token ' + apiKey },
        muteHttpExceptions: true
      });
      const code = response.getResponseCode();
      if (code !== 200) throw new Error(`Cuelinks API returned HTTP ${code}: ${response.getContentText().slice(0, 200)}`);
      json = JSON.parse(response.getContentText());
    } catch (err) {
      logError_(`fetchCuelinksOffers page ${page}`, err);
      break;
    }

    totalPages = (json.meta && json.meta.total_pages) || 1;
    const offers = json.data || [];

    for (const offer of offers) {
      const allFullInner = Object.keys(quotaUsed).every(cat => quotaUsed[cat] >= quotaLimit);
      if (allFullInner) break pageLoop;

      try {
        const title = offer.title;
        const link = offer.tracking_url;
        const description = offer.description || '';
        const couponCode = offer.coupon_code || '';

        if (!title || !link) continue;

        let orig = offer.original_price;
        let deal = offer.discount_price;
        let percentOff = '';
        let isEstimated = false;

        const hasRealPrices = orig != null && deal != null
          && !isNaN(Number(orig)) && !isNaN(Number(deal))
          && Number(deal) < Number(orig);

        if (!hasRealPrices) {
          const extracted = extractPriceFromText_(`${title} ${description}`);
          if (extracted) {
            orig = extracted.original;
            deal = extracted.deal;
            isEstimated = true;
          } else {
            orig = '';
            deal = '';
            const structuredPct = offer.percent_off;
            percentOff = (structuredPct != null && !isNaN(Number(structuredPct)))
              ? Number(structuredPct)
              : extractPercentOff_(`${title} ${description}`);
            if (percentOff == null && !couponCode) continue;
          }
        }

        if (isDuplicateLink_(link, dedupSet)) continue;

        const haystack = `${title} ${description}`.toLowerCase();
        const category = classifyCategory_(haystack, categoryRegexes);
        if (!category) continue;
        if (quotaUsed[category] >= quotaLimit) continue;

        const displayName = isEstimated ? `${title.trim()} (Est.)` : title.trim();
        rowsToAppend.push([
          displayName, link.trim(),
          orig === '' ? '' : Number(orig), deal === '' ? '' : Number(deal),
          category, 'Pending',
          couponCode, percentOff === '' || percentOff == null ? '' : Number(percentOff)
        ]);
        quotaUsed[category]++;
        dedupSet.add(link.trim());

      } catch (err) {
        logError_(`fetchCuelinksOffers page ${page} item`, err);
      }
    }

    page++;
  }

  if (rowsToAppend.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, ROW_WIDTH).setValues(rowsToAppend);
  }

  Logger.log(`fetchCuelinksOffers: appended ${rowsToAppend.length} rows.`);
  return { total: rowsToAppend.length, byCategory: quotaUsed };
}

// =====================================================================
// POSTING ENGINE — gated queue + LockService + move to Recent_Posted
// =====================================================================

function processQueue_(triggerLabel) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log(`processQueue_ (${triggerLabel}): Lock unavailable, skipping concurrent execution.`);
    return;
  }

  try {
    setConfigValue_(`LAST_${triggerLabel}_RUN_TIMESTAMP`, new Date());

    const config = readConfig_();
    const isSaleMode = getConfigBool_(config, 'IS_SALE_MODE_ACTIVE', CONFIG_DEFAULTS.IS_SALE_MODE_ACTIVE);
    const intervalMinutes = getConfigNumber_(config, 'POSTING_INTERVAL_MINUTES', CONFIG_DEFAULTS.POSTING_INTERVAL_MINUTES);
    const lastPost = getConfigDate_(config, 'LAST_POST_TIMESTAMP');

    if (!isSaleMode && lastPost) {
      const elapsedMs = Date.now() - lastPost.getTime();
      const intervalMs = intervalMinutes * 60 * 1000;
      if (elapsedMs < intervalMs) {
        return; // cooldown active
      }
    }

    const sheet = getLiveQueueSheet_();
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow <= HEADER_ROW) return;

    const values = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, ROW_WIDTH).getValues();
    const pendingRowNumbers = [];
    values.forEach((row, idx) => {
      if (row[COL.STATUS - 1] === 'Pending') pendingRowNumbers.push(HEADER_ROW + 1 + idx);
    });
    if (pendingRowNumbers.length === 0) return;

    if (!isSaleMode) {
      const rowNum = pendingRowNumbers[0];
      const rowValues = sheet.getRange(rowNum, 1, 1, ROW_WIDTH).getValues()[0];
      const posted = postQueueRow_(sheet, rowNum, rowValues);
      if (posted) setConfigValue_('LAST_POST_TIMESTAMP', new Date());
      return;
    }

    // Sale mode: burst-process rows, capped for timeout safety
    const SALE_MODE_BURST_CAP = 25;
    const toProcess = pendingRowNumbers.slice(0, SALE_MODE_BURST_CAP).reverse();
    let anyPosted = false;

    toProcess.forEach(rowNum => {
      const rowValues = sheet.getRange(rowNum, 1, 1, ROW_WIDTH).getValues()[0];
      if (rowValues[COL.STATUS - 1] !== 'Pending') return;
      const posted = postQueueRow_(sheet, rowNum, rowValues);
      if (posted) anyPosted = true;
    });

    if (anyPosted) setConfigValue_('LAST_POST_TIMESTAMP', new Date());
  } finally {
    lock.releaseLock();
  }
}

/**
 * Validates and sends ONE row to Telegram.
 * On success: appends it to Recent_Posted (active dedup window) and deletes it from Live_Queue.
 * On failure: marks row with "Error: <reason>".
 */
function postQueueRow_(sheet, rowNum, rowValues) {
  const [name, link, origPrice, dealPrice, category, , couponCode, percentOffRaw] = rowValues;

  try {
    const missing = [];
    if (!name) missing.push('Product_Name');
    if (!link) missing.push('Raw_URL');
    if (!category) missing.push('Category');
    if (missing.length > 0) throw new Error('Missing field(s): ' + missing.join(', '));

    const categoryConfig = CATEGORY_MAP[category];
    if (!categoryConfig) throw new Error(`Undefined category "${category}" (expected Gym, Pets, or Vehicle)`);

    const props = PropertiesService.getScriptProperties();
    const botToken = props.getProperty('BOT_TOKEN');
    const chatId = props.getProperty(categoryConfig.propKey);
    if (!botToken) throw new Error('BOT_TOKEN not set in Script Properties');
    if (!chatId) throw new Error(`${categoryConfig.propKey} not set in Script Properties`);

    const hasOrig = origPrice !== '' && origPrice != null;
    const hasDeal = dealPrice !== '' && dealPrice != null;
    const hasCouponCode = couponCode !== '' && couponCode != null;
    const hasPercentOff = percentOffRaw !== '' && percentOffRaw != null && !isNaN(Number(percentOffRaw));

    let message;

    if (hasOrig || hasDeal) {
      if (!hasOrig || !hasDeal) throw new Error('Original_Price and Deal_Price must both be set, or both left blank');

      const orig = Number(origPrice);
      const deal = Number(dealPrice);
      if (isNaN(orig) || orig <= 0) throw new Error('Invalid Original_Price');
      if (isNaN(deal) || deal <= 0) throw new Error('Invalid Deal_Price');
      if (deal >= orig) throw new Error('Deal_Price must be lower than Original_Price');

      const savingsPct = Math.round(((orig - deal) / orig) * 100);
      message = buildDealMessage_(categoryConfig, name, orig, deal, savingsPct, link, hasCouponCode ? couponCode : null);

    } else if (hasPercentOff || hasCouponCode) {
      const percentOff = hasPercentOff ? Math.round(Number(percentOffRaw)) : null;
      message = buildCouponMessage_(categoryConfig, name, hasCouponCode ? couponCode : null, percentOff, link);

    } else {
      throw new Error('No Original_Price/Deal_Price AND no Percent_Off/Coupon_Code — nothing to post');
    }

    const result = sendToTelegram_(botToken, chatId, message, link);
    if (!result.ok) throw new Error(result.error);

    appendToRecentPosted_([
      name, link,
      hasOrig ? Number(origPrice) : '', hasDeal ? Number(dealPrice) : '',
      category, new Date(),
      hasCouponCode ? couponCode : '', hasPercentOff ? Number(percentOffRaw) : ''
    ]);
    sheet.deleteRow(rowNum);
    return true;

  } catch (err) {
    sheet.getRange(rowNum, COL.STATUS).setValue('Error: ' + err.message);
    logError_('postQueueRow_ row ' + rowNum, err);
    return false;
  }
}

function appendToRecentPosted_(rowValues) {
  const sheet = getRecentPostedSheet_();
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, rowValues.length).setValues([rowValues]);
}

function buildDealMessage_(categoryConfig, name, orig, deal, savingsPct, link, couponCode) {
  const safeName = escapeMarkdownV2_(name);
  const origStr = escapeMarkdownV2_(`₹${orig.toLocaleString('en-IN')}`);
  const dealStr = escapeMarkdownV2_(`₹${deal.toLocaleString('en-IN')}`);
  const pctStr = escapeMarkdownV2_(`${savingsPct}%`);
  const safeLinkUrl = String(link).replace(/([)\\])/g, '\\$1');

  const lines = [
    `${categoryConfig.emoji} *${categoryConfig.label}* ${categoryConfig.emoji}`,
    ``,
    `*${safeName}*`,
    ``,
    `~${origStr}~  ➡️  *${dealStr}*`,
    `🔥 *${pctStr} OFF*`
  ];
  if (couponCode) {
    lines.push(`🏷️ Code: \`${escapeMarkdownV2Code_(couponCode)}\``);
  }
  lines.push(``, `🛒 [Buy Now](${safeLinkUrl})`);
  return lines.join('\n');
}

function buildCouponMessage_(categoryConfig, name, couponCode, percentOff, link) {
  const safeName = escapeMarkdownV2_(name);
  const safeLinkUrl = String(link).replace(/([)\\])/g, '\\$1');

  const lines = [
    `${categoryConfig.emoji} *${categoryConfig.label}* ${categoryConfig.emoji}`,
    ``,
    `*${safeName}*`
  ];
  if (percentOff) {
    lines.push(``, `🔥 *${escapeMarkdownV2_(percentOff + '%')} OFF*`);
  }
  if (couponCode) {
    lines.push(`🏷️ Code: \`${escapeMarkdownV2Code_(couponCode)}\``);
  }
  lines.push(``, `🛒 [Grab This Deal](${safeLinkUrl})`);
  return lines.join('\n');
}

function escapeMarkdownV2_(text) {
  return String(text).replace(/[_*\[\]()~`>#+\-=|{}.!]/g, '\\$&');
}

function escapeMarkdownV2Code_(text) {
  return String(text).replace(/([`\\])/g, '\\$1');
}

function sendToTelegram_(botToken, chatId, text, link) {
  const url = `https://api.telegram.org/bot${botToken}/sendMessage`;

  const payload = {
    chat_id: chatId,
    text: text,
    parse_mode: 'MarkdownV2',
    disable_web_page_preview: false,
    reply_markup: JSON.stringify({
      inline_keyboard: [[{ text: '🛒 Buy Now', url: link }]]
    })
  };

  const options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: payload,
    muteHttpExceptions: true
  };

  const maxAttempts = 3;
  let lastError = '';

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      const body = JSON.parse(response.getContentText());

      if (code === 200 && body.ok) return { ok: true };

      lastError = `HTTP ${code}: ${body.description || 'unknown error'}`;

      if (code === 429 && body.parameters && body.parameters.retry_after) {
        Utilities.sleep(body.parameters.retry_after * 1000);
        continue;
      }
      if (code >= 500) {
        Utilities.sleep(attempt * 1000);
        continue;
      }
      break;
    } catch (err) {
      lastError = err.message;
      Utilities.sleep(attempt * 1000);
    }
  }

  return { ok: false, error: lastError };
}

// =====================================================================
// ARCHIVING & MAINTENANCE (4-Hour Recent Migration + Daily Maintenance)
// =====================================================================

/**
 * Moves rows from Recent_Posted to Historical_Archive if they are older
 * than DEDUP_WINDOW_HOURS (default 4 hours).
 * Runs on a 4-hour time-driven trigger.
 */
function archiveRecentDeals() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    Logger.log('archiveRecentDeals: Could not acquire lock, skipping.');
    return;
  }

  try {
    setConfigValue_('LAST_ARCHIVE_RUN_TIMESTAMP', new Date());
    const config = readConfig_();
    const windowHours = getConfigNumber_(config, 'DEDUP_WINDOW_HOURS', CONFIG_DEFAULTS.DEDUP_WINDOW_HOURS);
    const cutoffTime = Date.now() - (windowHours * 60 * 60 * 1000);

    const recentSheet = getRecentPostedSheet_();
    const lastRow = recentSheet.getLastRow();
    if (lastRow <= HEADER_ROW) return;

    const values = recentSheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, ROW_WIDTH).getValues();
    const stillRecent = [];
    const toArchive = [];

    values.forEach(row => {
      const sentAtVal = row[ARCHIVE_COL.SENT_AT - 1];
      const sentAtDate = (sentAtVal instanceof Date) ? sentAtVal : new Date(sentAtVal);
      const isValidDate = !isNaN(sentAtDate.getTime());

      if (isValidDate && sentAtDate.getTime() < cutoffTime) {
        toArchive.push(row);
      } else {
        stillRecent.push(row);
      }
    });

    if (toArchive.length === 0) {
      Logger.log(`archiveRecentDeals: 0 rows older than ${windowHours}h to archive.`);
      return;
    }

    // 1. Batch append expired rows to Historical_Archive
    const archiveSheet = getArchiveSheet_();
    const archiveStart = archiveSheet.getLastRow() + 1;
    archiveSheet.getRange(archiveStart, 1, toArchive.length, ROW_WIDTH).setValues(toArchive);

    // 2. Rewrite Recent_Posted with only active, unexpired rows
    recentSheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, ROW_WIDTH).clearContent();
    if (stillRecent.length > 0) {
      recentSheet.getRange(HEADER_ROW + 1, 1, stillRecent.length, ROW_WIDTH).setValues(stillRecent);
    }

    Logger.log(`archiveRecentDeals: Moved ${toArchive.length} rows to Historical_Archive. ${stillRecent.length} rows remaining in Recent_Posted.`);
  } catch (err) {
    logError_('archiveRecentDeals', err);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Daily early morning maintenance routine:
 * 1. Purges Logs older than LOG_RETENTION_HOURS (default 24h) and caps total log rows at LOG_MAX_ROWS (default 1000).
 * 2. Purges Historical_Archive entries older than ARCHIVE_RETENTION_DAYS (default 30 days).
 * Runs on a daily time-driven trigger (e.g. 4:00 AM - 5:00 AM).
 */
function runDailyMaintenance() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    Logger.log('runDailyMaintenance: Could not acquire lock, skipping.');
    return;
  }

  try {
    setConfigValue_('LAST_MAINTENANCE_RUN_TIMESTAMP', new Date());
    const config = readConfig_();

    cleanLogs_(config);
    cleanHistoricalArchive_(config);

    Logger.log('runDailyMaintenance completed successfully.');
  } catch (err) {
    logError_('runDailyMaintenance', err);
  } finally {
    lock.releaseLock();
  }
}

function cleanLogs_(config) {
  const ss = SpreadsheetApp.getActive();
  const logSheet = ss.getSheetByName(LOGS_SHEET);
  if (!logSheet) return;

  const lastRow = logSheet.getLastRow();
  if (lastRow <= HEADER_ROW) return;

  const retentionHours = getConfigNumber_(config, 'LOG_RETENTION_HOURS', CONFIG_DEFAULTS.LOG_RETENTION_HOURS);
  const maxRows = getConfigNumber_(config, 'LOG_MAX_ROWS', CONFIG_DEFAULTS.LOG_MAX_ROWS);
  const cutoffTime = Date.now() - (retentionHours * 60 * 60 * 1000);

  const values = logSheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, 3).getValues();

  // 1. Filter out entries older than retentionHours
  let freshLogs = values.filter(row => {
    const ts = (row[0] instanceof Date) ? row[0] : new Date(row[0]);
    return !isNaN(ts.getTime()) && ts.getTime() >= cutoffTime;
  });

  // 2. Keep at most maxRows (latest)
  if (freshLogs.length > maxRows) {
    freshLogs = freshLogs.slice(freshLogs.length - maxRows);
  }

  // 3. Batch rewrite
  logSheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, 3).clearContent();
  if (freshLogs.length > 0) {
    logSheet.getRange(HEADER_ROW + 1, 1, freshLogs.length, 3).setValues(freshLogs);
  }

  // 4. Shrink excess empty rows if sheet is bloated
  const currentMax = logSheet.getMaxRows();
  const targetRows = Math.max(freshLogs.length + HEADER_ROW + 10, 100);
  if (currentMax > targetRows + 100) {
    logSheet.deleteRows(targetRows + 1, currentMax - targetRows);
  }
}

function cleanHistoricalArchive_(config) {
  const archiveSheet = getArchiveSheet_();
  if (!archiveSheet) return;

  const lastRow = archiveSheet.getLastRow();
  if (lastRow <= HEADER_ROW) return;

  const retentionDays = getConfigNumber_(config, 'ARCHIVE_RETENTION_DAYS', CONFIG_DEFAULTS.ARCHIVE_RETENTION_DAYS);
  const cutoffTime = Date.now() - (retentionDays * 24 * 60 * 60 * 1000);

  const values = archiveSheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, ROW_WIDTH).getValues();
  const retainedRows = values.filter(row => {
    const sentAtVal = row[ARCHIVE_COL.SENT_AT - 1];
    const sentAt = (sentAtVal instanceof Date) ? sentAtVal : new Date(sentAtVal);
    return isNaN(sentAt.getTime()) || sentAt.getTime() >= cutoffTime;
  });

  if (retainedRows.length !== values.length) {
    archiveSheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, ROW_WIDTH).clearContent();
    if (retainedRows.length > 0) {
      archiveSheet.getRange(HEADER_ROW + 1, 1, retainedRows.length, ROW_WIDTH).setValues(retainedRows);
    }
    Logger.log(`cleanHistoricalArchive_: Purged ${values.length - retainedRows.length} rows older than ${retentionDays} days.`);
  }
}

// =====================================================================
// TRIGGERS & ERROR LOGGING
// =====================================================================

function onEditInstallable(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== LIVE_QUEUE_SHEET) return;
    if (e.range.getRow() <= HEADER_ROW) return;
    processQueue_('ONEDIT');
  } catch (err) {
    logError_('onEditInstallable', err);
  }
}

function sweepPendingRows() {
  try {
    processQueue_('SWEEP');
  } catch (err) {
    logError_('sweepPendingRows', err);
  }
}

/**
 * One-time setup: installs the onEdit trigger.
 */
function createTrigger() {
  const ss = SpreadsheetApp.getActive();

  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onEditInstallable') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('onEditInstallable')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  Logger.log('----------------------------------------------------');
  Logger.log('SUCCESS: onEdit trigger installed.');
  Logger.log('Now manually set up the 4 time-driven triggers:');
  Logger.log('  1. sweepPendingRows    -> Time-driven (Minutes timer: every 1 to 5 min)');
  Logger.log('  2. runAllIntakes       -> Time-driven (Minutes/Hour timer: every 30 to 120 min)');
  Logger.log('  3. archiveRecentDeals  -> Time-driven (Hour timer: every 4 hours)');
  Logger.log('  4. runDailyMaintenance -> Time-driven (Day timer: daily, 4:00 AM - 5:00 AM)');
  Logger.log('----------------------------------------------------');
}

function logError_(context, err) {
  Logger.log(`[${context}] ${err.message}`);
  try {
    const ss = SpreadsheetApp.getActive();
    let logSheet = ss.getSheetByName(LOGS_SHEET);
    if (!logSheet) {
      logSheet = ss.insertSheet(LOGS_SHEET);
      logSheet.appendRow(['Timestamp', 'Context', 'Error']);
    }
    logSheet.appendRow([new Date(), context, err.message]);
  } catch (e) {
    // swallow logging errors so they never break the main flow
  }
}
