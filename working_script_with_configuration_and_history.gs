/**
 * Sheets -> Telegram MULTI-CHANNEL pipeline
 * Light-Queue / Configuration-Dashboard / Historical-Archive architecture
 * -------------------------------------------------------------------------
 * THREE TABS REQUIRED (exact names):
 *
 *   1. Live_Queue        — active, lightweight. Columns A-H:
 *        A: Product_Name | B: Raw_URL | C: Original_Price | D: Deal_Price
 *        E: Category (Gym / Pets / Vehicle) | F: Status (Pending / Error: ...)
 *        G: Coupon_Code (optional) | H: Percent_Off (optional)
 *      C/D (prices) OR H (percent) should be populated — NOT both required.
 *      A row with real/estimated prices posts as a price-comparison deal;
 *      a row with only a Percent_Off (and blank C/D) posts as a coupon-
 *      style message instead. See postQueueRow_ for the exact branching.
 *      NOTE: successfully-posted rows are DELETED from here and moved to
 *      Historical_Archive — Live_Queue never accumulates "Sent" rows.
 *
 *   2. Configuration     — key/value dashboard. Column A = Key, B = Value.
 *      Row 1: headers (Key, Value)
 *      Row 2: POSTING_INTERVAL_MINUTES   e.g. 5
 *      Row 3: MAX_DEALS_PER_INTAKE_RUN   e.g. 2
 *      Row 4: IS_SALE_MODE_ACTIVE        TRUE or FALSE
 *      Row 5: LAST_POST_TIMESTAMP        (leave blank — script fills this in)
 *      Additional rows (LAST_INTAKE_RUN_TIMESTAMP, LAST_SWEEP_RUN_TIMESTAMP,
 *      LAST_ONEDIT_RUN_TIMESTAMP) are auto-created by the script the first
 *      time each trigger runs — you don't need to pre-add them.
 *
 *   3. Historical_Archive — permanent log of every deal ever posted.
 *      Columns A-H: Product_Name | Raw_URL | Original_Price | Deal_Price
 *      | Category | Sent_At | Coupon_Code | Percent_Off
 *      This is also the second half of the dedup index (Live_Queue +
 *      Historical_Archive Raw_URL columns = full de-dup history).
 *
 * SETUP:
 * 1. Create the three tabs above with exact names and header rows.
 * 2. Extensions > Apps Script, paste this file in as Code.gs.
 * 3. Project Settings > Script Properties, add:
 *      BOT_TOKEN         = <your bot token from @BotFather>
 *      CHANNEL_GYM       = @YourGymChannel   (or numeric -100... id)
 *      CHANNEL_PETS      = @YourPetChannel
 *      CHANNEL_VEHICLE   = @YourVehicleChannel
 *      CUELINKS_API_KEY  = <your key from https://www.cuelinks.com/api-key>
 * 4. Run `createTrigger` once from the editor (grants permissions, installs
 *    the onEdit trigger).
 * 5. Set up TWO time-driven triggers manually (Triggers icon > Add Trigger):
 *      - sweepPendingRows -> every 1 minute is FINE now (see note below)
 *      - runAllIntakes    -> every 30-120 minutes (NOT 1 minute — see note)
 */

// =====================================================================
// SHEET NAMES & LAYOUT
// =====================================================================

const LIVE_QUEUE_SHEET = 'Live_Queue';
const CONFIG_SHEET = 'Configuration';
const ARCHIVE_SHEET = 'Historical_Archive';
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
  return SpreadsheetApp.getActive().getSheetByName(ARCHIVE_SHEET);
}

// =====================================================================
// CONFIGURATION DASHBOARD — read/write helpers
// =====================================================================

// Fallbacks used only if a key is missing/blank in the Configuration tab —
// the tab itself is always the source of truth when a value is present.
const CONFIG_DEFAULTS = {
  POSTING_INTERVAL_MINUTES: 5,
  MAX_DEALS_PER_INTAKE_RUN: 2,
  IS_SALE_MODE_ACTIVE: false
};

// Sale mode widens intake volume. Configuration only defines ONE
// MAX_DEALS_PER_INTAKE_RUN value (no separate sale-mode number was
// specified), so this multiplies the base value when sale mode is on.
// Optional override: add a MAX_DEALS_PER_INTAKE_RUN_SALE row to the
// Configuration tab for an exact number instead of a multiplier.
const SALE_MODE_INTAKE_MULTIPLIER = 7;

/** Reads the whole Configuration tab into a plain {key: value} object. */
function readConfig_() {
  const sheet = getConfigSheet_();
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
 * already present; otherwise self-heals by appending a new row — this is
 * how LAST_POST_TIMESTAMP, LAST_INTAKE_RUN_TIMESTAMP, etc. get created
 * automatically on first run without you having to pre-add them.
 */
function setConfigValue_(key, value) {
  const sheet = getConfigSheet_();
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
 * See SALE_MODE_INTAKE_MULTIPLIER note above for the sale-mode assumption.
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
// DEDUPLICATION INDEX — Live_Queue + Historical_Archive, O(1) lookups
// =====================================================================

/**
 * Builds a single Set of every Raw_URL that exists anywhere in the
 * pipeline right now — currently queued (Live_Queue) OR ever posted
 * (Historical_Archive). One column-only read per sheet, no nested loops.
 */
function buildDedupIndex_() {
  const set = new Set();
  [getLiveQueueSheet_(), getArchiveSheet_()].forEach(sheet => {
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

/**
 * Precompiled word-boundary regexes. Plain substring matching was
 * matching "cat" inside "category" and "car" inside "card" — \b
 * boundaries mean a keyword only matches as a standalone word.
 *
 * An optional trailing "s" is allowed before the closing boundary so
 * plurals still match ("dog" -> "dogs", "protein" -> "Proteins") without
 * reopening the substring-match bug: "category" still correctly fails
 * because "categor" isn't followed by a word-boundary after the "s?".
 */
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

/**
 * Fallback price extraction for offers with no structured original_price/
 * discount_price. Only succeeds when the text gives us BOTH a percent-off
 * figure AND a clear final/starting price ("only pay Rs. X", "starting at
 * Rs. X", "prices from Rs. X", "just Rs. X") — enough to back-calculate
 * original = deal / (1 - pct/100). Deliberately conservative: vaguer text
 * ("flat Rs. X off" with no base price) returns null rather than guessing.
 */
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

/**
 * Lighter fallback for the coupon-style message path: just grabs a
 * percent-off figure, with no requirement for an accompanying price
 * (unlike extractPriceFromText_, which needs both). Used only when the
 * stricter price extraction has already failed — this is Tier 3.
 */
function extractPercentOff_(text) {
  const m = text.match(/(\d{1,3})\s*%\s*(?:off|discount|extra)?/i);
  if (!m) return null;
  const pct = Number(m[1]);
  if (isNaN(pct) || pct <= 0 || pct >= 100) return null;
  return pct;
}

// =====================================================================
// XML/RSS GENERIC INTAKE ENGINE — for EarnKaro / INRDeals if confirmed
// =====================================================================

const DEAL_SOURCES = [
  {
    name: 'EarnKaro',
    enabled: false, // flip true once you've confirmed a real feed URL exists
    feedUrl: 'PASTE_EARNKARO_FEED_URL_HERE_IF_CONFIRMED',
    tags: { itemNode: 'item', titleTag: 'title', linkTag: 'link', descriptionTag: 'description' }
  },
  {
    name: 'INRDeals',
    enabled: false, // flip true once you've confirmed a real feed URL exists
    feedUrl: 'PASTE_INRDEALS_FEED_URL_HERE_IF_CONFIRMED',
    tags: { itemNode: 'item', titleTag: 'title', linkTag: 'link', descriptionTag: 'description' }
  }
];

function fetchDealsFromSource_(source, config) {
  const sheet = getLiveQueueSheet_();

  if (!source.feedUrl || source.feedUrl.indexOf('PASTE_') === 0) {
    logError_(`fetchDealsFromSource_ (${source.name})`, new Error('feedUrl not configured'));
    return;
  }

  let xmlText;
  try {
    const response = UrlFetchApp.fetch(source.feedUrl, { muteHttpExceptions: true });
    const code = response.getResponseCode();
    if (code !== 200) throw new Error(`Feed fetch failed with HTTP ${code}`);
    xmlText = response.getContentText();
  } catch (err) {
    logError_(`fetchDealsFromSource_ (${source.name}) fetch`, err);
    return;
  }

  let items;
  try {
    const document = XmlService.parse(xmlText);
    items = findAllDescendants_(document.getRootElement(), source.tags.itemNode);
  } catch (err) {
    logError_(`fetchDealsFromSource_ (${source.name}) parse`, err);
    return;
  }

  if (!items || items.length === 0) {
    logError_(`fetchDealsFromSource_ (${source.name})`, new Error('Feed parsed but returned 0 items — check tags.itemNode matches this feed'));
    return;
  }

  const dedupSet = buildDedupIndex_();
  const quotaLimit = getIntakeQuota_(config);
  const quotaUsed = { 'Gym': 0, 'Pets': 0, 'Vehicle': 0 };
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
      const category = classifyCategory_(haystack);
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

  Logger.log(`fetchDealsFromSource_ (${source.name}): appended ${rowsToAppend.length} rows. Quota used: ${JSON.stringify(quotaUsed)} / limit ${quotaLimit}`);
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
// CUELINKS — real JSON REST intake (developers.cuelinks.com/pub_api/v3)
// =====================================================================

const CUELINKS_API_BASE = 'https://developers.cuelinks.com/pub_api/v3';
const CUELINKS_PAGE_SIZE = 100;
const CUELINKS_MAX_PAGES = 10;

/**
 * Runs Cuelinks + every enabled XML source. Wire THIS to your intake
 * time-driven trigger (recommended: every 30-120 min, NOT 1 minute —
 * unlike the posting side, every invocation of this one does real
 * paginated API work).
 */
function runAllIntakes() {
  setConfigValue_('LAST_INTAKE_RUN_TIMESTAMP', new Date());
  const config = readConfig_();

  try {
    fetchCuelinksOffers(config);
  } catch (err) {
    logError_('runAllIntakes (Cuelinks)', err);
  }

  DEAL_SOURCES.forEach(source => {
    if (!source.enabled) return;
    try {
      fetchDealsFromSource_(source, config);
    } catch (err) {
      logError_(`runAllIntakes (${source.name})`, err);
    }
  });
}

function fetchCuelinksOffers(config) {
  const sheet = getLiveQueueSheet_();

  const apiKey = PropertiesService.getScriptProperties().getProperty('CUELINKS_API_KEY');
  if (!apiKey) {
    logError_('fetchCuelinksOffers', new Error('CUELINKS_API_KEY not set in Script Properties'));
    return;
  }

  const dedupSet = buildDedupIndex_();
  const quotaLimit = getIntakeQuota_(config);
  const quotaUsed = { 'Gym': 0, 'Pets': 0, 'Vehicle': 0 };
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
            // Tier 3: no derivable price at all — fall back to a
            // coupon-style post if there's at least a percent-off to show
            // (structured field first, then a lighter regex fallback that
            // doesn't require an accompanying price like extractPriceFromText_ does).
            orig = '';
            deal = '';
            const structuredPct = offer.percent_off;
            percentOff = (structuredPct != null && !isNaN(Number(structuredPct)))
              ? Number(structuredPct)
              : extractPercentOff_(`${title} ${description}`);
            if (percentOff == null && !couponCode) continue; // nothing worth posting
          }
        }

        if (isDuplicateLink_(link, dedupSet)) continue;

        const haystack = `${title} ${description}`.toLowerCase();
        const category = classifyCategory_(haystack);
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

  Logger.log(`fetchCuelinksOffers: appended ${rowsToAppend.length} rows across ${page - 1} page(s). Quota used: ${JSON.stringify(quotaUsed)} / limit ${quotaLimit}`);
}

// =====================================================================
// POSTING ENGINE — dynamic time-gating + move-on-success queue
// =====================================================================

/**
 * Shared entry point for BOTH the onEdit trigger and the time-driven
 * sweep trigger. triggerLabel is used to record a per-trigger last-run
 * timestamp in Configuration (e.g. LAST_SWEEP_RUN_TIMESTAMP).
 *
 * Gating: if IS_SALE_MODE_ACTIVE is false and POSTING_INTERVAL_MINUTES
 * hasn't elapsed since LAST_POST_TIMESTAMP, this returns immediately —
 * no Live_Queue scan, no Telegram calls, no quota consumed beyond the
 * trigger invocation itself. Once the cooldown clears, it posts exactly
 * ONE row (oldest Pending first) and resets the cooldown.
 *
 * If IS_SALE_MODE_ACTIVE is true, gating is bypassed entirely and every
 * currently-Pending row (up to a safety cap) is posted in one run.
 */
function processQueue_(triggerLabel) {
  setConfigValue_(`LAST_${triggerLabel}_RUN_TIMESTAMP`, new Date());

  const config = readConfig_();
  const isSaleMode = getConfigBool_(config, 'IS_SALE_MODE_ACTIVE', CONFIG_DEFAULTS.IS_SALE_MODE_ACTIVE);
  const intervalMinutes = getConfigNumber_(config, 'POSTING_INTERVAL_MINUTES', CONFIG_DEFAULTS.POSTING_INTERVAL_MINUTES);
  const lastPost = getConfigDate_(config, 'LAST_POST_TIMESTAMP');

  if (!isSaleMode && lastPost) {
    const elapsedMs = Date.now() - lastPost.getTime();
    const intervalMs = intervalMinutes * 60 * 1000;
    if (elapsedMs < intervalMs) {
      return; // cooldown still active — exit cleanly, no further work
    }
  }

  const sheet = getLiveQueueSheet_();
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

  // Sale mode: burst-process all eligible rows (safety-capped), processing
  // from the highest row number down so deleting a row never invalidates
  // the row numbers of rows still waiting to be processed below it.
  const SALE_MODE_BURST_CAP = 50;
  const toProcess = pendingRowNumbers.slice(0, SALE_MODE_BURST_CAP).reverse();
  let anyPosted = false;

  toProcess.forEach(rowNum => {
    const rowValues = sheet.getRange(rowNum, 1, 1, ROW_WIDTH).getValues()[0];
    if (rowValues[COL.STATUS - 1] !== 'Pending') return; // safety re-check
    const posted = postQueueRow_(sheet, rowNum, rowValues);
    if (posted) anyPosted = true;
  });

  if (anyPosted) setConfigValue_('LAST_POST_TIMESTAMP', new Date());
}

/**
 * Validates, formats, and sends ONE row. On success: appends it to
 * Historical_Archive and deletes it from Live_Queue (move-on-success —
 * Live_Queue never accumulates "Sent" rows). On failure: writes
 * "Error: <reason>" into the Status column in place, row stays put.
 * Returns true on success, false on failure.
 *
 * Two message shapes, chosen by what data the row actually has:
 *   - Original_Price + Deal_Price both present & valid -> price-comparison
 *     deal message (unchanged from before).
 *   - Neither present, but Percent_Off (and/or Coupon_Code) present ->
 *     coupon-style message instead. A row with NEITHER prices NOR a
 *     percent is a genuine data error (nothing worth posting), not a
 *     silently-skipped blank state.
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
      // Partial price data (one filled, one blank) is a real data error —
      // NOT the same as both intentionally blank for a coupon-style row.
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

    appendToArchive_([
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

function appendToArchive_(rowValues) {
  const sheet = getArchiveSheet_();
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, rowValues.length).setValues([rowValues]);
}

/**
 * MarkdownV2 price-comparison message with dynamic category emoji,
 * strikethrough original price, and a Buy Now hyperlink embedded in the
 * text (not just the inline button) so Telegram still generates a link
 * preview image. Coupon code line included when present, same as the
 * coupon-style message below.
 */
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

/**
 * Coupon-style message for offers with no derivable Original/Deal price —
 * shows percent-off (when known) and the coupon code (when known) instead
 * of a price comparison. Requires at least one of the two to be present
 * (enforced by the caller) so this never posts a bare, valueless message.
 */
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

/**
 * Escapes only what's special INSIDE a MarkdownV2 code span (backtick and
 * backslash) — using the full escapeMarkdownV2_ here would wrongly show
 * literal backslashes in coupon codes containing "-" or other symbols.
 */
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
// TRIGGERS
// =====================================================================

/**
 * Both entry points funnel through the same gated processQueue_(), so
 * pacing is enforced identically no matter what fired the trigger. An
 * edit anywhere in Live_Queue attempts to post the oldest eligible row,
 * not necessarily the row you just edited — this keeps posting strictly
 * oldest-first and avoids two different code paths with different rules.
 */
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
 * One-time setup: installs the onEdit trigger. Time-driven triggers
 * (sweepPendingRows, runAllIntakes) are set up manually via the Triggers
 * UI so you can choose different intervals for each — see the note at
 * the top of this file for recommended intervals.
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

  Logger.log('onEdit trigger created. Set up sweepPendingRows and runAllIntakes as time-driven triggers manually via the Triggers UI.');
}

function logError_(context, err) {
  Logger.log(`[${context}] ${err.message}`);
  try {
    const ss = SpreadsheetApp.getActive();
    let logSheet = ss.getSheetByName('Logs');
    if (!logSheet) {
      logSheet = ss.insertSheet('Logs');
      logSheet.appendRow(['Timestamp', 'Context', 'Error']);
    }
    logSheet.appendRow([new Date(), context, err.message]);
  } catch (e) {
    // swallow logging errors so they never break the main flow
  }
}
