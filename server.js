require('dotenv').config();
const express = require('express');
const line = require('@line/bot-sdk');
const OpenAI = require('openai');
const { DocumentAnalysisClient, AzureKeyCredential } = require('@azure/ai-form-recognizer');
const jpeg = require('jpeg-js');
const fs = require('fs').promises;
const fsSync = require('fs');
const path = require('path');
const https = require('https');
const crypto = require('crypto');
const db = require('./db');

// Debug logging to file
const logFile = fsSync.createWriteStream(path.join(__dirname, 'debug.log'), { flags: 'a' });
const origLog = console.log;
const origError = console.error;
console.log = (...args) => { logFile.write(new Date().toISOString() + ' ' + args.join(' ') + '\n'); origLog(...args); };
console.error = (...args) => { logFile.write(new Date().toISOString() + ' ERROR ' + args.join(' ') + '\n'); origError(...args); };

const config = {
  channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.LINE_CHANNEL_SECRET,
};

const client = new line.messagingApi.MessagingApiClient({
  channelAccessToken: config.channelAccessToken,
});

const blobClient = new line.messagingApi.MessagingApiBlobClient({
  channelAccessToken: config.channelAccessToken,
});

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
  timeout: 30000,   // fail fast (30s) instead of the SDK default 10min — avoids silent multi-minute hangs
  maxRetries: 2,
  // Use Node's native fetch (undici) instead of the SDK's bundled node-fetch v2.
  // On production every chat-completions POST failed with "Premature close"
  // (ERR_STREAM_PREMATURE_CLOSE) while raw https and native fetch both worked —
  // confirmed via /diag. It's a node-fetch v2 connection-handling bug on this
  // host, not a network/firewall issue. httpAgent is intentionally omitted: it
  // only applies to node-fetch, and there is no IPv6 route to pin anyway.
  fetch: (...args) => globalThis.fetch(...args),
});

const azureClient = new DocumentAnalysisClient(
  process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT,
  new AzureKeyCredential(process.env.AZURE_DOCUMENT_INTELLIGENCE_KEY)
);

let latestOCRResult = {
  timestamp: null,
  extractedText: '',
  tableData: null,
  rawResult: null
};

// Notification group IDs (comma-separated in .env)
const NOTIFICATION_GROUP_IDS = process.env.NOTIFICATION_GROUP_IDS
  ? process.env.NOTIFICATION_GROUP_IDS.split(',').map(id => id.trim()).filter(id => id)
  : [];

// Auto-send notification after successful extraction (default: true)
const AUTO_NOTIFY = process.env.AUTO_NOTIFY?.toLowerCase() !== 'false';

// Groups that receive operational alerts from the watchdog (a day with no data).
// Deliberately separate from NOTIFICATION_GROUP_IDS: those are the customer-facing
// "report recorded" groups and should not receive ops noise.
const ALERT_GROUP_IDS = process.env.ALERT_GROUP_IDS
  ? process.env.ALERT_GROUP_IDS.split(',').map(id => id.trim()).filter(id => id)
  : [];

// LINE user IDs that receive developer-facing alerts — currently the reconciliation
// warning raised when a column's per-CDC sum disagrees with the ยอดรวม printed on the
// plan. These go to a person, not a group: the record is still saved and the customer
// notification still goes out, so this must not become noise in a shared room.
const DEV_ALERT_USER_IDS = process.env.DEV_ALERT_USER_IDS
  ? process.env.DEV_ALERT_USER_IDS.split(',').map(id => id.trim()).filter(id => id)
  : [];

// How many days back the watchdog looks for missing data, and the hour (Bangkok time)
// after which the current day is expected to have been reported.
const WATCHDOG_LOOKBACK_DAYS = parseInt(process.env.WATCHDOG_LOOKBACK_DAYS || '7', 10);
const WATCHDOG_CUTOFF_HOUR = parseInt(process.env.WATCHDOG_CUTOFF_HOUR || '15', 10);

console.log('[CONFIG] Notification groups configured:', NOTIFICATION_GROUP_IDS.length);
console.log('[CONFIG] Notification group IDs:', NOTIFICATION_GROUP_IDS);
console.log('[CONFIG] Auto-notify:', AUTO_NOTIFY);
console.log('[CONFIG] Alert groups configured:', ALERT_GROUP_IDS.length);
if (ALERT_GROUP_IDS.length === 0) {
  console.warn('[CONFIG] WARNING: ALERT_GROUP_IDS is empty — the missing-data watchdog will log but cannot notify anyone');
}
console.log('[CONFIG] Dev alert users configured:', DEV_ALERT_USER_IDS.length);
if (DEV_ALERT_USER_IDS.length === 0) {
  console.warn('[CONFIG] WARNING: DEV_ALERT_USER_IDS is empty — reconciliation mismatches will be logged but nobody is told');
}

// Test database connection on startup
db.testConnection().then(connected => {
  if (connected) {
    console.log('[DB] MySQL database connected successfully');
  } else {
    console.error('[DB] WARNING: MySQL connection failed - check your database configuration');
  }
});

// Load detection logs from MySQL
async function loadDetectionLogs() {
  try {
    return await db.getDetectionLogs(100);
  } catch (error) {
    console.error('[LOG] Error loading detection logs:', error);
    return [];
  }
}

// Save detection log to MySQL
async function saveDetectionLog(logEntry) {
  try {
    await db.saveDetectionLog(logEntry);
    console.log('[LOG] Detection log saved:', logEntry.status);
  } catch (error) {
    console.error('[LOG] Error saving detection log:', error);
  }
}

// Load daily records from MySQL
async function loadDailyRecords() {
  try {
    return await db.loadDailyRecords();
  } catch (error) {
    console.error('[RECORD] Error loading daily records:', error);
    // Return empty arrays for all categories
    const empty = {};
    for (const cat of db.CATEGORIES) {
      empty[cat] = [];
    }
    return empty;
  }
}

// Save a single daily record to MySQL
async function saveDailyRecord(record, category) {
  try {
    await db.saveDailyRecord(record, category);
  } catch (error) {
    console.error('[RECORD] Error saving daily record:', error);
    throw error;
  }
}

// Send notification to configured groups about data update
async function sendNotificationToGroups(date, categories) {
  console.log(`[NOTIFICATION] sendNotificationToGroups called with date: ${date}, categories:`, categories);
  console.log(`[NOTIFICATION] NOTIFICATION_GROUP_IDS:`, NOTIFICATION_GROUP_IDS);

  if (NOTIFICATION_GROUP_IDS.length === 0) {
    console.log('[NOTIFICATION] No notification groups configured');
    return;
  }

  const message = `Report for ${date} has been recorded\n\nCategories: ${categories.join(', ')}\n\nDaily report: https://shinsen.yushi-marketing.com/daily-report\nMTD / YTD report: https://shinsen.yushi-marketing.com/mtd-report`;
  console.log(`[NOTIFICATION] Message to send: ${message}`);

  for (const groupId of NOTIFICATION_GROUP_IDS) {
    console.log(`[NOTIFICATION] Attempting to send to group: ${groupId}`);
    try {
      await client.pushMessage({
        to: groupId,
        messages: [{
          type: 'text',
          text: message,
        }],
      });
      console.log(`[NOTIFICATION] Successfully sent notification to group: ${groupId}`);
    } catch (error) {
      console.error(`[NOTIFICATION] Failed to send to group ${groupId}:`, error.message);
      console.error(`[NOTIFICATION] Full error:`, error);
    }
  }
  console.log('[NOTIFICATION] sendNotificationToGroups completed');
}

// Push an operational alert to the alert groups. Never throws: an alert that fails
// to send must not take down whatever called it.
async function sendAlertToGroups(text) {
  if (ALERT_GROUP_IDS.length === 0) {
    console.warn('[ALERT] No alert groups configured, alert not sent:', text.split('\n')[0]);
    return false;
  }

  let sent = false;
  for (const groupId of ALERT_GROUP_IDS) {
    try {
      await client.pushMessage({ to: groupId, messages: [{ type: 'text', text }] });
      console.log(`[ALERT] Sent alert to group: ${groupId}`);
      sent = true;
    } catch (error) {
      console.error(`[ALERT] Failed to send alert to group ${groupId}:`, error.message);
    }
  }
  return sent;
}

// Push a developer-facing alert to a private LINE chat. Same never-throws contract as
// sendAlertToGroups: this fires alongside a successful save, so a failed alert must not
// turn a recorded day into a failed one.
async function sendDevAlert(text) {
  if (DEV_ALERT_USER_IDS.length === 0) {
    console.warn('[DEV-ALERT] No dev users configured, alert not sent:', text.split('\n')[0]);
    return false;
  }

  let sent = false;
  for (const userId of DEV_ALERT_USER_IDS) {
    try {
      await client.pushMessage({ to: userId, messages: [{ type: 'text', text }] });
      console.log(`[DEV-ALERT] Sent alert to user: ${userId}`);
      sent = true;
    } catch (error) {
      console.error(`[DEV-ALERT] Failed to send alert to user ${userId}:`, error.message);
    }
  }
  return sent;
}

// Check if date already exists in records (uses MySQL)
async function isDateRecorded(date, category) {
  try {
    return await db.isDateRecorded(date, category);
  } catch (error) {
    console.error('[RECORD] Error checking if date recorded:', error);
    return false;
  }
}

// Reload auto-added juices. A failure keeps the last loaded list (built-ins at minimum):
// a dashboard or an OCR run must not fail because this lookup did.
async function refreshCategories() {
  try {
    await db.refreshCategories();
  } catch (error) {
    console.error('[CATEGORY] Could not refresh auto-added juices:', error.message);
  }
}

// Product detection configuration - keywords to identify each product in column headers
// Detection uses longest-matching-keyword scoring: when multiple products match the same cell,
// the product with the longest matching keyword wins (e.g. "มุกป๊อป" beats "น้ำส้ม" in
// "น้ำส้มมุกป๊อป" — pop wins over orange).
const PRODUCT_DETECTION = {
  orange: {
    keywords: ['น้ำส้ม', 'orange', 'ส้ม'],
    dbCategory: 'orange'
  },
  yuzu: {
    keywords: ['ยูซุ', 'yuzu'],
    dbCategory: 'yuzu'
  },
  pop: {
    keywords: ['shinsen pop', 'มุกป๊อป', 'ป๊อป', 'pop'],
    dbCategory: 'pop'
  },
  mixed: {
    keywords: ['น้ำผลไม้รวม', 'ผลไม้รวม', 'mixed fruit'],
    dbCategory: 'mixed'
  },
  tomato: {
    keywords: ['tomato', 'มะเขือเทศ'],
    dbCategory: 'tomato'
  },
  pineapple: {
    keywords: ['สับปะรด', 'pineapple'],
    dbCategory: 'pineapple'
  },
  watermelon: {
    keywords: ['แตงโม', 'watermelon'],
    dbCategory: 'watermelon'
  }
};

// Built-in juices plus auto-added ones (db.CUSTOM_KEYWORDS, from product_categories), each
// with any spellings merged into it on the juice page (db.JUICE_ALIASES).
function getProductDetection() {
  const detection = {};
  for (const [key, config] of Object.entries(PRODUCT_DETECTION)) {
    detection[key] = { ...config, keywords: [...config.keywords] };
  }
  for (const [category, keywords] of Object.entries(db.CUSTOM_KEYWORDS)) {
    if (!detection[category]) detection[category] = { keywords: [...keywords], dbCategory: category };
  }
  for (const [keyword, category] of Object.entries(db.JUICE_ALIASES)) {
    if (category && detection[category]) detection[category].keywords.push(keyword);
  }
  return detection;
}

// Helper function to parse numbers from OCR
// Handles: commas, periods as thousand separators (OCR misreads), newlines/whitespace
function parseOCRNumber(str) {
  if (!str) return 0;
  let s = str.toString().trim();

  // Remove whitespace and newlines
  s = s.replace(/[\s]/g, '');

  // Remove periods and commas (they are always misread thousand separators, noise, or formatting)
  s = s.replace(/[\.,]/g, '');

  return parseFloat(s) || 0;
}

// Locate the ตะกร้า (crate) columns from the header rows. Anchoring on the crate
// headers — rather than a fixed offset — keeps the layout correct no matter how many
// juice columns are added or removed, even a juice we have no detection keyword for.
// Headers can be split across the first few rows (merged cells), so accumulate per column.
function detectCrateColumns(table) {
  const headerByCol = new Map();
  for (let rowIdx = 0; rowIdx < Math.min(4, table.length); rowIdx++) {
    const row = table[rowIdx];
    if (!row) continue;
    for (let colIdx = 0; colIdx < row.length; colIdx++) {
      const cellText = row[colIdx] ? row[colIdx].toString() : '';
      if (!cellText) continue;
      headerByCol.set(colIdx, (headerByCol.get(colIdx) || '') + ' ' + cellText);
    }
  }

  let fullCratesCol = null, partialCol = null, totalCratesCol = null, firstCrateCol = null;
  for (const [colIdx, text] of headerByCol) {
    if (!text.includes('ตะกร้า')) continue;                       // product headers never say ตะกร้า
    if (firstCrateCol === null || colIdx < firstCrateCol) firstCrateCol = colIdx;
    if (fullCratesCol === null && text.includes('เต็ม')) fullCratesCol = colIdx;
    if (partialCol === null && text.includes('เศษ')) partialCol = colIdx;
    if (totalCratesCol === null && text.includes('รวม')) totalCratesCol = colIdx;
  }

  // Crate columns always run เต็ม → เศษ → รวม in order. If OCR garbled a sub-label,
  // fall back to that fixed order counting from the first ตะกร้า column.
  if (firstCrateCol !== null && (fullCratesCol === null || partialCol === null || totalCratesCol === null)) {
    fullCratesCol = firstCrateCol;
    partialCol = firstCrateCol + 1;
    totalCratesCol = firstCrateCol + 2;
  }

  return { firstCrateCol, fullCratesCol, partialCol, totalCratesCol };
}

// Derive product/crate column positions for a plan. The layout is
// [vendor][warehouse][juice...][ตะกร้าเต็ม][ตะกร้าเศษ][ตะกร้ารวม][รถ]... where the number
// of juice columns varies (2 = orange+pineapple, 3 = +pop, etc). Prefer anchoring the
// crate columns on their headers so add/remove of any juice — even an unknown one — keeps
// the crate positions (and the digit-recovery that depends on them) correct. Falls back to
// "right of the last detected juice column" when the crate headers can't be read.
function deriveLayout(productColumns, table = null) {
  const detected = (productColumns && productColumns.length) ? [...productColumns] : [2, 3, 4];

  const crates = table ? detectCrateColumns(table) : null;
  if (crates && crates.fullCratesCol != null && crates.partialCol != null && crates.totalCratesCol != null) {
    // Every column between warehouse (C1) and the first crate column holds a juice count,
    // so this stays correct even for a juice type we have no keyword for.
    const productCols = [];
    for (let c = 2; c < crates.firstCrateCol; c++) productCols.push(c);
    return {
      productCols: productCols.length ? productCols : detected,
      fullCratesCol: crates.fullCratesCol,
      partialCol: crates.partialCol,
      totalCratesCol: crates.totalCratesCol
    };
  }

  // Fallback: no crate headers found — assume crates sit right of the last juice column.
  const lastProductCol = Math.max(...detected);
  return {
    productCols: detected,
    fullCratesCol: lastProductCol + 1,
    partialCol: lastProductCol + 2,
    totalCratesCol: lastProductCol + 3
  };
}

// Get product value from a row, with fallback to crate-based reconstruction if the value is misread (non-numeric)
function getRowProductValue(row, columnIndex, rowIndex = -1, table = null, totalsRowIndex = -1, yodruamTotal = 0, layout = null) {
  if (!row) return 0;

  const lay = layout || deriveLayout(null);
  const rawValText = row[columnIndex] ? row[columnIndex].toString().trim() : '';
  const fullCrates = row[lay.fullCratesCol] ? parseOCRNumber(row[lay.fullCratesCol]) : 0;
  const partialBottles = row[lay.partialCol] ? parseOCRNumber(row[lay.partialCol]) : 0;
  const reconstructed = (fullCrates * 35) + partialBottles;

  if (rawValText !== '') {
    let value = parseOCRNumber(rawValText);
    if (value === 0 && reconstructed > 0) {
      // Check if this column is the active product column in this row (to avoid double-counting other products)
      let isOtherProductNonEmpty = false;
      for (const c of lay.productCols) {
        if (c === columnIndex) continue;
        const otherVal = row[c] ? row[c].toString().trim() : '';
        if (otherVal !== '' && otherVal !== '0') {
          isOtherProductNonEmpty = true;
          break;
        }
      }

      if (!isOtherProductNonEmpty) {
        console.log(`[OCR-FIX] Row ${rowIndex}: Reconstructed value using crates: ${reconstructed} (Full: ${fullCrates}, Partial: ${partialBottles}) because original was "${rawValText}"`);
        value = reconstructed;
      }
    }
    return value;
  }

  // If the cell is completely empty, check if we can infer it using grand total difference.
  // These plans normally have NO totals row inside the table (the ยอดรวม row is OCR'd from
  // raw page text), so totalsRowIndex is -1 in the common case — we must still run this.
  // The inner loop already excludes row `totalsRowIndex` correctly (-1 matches no row).
  if (reconstructed > 0 && table && yodruamTotal > 0) {
    // Check if other product columns in this row are empty/0
    let isOtherProductNonEmpty = false;
    for (const c of lay.productCols) {
      if (c === columnIndex) continue;
      const otherVal = row[c] ? row[c].toString().trim() : '';
      if (otherVal !== '' && otherVal !== '0') {
        isOtherProductNonEmpty = true;
        break;
      }
    }

    if (!isOtherProductNonEmpty) {
      // Calculate sum of all other rows (excluding this row and totals row)
      let sumOfOtherRows = 0;
      for (let i = 0; i < table.length; i++) {
        if (i === totalsRowIndex || i === rowIndex || !table[i]) continue;
        const rValText = table[i][columnIndex] ? table[i][columnIndex].toString().trim() : '';
        if (rValText !== '') {
          // Pass null for table to prevent infinite recursion
          sumOfOtherRows += getRowProductValue(table[i], columnIndex, i, null, -1, 0, lay);
        }
      }

      const diff = yodruamTotal - sumOfOtherRows;
      if (diff === reconstructed) {
        console.log(`[OCR-FIX] Row ${rowIndex}: Assigned empty cell to col ${columnIndex} because grand total diff (${diff}) matches reconstructed value (${reconstructed})`);
        return reconstructed;
      }
    }
  }

  return 0;
}

// CDC names to track
const CDC_NAMES = [
  'บางบัวทอง',
  'นครราชสีมา',
  'นครสวรรค์',
  'ชลบุรี',
  'มหาชัย',
  'สุวรรณภูมิ',
  'หาดใหญ่',
  'ภูเก็ต',
  'เชียงใหม่',
  'สุราษฎร์',
  'ขอนแก่น'
];

// The คลังส่งสินค้า (C1) that ships nearly every vendor on the plan. validateTableData
// uses it as the "is this sheet actually filled in?" probe.
const BANGBUATHONG_WAREHOUSE = 'บางบัวทอง';

// Map short names to full names for consistency
const CDC_NAME_MAPPING = {
  'บางบัวทอง': 'คลังบางบัวทอง',
  'นครราชสีมา': 'นครราชสีมา',
  'นครสวรรค์': 'นครสวรรค์',
  'ชลบุรี': 'ชลบุรี',
  'มหาชัย': 'คลังมหาชัย',
  'สุวรรณภูมิ': 'คลังสุวรรณภูมิ',
  'หาดใหญ่': 'หาดใหญ่',
  'ภูเก็ต': 'ภูเก็ต',
  'เชียงใหม่': 'เชียงใหม่',
  'สุราษฎร์': 'สุราษฎร์',
  'ขอนแก่น': 'ขอนแก่น'
};

// Detect product columns from table headers using longest-keyword-match scoring.
// For each cell, find the product with the longest matching keyword. Then each column
// is "won" by its highest-scoring product. Each product is assigned to its best-won column.
// This resolves substring collisions like "น้ำส้ม" ⊂ "น้ำส้มมุกป๊อป" — pop's "มุกป๊อป" (7)
// beats orange's "น้ำส้ม" (6), so pop wins that column.
function detectProductColumns(table, detection = getProductDetection()) {
  console.log('[DETECT] Scanning table for product columns (longest-match scoring)...');

  // Stage 1: collect best (productKey, score) per (rowIdx, colIdx) cell
  const cellMatches = [];
  for (let rowIdx = 0; rowIdx < Math.min(10, table.length); rowIdx++) {
    const row = table[rowIdx];
    if (!row) continue;

    for (let colIdx = 0; colIdx < row.length; colIdx++) {
      const cellRaw = row[colIdx];
      const cellText = cellRaw ? cellRaw.toString().toLowerCase().trim() : '';
      if (!cellText) continue;

      let best = null;
      for (const [productKey, productConfig] of Object.entries(detection)) {
        let longest = 0;
        for (const keyword of productConfig.keywords) {
          const kw = keyword.toLowerCase();
          if (cellText.includes(kw) && kw.length > longest) {
            longest = kw.length;
          }
        }
        if (longest > 0 && (!best || longest > best.score)) {
          best = { productKey, score: longest, dbCategory: productConfig.dbCategory };
        }
      }

      if (best) {
        cellMatches.push({
          rowIdx, colIdx,
          productKey: best.productKey,
          score: best.score,
          dbCategory: best.dbCategory,
          cellText: cellRaw
        });
        console.log(`[DETECT] Row ${rowIdx} Col ${colIdx} "${cellRaw}" → ${best.productKey} (score ${best.score})`);
      }
    }
  }

  // Stage 2: for each column, the highest-scoring product wins
  const colWinner = new Map(); // colIdx → match
  for (const m of cellMatches) {
    const existing = colWinner.get(m.colIdx);
    if (!existing || m.score > existing.score) {
      colWinner.set(m.colIdx, m);
    }
  }

  // Stage 3: each product is assigned to the best column it won
  const productBest = {}; // productKey → match
  for (const m of colWinner.values()) {
    const existing = productBest[m.productKey];
    if (!existing || m.score > existing.score) {
      productBest[m.productKey] = m;
    }
  }

  const detectedProducts = {};
  for (const [productKey, m] of Object.entries(productBest)) {
    detectedProducts[productKey] = {
      column: m.colIdx,
      dbCategory: m.dbCategory,
      foundIn: `Row ${m.rowIdx}, Col ${m.colIdx}: "${m.cellText}"`
    };
    console.log(`[DETECT] Assigned ${productKey} → column ${m.colIdx} (Row ${m.rowIdx}: "${m.cellText}", score ${m.score})`);
  }

  // Fallback: recover orange when Azure OCR garbles "ส้ม" (observed: "ẫu", "ลม", etc.).
  // Plan ส่งสินค้า has a fixed column order — orange sits immediately left of pop/pineapple/watermelon,
  // and every product column shares the "จำนวน น้ำ ___ (ขวด)" header pattern.
  detectOrangeByPosition(table, detectedProducts);

  console.log(`[DETECT] Detected ${Object.keys(detectedProducts).length} products:`, Object.keys(detectedProducts));
  return detectedProducts;
}

// Fallback when keyword matching misses orange because Azure OCR misread "ส้ม".
// Looks for an unclaimed column to the left of detected products whose header rows
// match the canonical product-column shape ("จำนวน"/"น้ำ" + "(ขวด)").
function detectOrangeByPosition(table, detectedProducts) {
  if (detectedProducts.orange) return;

  const claimedCols = new Set(Object.values(detectedProducts).map(p => p.column));
  if (claimedCols.size === 0) return;

  const minClaimedCol = Math.min(...claimedCols);

  // Accumulate header text per column from the first few rows
  const headerByCol = new Map();
  for (let rowIdx = 0; rowIdx < Math.min(3, table.length); rowIdx++) {
    const row = table[rowIdx];
    if (!row) continue;
    for (let colIdx = 0; colIdx < row.length; colIdx++) {
      if (claimedCols.has(colIdx)) continue;
      if (colIdx >= minClaimedCol) continue; // must sit left of detected products
      if (colIdx < 2) continue;              // C0/C1 are vendor/warehouse, never products
      const cellText = row[colIdx] ? row[colIdx].toString() : '';
      if (!cellText) continue;
      headerByCol.set(colIdx, (headerByCol.get(colIdx) || '') + ' ' + cellText);
    }
  }

  // Pick the rightmost candidate (closest to detected products = most likely the missing product)
  let chosenCol = null;
  for (const [colIdx, text] of headerByCol) {
    const hasKhuat = text.includes('ขวด');
    const hasProductPrefix = text.includes('จำนวน') || text.includes('น้ำ');
    if (hasKhuat && hasProductPrefix) {
      if (chosenCol === null || colIdx > chosenCol) chosenCol = colIdx;
    }
  }

  if (chosenCol !== null) {
    detectedProducts.orange = {
      column: chosenCol,
      dbCategory: 'orange',
      foundIn: `Position fallback col ${chosenCol} (OCR likely misread "ส้ม"; header: "${headerByCol.get(chosenCol).trim().substring(0, 40)}")`
    };
    console.log(`[DETECT] Fallback: assigned orange → column ${chosenCol} (OCR garbled header, position-inferred)`);
  }
}

// Find juice columns that carry bottles but matched no PRODUCT_DETECTION keyword.
// layout.productCols spans every column between the warehouse and the crate columns, so a
// juice the sheet adds before we add its keyword shows up here as an unclaimed column —
// the recording loop only walks detectedProducts, so without this its bottles are dropped
// silently (which is how น้ำแตงโม went unrecorded until it got a keyword).
// Bottles are summed straight from the cells, without crate reconstruction: that fills a
// blank cell from its row's crates, which here would just re-count another juice's row.
function detectUnknownJuiceColumns(table, detectedProducts, layout, yodruamTotals = {}) {
  const claimedCols = new Set(Object.values(detectedProducts).map(p => p.column));
  const unknown = [];

  for (const colIdx of layout.productCols) {
    if (claimedCols.has(colIdx)) continue;

    // Header text comes only from the rows above the first vendor row, so a totals row
    // inside the table can't leak its number into the column name.
    const headerParts = [];
    let bottles = 0;
    let inBody = false;
    for (const row of table) {
      if (!row) continue;
      const c0 = row[0] ? row[0].toString() : '';
      const cell = row[colIdx] ? row[colIdx].toString().trim() : '';
      if (/FC\d+/.test(c0)) {
        inBody = true;
        bottles += parseOCRNumber(cell);
      } else if (!inBody && cell && !headerParts.includes(cell)) {
        headerParts.push(cell);
      }
    }

    // ยอดรวม is the figure on the sheet, so prefer it over our own cell sum when OCR found it
    const total = yodruamTotals[colIdx] || bottles;
    if (total > 0) {
      unknown.push({ column: colIdx, header: headerParts.join(' ') || `C${colIdx}`, total });
      console.warn(`[DETECT] Unknown juice column ${colIdx} "${headerParts.join(' ')}" carries ${total} bottles — not recorded`);
    }
  }

  return unknown;
}

function rgbToHsl(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return { h: 0, s: 0, l };
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? ((g - b) / d + (g < b ? 6 : 0)) / 6
    : max === g ? ((b - r) / d + 2) / 6
      : ((r - g) / d + 4) / 6;
  return { h, s, l };
}

function hslToHex(h, s, l) {
  const f = n => {
    const k = (n + h * 12) % 12;
    const a = s * Math.min(l, 1 - l);
    const c = l - a * Math.max(-1, Math.min(k - 3, 9 - k, 1));
    return Math.round(c * 255).toString(16).padStart(2, '0');
  };
  return `#${f(0)}${f(8)}${f(4)}`;
}

// A very light version of a colour, for the section background behind a juice's table.
function tintColor(hex, amount = 0.9) {
  const n = parseInt(hex.slice(1), 16);
  const mix = c => Math.round(c + (255 - c) * amount).toString(16).padStart(2, '0');
  return `#${mix(n >> 16)}${mix((n >> 8) & 255)}${mix(n & 255)}`;
}

// Background colour of a juice column, read off the plan image, so the dashboard shows a
// juice added automatically in the colour its column has on the sheet. Only the cells Azure
// placed in this column are sampled, inset from their edges, and the most common saturated
// colour wins — cell text, grid lines and JPEG noise don't decide it. Too light a colour is
// darkened (same hue) because the dashboard puts white text on it and uses it as text on white.
// null — the caller falls back to a default — for anything but a JPEG, or a column with no
// real colour (white/grey/black), which is what a column OCR mis-assigned would look like.
function sampleColumnColor(imageBuffer, rawResult, columnIndex) {
  try {
    if (!imageBuffer || imageBuffer[0] !== 0xFF || imageBuffer[1] !== 0xD8) return null;
    const table = rawResult?.tables?.[0];
    const page = rawResult?.pages?.[0];
    if (!table || !page) return null;

    const img = jpeg.decode(imageBuffer, { useTArray: true, maxMemoryUsageInMB: 256 });
    // Polygons come in the page's unit; for an image that is pixels, but scale anyway.
    const sx = img.width / (page.width || img.width);
    const sy = img.height / (page.height || img.height);

    const bins = new Map();   // colour quantised to 4 bits/channel → summed pixels
    for (const cell of table.cells) {
      if (cell.columnIndex !== columnIndex) continue;
      const poly = cell.boundingRegions?.[0]?.polygon;
      if (!poly || poly.length < 4) continue;
      const xs = poly.map(p => p.x * sx);
      const ys = poly.map(p => p.y * sy);
      const ix = (Math.max(...xs) - Math.min(...xs)) * 0.15;
      const iy = (Math.max(...ys) - Math.min(...ys)) * 0.15;
      const x0 = Math.max(0, Math.round(Math.min(...xs) + ix));
      const x1 = Math.min(img.width - 1, Math.round(Math.max(...xs) - ix));
      const y0 = Math.max(0, Math.round(Math.min(...ys) + iy));
      const y1 = Math.min(img.height - 1, Math.round(Math.max(...ys) - iy));
      for (let y = y0; y <= y1; y += 2) {
        for (let x = x0; x <= x1; x += 2) {
          const i = (y * img.width + x) * 4;
          const r = img.data[i], g = img.data[i + 1], b = img.data[i + 2];
          const { s, l } = rgbToHsl(r, g, b);
          if (s < 0.25 || l > 0.92 || l < 0.12) continue;
          const key = ((r >> 4) << 8) | ((g >> 4) << 4) | (b >> 4);
          const bin = bins.get(key) || { n: 0, r: 0, g: 0, b: 0 };
          bin.n++; bin.r += r; bin.g += g; bin.b += b;
          bins.set(key, bin);
        }
      }
    }

    let best = null;
    for (const bin of bins.values()) if (!best || bin.n > best.n) best = bin;
    if (!best || best.n < 50) return null;
    const { h, s, l } = rgbToHsl(best.r / best.n, best.g / best.n, best.b / best.n);
    return hslToHex(h, s, Math.min(l, 0.55));
  } catch (error) {
    console.warn(`[JUICE] Could not read the colour of column ${columnIndex}:`, error.message);
    return null;
  }
}

// "จำนวน น้ำ มะม่วง (ขวด)" → "มะม่วง": drop the words every juice header shares, leaving the
// part that names this juice — the keyword future plans will be matched on. Only a leading
// "น้ำ" is stripped, so a name like "มะพร้าวน้ำหอม" keeps the น้ำ its header actually shows.
function juiceKeywordFromHeader(header) {
  return String(header || '')
    .replace(/\(.*?\)/g, ' ')
    .split(/\s+/)
    .filter(t => t && !['จำนวน', 'น้ำ', 'ขวด'].includes(t))
    .join(' ')
    .replace(/^น้ำ\s*/, '')
    .replace(/[<>&"'`]/g, '')
    .trim();
}

// Header text of the non-juice columns. A juice keyword contained in any of these would
// make detectProductColumns claim a crate/vendor/warehouse column as that juice.
const NON_JUICE_HEADER_WORDS = ['vendor', 'คลังส่งสินค้า', 'ตะกร้าเต็ม', 'ตะกร้าเศษ', 'ตะกร้ารวม',
  'จำนวนรถ', 'รายชื่อผู้ส่งสินค้า', 'เวลาโหลดสินค้า', 'เวลาจัดส่งสินค้า', 'ยอดรวม', ...CDC_NAMES];

// The juice (built-in or auto-added) that already answers to this exact keyword.
function categoryForKeyword(keyword) {
  const kw = String(keyword || '').toLowerCase();
  for (const [category, config] of Object.entries(getProductDetection())) {
    if (config.keywords.some(k => k.toLowerCase() === kw)) return category;
  }
  return null;
}

// Why a keyword can't name a juice, or null when it can. selfCategory is the juice being
// renamed, whose own keywords don't count as a clash.
function juiceKeywordProblem(keyword, selfCategory = null) {
  if (!keyword || keyword.length < 2) return 'อ่านชื่อน้ำจากหัวคอลัมน์ไม่ได้';
  if (keyword.length > 30) return `ชื่อ "${keyword}" ยาวเกินไป`;
  const existing = categoryForKeyword(keyword);
  if (existing && existing !== selfCategory) {
    return `"${keyword}" เป็นชื่อของ ${db.CATEGORY_NAMES[existing] || existing} อยู่แล้ว`;
  }
  const kw = keyword.toLowerCase();
  if (NON_JUICE_HEADER_WORDS.some(w => w.toLowerCase().includes(kw))) {
    return `"${keyword}" ซ้ำกับหัวคอลัมน์อื่นในตาราง`;
  }
  return null;
}

// Validate table has data (check FC33 หาดใหญ่ or FC07 ภูเก็ต in any detected product column)
function validateTableData(table, detectedProducts) {
  // Get the first detected product's column for validation
  const productColumns = Object.values(detectedProducts).map(p => p.column);
  if (productColumns.length === 0) {
    return { valid: false, reason: 'No product columns detected' };
  }

  let hadyaiSum = 0;
  let phuketSum = 0;
  const layout = deriveLayout(productColumns, table);

  // An unfinished plan sheet shows up as the บางบัวทอง dispatch block sitting at 0 for
  // every juice — that block is the bulk of the sheet and is never legitimately empty on
  // a real shipping day. A single juice reading 0 while another juice on the same rows
  // carries real counts is normal (some days ship no น้ำสับปะรด at all), so only an
  // all-juice zero counts as "not filled in yet".
  // layout.productCols spans every column between the warehouse and the crate columns, so
  // a juice added to the sheet later is covered without touching this code.
  let bbtRows = 0;
  const bbtSums = new Map(layout.productCols.map(c => [c, 0]));

  // Check all detected product columns for validation
  table.forEach((row, rowIndex) => {
    if (!row || row.length < 3) return;

    const c0Cell = row[0] ? row[0].toString().trim() : '';
    const isVendorRow = /FC\d+/.test(c0Cell);
    const warehouse = row[1] ? row[1].toString() : '';

    // Sum values from all product columns for validation
    productColumns.forEach(colIdx => {
      const value = getRowProductValue(row, colIdx, rowIndex, null, -1, 0, layout);

      if (c0Cell.includes('FC33') && c0Cell.includes('หาดใหญ่')) {
        hadyaiSum += value;
      }
      if (c0Cell.includes('FC07')) {
        phuketSum += value;
      }
    });

    if (isVendorRow && warehouse.includes(BANGBUATHONG_WAREHOUSE)) {
      bbtRows++;
      for (const colIdx of layout.productCols) {
        bbtSums.set(colIdx, bbtSums.get(colIdx) + getRowProductValue(row, colIdx, rowIndex, null, -1, 0, layout));
      }
    }
  });

  console.log(`[VALIDATION] FC33 หาดใหญ่ sum: ${hadyaiSum}, FC07 ภูเก็ต sum: ${phuketSum}`);

  // Skipped when no บางบัวทอง row matched at all: that means OCR garbled the warehouse
  // name, not that the block is empty, and rejecting on it would throw away a good plan.
  if (bbtRows > 0) {
    const perCol = [...bbtSums.entries()].map(([c, v]) => `col${c}=${v}`).join(', ');
    console.log(`[VALIDATION] คลังส่งสินค้าบางบัวทอง (${bbtRows} vendor rows): ${perCol}`);
    if ([...bbtSums.values()].every(v => v === 0)) {
      return {
        valid: false,
        reason: `คลังส่งสินค้าบางบัวทอง is 0 across all ${bbtSums.size} juice column(s) over ${bbtRows} vendor row(s) — plan not filled in yet.`,
        hadyaiSum,
        phuketSum
      };
    }
  }

  if (hadyaiSum === 0 && phuketSum === 0) {
    return { valid: false, reason: 'Both FC33 หาดใหญ่ and FC07 ภูเก็ต sums are 0 (validation failed)', hadyaiSum, phuketSum };
  }

  return { valid: true, hadyaiSum, phuketSum };
}

// Extract date from table or text
function extractDate(table, extractedText) {
  let dateStr = null;

  // First, try to find date in the extracted text
  if (extractedText) {
    console.log('[DATE] Searching for date in extracted text...');
    const textDateMatch = extractedText.match(/(?:วันที่\s*)?(\d{1,2}\/\d{1,2}\/\d{4})/);
    if (textDateMatch) {
      dateStr = textDateMatch[1];
      console.log(`[DATE] Found date in extracted text: ${dateStr}`);
      return dateStr;
    }
  }

  // Search in table
  console.log('[DATE] Searching for date in table...');
  for (let i = 0; i < table.length; i++) {
    const row = table[i];
    if (!row) continue;

    for (let col = 0; col < row.length; col++) {
      const cellText = row[col] ? row[col].toString() : '';
      const dateMatch = cellText.match(/(?:วันที่\s*)?(\d{1,2}\/\d{1,2}\/\d{4})/);
      if (dateMatch) {
        dateStr = dateMatch[1];
        console.log(`[DATE] Found date: ${dateStr} in row ${i}, column ${col}`);
        return dateStr;
      }
    }
  }

  return null;
}

// Extract CDC totals for a specific product column
function extractCDCTotals(table, columnIndex, yodruamTotal = 0, layout = null) {
  const lay = layout || deriveLayout(null);
  const cdcTotals = {};
  let totalSum = 0;
  let totalsRowIndex = -1;

  // Initialize all CDC totals to 0
  Object.values(CDC_NAME_MAPPING).forEach(cdcFullName => {
    cdcTotals[cdcFullName] = 0;
  });

  // Find the total sum from the table
  for (let i = 0; i < table.length; i++) {
    const row = table[i];
    if (!row || row.length < 2) continue;

    const c0 = row[0] ? row[0].toString().trim() : '';
    const c1 = row[1] ? row[1].toString().trim() : '';

    const totalIndicators = ['ยอดรวม', 'รวม', 'total', 'grand total', 'sum'];
    const foundTotal = totalIndicators.some(indicator =>
      c0.toLowerCase().includes(indicator.toLowerCase()) || c1.toLowerCase().includes(indicator.toLowerCase())
    );

    if (foundTotal && row[columnIndex]) {
      totalSum = parseOCRNumber(row[columnIndex]);
      totalsRowIndex = i;
      console.log(`[TOTAL] Found total sum: ${totalSum} in row ${i}`);
      break;
    }
  }

  // If no total row found by indicators, check the last row with multiple large values
  if (totalSum === 0) {
    for (let i = table.length - 1; i >= Math.max(0, table.length - 5); i--) {
      const row = table[i];
      if (!row) continue;

      let largeValueCount = 0;
      for (const col of lay.productCols) {
        if (row[col]) {
          const val = parseOCRNumber(row[col]);
          if (val > 100) largeValueCount++;
        }
      }

      if (largeValueCount >= 2) {
        totalSum = parseOCRNumber(row[columnIndex]);
        totalsRowIndex = i;
        console.log(`[TOTAL] Detected totals row by pattern at row ${i}, total: ${totalSum}`);
        break;
      }
    }
  }

  // Fallback: calculate by summing (excluding detected totals row)
  if (totalSum === 0) {
    for (let i = 0; i < table.length; i++) {
      if (i === totalsRowIndex) continue;
      const row = table[i];
      if (!row || !row[columnIndex]) continue;
      const value = getRowProductValue(row, columnIndex, i, null, -1, 0, lay);
      if (value > 0) totalSum += value;
    }
  }

  console.log(`[CDC] Extracting CDC values, excluding totals row index: ${totalsRowIndex}`);

  // Extract CDC values row by row, tracking individual contributions for ratio correction
  const CRATE_TOTAL_COL = lay.totalCratesCol; // ตะกร้า รวม (total crates) — right of the last product column
  const cdcRows = [];
  const unknownVendors = [];

  table.forEach((row, rowIndex) => {
    if (!row || row.length < 2) return;
    if (rowIndex === totalsRowIndex) return;

    // Each row maps to exactly one CDC, derived from the Vendor cell (C0).
    // C1 (source warehouse) is unreliable: fill-down can carry a previous vendor's
    // warehouse name into the next vendor's first row, and a single row's value
    // ends up double-counted across two CDCs.
    const c0 = row[0] ? row[0].toString() : '';
    if (!c0) return;

    let bestCdcName = null;
    let bestScore = 0;
    for (const cdcName of CDC_NAMES) {
      if (c0.includes(cdcName) && cdcName.length > bestScore) {
        bestCdcName = cdcName;
        bestScore = cdcName.length;
      }
    }
    if (!bestCdcName) {
      // CDC_NAMES is a fixed list of the eleven branches, and daily_records has a column
      // per branch — so a row we cannot place has nowhere to go and is dropped. Dropping
      // it silently is the problem: a branch added to the plan, or one whose Thai name
      // OCR mangled, quietly subtracts its bottles from the day. Only vendor rows are
      // worth reporting; header and spacer rows fail this match too and mean nothing.
      if (/FC\d+/.test(c0)) {
        const label = c0.replace(/\s+/g, ' ').trim();
        if (!unknownVendors.includes(label)) {
          unknownVendors.push(label);
          console.warn(`[CDC] Unrecognized branch in row ${rowIndex}: "${label}" — its bottles are not counted in any CDC`);
        }
      }
      return;
    }

    const fullCdcName = CDC_NAME_MAPPING[bestCdcName];
    const value = getRowProductValue(row, columnIndex, rowIndex, table, totalsRowIndex, yodruamTotal, lay);
    const crateTotal = row[CRATE_TOTAL_COL] ? parseOCRNumber(row[CRATE_TOTAL_COL]) : 0;
    if (value > 0) {
      cdcRows.push({ rowIndex, cdcName: fullCdcName, value, crateTotal });
    }
  });

  // Ratio correction: fix truncated values where bottle/crate ratio is suspiciously low
  // Normal range: ~25-42 bottles per crate. Below 20 indicates a missing digit.
  // Only correct when crates >= 2 to avoid false positives on genuinely small orders
  const rawSum = cdcRows.reduce((s, r) => s + r.value, 0);
  const targetTotal = yodruamTotal > 0 ? yodruamTotal : totalSum;

  if (rawSum !== targetTotal && targetTotal > 0) {
    const corrected = []; // entries we adjusted, with their in-range alternatives

    for (const entry of cdcRows) {
      if (entry.crateTotal < 2) continue;
      const ratio = entry.value / entry.crateTotal;
      if (ratio >= 20 || entry.value <= 1) continue;

      // Value appears truncated (a digit was dropped) - try appending digit 0-9
      // and keep EVERY candidate whose bottle/crate ratio is plausible. Several
      // digits usually qualify (crates=2 admits 50..59), so the ratio alone can't
      // pin the exact digit - that disambiguation happens against ยอดรวม below.
      const candidates = [];
      for (let d = 0; d <= 9; d++) {
        const candidate = entry.value * 10 + d;
        const candRatio = candidate / entry.crateTotal;
        if (candRatio >= 25 && candRatio <= 42) candidates.push(candidate);
      }
      if (candidates.length === 0) continue;

      const original = entry.value;
      entry.value = candidates[0]; // provisional: smallest in-range value
      corrected.push({ entry, candidates, original });
      console.log(`[RATIO-FIX] Row ${entry.rowIndex} (${entry.cdcName}): ${original} → ${entry.value} (ratio ${(original / entry.crateTotal).toFixed(1)} → ${(entry.value / entry.crateTotal).toFixed(1)}, crates=${entry.crateTotal}, candidates=[${candidates.join(',')}])`);
    }

    // Pin the appended digit using the ยอดรวม grand total: re-pick each corrected
    // entry's candidate to drive the corrected sum toward ยอดรวม. Without this, the
    // first in-range digit wins (e.g. 5 → 50) even when another candidate (51) makes
    // the column total reconcile exactly.
    for (const c of corrected) {
      let diff = targetTotal - cdcRows.reduce((s, r) => s + r.value, 0);
      if (diff === 0) break;
      let best = c.entry.value;
      let bestDiff = Math.abs(diff);
      for (const cand of c.candidates) {
        const newDiff = Math.abs(diff - (cand - c.entry.value));
        if (newDiff < bestDiff) { bestDiff = newDiff; best = cand; }
      }
      if (best !== c.entry.value) {
        console.log(`[RATIO-FIX] Grand-total adjust Row ${c.entry.rowIndex} (${c.entry.cdcName}): ${c.entry.value} → ${best} (ratio ${(best / c.entry.crateTotal).toFixed(1)})`);
        c.entry.value = best;
      }
    }

    const correctedSum = cdcRows.reduce((s, r) => s + r.value, 0);

    // The whole premise of this pass is that a digit was dropped, so a correct fix always
    // moves the column *toward* ยอดรวม. Landing further away means the ratio was wrong for
    // some other reason — a merged crate cell, a misread crate count — and appending a
    // digit turned a good value into one 10x too large. Throw the pass away rather than
    // record that.
    if (Math.abs(targetTotal - correctedSum) > Math.abs(targetTotal - rawSum)) {
      for (const c of corrected) c.entry.value = c.original;
      console.log(`[RATIO-FIX] Reverted: ${rawSum} → ${correctedSum} moved away from ยอดรวม=${targetTotal}; keeping the raw values`);
    } else if (correctedSum !== rawSum) {
      console.log(`[RATIO-FIX] Corrected: ${rawSum} → ${correctedSum} (ยอดรวม=${targetTotal}, remaining diff=${targetTotal - correctedSum})`);
    }
  }

  // Build CDC totals from (possibly corrected) row values
  for (const entry of cdcRows) {
    cdcTotals[entry.cdcName] += entry.value;
  }

  return { cdcTotals, totalSum, unknownVendors };
}

// Extract Khon Kaen Laos value for orange category
function extractKhonKaenLaos(table, columnIndex, productColumnIndices) {
  // 1. Search for a row that explicitly contains both 'ขอนแก่น' and 'ลาว'/'laos'
  for (let i = 0; i < table.length; i++) {
    const row = table[i];
    if (!row || !row[0]) continue;

    const c0 = row[0].toString().toLowerCase().trim();
    if (c0.includes('ขอนแก่น') && (c0.includes('ลาว') || c0.includes('laos') || c0.includes('lao')) && row[columnIndex]) {
      const value = parseOCRNumber(row[columnIndex]);
      console.log(`[LAOS] Found explicit ขอนแก่น Laos row at index ${i}: "${row[0]}" with value ${value}`);
      return value;
    }
  }

  // 2. Group Khon Kaen rows into logical slots
  const khonKaenRows = [];
  table.forEach((row, idx) => {
    if (!row || !row[0]) return;
    const c0 = row[0].toString().trim();
    if (c0.includes('ขอนแก่น')) {
      khonKaenRows.push({ rowIndex: idx, row });
    }
  });

  const slots = [];
  let currentSlot = [];
  let lastColIndex = -1;

  for (const entry of khonKaenRows) {
    const row = entry.row;
    const rowIndex = entry.rowIndex;

    let rowProductCols = [];
    for (const colIdx of productColumnIndices) {
      if (row[colIdx] !== undefined && row[colIdx] !== null && row[colIdx] !== '') {
        const val = parseOCRNumber(row[colIdx]);
        if (val > 0) {
          rowProductCols.push(colIdx);
        }
      }
    }

    if (rowProductCols.length === 0) continue;

    const minColIdx = Math.min(...rowProductCols);
    if (minColIdx <= lastColIndex) {
      slots.push(currentSlot);
      currentSlot = [];
    }

    currentSlot.push({ rowIndex, row, rowProductCols });
    lastColIndex = Math.max(...rowProductCols);
  }

  if (currentSlot.length > 0) {
    slots.push(currentSlot);
  }

  console.log(`[LAOS] Total logical slots found for Khon Kaen: ${slots.length}`);
  slots.forEach((s, idx) => {
    console.log(`  Slot ${idx + 1}:`);
    s.forEach(e => {
      const productVals = productColumnIndices.map(col => `${col}: "${e.row[col] || ''}"`).join(', ');
      console.log(`    Row ${e.rowIndex}: [${productVals}]`);
    });
  });

  // If there are exactly 3 slots, the last one is Laos
  if (slots.length === 3) {
    const laosSlot = slots[2];
    for (const entry of laosSlot) {
      if (entry.row[columnIndex]) {
        const val = parseOCRNumber(entry.row[columnIndex]);
        if (val > 0) {
          return val;
        }
      }
    }
  }

  return 0;
}

// Extract ยอดรวม (totals) from raw OCR page text, mapped to table column indices
// The ยอดรวม row is often outside the table structure but always present in raw text
function extractYodruamTotals(rawResult) {
  if (!rawResult || !rawResult.pages) return {};

  for (const page of rawResult.pages) {
    if (!page.lines) continue;

    // Find the ยอดรวม line
    let yodruamY = null;
    for (const line of page.lines) {
      if (line.content.includes('ยอดรวม') && line.polygon) {
        yodruamY = line.polygon[0].y;
        break;
      }
    }

    if (yodruamY === null) continue;

    // Find all numeric values on the same y-level
    const nearLines = page.lines.filter(l =>
      l.polygon && Math.abs(l.polygon[0].y - yodruamY) < 15 &&
      /[\d,.]/.test(l.content) && !l.content.includes('ยอดรวม')
    );

    // Sort by x-position (left to right)
    nearLines.sort((a, b) => a.polygon[0].x - b.polygon[0].x);

    // The ยอดรวม numbers correspond to columns in order starting from C2
    // C0=Vendor, C1=Warehouse are text columns; C2+ are numeric data columns
    const FIRST_NUMERIC_COLUMN = 2;
    const yodruamByColumn = {};
    nearLines.forEach((line, idx) => {
      const col = FIRST_NUMERIC_COLUMN + idx;
      const value = parseOCRNumber(line.content);
      yodruamByColumn[col] = value;
      console.log(`[ยอดรวม] Column ${col}: ${value} (raw="${line.content}", x=${line.polygon[0].x.toFixed(0)})`);
    });

    return yodruamByColumn;
  }

  return {};
}

// Record daily data - dynamically detects products from column headers
async function recordDailyData(tableData, extractedText = '', rawResult = null, imageBuffer = null) {
  if (!tableData || tableData.length === 0) {
    console.log('No table data to record');
    return { success: false, reason: 'No table data found' };
  }

  const table = tableData[0];

  // Auto-added juices live in the DB; another process may have added or changed one since boot.
  await refreshCategories();

  // Step 1: Detect product columns from headers
  const detectedProducts = detectProductColumns(table);

  if (Object.keys(detectedProducts).length === 0) {
    console.log('[RECORD] No products detected in table headers');
    return { success: false, reason: 'No product columns detected in table headers' };
  }

  // Step 2: Validate table has data
  const validation = validateTableData(table, detectedProducts);
  if (!validation.valid) {
    console.log(`[VALIDATION] Failed: ${validation.reason}`);
    return { success: false, reason: validation.reason };
  }

  console.log(`[VALIDATION] Passed - FC33 หาดใหญ่: ${validation.hadyaiSum}, FC07 ภูเก็ต: ${validation.phuketSum}`);

  // Step 3: Extract date
  const dateStr = extractDate(table, extractedText);
  if (!dateStr) {
    console.log('[DATE] No date found');
    return { success: false, reason: 'No date found in table or extracted text' };
  }

  console.log(`[DATE] Using date: ${dateStr}`);

  // Step 4: Extract ยอดรวม (totals) from raw OCR page text
  // The ยอดรวม row is often outside the table structure but reliably OCR'd from raw text
  const yodruamTotals = rawResult ? extractYodruamTotals(rawResult) : {};
  if (Object.keys(yodruamTotals).length > 0) {
    console.log(`[ยอดรวม] Extracted totals from raw page text:`, yodruamTotals);
  } else {
    console.log(`[ยอดรวม] No totals found in raw page text`);
  }

  // Step 5: Extract data for each detected product
  const results = [];
  const reconcileIssues = [];
  const unknownBranches = new Set();

  // Crate columns are located from the table headers, so a plan with any number of juice
  // columns (added or removed) still finds ตะกร้า counts in the right place.
  const layout = deriveLayout(Object.values(detectedProducts).map(p => p.column), table);

  for (const [productKey, productInfo] of Object.entries(detectedProducts)) {
    const category = productInfo.dbCategory;
    const columnIndex = productInfo.column;

    // Check if already recorded
    const alreadyRecorded = await isDateRecorded(dateStr, category);
    if (alreadyRecorded) {
      console.log(`Date ${dateStr} already recorded for ${category}`);
      continue;
    }

    console.log(`[RECORD] Extracting ${category} from column ${columnIndex}`);

    // Extract CDC totals, passing ยอดรวม for ratio-based correction of truncated values
    const yodruamValue = yodruamTotals[columnIndex] || 0;
    const { cdcTotals, totalSum, unknownVendors } = extractCDCTotals(table, columnIndex, yodruamValue, layout);
    for (const v of unknownVendors) unknownBranches.add(v);

    // Use ยอดรวม from raw text as authoritative total when available
    let finalTotalSum = totalSum;
    if (yodruamValue && yodruamValue > 0) {
      if (yodruamValue !== totalSum) {
        console.log(`[ยอดรวม] ${category}: Cell sum=${totalSum} differs from ยอดรวม=${yodruamValue} (diff=${yodruamValue - totalSum}). Using ยอดรวม as authoritative total.`);
      } else {
        console.log(`[ยอดรวม] ${category}: Cell sum matches ยอดรวม=${yodruamValue} ✓`);
      }
      finalTotalSum = yodruamValue;
    }

    // Extract Khon Kaen Laos for orange
    const productColumnIndices = Object.values(detectedProducts).map(p => p.column).sort((a, b) => a - b);
    const khonKaenLaosValue = category === 'orange' ? extractKhonKaenLaos(table, columnIndex, productColumnIndices) : 0;

    // Reconcile against the ยอดรวม printed on the plan. That figure is an independent
    // checksum the sheet hands us for free, and every per-CDC breakdown that came out
    // right has matched it exactly (30/30 columns across 11 archived plans). A mismatch
    // means the extraction went wrong somewhere — a merged cell read as one number, a
    // ratio "fix" that appended a digit, a branch whose name OCR mangled — and without
    // this check that lands in the dashboard looking like a normal day.
    // The record is still saved: a number that is probably right beats a missing day,
    // and the alert goes to a person who can compare it against the image.
    const cdcSum = Object.values(cdcTotals).reduce((a, b) => a + b, 0);
    if (finalTotalSum > 0 && cdcSum !== finalTotalSum) {
      const diff = cdcSum - finalTotalSum;
      console.warn(`[RECONCILE] ${category}: per-CDC sum ${cdcSum} != ยอดรวม ${finalTotalSum} (diff ${diff > 0 ? '+' : ''}${diff})`);
      reconcileIssues.push({ category, cdcSum, expected: finalTotalSum, diff, cdcTotals });
    } else {
      console.log(`[RECONCILE] ${category}: per-CDC sum ${cdcSum} matches ยอดรวม ✓`);
    }

    // Build record
    const dailyRecord = {
      date: dateStr,
      timestamp: new Date().toISOString(),
      fc33HadyaiSum: validation.hadyaiSum,
      totalSum: finalTotalSum,
      cdcTotals: cdcTotals,
      khonKaenLaos: category === 'orange' ? khonKaenLaosValue : undefined,
      khonKaenCambodia: category === 'orange' ? 0 : undefined
    };

    // Save to MySQL
    await saveDailyRecord(dailyRecord, category);
    console.log(`Recorded daily data for ${dateStr} in ${category} category (column ${columnIndex})`);
    results.push({ category, record: dailyRecord });
  }

  // Only checked on a run that recorded something — once per plan per day. A re-send of an
  // already-recorded plan records nothing, so it does not repeat the alert.
  const unknownJuices = results.length > 0
    ? detectUnknownJuiceColumns(table, detectedProducts, layout, yodruamTotals)
    : [];

  // A juice column whose name reads cleanly is added and recorded right away — translated,
  // coloured from the image — and the dev is only told, with the juice page link to fix
  // anything. One whose name can't be used is kept unrecorded for the dev to name there:
  // auto-adding it would leave nothing to match tomorrow's plan, so every day would add
  // another juice. A name deleted on the juice page (alias → null) is skipped silently.
  const juiceLines = [];
  let juicesWaiting = 0;
  for (const j of unknownJuices) {
    const keyword = juiceKeywordFromHeader(j.header);
    if (keyword && db.JUICE_ALIASES[keyword] === null) {
      console.log(`[JUICE] Skipping column ${j.column} "${keyword}" — deleted on the juice page`);
      continue;
    }
    const problem = juiceKeywordProblem(keyword);
    try {
      const englishName = problem ? null : await translateJuiceName(`น้ำ${keyword}`);
      const color = sampleColumnColor(imageBuffer, rawResult, j.column);
      const pending = await db.savePendingJuice({
        date: dateStr,
        header: j.header,
        keyword,
        englishName,
        color,
        total: j.total,
        context: { table, layout, column: j.column, yodruamTotal: yodruamTotals[j.column] || 0, hadyaiSum: validation.hadyaiSum }
      });
      if (!pending || pending.status !== 'pending') continue;   // this plan was handled before

      if (problem) {
        juicesWaiting++;
        juiceLines.push(`  ⚠️ "${j.header}" ${j.total.toLocaleString()} ขวด — ${problem} ยังไม่ถูกบันทึก`);
        continue;
      }
      const added = await addJuiceFromPending(pending, {}, 'auto');
      if (added.error) {
        juicesWaiting++;
        juiceLines.push(`  ⚠️ "${j.header}" ${j.total.toLocaleString()} ขวด — ${added.error} ยังไม่ถูกบันทึก`);
      } else {
        juiceLines.push(`  🆕 เพิ่ม "${added.thaiName}"${added.englishName ? ` (${added.englishName})` : ''} อัตโนมัติ\n`
          + added.lines.map(l => `     ${l}`).join('\n'));
      }
    } catch (error) {
      console.error('[JUICE] Could not add unknown juice:', error);
      juicesWaiting++;
      juiceLines.push(`  ⚠️ "${j.header}" ${j.total.toLocaleString()} ขวด — เพิ่มอัตโนมัติไม่สำเร็จ (${error.message})`);
    }
  }

  // Raised after saving, and deliberately not awaited into the success/failure result:
  // the day is recorded either way, this only tells a human to go look at the image.
  // An unplaceable branch is worth its own alert even when the totals happen to add up,
  // because the fix is a code change (add it to CDC_NAMES and daily_records), not a
  // re-scan of the image.
  if (reconcileIssues.length > 0 || unknownBranches.size > 0 || juiceLines.length > 0) {
    const sections = [];

    if (juiceLines.length > 0) {
      sections.push(
        `[น้ำชนิดใหม่]\n`
        + juiceLines.join('\n')
        + `\n  ตรวจหรือแก้ชื่อ สี${juicesWaiting ? ' และตั้งชื่อน้ำที่ยังไม่ถูกบันทึก' : ''}:\n  ${juicePageUrl()}`
      );
    }

    if (unknownBranches.size > 0) {
      sections.push(
        `[สาขาที่ระบบไม่รู้จัก]\n`
        + [...unknownBranches].map(v => `  ${v}`).join('\n')
        + `\n  ← ขวดของแถวนี้ไม่ถูกนับในสาขาใดเลย ต้องเพิ่มสาขาในโค้ด`
      );
    }

    for (const i of reconcileIssues) {
      const breakdown = Object.entries(i.cdcTotals)
        .filter(([, v]) => v > 0)
        .map(([name, v]) => `  ${name}: ${v.toLocaleString()}`)
        .join('\n');
      sections.push(
        `[${db.CATEGORY_NAMES[i.category] || i.category}]\n`
        + `รวมรายสาขา: ${i.cdcSum.toLocaleString()}\n`
        + `ยอดรวมในภาพ: ${i.expected.toLocaleString()}\n`
        + `ต่างกัน: ${i.diff > 0 ? '+' : ''}${i.diff.toLocaleString()}\n`
        + breakdown
      );
    }

    const headline = reconcileIssues.length > 0
      ? '⚠️ OCR ไม่ตรงยอดรวม'
      : unknownBranches.size > 0
        ? '⚠️ OCR เจอสาขาที่ไม่รู้จัก'
        : juicesWaiting > 0
          ? '⚠️ พบน้ำชนิดใหม่ที่อ่านชื่อไม่ได้'
          : '🆕 พบน้ำชนิดใหม่';

    const note = (reconcileIssues.length > 0 || unknownBranches.size > 0)
      ? 'บันทึกลงฐานข้อมูลแล้ว แต่ตัวเลขอาจผิด กรุณาเทียบกับภาพต้นฉบับ'
      : juicesWaiting > 0
        ? 'น้ำที่อ่านชื่อได้บันทึกแล้ว แต่ยังมีน้ำที่ต้องตั้งชื่อก่อนจึงจะบันทึก'
        : 'เพิ่มและบันทึกให้อัตโนมัติแล้ว ถ้าชื่อหรือสีผิด แก้ได้ที่ลิงก์ด้านล่าง';

    await sendDevAlert(
      `${headline} — ${dateStr}\n\n`
      + `${note}\n\n`
      + sections.join('\n\n')
      + `\n\nhttps://shinsen.yushi-marketing.com/daily-report`
    );
  }

  if (results.length === 0) {
    return { success: false, reason: 'All detected products already recorded for this date' };
  }

  return { success: true, date: dateStr, results: results };
}

// Turn Azure's flat cell list into row arrays. Split out of performOCR so the merged-cell
// handling can be exercised offline against a saved Azure response.
function buildTablesFromResult(result) {
  const tableData = [];
  if (!result || !result.tables) return tableData;

  for (const table of result.tables) {
    const tableRows = [];
    const maxRow = Math.max(...table.cells.map(c => c.rowIndex)) + 1;
    const maxCol = Math.max(...table.cells.map(c => c.columnIndex)) + 1;

    for (let i = 0; i < maxRow; i++) {
      tableRows[i] = new Array(maxCol).fill('');
    }

    for (const cell of table.cells) {
      const content = cell.content || '';
      const rowIdx = cell.rowIndex;
      const colIdx = cell.columnIndex;
      const rowSpan = cell.rowSpan || 1;

      // Azure occasionally merges adjacent Vendor cells (e.g., FC03 + FC15) into
      // a single multi-row cell. Split the content by FC code so fill-down later
      // gives each row exactly one vendor instead of leaving both names in every row.
      if (colIdx === 0 && rowSpan > 1) {
        const fcParts = content
          .split(/(?=FC\d+)/)
          .map(p => p.trim())
          .filter(p => /^FC\d+/.test(p));
        if (fcParts.length > 1) {
          const K = fcParts.length;
          for (let i = 0; i < rowSpan; i++) {
            const fcIdx = Math.min(K - 1, Math.floor((i * K) / rowSpan));
            tableRows[rowIdx + i][colIdx] = fcParts[fcIdx];
          }
          console.log(`[PREPROCESS] Split merged Vendor cell at row ${rowIdx} into ${K} FCs across ${rowSpan} rows: [${fcParts.join(' | ')}]`);
          continue;
        }
      }

      // Azure does the same to a stack of *numeric* cells, joining the values with a
      // space — ตะกร้ารวม 29 above 3 arrives as one rowSpan-2 cell reading "29 3".
      // Left whole, parseOCRNumber strips the space and reads 293, the bottle/crate
      // ratio then looks impossible, and the ratio-fix below "restores" a digit that
      // was never dropped (1,012 → 10,120). Hand each spanned row its own value.
      //
      // Each spanned row contributes at most one line to the cell, so N numbers means
      // N of those rows had a value and concatenating them is always wrong. One number
      // fewer than rowSpan just means a row was blank; the values still belong to the
      // rows from the top down, and even a wrong guess about which row beats reading
      // 29 and 3 as 293. More numbers than rows is the opposite situation — one value
      // that OCR split on a space ("8 566") — so that one is left alone to be rejoined.
      if (colIdx > 0 && rowSpan > 1) {
        const numParts = content.trim().split(/\s+/).filter(Boolean);
        const allNumeric = numParts.length > 1 && numParts.every(p => /^[\d.,]+$/.test(p));
        if (allNumeric && numParts.length <= rowSpan) {
          for (let i = 0; i < numParts.length && tableRows[rowIdx + i]; i++) {
            tableRows[rowIdx + i][colIdx] = numParts[i];
          }
          const how = numParts.length === rowSpan ? 'across' : `over the first ${numParts.length} of`;
          console.log(`[PREPROCESS] Split merged numeric cell at row ${rowIdx} col ${colIdx} ${how} ${rowSpan} rows: [${numParts.join(' | ')}]`);
          continue;
        }
        if (allNumeric) {
          console.warn(`[PREPROCESS] Merged cell at row ${rowIdx} col ${colIdx} holds ${numParts.length} numbers over only ${rowSpan} rows ("${content}") — leaving it for parseOCRNumber to rejoin`);
        }
      }

      tableRows[rowIdx][colIdx] = content;
    }

    tableData.push(tableRows);
  }

  return tableData;
}

// Preprocess table data: fill down C0 and C1 values to handle merged cells
function preprocessTableData(tableData) {
  console.log('[PREPROCESS] Filling down C0 and C1 values...');

  const processedTables = tableData.map(table => {
    if (!table || table.length === 0) return table;

    // Create a copy of the table
    const processedTable = table.map(row => [...row]);

    // Fill down column 0 (C0)
    let lastC0Value = null;
    for (let i = 0; i < processedTable.length; i++) {
      if (processedTable[i] && processedTable[i][0]) {
        const value = processedTable[i][0].toString().trim();
        if (value) {
          lastC0Value = value;
        }
      } else if (lastC0Value) {
        // Fill down the last non-empty value
        if (!processedTable[i]) processedTable[i] = [];
        processedTable[i][0] = lastC0Value;
        console.log(`[PREPROCESS] Row ${i} C0: filled with "${lastC0Value}"`);
      }
    }

    // Fill down column 1 (C1)
    let lastC1Value = null;
    for (let i = 0; i < processedTable.length; i++) {
      if (processedTable[i] && processedTable[i][1]) {
        const value = processedTable[i][1].toString().trim();
        if (value) {
          lastC1Value = value;
        }
      } else if (lastC1Value) {
        // Fill down the last non-empty value
        if (!processedTable[i]) processedTable[i] = [];
        processedTable[i][1] = lastC1Value;
        console.log(`[PREPROCESS] Row ${i} C1: filled with "${lastC1Value}"`);
      }
    }

    // Log all rows to verify table structure and identify the total row
    console.log(`[PREPROCESS] Table has ${processedTable.length} rows after preprocessing`);
    console.log('[PREPROCESS] All rows (first 4 columns):');
    processedTable.forEach((row, idx) => {
      const c0 = row[0] ? row[0].toString().substring(0, 30) : '';
      const c1 = row[1] ? row[1].toString().substring(0, 30) : '';
      const c2 = row[2] ? row[2].toString().substring(0, 15) : '';
      const c3 = row[3] ? row[3].toString().substring(0, 15) : '';
      console.log(`[PREPROCESS] Row ${idx}: C0="${c0}" | C1="${c1}" | C2="${c2}" | C3="${c3}"`);
    });

    return processedTable;
  });

  console.log('[PREPROCESS] Fill down completed');
  return processedTables;
}

// Detect category (orange or yuzu) from table data
function detectCategory(table) {
  // Check headers or specific cells to determine category
  // This is a placeholder - adjust based on actual data structure
  if (table[0]) {
    const headerText = table[0].join(' ').toLowerCase();
    if (headerText.includes('yuzu')) return 'yuzu';
    if (headerText.includes('orange')) return 'orange';
  }

  // Default to orange if can't determine
  return 'orange';
}

const app = express();

// Logging middleware
app.use((req, res, next) => {
  // The juice page's token is its only protection — keep it out of the logs.
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.path.replace(/^\/juices\/[^/]+/, '/juices/***')}`);
  next();
});

// LINE webhook - must be before express.json() to preserve raw body for signature validation
app.post('/webhook', line.middleware(config), async (req, res) => {
  const events = req.body.events || [];
  console.log('[WEBHOOK] Received webhook request');
  console.log('[WEBHOOK] Number of events:', events.length);
  if (req.body.destination) console.log('[WEBHOOK] Destination:', req.body.destination);

  // Acknowledge before processing. A full image run (Vision + Azure OCR) takes ~15s,
  // far past LINE's webhook timeout — answering only after it finished made LINE
  // record the delivery as failed, and with redelivery enabled it would resend the
  // same event while the first run was still going. handleEvent is idempotent per
  // message ID, so a redelivery that does slip through is a no-op.
  //
  // Answering early is not free, though: Passenger reaps this process once no
  // request is in flight, which killed image runs mid-OCR and lost the screenshot
  // outright. So every image event is persisted to pending_events *before* the
  // response — an INSERT fast enough to always beat the teardown — and the run
  // itself becomes retryable work rather than something that must survive.
  const queued = [];
  const inline = [];
  for (const event of events) {
    const messageId = event.type === 'message' && event.message?.type === 'image'
      ? event.message.id
      : null;
    if (!messageId) {
      inline.push(event);
      continue;
    }
    try {
      const isNew = await db.enqueueEvent(messageId, event);
      console.log(`[QUEUE] ${isNew ? 'Queued' : 'Already known'} image event ${messageId}`);
      queued.push(messageId);
    } catch (error) {
      // A queue write failure must not drop the image — fall back to the old
      // best-effort path, which at least works when the process is not reaped.
      console.error('[QUEUE] Enqueue failed, processing inline instead:', error.message);
      inline.push(event);
    }
  }

  res.json({ success: true, queued: queued.length });

  // Best effort: when the process does survive, the image is done within seconds of
  // being posted. When it does not, the queue row is still there for /process-queue.
  const work = inline.map(event =>
    handleEvent(event).catch(async err => {
      console.error('[ERROR] Unhandled error in handleEvent:', err);
      // Only reachable for events the queue never took (non-image events, or an
      // image whose enqueue failed). There is no retry behind this path, so the
      // failure has to be written down here or it is gone.
      if (event.message?.type === 'image') {
        await saveDetectionLog({
          timestamp: new Date().toISOString(),
          messageId: event.message.id,
          groupId: event.source?.groupId || null,
          userId: event.source?.userId || null,
          status: 'error',
          reason: `Processing failed (not queued): ${err?.message || err}`
        }).catch(logError => console.error('[ERROR] Could not write error log:', logError.message));
      }
      return null;
    })
  );
  if (queued.length) {
    work.push(drainQueue({ reason: 'webhook' }).catch(err => {
      console.error('[QUEUE] Webhook drain failed:', err.message);
      return null;
    }));
  }
  Promise.all(work).then(() => console.log('[WEBHOOK] Finished processing', events.length, 'event(s)'));
});

// Add JSON parsing middleware for other endpoints (after LINE webhook)
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

async function handleEvent(event) {
  console.log('[EVENT] Event type:', event.type);
  if (event.deliveryContext?.isRedelivery) {
    console.log('[EVENT] This is a REDELIVERY from LINE');
  }

  // Log source information (group ID, user ID, etc.)
  const sourceInfo = {
    type: event.source?.type,
    groupId: event.source?.groupId || null,
    userId: event.source?.userId || null
  };

  if (event.source) {
    console.log('[EVENT] Source type:', event.source.type);
    if (event.source.groupId) {
      console.log('[EVENT] Group ID:', event.source.groupId);
    }
    if (event.source.userId) {
      console.log('[EVENT] User ID:', event.source.userId);
    }
  }

  if (event.type !== 'message') {
    console.log('[EVENT] Skipping non-message event');
    return null;
  }

  const message = event.message;
  console.log('[EVENT] Message type:', message.type);

  if (message.type === 'image') {
    console.log('[IMAGE] Image message detected:', message.id);

    // Every finished run writes a detection log keyed by message ID, so an existing
    // log means this image was already handled — skip it. Without this, LINE's
    // webhook redelivery would re-run OCR and re-notify for the same screenshot.
    try {
      if (await db.hasProcessedMessage(message.id)) {
        console.log('[IMAGE] Message already processed, skipping (redelivery):', message.id);
        return null;
      }
    } catch (dedupeError) {
      // A dedupe lookup failure must not block processing — worst case is a duplicate.
      console.error('[IMAGE] Dedupe check failed, processing anyway:', dedupeError.message);
    }

    try {
      console.log('[IMAGE] Fetching image content...');
      const imageBuffer = await getImageContent(message.id);
      console.log('[IMAGE] Image size:', imageBuffer.length, 'bytes');

      console.log('[IMAGE] Detecting if Excel screenshot...');
      const isExcelScreenshot = await detectExcelScreenshot(imageBuffer);
      console.log('[IMAGE] Is Excel screenshot:', isExcelScreenshot);

      if (isExcelScreenshot) {
        console.log('[OCR] Excel screenshot detected, performing OCR...');
        const { extractedText, recordResult } = await performOCR(imageBuffer, message.id, sourceInfo);
        console.log('[OCR] Extraction completed, text length:', extractedText.length);

        // Only send reply if data was successfully recorded
        if (recordResult && recordResult.success) {
          const replyMessage = `Report for ${recordResult.date} has been recorded\n\nDaily report: https://shinsen.yushi-marketing.com/daily-report\nMTD / YTD report: https://shinsen.yushi-marketing.com/mtd-report`;

          // The reply is best-effort. A redelivered event carries a reply token that
          // has already expired, and letting that throw would skip the group
          // notification below for data that was, in fact, recorded successfully.
          try {
            await client.replyMessage({
              replyToken: event.replyToken,
              messages: [{
                type: 'text',
                text: replyMessage,
              }],
            });
            console.log('Success message sent to user');
          } catch (replyError) {
            console.warn('[REPLY] Could not reply (expired token on redelivery?):', replyError.message);
          }

          // Send notifications to configured groups (if AUTO_NOTIFY is enabled)
          if (AUTO_NOTIFY) {
            const categories = recordResult.results.map(r => r.category);
            await sendNotificationToGroups(recordResult.date, categories);
          } else {
            console.log('[NOTIFICATION] Auto-notify disabled, skipping group notification');
          }
        } else {
          console.log('Data not recorded, no reply sent');
        }
      } else {
        // Log failed detection - not an Excel screenshot
        await saveDetectionLog({
          timestamp: new Date().toISOString(),
          messageId: message.id,
          groupId: sourceInfo.groupId,
          userId: sourceInfo.userId,
          status: 'failed',
          reason: 'Not an Excel screenshot (failed GPT-4 Vision detection)'
        });
        console.log('Image not detected as Excel screenshot, no reply sent');
      }

      console.log('Detection result:', { isExcelScreenshot });
    } catch (error) {
      // Silent mode towards the user, but the error must reach the caller: the
      // queue can only retry what it is told failed. Swallowing it here would mark
      // a LINE blob-download blip as finished work. The detection log for a run
      // that never reaches a verdict is written once, by whoever gives up on it.
      console.error('Error processing image:', error);
      throw error;
    }
  } else {
    console.log('Non-image message received:', message.type);
  }

  return null;
}

// ---------- New juices (auto-added from plans) ----------

// English label for a new juice, shown where the daily report uses English names for
// built-in juices (summary cards, chart legend, table title). A slow or failed call only
// costs the label: it gives up after 5s without retrying and returns null, and the report
// falls back to the Thai name.
async function translateJuiceName(thaiName) {
  try {
    const response = await openai.chat.completions.create({
      model: 'gpt-4o',
      max_tokens: 20,
      temperature: 0,
      messages: [{
        role: 'user',
        content: 'Translate this Thai juice product name into a short English product name: '
          + 'the fruit or flavour only, Title Case, at most 3 words, without the word "Juice", '
          + 'no quotes or punctuation. Examples: น้ำส้ม → Orange, น้ำผลไม้รวม → Mixed Fruit.\n'
          + `Thai: ${thaiName}`
      }]
    }, { timeout: 5000, maxRetries: 0 });   // measured 1.5–2.5s; 3s cut too many off
    return cleanEnglishName(response.choices?.[0]?.message?.content);
  } catch (error) {
    console.warn(`[JUICE] Could not translate "${thaiName}":`, error.message);
    return null;
  }
}

// Letters, digits and simple joiners only — the label goes into report HTML and chart JS.
// Thai text is refused rather than stripped, so a name typed in Thai is reported, not mangled.
function cleanEnglishName(text) {
  const raw = String(text || '');
  if (/[฀-๿]/.test(raw)) return null;
  const english = raw.replace(/[^A-Za-z0-9 &'-]/g, '').replace(/\s+/g, ' ').trim().slice(0, 40);
  return english || null;
}

// Add a waiting juice column as a new juice and record it — plus any other plan still
// waiting on the same name — from the OCR table kept with it. Used for the automatic add
// (by 'auto') and for a name given on the juice page (by 'web'); thaiName / englishName
// override what was read off the header and the translation made from it.
// Returns { thaiName, englishName, lines } or { error }.
async function addJuiceFromPending(pending, { thaiName: givenName = null, englishName: givenEnglish = null, icon = null } = {}, by = 'auto') {
  await refreshCategories();
  const keyword = givenName != null ? juiceKeywordFromHeader(givenName) : pending.keyword;

  // A given name that is already a juice means "OCR misread this header — it's that one".
  const existing = givenName != null ? categoryForKeyword(keyword) : null;
  if (existing) return recordPendingAsExisting(pending, existing, by);

  const problem = juiceKeywordProblem(keyword);
  if (problem) return { error: problem };

  const category = `juice_${pending.id}`;
  const thaiName = `น้ำ${keyword}`;
  // When the name was corrected, keep matching on what OCR read too — the next plan will
  // most likely be read the same way.
  const keywords = [keyword];
  if (pending.keyword !== keyword && !juiceKeywordProblem(pending.keyword)) keywords.push(pending.keyword);

  if (!(await db.resolvePendingJuice(pending.id, 'added', category, by))) {
    return { error: 'รายการนี้ถูกจัดการไปแล้ว' };
  }
  const englishName = givenEnglish
    || (keyword === pending.keyword && pending.englishName)
    || await translateJuiceName(thaiName);
  try {
    await db.addCategory(category, { thaiName, englishName, color: pending.color, icon, keywords }, by);
  } catch (error) {
    await db.reopenPendingJuice(pending.id);
    throw error;
  }
  console.log(`[JUICE] Added ${category} "${thaiName}" / ${englishName || '(no English name)'} / ${pending.color || 'default colour'} (keywords: ${keywords.join(', ')}) by ${by}`);

  return { thaiName, englishName, lines: await recordPendingFamily(pending, category, by) };
}

// Record a waiting column under a juice that already exists, and teach that juice the
// spelling OCR read, so the next plan read the same way matches on its own.
async function recordPendingAsExisting(pending, category, by) {
  if (!(await db.resolvePendingJuice(pending.id, 'added', category, by))) {
    return { error: 'รายการนี้ถูกจัดการไปแล้ว' };
  }
  if (!juiceKeywordProblem(pending.keyword)) await db.addJuiceAlias(pending.keyword, category);
  console.log(`[JUICE] Pending #${pending.id} "${pending.keyword}" recorded as existing ${category} by ${by}`);
  return {
    thaiName: db.CATEGORY_NAMES[category] || category,
    englishName: db.CATEGORY_ENGLISH_NAMES[category] || null,
    existing: true,
    lines: await recordPendingFamily(pending, category, by)
  };
}

// This plan's column, plus other plans still waiting on the same name. Columns whose name
// couldn't be read share an empty keyword without being the same juice, so those are
// only ever recorded one at a time.
async function recordPendingFamily(pending, category, by) {
  const lines = [await recordPendingJuice(pending, category)];
  if (!pending.keyword) return lines;
  for (const other of await db.getPendingJuicesByKeyword(pending.keyword, pending.id)) {
    if (await db.resolvePendingJuice(other.id, 'added', category, by)) {
      lines.push(await recordPendingJuice(other, category));
    }
  }
  return lines;
}

// Record a waiting column the way recordDailyData records a known juice: per-CDC from the
// table, with the ยอดรวม printed on the plan as the authoritative total.
// Never throws — the juice is already added, so one bad day must not hide the others.
async function recordPendingJuice(pending, category) {
  const { table, layout, column, yodruamTotal, hadyaiSum } = pending.context;
  try {
    if (await isDateRecorded(pending.date, category)) return `${pending.date}: มีข้อมูลอยู่แล้ว`;
    const { cdcTotals, totalSum } = extractCDCTotals(table, column, yodruamTotal, layout);
    const finalTotalSum = yodruamTotal > 0 ? yodruamTotal : totalSum;
    await saveDailyRecord({
      date: pending.date,
      timestamp: new Date().toISOString(),
      fc33HadyaiSum: hadyaiSum,
      totalSum: finalTotalSum,
      cdcTotals
    }, category);
    const cdcSum = Object.values(cdcTotals).reduce((a, b) => a + b, 0);
    const check = cdcSum === finalTotalSum
      ? '✓ รายสาขาตรงยอดรวม'
      : `⚠️ รายสาขารวม ${cdcSum.toLocaleString()} ไม่ตรงยอดรวม กรุณาเทียบกับภาพ`;
    console.log(`[JUICE] Recorded ${category} for ${pending.date}: ${finalTotalSum}`);
    return `${pending.date}: บันทึก ${finalTotalSum.toLocaleString()} ขวด ${check}`;
  } catch (error) {
    console.error(`[JUICE] Could not record ${category} for ${pending.date}:`, error);
    return `${pending.date}: ❌ บันทึกไม่สำเร็จ (${error.message})`;
  }
}

// An image event gets this many tries before it is buried as 'failed'. Generous
// because most retries exist to recover from the process being reaped, not from
// anything wrong with the image itself.
const QUEUE_MAX_ATTEMPTS = 5;

// One drain at a time per process — the webhook and a cron hit can easily overlap,
// and both claiming at once would just fight over the same rows.
let drainInFlight = null;

// Work the queue. Callers that need the work to actually finish (the cron endpoint)
// must await this *inside* their request: an in-flight request is the only thing
// that stops Passenger from reaping the process mid-OCR.
async function drainQueue({ reason = 'manual', limit = 3 } = {}) {
  if (drainInFlight) {
    console.log(`[QUEUE] Drain already in flight, joining it (${reason})`);
    return drainInFlight;
  }

  drainInFlight = (async () => {
    const claimToken = `${process.pid}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    const claimed = await db.claimPendingEvents(claimToken, { limit, maxAttempts: QUEUE_MAX_ATTEMPTS });

    if (!claimed.length) {
      console.log(`[QUEUE] Nothing to drain (${reason})`);
      return { reason, claimed: 0, processed: 0, retrying: 0, failed: 0 };
    }

    console.log(`[QUEUE] Draining ${claimed.length} event(s) (${reason})`);
    const summary = { reason, claimed: claimed.length, processed: 0, retrying: 0, failed: 0 };

    for (const item of claimed) {
      try {
        await handleEvent(item.event);
        await db.completeEvent(item.messageId);
        summary.processed++;
        console.log(`[QUEUE] Done: ${item.messageId} (attempt ${item.attempts})`);
      } catch (error) {
        const row = await db.releaseEvent(item.messageId, error.message, QUEUE_MAX_ATTEMPTS);
        if (row?.status === 'failed') {
          summary.failed++;
          console.error(`[QUEUE] Giving up on ${item.messageId} after ${row.attempts} attempts: ${error.message}`);
          // Bury it where the operators actually look. Without this, an image that
          // exhausts its retries would go quiet again — the exact failure mode the
          // queue exists to end.
          await saveDetectionLog({
            timestamp: new Date().toISOString(),
            messageId: item.messageId,
            groupId: item.event?.source?.groupId || null,
            userId: item.event?.source?.userId || null,
            status: 'error',
            reason: `Gave up after ${row.attempts} attempts: ${error.message}`
          }).catch(logError => console.error('[QUEUE] Could not log give-up:', logError.message));
        } else {
          summary.retrying++;
          console.error(`[QUEUE] Retry queued for ${item.messageId} (attempt ${item.attempts}): ${error.message}`);
        }
      }
    }

    return summary;
  })().finally(() => { drainInFlight = null; });

  return drainInFlight;
}

async function getImageContent(messageId) {
  try {
    const stream = await blobClient.getMessageContent(messageId);
    const chunks = [];

    for await (const chunk of stream) {
      chunks.push(chunk);
    }

    return Buffer.concat(chunks);
  } catch (error) {
    console.error('Error getting image content:', error);
    throw error;
  }
}

async function detectExcelScreenshot(imageBuffer) {
  try {
    const base64Image = imageBuffer.toString('base64');

    const response = await openai.chat.completions.create({
      model: 'gpt-4o',
      max_tokens: 100,
      messages: [
        {
          role: 'user',
          content: [
            {
              type: 'image_url',
              image_url: {
                url: `data:image/jpeg;base64,${base64Image}`,
              },
            },
            {
              type: 'text',
              text: 'Does this image contain a spreadsheet with a table grid layout? Look for: cells arranged in rows and columns, gridlines, tabular data structure, column/row organization, or any spreadsheet-like table format (including Excel, Google Sheets, printed spreadsheets, or any tabular data displays). Answer with only "YES" if you see a spreadsheet/table grid, or "NO" if you do not.',
            },
          ],
        },
      ],
    });

    const answer = response.choices[0].message.content.trim().toUpperCase();
    console.log('GPT-4 Vision response:', answer);

    return answer === 'YES';
  } catch (error) {
    console.error('Error in Excel detection:', error);
    throw error;
  }
}

async function performOCR(imageBuffer, messageId = 'unknown', sourceInfo = null) {
  try {
    console.log('Starting Azure OCR...');

    const poller = await azureClient.beginAnalyzeDocument('prebuilt-layout', imageBuffer);
    const result = await poller.pollUntilDone();

    let extractedText = '';
    let tableData = [];

    if (result.pages) {
      for (const page of result.pages) {
        if (page.lines) {
          for (const line of page.lines) {
            extractedText += line.content + '\n';
          }
        }
      }
    }

    tableData = buildTablesFromResult(result);

    // Preprocess table data: fill down C0 and C1 to help with extraction
    if (tableData && tableData.length > 0) {
      tableData = preprocessTableData(tableData);
    }

    latestOCRResult = {
      timestamp: new Date(),
      extractedText: extractedText,
      tableData: tableData,
      rawResult: result,
      messageId: messageId
    };

    console.log('OCR completed. Text length:', extractedText.length);
    console.log('Tables found:', tableData.length);

    // Record daily data if conditions are met
    let recordResult = null;
    try {
      recordResult = await recordDailyData(tableData, extractedText, result, imageBuffer);
      if (recordResult && recordResult.success) {
        console.log('Daily data recorded successfully');

        // Save the valid image to 'stored_images' folder on the server
        try {
          const imagesDir = path.join(__dirname, 'stored_images');
          await fs.mkdir(imagesDir, { recursive: true });

          // recordResult.date is "DD/MM/YYYY" format
          const formattedDate = recordResult.date.replace(/\//g, '-');
          const fileName = `${formattedDate}.jpg`;
          const filePath = path.join(imagesDir, fileName);
          await fs.writeFile(filePath, imageBuffer);
          console.log(`[IMAGE] Saved successful screenshot to ${filePath}`);
        } catch (imgError) {
          console.error('[IMAGE] Error saving successful screenshot:', imgError);
        }

        // Log successful extraction
        await saveDetectionLog({
          timestamp: new Date().toISOString(),
          messageId: latestOCRResult.messageId || 'unknown',
          groupId: sourceInfo?.groupId || null,
          userId: sourceInfo?.userId || null,
          status: 'success',
          date: recordResult.date,
          categories: recordResult.results.map(r => r.category),
          recordsCreated: recordResult.results.length
        });
      } else if (recordResult && !recordResult.success) {
        console.log('Daily data not recorded:', recordResult.reason);
        // Log failed extraction with reason
        await saveDetectionLog({
          timestamp: new Date().toISOString(),
          messageId: latestOCRResult.messageId || 'unknown',
          groupId: sourceInfo?.groupId || null,
          userId: sourceInfo?.userId || null,
          status: 'failed',
          reason: recordResult.reason
        });
      }
    } catch (error) {
      console.error('Error recording daily data:', error);
      // Log error
      await saveDetectionLog({
        timestamp: new Date().toISOString(),
        messageId: latestOCRResult.messageId || 'unknown',
        groupId: sourceInfo?.groupId || null,
        userId: sourceInfo?.userId || null,
        status: 'error',
        reason: `Error: ${error.message}`
      });
    }

    return { extractedText, recordResult };
  } catch (error) {
    console.error('Error in OCR:', error);
    throw error;
  }
}

function transformTableData(tableData, columnIndex = 2) {
  if (!tableData || tableData.length === 0) {
    return null;
  }

  // Define CDC order (this will be the display order)
  const cdcOrder = [
    'คลังบางบัวทอง',
    'นครราชสีมา',
    'นครสวรรค์',
    'ชลบุรี',
    'คลังมหาชัย',
    'คลังสุวรรณภูมิ',
    'หาดใหญ่',
    'ภูเก็ต',
    'เชียงใหม่',
    'สุราษฎร์',
    'ขอนแก่น'
  ];

  // CDC names to search for (short names)
  const cdcNames = [
    'บางบัวทอง',
    'นครราชสีมา',
    'นครสวรรค์',
    'ชลบุรี',
    'มหาชัย',
    'สุวรรณภูมิ',
    'หาดใหญ่',
    'ภูเก็ต',
    'เชียงใหม่',
    'สุราษฎร์',
    'ขอนแก่น'
  ];

  // Map short names to full names
  const cdcNameMapping = {
    'บางบัวทอง': 'คลังบางบัวทอง',
    'นครราชสีมา': 'นครราชสีมา',
    'นครสวรรค์': 'นครสวรรค์',
    'ชลบุรี': 'ชลบุรี',
    'มหาชัย': 'คลังมหาชัย',
    'สุวรรณภูมิ': 'คลังสุวรรณภูมิ',
    'หาดใหญ่': 'หาดใหญ่',
    'ภูเก็ต': 'ภูเก็ต',
    'เชียงใหม่': 'เชียงใหม่',
    'สุราษฎร์': 'สุราษฎร์',
    'ขอนแก่น': 'ขอนแก่น'
  };

  const transformed = [];

  if (tableData.length === 0) return transformed;

  const table = tableData[0];

  // Crate columns are located from the table headers, so any number of juice columns works.
  const layout = deriveLayout(Object.values(detectProductColumns(table)).map(p => p.column), table);

  // Initialize totals for all CDC locations
  const cdcTotals = {};
  cdcOrder.forEach(cdc => {
    cdcTotals[cdc] = 0;
  });

  // Each row maps to exactly one CDC, derived from the Vendor cell (C0).
  // Using C0 (with the FC code) avoids double-counting when C1 fill-down carries
  // a previous vendor's warehouse into a later row.
  let totalsRowIndex = -1;
  for (let i = 0; i < table.length; i++) {
    const row = table[i];
    if (!row || row.length < 2) continue;
    const c0 = row[0] ? row[0].toString().trim() : '';
    const totalIndicators = ['ยอดรวม', 'รวม', 'total', 'grand total', 'sum'];
    const foundTotal = totalIndicators.some(indicator =>
      c0.toLowerCase().includes(indicator.toLowerCase())
    );
    if (foundTotal) {
      totalsRowIndex = i;
      break;
    }
  }

  let yodruamValue = 0;
  if (totalsRowIndex !== -1 && table[totalsRowIndex][columnIndex]) {
    yodruamValue = parseOCRNumber(table[totalsRowIndex][columnIndex]);
  }

  // Each row maps to exactly one CDC, derived from the Vendor cell (C0).
  // Using C0 (with the FC code) avoids double-counting when C1 fill-down carries
  // a previous vendor's warehouse into a later row.
  table.forEach((row, rowIndex) => {
    if (!row || row.length < 2) return;
    if (rowIndex === totalsRowIndex) return;

    const c0 = row[0] ? row[0].toString() : '';
    if (!c0) return;

    let bestCdcName = null;
    let bestScore = 0;
    for (const cdcName of cdcNames) {
      if (c0.includes(cdcName) && cdcName.length > bestScore) {
        bestCdcName = cdcName;
        bestScore = cdcName.length;
      }
    }
    if (!bestCdcName) return;

    const value = getRowProductValue(row, columnIndex, rowIndex, table, totalsRowIndex, yodruamValue, layout);
    if (value > 0) {
      cdcTotals[cdcNameMapping[bestCdcName]] += value;
    }
  });

  // Build transformed array in the defined order
  cdcOrder.forEach(cdc => {
    transformed.push({
      cdc: cdc,
      value: cdcTotals[cdc]
    });
  });

  return transformed;
}

function generateHTMLTable(tableData) {
  if (!tableData || tableData.length === 0) {
    return '<p>No table data available</p>';
  }

  let html = '<style>table { border-collapse: collapse; width: 100%; margin: 20px 0; } th, td { border: 1px solid #ddd; padding: 8px; text-align: left; position: relative; } th { background-color: #4CAF50; color: white; } tr:nth-child(even) { background-color: #f2f2f2; } .cell-position { font-size: 9px; color: #999; position: absolute; top: 2px; right: 2px; font-weight: normal; }</style>';

  tableData.forEach((table, tableIndex) => {
    html += `<h3>Table ${tableIndex + 1}</h3>`;
    html += '<table>';

    table.forEach((row, rowIndex) => {
      html += '<tr>';
      row.forEach((cell, colIndex) => {
        const tag = rowIndex === 0 ? 'th' : 'td';
        const position = `R${rowIndex}C${colIndex}`;
        html += `<${tag}><span class="cell-position">${position}</span>${cell || ''}</${tag}>`;
      });
      html += '</tr>';
    });

    html += '</table>';
  });

  return html;
}

function generateTransformedTable(transformedData, category = 'น้ำส้ม') {
  if (!transformedData || transformedData.length === 0) {
    return `<p>No transformed data available for ${category}</p>`;
  }

  const headerColor = category === 'น้ำส้ม' ? '#FF8C00' : '#FFD700';
  const categoryLabel = category === 'น้ำส้ม' ? 'ขวดน้ำส้ม' : 'ขวดยูซุ';

  let html = `<style>
    .category-section { margin: 30px 0; }
    .category-title {
      font-size: 24px;
      font-weight: bold;
      color: ${headerColor};
      margin-bottom: 15px;
      padding: 10px;
      background-color: ${headerColor}20;
      border-radius: 5px;
    }
    table {
      border-collapse: collapse;
      width: 100%;
      margin: 20px 0;
    }
    th, td {
      border: 1px solid #ddd;
      padding: 8px;
      text-align: left;
    }
    th {
      background-color: ${headerColor};
      color: white;
    }
    tr:nth-child(even) {
      background-color: #f2f2f2;
    }
  </style>`;

  html += `<div class="category-section">`;
  html += `<div class="category-title">${category}</div>`;
  html += '<table>';
  html += `<tr><th>CDC</th><th>${categoryLabel}</th></tr>`;

  transformedData.forEach(item => {
    html += `<tr><td>${item.cdc}</td><td>${item.value}</td></tr>`;
  });

  html += '</table>';
  html += '</div>';

  return html;
}

app.get('/latest-ocr', (req, res) => {
  if (!latestOCRResult.timestamp) {
    res.send(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>Latest OCR Result</title>
        <meta charset="utf-8">
      </head>
      <body>
        <h1>No OCR results yet</h1>
        <p>Send an Excel screenshot to the LINE bot to see results here.</p>
      </body>
      </html>
    `);
    return;
  }

  const tableHTML = generateHTMLTable(latestOCRResult.tableData);

  res.send(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>Latest OCR Result</title>
      <meta charset="utf-8">
      <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .info { background-color: #e7f3fe; border-left: 6px solid #2196F3; padding: 10px; margin-bottom: 20px; }
        .text-result { background-color: #f4f4f4; padding: 15px; border-radius: 5px; white-space: pre-wrap; max-height: 300px; overflow-y: auto; }
        .nav { margin-bottom: 20px; }
        .nav a { padding: 10px 15px; background-color: #4CAF50; color: white; text-decoration: none; margin-right: 10px; border-radius: 5px; }
        .nav a:hover { background-color: #45a049; }
      </style>
    </head>
    <body>
      ${reportNav()}

      <h1>Latest OCR Result</h1>
      <div class="info">
        <strong>Processed at:</strong> ${latestOCRResult.timestamp.toLocaleString()}<br>
        <strong>Tables found:</strong> ${latestOCRResult.tableData ? latestOCRResult.tableData.length : 0}
      </div>

      <h2>Table Data</h2>
      ${tableHTML}

      <h2>Raw Extracted Text</h2>
      <div class="text-result">${latestOCRResult.extractedText || 'No text extracted'}</div>
    </body>
    </html>
  `);
});

app.get('/transformed-data', (req, res) => {
  if (!latestOCRResult.timestamp) {
    res.send(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>Transformed Data</title>
        <meta charset="utf-8">
      </head>
      <body>
        <h1>No OCR results yet</h1>
        <p>Send an Excel screenshot to the LINE bot to see results here.</p>
      </body>
      </html>
    `);
    return;
  }

  // Transform data for both Orange (C2) and Yuzu (C3)
  const orangeData = transformTableData(latestOCRResult.tableData, 2);
  const yuzuData = transformTableData(latestOCRResult.tableData, 3);

  const orangeHTML = generateTransformedTable(orangeData, 'น้ำส้ม');
  const yuzuHTML = generateTransformedTable(yuzuData, 'ยูซุ');

  res.send(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>Transformed Data</title>
      <meta charset="utf-8">
      <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .info { background-color: #e7f3fe; border-left: 6px solid #2196F3; padding: 10px; margin-bottom: 20px; }
        .nav { margin-bottom: 20px; }
        .nav a { padding: 10px 15px; background-color: #4CAF50; color: white; text-decoration: none; margin-right: 10px; border-radius: 5px; }
        .nav a:hover { background-color: #45a049; }
      </style>
    </head>
    <body>
      ${reportNav()}

      <h1>Transformed Data (CDC Summary)</h1>
      <div class="info">
        <strong>Processed at:</strong> ${latestOCRResult.timestamp.toLocaleString()}<br>
        <strong>CDC locations:</strong> ${orangeData ? orangeData.length : 0}
      </div>

      ${orangeHTML}
      ${yuzuHTML}
    </body>
    </html>
  `);
});

// The bare domain has no page of its own — send it to the report people actually
// open. Note this only takes effect once the Plesk default index.html is gone from
// the document root: the web server serves static files before handing anything to
// Passenger, so that file wins over this route while it exists.
app.get('/', (req, res) => res.redirect('/daily-report'));

// Pages that list juices pick up ones added or changed by another process.
app.use(['/daily-report', '/mtd-report', '/targets'], (req, res, next) => {
  refreshCategories().then(() => next());
});

app.get('/daily-report', async (req, res) => {
  try {
    // Get current date for default filter
    const now = new Date();
    const currentMonth = now.getMonth() + 1; // 1-12
    const currentYear = now.getFullYear();

    // Get filter parameters from query string (default to current month/year)
    const selectedMonth = parseInt(req.query.month) || currentMonth;
    const selectedYear = parseInt(req.query.year) || currentYear;

    // Fetch records directly filtered by month/year from MySQL
    const records = await db.getDailyRecordsByMonth(selectedMonth, selectedYear);

    // Generate table HTML for each category
    // Names, icon and colour come from the same place as every other page, so what is set
    // on the juice page shows here too.
    const categoryData = {};
    for (const cat of db.CATEGORIES) {
      const thaiName = db.CATEGORY_NAMES[cat] || cat;
      const color = juiceColor(cat);
      categoryData[cat] = {
        records: records[cat] || [],
        name: db.CATEGORY_ENGLISH_NAMES[cat] || thaiName,
        thaiName,
        icon: juiceIcon(cat),
        bgColor: (!db.JUICE_COLORS[cat] && CATEGORY_TINTS[cat]) || tintColor(color),
        borderColor: color
      };
    }

    // Generate tables for each category
    const categoryTables = {};
    for (const [cat, data] of Object.entries(categoryData)) {
      categoryTables[cat] = generateDailyRecordsTable(data.records, data.name);
    }

    // Calculate totals
    let totalRecords = 0;
    for (const cat of Object.keys(categoryData)) {
      totalRecords += categoryData[cat].records.length;
    }

    // Prepare chart data - get all unique dates and build datasets
    const allDates = new Set();
    const activeCategories = Object.entries(categoryData).filter(([cat, data]) => data.records.length > 0);

    activeCategories.forEach(([cat, data]) => {
      data.records.forEach(record => allDates.add(record.date));
    });

    // Sort dates chronologically
    const sortedDates = Array.from(allDates).sort((a, b) => {
      const [dayA, monthA, yearA] = a.split('/').map(Number);
      const [dayB, monthB, yearB] = b.split('/').map(Number);
      return new Date(yearA, monthA - 1, dayA) - new Date(yearB, monthB - 1, dayB);
    });

    // Build chart datasets
    const chartDatasets = activeCategories.map(([cat, data]) => {
      const dateToTotal = {};
      data.records.forEach(record => {
        dateToTotal[record.date] = record.totalSum || 0;
      });

      return {
        label: data.name,
        data: sortedDates.map(date => dateToTotal[date] || 0),
        backgroundColor: data.borderColor,
        borderColor: data.borderColor,
        borderWidth: 1
      };
    });

    const chartLabels = sortedDates.map(date => {
      const [day] = date.split('/');
      return day;
    });

    res.send(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>Daily Records Report</title>
        <meta charset="utf-8">
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
          }
          .container {
            max-width: 100%;
            margin: 0 auto;
            background-color: white;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          }
          h1 {
            text-align: center;
            color: #333;
            margin-bottom: 30px;
          }
          .category-section {
            margin-bottom: 40px;
            padding: 20px;
            border-radius: 8px;
          }
          .orange-section {
            background-color: #fff3e0;
            border-left: 5px solid #ff9800;
          }
          .yuzu-section {
            background-color: #fffde7;
            border-left: 5px solid #fdd835;
          }
          .category-title {
            font-size: 24px;
            font-weight: bold;
            margin-bottom: 15px;
            display: flex;
            align-items: center;
            gap: 10px;
          }
          .orange-title {
            color: #ff9800;
          }
          .yuzu-title {
            color: #f9a825;
          }
          table {
            border-collapse: collapse;
            width: 100%;
            background-color: white;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            font-size: 14px;
          }
          th, td {
            border: 1px solid #ccc;
            padding: 10px 8px;
            text-align: center;
          }
          th {
            background-color: #4CAF50;
            color: white;
            font-weight: bold;
            white-space: nowrap;
          }
          td:first-child {
            text-align: left;
            white-space: nowrap;
          }
          tr:nth-child(even) {
            background-color: #f9f9f9;
          }
          tr:hover {
            background-color: #f1f1f1;
          }
          .no-data {
            padding: 20px;
            text-align: center;
            color: #666;
            font-style: italic;
          }
          .nav {
            margin-bottom: 20px;
            text-align: center;
          }
          .nav a {
            padding: 10px 20px;
            background-color: #4CAF50;
            color: white;
            text-decoration: none;
            margin: 0 5px;
            border-radius: 5px;
            display: inline-block;
          }
          .nav a:hover {
            background-color: #45a049;
          }
          .summary {
            display: flex;
            justify-content: space-around;
            margin-bottom: 30px;
            flex-wrap: wrap;
          }
          .summary-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 10px;
            min-width: 200px;
            text-align: center;
            margin: 10px;
          }
          .summary-number {
            font-size: 36px;
            font-weight: bold;
          }
          .summary-label {
            font-size: 14px;
            opacity: 0.9;
          }
          .filter-section {
            margin-bottom: 30px;
            text-align: center;
            padding: 20px;
            background-color: #f0f0f0;
            border-radius: 8px;
          }
          .filter-section label {
            font-weight: bold;
            font-size: 16px;
            margin-right: 10px;
          }
          .filter-section select {
            padding: 8px 15px;
            font-size: 16px;
            border: 2px solid #4CAF50;
            border-radius: 5px;
            background-color: white;
            cursor: pointer;
            transition: border-color 0.3s;
          }
          .filter-section select:hover {
            border-color: #45a049;
          }
          .filter-section select:focus {
            outline: none;
            border-color: #2196F3;
          }
        </style>
      </head>
      <body>
        <div class="container">
          ${reportNav()}

          <h1>📊 Daily Records Report</h1>

          <div class="filter-section">
            <label for="month-filter">Month:</label>
            <select id="month-filter" onchange="applyFilter()">
              <option value="1" ${selectedMonth === 1 ? 'selected' : ''}>January</option>
              <option value="2" ${selectedMonth === 2 ? 'selected' : ''}>February</option>
              <option value="3" ${selectedMonth === 3 ? 'selected' : ''}>March</option>
              <option value="4" ${selectedMonth === 4 ? 'selected' : ''}>April</option>
              <option value="5" ${selectedMonth === 5 ? 'selected' : ''}>May</option>
              <option value="6" ${selectedMonth === 6 ? 'selected' : ''}>June</option>
              <option value="7" ${selectedMonth === 7 ? 'selected' : ''}>July</option>
              <option value="8" ${selectedMonth === 8 ? 'selected' : ''}>August</option>
              <option value="9" ${selectedMonth === 9 ? 'selected' : ''}>September</option>
              <option value="10" ${selectedMonth === 10 ? 'selected' : ''}>October</option>
              <option value="11" ${selectedMonth === 11 ? 'selected' : ''}>November</option>
              <option value="12" ${selectedMonth === 12 ? 'selected' : ''}>December</option>
            </select>

            <label for="year-filter" style="margin-left: 20px;">Year:</label>
            <select id="year-filter" onchange="applyFilter()">
              <option value="2024" ${selectedYear === 2024 ? 'selected' : ''}>2024</option>
              <option value="2025" ${selectedYear === 2025 ? 'selected' : ''}>2025</option>
              <option value="2026" ${selectedYear === 2026 ? 'selected' : ''}>2026</option>
            </select>
          </div>

          <div class="summary">
            ${Object.entries(categoryData).filter(([cat, data]) => data.records.length > 0).map(([cat, data]) => `
              <div class="summary-card" style="background: linear-gradient(135deg, ${data.borderColor} 0%, ${data.borderColor}dd 100%);">
                <div class="summary-number">${data.records.length}</div>
                <div class="summary-label">${data.icon} ${data.name}</div>
              </div>
            `).join('')}
            <div class="summary-card">
              <div class="summary-number">${totalRecords}</div>
              <div class="summary-label">Total Records</div>
            </div>
          </div>

          ${totalRecords === 0 ? '<div class="no-data" style="padding: 40px; text-align: center; color: #666;">No records found for this month</div>' : ''}

          ${totalRecords > 0 ? `
          <div class="chart-container" style="background: white; padding: 20px; border-radius: 8px; margin-bottom: 30px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
            <h3 style="margin-top: 0; color: #333;">Daily Sales Chart</h3>
            <canvas id="dailyChart" height="100"></canvas>
          </div>
          ` : ''}

          ${Object.entries(categoryData).filter(([cat, data]) => data.records.length > 0).map(([cat, data]) => `
            <div class="category-section" style="background-color: ${data.bgColor}; border-left: 5px solid ${data.borderColor};">
              <div class="category-title" style="color: ${data.borderColor};">
                ${data.icon} ${data.thaiName} Records
              </div>
              ${categoryTables[cat]}
            </div>
          `).join('')}
        </div>

        <script>
          function applyFilter() {
            const month = document.getElementById('month-filter').value;
            const year = document.getElementById('year-filter').value;
            window.location.href = '/daily-report?month=' + month + '&year=' + year;
          }

          // Render stacked bar chart
          ${totalRecords > 0 ? `
          const ctx = document.getElementById('dailyChart').getContext('2d');
          new Chart(ctx, {
            type: 'bar',
            data: {
              labels: ${JSON.stringify(chartLabels)},
              datasets: ${JSON.stringify(chartDatasets)}
            },
            options: {
              responsive: true,
              scales: {
                x: {
                  stacked: true,
                  title: {
                    display: true,
                    text: 'Day of Month'
                  }
                },
                y: {
                  stacked: true,
                  title: {
                    display: true,
                    text: 'Total Units'
                  },
                  beginAtZero: true
                }
              },
              plugins: {
                legend: {
                  position: 'top'
                },
                tooltip: {
                  mode: 'index',
                  intersect: false
                }
              }
            }
          });
          ` : ''}
        </script>
      </body>
      </html>
    `);
  } catch (error) {
    console.error('Error generating daily report:', error);
    res.status(500).send('Error generating report');
  }
});

function generateDailyRecordsTable(records, category) {
  if (!records || records.length === 0) {
    return '<div class="no-data">No records available for ' + category + '</div>';
  }

  // Sort records by date in ascending order
  const sortedRecords = [...records].sort((a, b) => {
    // Parse DD/MM/YYYY format
    const parseDate = (dateStr) => {
      const [day, month, year] = dateStr.split('/').map(Number);
      return new Date(year, month - 1, day);
    };

    return parseDate(a.date) - parseDate(b.date);
  });

  // Column headers matching CDC locations (in proper order)
  const cdcColumns = [
    'คลังบางบัวทอง',
    'นครราชสีมา',
    'นครสวรรค์',
    'ชลบุรี',
    'หาดใหญ่',
    'ภูเก็ต',
    'เชียงใหม่',
    'สุราษฎร์',
    'ขอนแก่น',
    'คลังมหาชัย',
    'คลังสุวรรณภูมิ'
  ];

  let html = '<div style="overflow-x: auto;"><table>';

  // Header row
  html += '<tr>';
  html += '<th>วันที่</th>';
  cdcColumns.forEach((header) => {
    if (header === 'คลังสุวรรณภูมิ') {
      // Highlight คลังสุวรรณภูมิ column
      html += `<th style="background-color: #2196F3;">${header}</th>`;
    } else {
      html += `<th>${header}</th>`;
    }
  });

  // Add special columns for Orange category only
  if (category.toLowerCase() === 'orange') {
    html += '<th style="background-color: #FF9800;">ขอนแก่น<br>Laos<br>4 สาขา</th>';
    html += '<th style="background-color: #FF9800;">ขอนแก่น<br>Cambodia<br>85 สาขา</th>';
  }

  html += '<th style="background-color: #4CAF50;">รวม</th>'; // Total sum column
  html += '<th style="background-color: #9c27b0;">Recorded At</th></tr>';

  // Initialize sums for averaging
  const cdcSums = {};
  cdcColumns.forEach(cdc => cdcSums[cdc] = 0);
  let totalSumSum = 0;
  let laosSum = 0;
  let cambodiaSum = 0;
  const recordCount = sortedRecords.length;

  // Use the already sorted records (by date ascending)
  sortedRecords.forEach(record => {
    const recordedDate = new Date(record.timestamp).toLocaleString('th-TH', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit'
    });

    html += '<tr>';

    // Date column
    html += `<td style="background-color: #fff9c4;"><strong>${record.date}</strong></td>`;

    // Display CDC totals
    if (record.cdcTotals) {
      cdcColumns.forEach(cdc => {
        const value = record.cdcTotals[cdc] || 0;
        cdcSums[cdc] += value; // Accumulate for average
        const formattedValue = value.toLocaleString('en-US');

        if (cdc === 'คลังสุวรรณภูมิ') {
          // Highlight คลังสุวรรณภูมิ column
          html += `<td style="background-color: #e3f2fd; font-weight: bold;">${formattedValue}</td>`;
        } else {
          html += `<td>${formattedValue}</td>`;
        }
      });
    } else {
      // Fallback for old records without cdcTotals
      cdcColumns.forEach(() => {
        html += '<td>-</td>';
      });
    }

    // Add special columns for Orange category only
    if (category.toLowerCase() === 'orange') {
      // ขอนแก่น Laos column
      const laosValue = record.khonKaenLaos || 0;
      laosSum += laosValue; // Accumulate for average
      const formattedLaos = laosValue.toLocaleString('en-US');
      html += `<td style="background-color: #FFE0B2; font-weight: bold;">${formattedLaos}</td>`;

      // ขอนแก่น Cambodia column (always blank/0)
      const cambodiaValue = record.khonKaenCambodia || 0;
      cambodiaSum += cambodiaValue; // Accumulate for average
      const formattedCambodia = cambodiaValue.toLocaleString('en-US');
      html += `<td style="background-color: #FFE0B2;">${formattedCambodia}</td>`;
    }

    // Total sum column
    const totalValue = record.totalSum || 0;
    totalSumSum += totalValue; // Accumulate for average
    const formattedTotal = totalValue.toLocaleString('en-US');
    html += `<td style="background-color: #c8e6c9; font-weight: bold;">${formattedTotal}</td>`;

    // Recorded timestamp
    html += `<td style="background-color: #f3e5f5; font-size: 0.9em;">${recordedDate}</td>`;
    html += '</tr>';
  });

  // Add Month to Date Sum and Average rows
  if (recordCount > 0) {
    // MTD Sum row
    html += '<tr style="background-color: #d1c4e9; font-weight: bold; border-top: 3px solid #333;">';
    html += '<td style="background-color: #b39ddb;">MTD Sum<br>(' + recordCount + ' days)</td>';

    // CDC column sums
    cdcColumns.forEach(cdc => {
      const sum = cdcSums[cdc];
      const formattedSum = sum.toLocaleString('en-US');
      if (cdc === 'คลังสุวรรณภูมิ') {
        html += `<td style="background-color: #b39ddb;">${formattedSum}</td>`;
      } else {
        html += `<td>${formattedSum}</td>`;
      }
    });

    // Orange category special columns sums
    if (category.toLowerCase() === 'orange') {
      html += `<td style="background-color: #ce93d8;">${laosSum.toLocaleString('en-US')}</td>`;
      html += `<td style="background-color: #ce93d8;">${cambodiaSum.toLocaleString('en-US')}</td>`;
    }

    // Total sum
    html += `<td style="background-color: #b39ddb;">${totalSumSum.toLocaleString('en-US')}</td>`;

    // Empty cell for timestamp column
    html += '<td style="background-color: #b39ddb;">-</td>';
    html += '</tr>';

    // MTD Average row
    html += '<tr style="background-color: #e0e0e0; font-weight: bold;">';
    html += '<td style="background-color: #bdbdbd;">MTD Avg</td>';

    // CDC column averages
    cdcColumns.forEach(cdc => {
      const avg = Math.round(cdcSums[cdc] / recordCount);
      const formattedAvg = avg.toLocaleString('en-US');
      if (cdc === 'คลังสุวรรณภูมิ') {
        html += `<td style="background-color: #bbdefb;">${formattedAvg}</td>`;
      } else {
        html += `<td>${formattedAvg}</td>`;
      }
    });

    // Orange category special columns averages
    if (category.toLowerCase() === 'orange') {
      const laosAvg = Math.round(laosSum / recordCount);
      const cambodiaAvg = Math.round(cambodiaSum / recordCount);
      html += `<td style="background-color: #FFCC80;">${laosAvg.toLocaleString('en-US')}</td>`;
      html += `<td style="background-color: #FFCC80;">${cambodiaAvg.toLocaleString('en-US')}</td>`;
    }

    // Total average
    const totalAvg = Math.round(totalSumSum / recordCount);
    html += `<td style="background-color: #a5d6a7;">${totalAvg.toLocaleString('en-US')}</td>`;

    // Empty cell for timestamp column
    html += '<td style="background-color: #bdbdbd;">-</td>';
    html += '</tr>';
  }

  html += '</table></div>';
  return html;
}

// Test endpoint for uploading images directly (bypass LINE webhook)
app.post('/test-ocr', express.raw({ type: 'image/*', limit: '10mb' }), async (req, res) => {
  try {
    console.log('[TEST-OCR] Received image for testing');
    const imageBuffer = req.body;
    console.log('[TEST-OCR] Image size:', imageBuffer.length, 'bytes');

    console.log('[TEST-OCR] Detecting if Excel screenshot...');
    const isExcelScreenshot = await detectExcelScreenshot(imageBuffer);
    console.log('[TEST-OCR] Is Excel screenshot:', isExcelScreenshot);

    if (isExcelScreenshot) {
      console.log('[TEST-OCR] Performing OCR...');
      // performOCR returns { extractedText, recordResult } — assigning the whole
      // object to extractedText made the page call .substring() on an object and
      // report an error for uploads that had in fact been recorded correctly.
      const { extractedText, recordResult } = await performOCR(imageBuffer);

      res.json({
        success: true,
        isExcelScreenshot: true,
        extractedText: extractedText,
        recorded: !!(recordResult && recordResult.success),
        date: recordResult?.date || null,
        categories: recordResult?.results?.map(r => r.category) || [],
        reason: recordResult?.reason || null,
        tableData: latestOCRResult.tableData,
        cdcTotals: latestOCRResult.tableData ? transformTableData(latestOCRResult.tableData) : null
      });
    } else {
      res.json({
        success: true,
        isExcelScreenshot: false,
        message: 'Not detected as an Excel screenshot'
      });
    }
  } catch (error) {
    console.error('[ERROR] Test OCR error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// ---------- Juice page (linked only from dev alerts) ----------

// No login: the page is reached only through the link in dev alerts. Its token is derived
// from the LINE channel secret — stable across restarts, nothing to configure on Plesk —
// and a wrong token gets a plain 404, as if the page didn't exist.
const JUICE_PAGE_TOKEN = crypto.createHmac('sha256', process.env.LINE_CHANNEL_SECRET || '')
  .update('juice-page').digest('hex').slice(0, 24);

function juicePageUrl() {
  return `https://shinsen.yushi-marketing.com/juices/${JUICE_PAGE_TOKEN}`;
}

function requireJuiceToken(req, res, next) {
  const given = Buffer.from(String(req.params.token || ''));
  const expected = Buffer.from(JUICE_PAGE_TOKEN);
  if (given.length !== expected.length || !crypto.timingSafeEqual(given, expected)) {
    return res.sendStatus(404);
  }
  // The token is in the URL; don't hand it to anything the page links to.
  res.set('Referrer-Policy', 'no-referrer');
  next();
}

function backToJuicePage(res, msg) {
  res.redirect(303, `/juices/${JUICE_PAGE_TOKEN}?msg=${encodeURIComponent(msg)}`);
}

// Icons a juice can be given on the juice page. A fixed list rather than free text: the value
// lands in report HTML, and picking from a list beats hunting for an emoji anyway.
const JUICE_ICON_CHOICES = [
  ['🧃', 'กล่องน้ำผลไม้'], ['🍊', 'ส้ม'], ['🍋', 'มะนาว / ยูซุ'], ['🍎', 'แอปเปิ้ลแดง'],
  ['🍏', 'แอปเปิ้ลเขียว'], ['🍐', 'ลูกแพร์'], ['🍑', 'พีช'], ['🍒', 'เชอร์รี่'],
  ['🍓', 'สตรอเบอร์รี่'], ['🫐', 'บลูเบอร์รี่'], ['🍇', 'องุ่น'], ['🍉', 'แตงโม'],
  ['🍈', 'เมลอน'], ['🍌', 'กล้วย'], ['🍍', 'สับปะรด'], ['🥭', 'มะม่วง'],
  ['🥥', 'มะพร้าว'], ['🥝', 'กีวี่'], ['🍅', 'มะเขือเทศ'], ['🥕', 'แครอท'],
  ['🍵', 'ชา'], ['🍹', 'น้ำผลไม้ปั่น'], ['🥤', 'แก้วน้ำ']
];

// A submitted icon if it is one of the choices, otherwise `fallback`.
function iconFromForm(value, fallback = null) {
  return JUICE_ICON_CHOICES.some(([emoji]) => emoji === value) ? value : fallback;
}

function iconOptions(selected) {
  return JUICE_ICON_CHOICES
    .map(([emoji, label]) => `<option value="${emoji}"${emoji === selected ? ' selected' : ''}>${emoji} ${label}</option>`)
    .join('');
}

// English name from a form field: blank means "none", anything else must clean up.
function englishFromForm(value) {
  const raw = String(value || '').trim();
  if (!raw) return { englishName: null };
  const englishName = cleanEnglishName(raw);
  return englishName ? { englishName } : { error: 'ชื่ออังกฤษต้องเป็นตัวอักษรภาษาอังกฤษ เช่น Lychee' };
}

// A built-in juice's display name is free text (they aren't all "น้ำ…": ยูซุ, Shinsen Pop);
// it only has to be safe to put in report HTML.
function displayNameFromForm(value) {
  return String(value || '').replace(/[<>&"'`]/g, '').replace(/\s+/g, ' ').trim().slice(0, 50);
}

function colorFromForm(value, fallback = null) {
  return /^#[0-9a-f]{6}$/i.test(value || '') ? value.toLowerCase() : fallback;
}

function renderJuicePage({ pendings, customs, stats, msg }) {
  const base = `/juices/${JUICE_PAGE_TOKEN}`;
  const detection = getProductDetection();
  const mergeTargets = Object.values(detection).map(d => d.dbCategory);
  const customByCategory = Object.fromEntries(customs.map(c => [c.category, c]));

  const pendingHtml = pendings.length ? pendings.map(p => `
    <div class="card" style="border-left-color:${p.color || '#bbb'}">
      <div class="title">${escapeHtml(p.header)}</div>
      <div class="meta">${escapeHtml(p.date)} · ${p.total.toLocaleString()} ขวด · ยังไม่ถูกบันทึก</div>
      <form method="post" action="${base}/pending/${p.id}/add" class="fields">
        <label>ชื่อไทย<input name="thai_name" required placeholder="เช่น ลิ้นจี่" value="${escapeHtml(p.keyword)}"></label>
        <label>ชื่ออังกฤษ<input name="english_name" placeholder="เว้นว่างให้แปลเอง" value="${escapeHtml(p.englishName || '')}"></label>
        <label class="icon">รูป<select name="icon">${iconOptions(NEW_JUICE_ICON)}</select></label>
        <button>เพิ่มและบันทึกยอด</button>
      </form>
      <div class="hint">ถ้าเป็นน้ำที่มีอยู่แล้ว ใส่ชื่อน้ำนั้น (เช่น แตงโม) ระบบจะบันทึกยอดเข้าน้ำนั้น</div>
      <form method="post" action="${base}/pending/${p.id}/ignore" onsubmit="return confirm('ทิ้งรายการนี้? ยอดจะไม่ถูกบันทึก')">
        <button class="secondary">ทิ้ง</button>
      </form>
    </div>`).join('') : '<p class="empty">ไม่มีรายการที่รอตั้งชื่อ</p>';

  // Every juice, built-in first, in dashboard order. Built-in ones only take display settings
  // (their keywords are in code, and the plan always carries them); auto-added ones can also
  // be merged into another juice or deleted.
  const juiceHtml = db.CATEGORIES.map(category => {
    const builtin = db.BUILTIN_CATEGORIES.includes(category);
    const row = customByCategory[category];
    const s = stats[category] || { days: 0, bottles: 0, lastDate: null };
    const color = juiceColor(category);
    const icon = juiceIcon(category);
    const thaiName = db.CATEGORY_NAMES[category] || category;
    const englishName = builtin ? db.CATEGORY_ENGLISH_NAMES[category] : (row?.english_name || '');
    const keywords = detection[category]?.keywords || [];
    const action = `${base}/category/${encodeURIComponent(category)}`;
    const origin = builtin
      ? `น้ำหลักในระบบ${db.DISPLAY_OVERRIDES[category] ? ' · แก้การแสดงผลแล้ว' : ''}`
      : `เพิ่มเมื่อ ${escapeHtml(db.toDisplayDate(row?.created_at) || '-')} (${row?.created_by === 'web' ? 'ตั้งชื่อจากหน้านี้' : 'อัตโนมัติ'})`;

    const actions = builtin
      ? (db.DISPLAY_OVERRIDES[category] ? `
        <form method="post" action="${action}/reset" onsubmit="return confirm('คืนชื่อ สี และรูปของน้ำนี้เป็นค่าเดิม?')">
          <button class="secondary">คืนค่าเดิม</button>
        </form>` : '')
      : `
        <form method="post" action="${action}/merge" onsubmit="return confirm('รวมเข้ากับน้ำที่เลือก? ข้อมูลทั้งหมดจะย้ายไป และน้ำนี้จะหายไป')">
          <select name="target" required>
            <option value="">รวมเข้ากับน้ำ…</option>
            ${mergeTargets.filter(t => t !== category).map(t => `<option value="${escapeHtml(t)}">${escapeHtml(db.CATEGORY_NAMES[t] || t)}</option>`).join('')}
          </select>
          <button class="secondary">รวม</button>
        </form>
        <form method="post" action="${action}/delete" onsubmit="return confirm('ลบน้ำนี้พร้อมข้อมูล ${s.days} วัน? คอลัมน์ชื่อนี้จะไม่ถูกเพิ่มอีก')">
          <button class="danger">ลบ</button>
        </form>`;

    return `
    <div class="card" style="border-left-color:${escapeHtml(color)}">
      <div class="title">${escapeHtml(icon)} ${escapeHtml(thaiName)}${englishName ? ` <span class="en">${escapeHtml(englishName)}</span>` : ''}</div>
      <div class="meta">${origin} · มีข้อมูล ${s.days} วัน · รวม ${s.bottles.toLocaleString()} ขวด · ล่าสุด ${escapeHtml(s.lastDate || '-')}</div>
      <div class="meta">คำที่ใช้จับคอลัมน์: ${keywords.map(k => escapeHtml(k)).join(', ')}</div>
      <form method="post" action="${action}/update" class="fields">
        <label>ชื่อไทย<input name="thai_name" required value="${escapeHtml(builtin ? thaiName : thaiName.replace(/^น้ำ/, ''))}"></label>
        <label>ชื่ออังกฤษ<input name="english_name" placeholder="${escapeHtml(builtin ? db.BUILTIN_ENGLISH_NAMES[category] : 'เว้นว่าง = ใช้ชื่อไทย')}" value="${escapeHtml(englishName)}"></label>
        <label class="icon">รูป<select name="icon">${iconOptions(icon)}</select></label>
        <label class="color">สี<input type="color" name="color" value="${escapeHtml(color)}"></label>
        <button>บันทึก</button>
      </form>
      ${builtin ? '<div class="hint">แก้แค่ชื่อและสีที่แสดงบน dashboard ไม่กระทบการอ่านหัวคอลัมน์</div>' : ''}
      ${actions ? `<div class="row">${actions}</div>` : ''}
    </div>`;
  }).join('');

  return `<!DOCTYPE html>
<html lang="th">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="robots" content="noindex, nofollow">
  <title>จัดการน้ำ</title>
  <style>
    body { font-family: -apple-system, 'Segoe UI', Tahoma, sans-serif; background: #f5f5f5; color: #222; margin: 0; padding: 16px; }
    main { max-width: 760px; margin: 0 auto; }
    h1 { font-size: 22px; margin: 4px 0 16px; }
    h2 { font-size: 17px; margin: 28px 0 10px; color: #444; }
    .msg { background: #fff; border: 1px solid #ddd; border-radius: 8px; padding: 10px 14px; margin-bottom: 8px; }
    .card { background: #fff; border-radius: 8px; border-left: 6px solid #bbb; padding: 14px 16px; margin-bottom: 12px; box-shadow: 0 1px 3px rgba(0,0,0,.08); }
    .title { font-weight: 600; font-size: 16px; }
    .en { color: #777; font-weight: 400; }
    .meta, .hint { color: #666; font-size: 13px; margin-top: 4px; }
    .fields { display: flex; flex-wrap: wrap; gap: 8px; align-items: flex-end; margin-top: 10px; }
    .fields label { display: flex; flex-direction: column; font-size: 13px; color: #555; gap: 3px; flex: 1 1 160px; }
    .fields label.color, .fields label.icon { flex: 0 0 auto; }
    input, select { font: inherit; padding: 7px 9px; border: 1px solid #ccc; border-radius: 6px; }
    input[type=color] { padding: 2px; width: 52px; height: 36px; }
    .row { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 10px; }
    .row form { display: flex; gap: 6px; }
    form { margin: 0; }
    button { font: inherit; padding: 7px 14px; border: 0; border-radius: 6px; background: #2e7d32; color: #fff; cursor: pointer; }
    button.secondary { background: #e0e0e0; color: #222; }
    button.danger { background: #c62828; }
    .card > form:last-child { margin-top: 8px; }
    .empty { color: #777; }
    .foot { margin-top: 28px; font-size: 13px; }
  </style>
</head>
<body>
<main>
  <h1>จัดการน้ำ</h1>
  ${msg ? `<div class="msg">${escapeHtml(msg)}</div>` : ''}
  <h2>รอตั้งชื่อ (${pendings.length})</h2>
  ${pendingHtml}
  <h2>น้ำทั้งหมด (${db.CATEGORIES.length})</h2>
  ${juiceHtml}
  <p class="foot"><a href="/daily-report">Daily report</a></p>
</main>
</body>
</html>`;
}

app.get('/juices/:token', requireJuiceToken, async (req, res) => {
  try {
    await refreshCategories();
    const [pendings, customs, stats] = await Promise.all([
      db.listPendingJuices(),
      db.listCustomCategories(),
      db.getCategoryStats(db.CATEGORIES)
    ]);
    res.send(renderJuicePage({ pendings, customs, stats, msg: req.query.msg }));
  } catch (error) {
    console.error('[JUICE-PAGE] Error rendering:', error);
    res.status(500).send(`เกิดข้อผิดพลาด: ${escapeHtml(error.message)}`);
  }
});

// Name a column that couldn't be added on its own, and record it.
app.post('/juices/:token/pending/:id/add', requireJuiceToken, async (req, res) => {
  try {
    const pending = await db.getPendingJuice(parseInt(req.params.id, 10));
    if (!pending || pending.status !== 'pending') return backToJuicePage(res, 'รายการนี้ถูกจัดการไปแล้ว');
    const thaiName = String(req.body.thai_name || '').trim();
    if (!thaiName) return backToJuicePage(res, '❌ กรุณาใส่ชื่อไทย');
    const english = englishFromForm(req.body.english_name);
    if (english.error) return backToJuicePage(res, `❌ ${english.error}`);

    const icon = iconFromForm(req.body.icon);
    const added = await addJuiceFromPending(pending, { thaiName, englishName: english.englishName, icon }, 'web');
    if (added.error) return backToJuicePage(res, `❌ ${added.error}`);
    backToJuicePage(res, `✅ ${added.existing ? 'บันทึกเข้า' : 'เพิ่ม'} "${added.thaiName}" แล้ว — ${added.lines.join(' / ')}`);
  } catch (error) {
    console.error('[JUICE-PAGE] Add failed:', error);
    backToJuicePage(res, `❌ ทำรายการไม่สำเร็จ: ${error.message}`);
  }
});

app.post('/juices/:token/pending/:id/ignore', requireJuiceToken, async (req, res) => {
  try {
    const ok = await db.resolvePendingJuice(parseInt(req.params.id, 10), 'ignored', null, 'web');
    backToJuicePage(res, ok ? 'ทิ้งรายการแล้ว ยอดนี้จะไม่ถูกบันทึก' : 'รายการนี้ถูกจัดการไปแล้ว');
  } catch (error) {
    console.error('[JUICE-PAGE] Ignore failed:', error);
    backToJuicePage(res, `❌ ทำรายการไม่สำเร็จ: ${error.message}`);
  }
});

// Change how a built-in juice looks on the dashboards. Only what differs from the look in
// code is stored, so a later change to that default still shows where nothing was set.
async function updateBuiltinJuice(category, body) {
  const thaiName = displayNameFromForm(body.thai_name);
  if (!thaiName) return '❌ กรุณาใส่ชื่อไทย';
  const english = englishFromForm(body.english_name);
  if (english.error) return `❌ ${english.error}`;
  const color = colorFromForm(body.color, juiceColor(category));
  const icon = iconFromForm(body.icon, juiceIcon(category));

  const settings = {
    thaiName: thaiName !== db.BUILTIN_CATEGORY_NAMES[category] ? thaiName : null,
    englishName: english.englishName && english.englishName !== db.BUILTIN_ENGLISH_NAMES[category] ? english.englishName : null,
    color: color !== CATEGORY_COLORS[category] ? color : null,
    icon: icon !== CATEGORY_ICONS[category] ? icon : null
  };
  if (Object.values(settings).every(v => v === null)) await db.resetBuiltinDisplay(category);
  else await db.setBuiltinDisplay(category, settings);
  console.log(`[JUICE-PAGE] Display of built-in ${category}:`, settings);
  return `✅ บันทึก "${db.CATEGORY_NAMES[category]}" แล้ว`;
}

// Fix the names, colour or icon of a juice. For an auto-added one a corrected Thai name joins
// its keywords rather than replacing them: tomorrow's plan may still be read the old way.
app.post('/juices/:token/category/:category/update', requireJuiceToken, async (req, res) => {
  try {
    await refreshCategories();
    const category = req.params.category;
    if (db.BUILTIN_CATEGORIES.includes(category)) return backToJuicePage(res, await updateBuiltinJuice(category, req.body));
    if (!db.CUSTOM_KEYWORDS[category]) return backToJuicePage(res, '❌ ไม่พบน้ำนี้');
    const keyword = juiceKeywordFromHeader(req.body.thai_name);
    const problem = juiceKeywordProblem(keyword, category);
    if (problem) return backToJuicePage(res, `❌ ${problem}`);
    const english = englishFromForm(req.body.english_name);
    if (english.error) return backToJuicePage(res, `❌ ${english.error}`);
    const color = colorFromForm(req.body.color, db.JUICE_COLORS[category] || null);
    const icon = iconFromForm(req.body.icon, db.JUICE_ICONS[category] || null);

    const thaiName = `น้ำ${keyword}`;
    await db.updateCategory(category, {
      thaiName,
      englishName: english.englishName,
      color,
      icon,
      keywords: [...new Set([...db.CUSTOM_KEYWORDS[category], keyword])]
    });
    console.log(`[JUICE-PAGE] Updated ${category}: "${thaiName}" / ${english.englishName || '-'} / ${color || '-'} / ${icon || '-'}`);
    backToJuicePage(res, `✅ บันทึก "${thaiName}" แล้ว`);
  } catch (error) {
    console.error('[JUICE-PAGE] Update failed:', error);
    backToJuicePage(res, `❌ ทำรายการไม่สำเร็จ: ${error.message}`);
  }
});

app.post('/juices/:token/category/:category/reset', requireJuiceToken, async (req, res) => {
  try {
    const category = req.params.category;
    if (!db.BUILTIN_CATEGORIES.includes(category)) return backToJuicePage(res, '❌ คืนค่าเดิมได้เฉพาะน้ำหลักในระบบ');
    await db.resetBuiltinDisplay(category);
    console.log(`[JUICE-PAGE] Display of built-in ${category} reset`);
    backToJuicePage(res, `✅ คืนค่าเดิมของ "${db.CATEGORY_NAMES[category]}" แล้ว`);
  } catch (error) {
    console.error('[JUICE-PAGE] Reset failed:', error);
    backToJuicePage(res, `❌ ทำรายการไม่สำเร็จ: ${error.message}`);
  }
});

app.post('/juices/:token/category/:category/merge', requireJuiceToken, async (req, res) => {
  try {
    await refreshCategories();
    const from = req.params.category;
    const to = String(req.body.target || '');
    const targets = Object.values(getProductDetection()).map(d => d.dbCategory);
    if (!db.CUSTOM_KEYWORDS[from]) return backToJuicePage(res, '❌ ไม่พบน้ำนี้');
    if (!targets.includes(to) || to === from) return backToJuicePage(res, '❌ กรุณาเลือกน้ำที่จะรวมเข้า');

    const fromName = db.CATEGORY_NAMES[from] || from;
    const toName = db.CATEGORY_NAMES[to] || to;
    const result = await db.mergeCategory(from, to);
    if (result.clash) {
      return backToJuicePage(res, `❌ รวมไม่ได้: ${fromName} กับ ${toName} มีข้อมูลวันเดียวกัน (${result.clash.join(', ')}) แปลว่าเป็นน้ำคนละชนิดในภาพเดียวกัน`);
    }
    console.log(`[JUICE-PAGE] Merged ${from} into ${to} (${result.moved} records)`);
    backToJuicePage(res, `✅ รวม ${fromName} เข้ากับ ${toName} แล้ว — ย้ายข้อมูล ${result.moved} วัน`);
  } catch (error) {
    console.error('[JUICE-PAGE] Merge failed:', error);
    backToJuicePage(res, `❌ ทำรายการไม่สำเร็จ: ${error.message}`);
  }
});

app.post('/juices/:token/category/:category/delete', requireJuiceToken, async (req, res) => {
  try {
    await refreshCategories();
    const category = req.params.category;
    if (!db.CUSTOM_KEYWORDS[category]) return backToJuicePage(res, '❌ ไม่พบน้ำนี้');
    const name = db.CATEGORY_NAMES[category] || category;
    const result = await db.deleteCategory(category);
    console.log(`[JUICE-PAGE] Deleted ${category} (${result.deleted} records)`);
    backToJuicePage(res, `🗑️ ลบ ${name} แล้ว (ลบข้อมูล ${result.deleted} วัน) — คอลัมน์ชื่อนี้จะไม่ถูกเพิ่มอีก`);
  } catch (error) {
    console.error('[JUICE-PAGE] Delete failed:', error);
    backToJuicePage(res, `❌ ทำรายการไม่สำเร็จ: ${error.message}`);
  }
});

// ---------- MTD / YTD Report ----------

const THAI_MONTHS = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];

function pad2(n) { return String(n).padStart(2, '0'); }
function ymd(y, m, d) { return `${y}-${pad2(m)}-${pad2(d)}`; }
function formatInt(n) { return (n == null || Number.isNaN(n)) ? '-' : Math.round(n).toLocaleString('en-US'); }
function formatThaiDate(y, m, d) { return `${d} ${THAI_MONTHS[m - 1]} ${y}`; }

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// The look the built-in juices ship with. What the DB holds wins (db.JUICE_ICONS /
// db.JUICE_COLORS): set on the juice page for any juice, and for an auto-added one also
// the colour read off its plan column. An auto-added juice with neither gets 🧃 and green.
const NEW_JUICE_ICON = '🧃';
function juiceIcon(category) {
  return db.JUICE_ICONS[category] || CATEGORY_ICONS[category] || NEW_JUICE_ICON;
}
function juiceColor(category) {
  return db.JUICE_COLORS[category] || CATEGORY_COLORS[category] || '#4CAF50';
}
// Hand-picked section backgrounds for the daily report; any other colour gets a computed tint.
const CATEGORY_TINTS = { orange: '#fff3e0', yuzu: '#fffde7', pop: '#e1f5fe', mixed: '#f3e5f5', tomato: '#ffebee', pineapple: '#f9fbe7', watermelon: '#fce4ec' };
const CATEGORY_ICONS = { orange: '🍊', yuzu: '🍋', pop: '🍹', mixed: '🍇', tomato: '🍅', pineapple: '🍍', watermelon: '🍉' };
const CATEGORY_COLORS = { orange: '#ff9800', yuzu: '#f9a825', pop: '#03a9f4', mixed: '#9c27b0', tomato: '#f44336', pineapple: '#cddc39', watermelon: '#e57373' };

// Pick progress-bar color based on achievement % (Volume vs Target).
function progressBarColor(pct) {
  if (pct >= 100) return '#1b7a2f';   // hit/exceeded target — dark green
  if (pct >= 80)  return '#4CAF50';   // close — green
  if (pct >= 50)  return '#ff9800';   // on the way — orange
  return '#e53935';                   // far off — red
}

function buildProgressCell(volume, target, yoy) {
  const yoyLine = yoy > 0
    ? `<div class="progress-yoy">ปีก่อน: ${formatInt(yoy)}</div>`
    : `<div class="progress-yoy" style="color:#aaa;">ปีก่อน: —</div>`;

  if (!target || target <= 0) {
    return `
      <td>
        <div class="progress-wrap">
          <div class="progress-head"><b>${formatInt(volume)}</b> <span style="color:#888;">/ ยังไม่ได้ตั้ง target</span></div>
          <div class="progress-track"><div class="progress-fill" style="width:0%;"></div></div>
          ${yoyLine}
        </div>
      </td>`;
  }

  const pct = (volume / target) * 100;
  const fillWidth = Math.min(pct, 100);
  const color = progressBarColor(pct);
  return `
    <td>
      <div class="progress-wrap">
        <div class="progress-head">
          <b>${formatInt(volume)}</b> <span style="color:#888;">/ ${formatInt(target)}</span>
          <span class="progress-pct" style="color:${color};">(${pct.toFixed(1)}%)</span>
        </div>
        <div class="progress-track"><div class="progress-fill" style="width:${fillWidth}%; background:${color};"></div></div>
        ${yoyLine}
      </div>
    </td>`;
}

function buildReportRow(category, volume, target, yoy) {
  const hasYoy = yoy > 0;
  let growthCell;
  if (!hasYoy) {
    growthCell = `<td style="color:#1b7a2f;font-weight:600;font-size:16px;">+${formatInt(volume)}</td>`;
  } else {
    const pct = ((volume - yoy) / yoy) * 100;
    const color = pct >= 0 ? '#1b7a2f' : '#c62828';
    const sign = pct >= 0 ? '+' : '';
    growthCell = `<td style="color:${color};font-weight:600;font-size:16px;">${sign}${pct.toFixed(1)}%</td>`;
  }
  const name = db.CATEGORY_NAMES[category] || category;
  const icon = juiceIcon(category);
  const color = juiceColor(category);
  return `
    <tr>
      <td style="text-align:left;font-weight:600;color:${color};white-space:nowrap;">${icon} ${escapeHtml(name)}</td>
      ${buildProgressCell(volume, target, yoy)}
      ${growthCell}
    </tr>`;
}

function buildReportTotalRow(volume, totalTarget, totalPy, accentColor) {
  const hasYoy = totalPy > 0;
  let growthCell;
  if (!hasYoy) {
    growthCell = `<td style="color:#1b7a2f;font-weight:700;font-size:16px;">+${formatInt(volume)}</td>`;
  } else {
    const pct = ((volume - totalPy) / totalPy) * 100;
    const color = pct >= 0 ? '#1b7a2f' : '#c62828';
    const sign = pct >= 0 ? '+' : '';
    growthCell = `<td style="color:${color};font-weight:700;font-size:16px;">${sign}${pct.toFixed(1)}%</td>`;
  }
  const accent = accentColor || '#333';
  return `
    <tr style="background:#f5f5f5;border-top:2px solid #ccc;">
      <td style="text-align:left;font-weight:700;color:${accent};white-space:nowrap;">รวม</td>
      ${buildProgressCell(volume, totalTarget, totalPy)}
      ${growthCell}
    </tr>`;
}

// Shared CSS/nav used by /mtd-report and /targets — mirrors the /daily-report look.
function reportSharedStyles() {
  return `
    body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
    .container { max-width: 100%; margin: 0 auto; background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    h1 { text-align: center; color: #333; margin-bottom: 10px; }
    .subtitle { text-align: center; color: #555; margin-bottom: 24px; }
    .nav { margin-bottom: 20px; text-align: center; }
    .nav a { padding: 10px 20px; background-color: #4CAF50; color: white; text-decoration: none; margin: 0 5px; border-radius: 5px; display: inline-block; }
    .nav a:hover { background-color: #45a049; }
    .filter-section { margin-bottom: 30px; text-align: center; padding: 20px; background-color: #f0f0f0; border-radius: 8px; }
    .filter-section label { font-weight: bold; font-size: 16px; margin-right: 10px; }
    .filter-section select { padding: 8px 15px; font-size: 16px; border: 2px solid #4CAF50; border-radius: 5px; background-color: white; cursor: pointer; transition: border-color 0.3s; }
    .filter-section select:hover { border-color: #45a049; }
    .filter-section select:focus { outline: none; border-color: #2196F3; }
    .summary { display: flex; justify-content: space-around; margin-bottom: 30px; flex-wrap: wrap; }
    .summary-card { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 10px; min-width: 200px; text-align: center; margin: 10px; }
    .summary-number { font-size: 32px; font-weight: bold; }
    .summary-label { font-size: 14px; opacity: 0.9; }
    .section { margin-bottom: 40px; padding: 20px; border-radius: 8px; background-color: #fafafa; border-left: 5px solid #4CAF50; }
    .section-title { font-size: 22px; font-weight: bold; margin-bottom: 15px; color: #333; }
    .chart-container { background: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    table { border-collapse: collapse; width: 100%; background-color: white; box-shadow: 0 1px 3px rgba(0,0,0,0.1); font-size: 14px; }
    th, td { border: 1px solid #ccc; padding: 10px 8px; text-align: center; }
    th { background-color: #4CAF50; color: white; font-weight: bold; white-space: nowrap; }
    tr:nth-child(even) { background-color: #f9f9f9; }
    tr:hover { background-color: #f1f1f1; }
    .no-data { padding: 20px; text-align: center; color: #666; font-style: italic; }
    .notice { background:#e8f5e9; border:1px solid #a5d6a7; color:#2e7d32; padding:10px 12px; border-radius:4px; margin: 0 auto 16px auto; max-width: 600px; text-align: center; }
    form.target-form table { max-width: 700px; margin: 0 auto; }
    form.target-form input[type=number] { width: 160px; padding: 8px; font-size: 14px; border: 2px solid #ccc; border-radius: 4px; }
    form.target-form input[type=number]:focus { outline: none; border-color: #4CAF50; }
    .save-btn { display: block; margin: 20px auto 0; padding: 12px 28px; font-size: 16px; background:#4CAF50; color:#fff; border:0; border-radius:5px; cursor:pointer; }
    .save-btn:hover { background:#45a049; }
    .progress-wrap { min-width: 240px; text-align: left; padding: 4px 0; }
    .progress-head { font-size: 14px; margin-bottom: 6px; }
    .progress-head b { font-size: 16px; }
    .progress-pct { margin-left: 6px; font-weight: 600; }
    .progress-track { width: 100%; height: 10px; background: #e0e0e0; border-radius: 5px; overflow: hidden; }
    .progress-fill { height: 100%; transition: width 0.4s ease; border-radius: 5px; }
    .progress-yoy { margin-top: 6px; font-size: 12px; color: #777; }
    .no-yoy { display: inline-block; color: #888; font-size: 13px; font-style: italic; }
  `;
}

function reportNav() {
  return `
    <div class="nav" style="display:flex; flex-wrap:wrap; justify-content:center; gap:10px; margin:0 0 24px 0;">
      <a href="/latest-ocr">Original Table</a>
      <a href="/transformed-data">Transformed Data</a>
      <a href="/daily-report">Daily Report</a>
      <a href="/mtd-report">MTD/YTD Report</a>
      <a href="/targets">Manage Targets</a>
      <a href="/detection-logs">Detection Logs</a>
      <a href="/send-notification">Send Notification</a>
      <a href="/test">Test OCR</a>
    </div>`;
}

app.get('/mtd-report', async (req, res) => {
  try {
    const now = new Date();
    const y = parseInt(req.query.year, 10) || now.getFullYear();
    const m = parseInt(req.query.month, 10) || (now.getMonth() + 1);
    if (m < 1 || m > 12) return res.status(400).send('Invalid month');

    const lastDayOfMonth = new Date(y, m, 0).getDate();
    const isCurrentMonth = (y === now.getFullYear() && m === now.getMonth() + 1);
    const isFuture = (y > now.getFullYear()) || (y === now.getFullYear() && m > now.getMonth() + 1);

    // Resolve "as of" day: current month uses the latest date that has data
    // (or today if no data yet), past/future months use the last day of month.
    let d;
    if (isCurrentMonth) {
      const latest = await db.getLatestDayWithData(y, m);
      d = latest || Math.min(now.getDate(), lastDayOfMonth);
    } else if (isFuture) {
      d = lastDayOfMonth;
    } else {
      d = lastDayOfMonth;
    }
    const prevY = y - 1;

    const mtdStart = ymd(y, m, 1);
    const mtdEnd = ymd(y, m, d);
    const ytdStart = ymd(y, 1, 1);
    const ytdEnd = ymd(y, m, d);
    const mtdPyStart = ymd(prevY, m, 1);
    const mtdPyEnd = ymd(prevY, m, d);
    const ytdPyStart = ymd(prevY, 1, 1);
    const ytdPyEnd = ymd(prevY, m, d);

    const [
      mtdAgg, ytdAgg, mtdPyAgg, ytdPyAgg,
      currentTargets, ytdTargets
    ] = await Promise.all([
      db.getAggregateByCategory(mtdStart, mtdEnd),
      db.getAggregateByCategory(ytdStart, ytdEnd),
      db.getAggregateByCategory(mtdPyStart, mtdPyEnd),
      db.getAggregateByCategory(ytdPyStart, ytdPyEnd),
      db.getTargets(y, m),
      db.getTargetsYTD(y, m)
    ]);

    // Show only categories that have sales in each respective period.
    // Canonical order first, then any new category names.
    const orderCategories = (keys) => {
      const setObj = new Set(keys);
      const arr = [];
      for (const c of db.CATEGORIES) if (setObj.has(c)) arr.push(c);
      for (const c of setObj) if (!db.CATEGORIES.includes(c)) arr.push(c);
      return arr;
    };
    const mtdOrdered = orderCategories(Object.keys(mtdAgg).filter(c => (mtdAgg[c] || 0) > 0));
    const ytdOrdered = orderCategories(Object.keys(ytdAgg).filter(c => (ytdAgg[c] || 0) > 0));

    const mtdRows = mtdOrdered.map(c => buildReportRow(c, mtdAgg[c] || 0, currentTargets[c] || 0, mtdPyAgg[c] || 0)).join('');
    const ytdRows = ytdOrdered.map(c => buildReportRow(c, ytdAgg[c] || 0, ytdTargets[c] || 0, ytdPyAgg[c] || 0)).join('');
    const sumTargetsFor = (obj, cats) => cats.reduce((s, c) => s + (Number(obj[c]) || 0), 0);
    const mtdTotalTarget = sumTargetsFor(currentTargets, mtdOrdered);
    const ytdTotalTarget = sumTargetsFor(ytdTargets, ytdOrdered);

    const asOfLabel = formatThaiDate(y, m, d);
    const pyLabel = formatThaiDate(prevY, m, d);

    const yearOptions = [];
    for (let yy = now.getFullYear() - 2; yy <= now.getFullYear() + 1; yy++) {
      yearOptions.push(`<option value="${yy}"${yy === y ? ' selected' : ''}>${yy}</option>`);
    }
    const monthOptions = THAI_MONTHS.map((name, i) =>
      `<option value="${i + 1}"${i + 1 === m ? ' selected' : ''}>${name}</option>`
    ).join('');

    const asOfNote = isCurrentMonth
      ? `<span style="color:#555;font-size:13px;">(เดือนปัจจุบัน — ข้อมูลถึงวันล่าสุดที่บันทึก)</span>`
      : '';

    const sumVals = obj => Object.values(obj).reduce((s, v) => s + (v || 0), 0);
    const mtdTotalVol = sumVals(mtdAgg);
    const ytdTotalVol = sumVals(ytdAgg);
    const mtdTotalPy = sumVals(mtdPyAgg);
    const ytdTotalPy = sumVals(ytdPyAgg);

    const hasAny = mtdOrdered.length > 0 || ytdOrdered.length > 0;
    const emptyRow = '<tr><td colspan="3" class="no-data">ไม่มีสินค้าที่มียอดขายในช่วงนี้</td></tr>';

    res.send(`<!DOCTYPE html>
<html>
<head>
  <title>MTD / YTD Report</title>
  <meta charset="utf-8">
  <style>${reportSharedStyles()}</style>
</head>
<body>
  <div class="container">
    ${reportNav()}
    <h1>📈 MTD / YTD Report</h1>
    <div class="subtitle">ข้อมูล ณ วันที่ <b>${asOfLabel}</b> เทียบปีก่อน <b>${pyLabel}</b><br>${asOfNote}</div>

    <div class="filter-section">
      <label for="month-filter">Month:</label>
      <select id="month-filter" onchange="applyFilter()">${monthOptions}</select>
      <label for="year-filter" style="margin-left: 20px;">Year:</label>
      <select id="year-filter" onchange="applyFilter()">${yearOptions}</select>
    </div>

    ${!hasAny ? '<div class="no-data" style="padding: 40px;">ไม่มีข้อมูลในช่วงที่เลือก</div>' : ''}

    ${hasAny ? `
    <div class="section">
      <div class="section-title">🗓️ Month-to-Date (${THAI_MONTHS[m - 1]} ${y})</div>
      <div style="overflow-x:auto;">
      <table>
        <thead>
          <tr>
            <th>Category</th>
            <th>Volume / Target (${THAI_MONTHS[m - 1]})</th>
            <th>vs ปีก่อน (${prevY})</th>
          </tr>
        </thead>
        <tbody>${mtdRows || emptyRow}${mtdOrdered.length > 0 ? buildReportTotalRow(mtdTotalVol, mtdTotalTarget, mtdTotalPy, '#333') : ''}</tbody>
      </table>
      </div>
    </div>

    <div class="section" style="border-left-color: #11998e;">
      <div class="section-title" style="color:#11998e;">📅 Year-to-Date (${y})</div>
      <div style="overflow-x:auto;">
      <table>
        <thead>
          <tr>
            <th>Category</th>
            <th>Volume / Target YTD (ม.ค.–${THAI_MONTHS[m - 1]})</th>
            <th>vs ปีก่อน (${prevY})</th>
          </tr>
        </thead>
        <tbody>${ytdRows || emptyRow}${ytdOrdered.length > 0 ? buildReportTotalRow(ytdTotalVol, ytdTotalTarget, ytdTotalPy, '#11998e') : ''}</tbody>
      </table>
      </div>
    </div>
    ` : ''}
  </div>

  <script>
    function applyFilter() {
      const month = document.getElementById('month-filter').value;
      const year = document.getElementById('year-filter').value;
      window.location.href = '/mtd-report?month=' + month + '&year=' + year;
    }
  </script>
</body>
</html>`);
  } catch (err) {
    console.error('[MTD] Error:', err);
    res.status(500).send('Error loading MTD report: ' + escapeHtml(err.message));
  }
});

// ---------- Target management ----------

app.get('/targets', async (req, res) => {
  try {
    const now = new Date();
    const y = parseInt(req.query.year, 10) || now.getFullYear();
    const m = parseInt(req.query.month, 10) || (now.getMonth() + 1);

    const [currentTargets, knownCats] = await Promise.all([
      db.getTargets(y, m),
      db.getKnownCategories()
    ]);

    const appearing = new Set([...Object.keys(currentTargets), ...knownCats, ...db.CATEGORIES]);
    const ordered = [];
    for (const c of db.CATEGORIES) if (appearing.has(c)) ordered.push(c);
    for (const c of appearing) if (!db.CATEGORIES.includes(c)) ordered.push(c);

    const rows = ordered.map(c => {
      const current = currentTargets[c] || 0;
      const name = db.CATEGORY_NAMES[c] || c;
      const icon = juiceIcon(c);
      const color = juiceColor(c);
      return `
        <tr>
          <td style="text-align:left;font-weight:600;color:${color};">${icon} ${escapeHtml(name)}</td>
          <td><input type="number" min="0" step="1" name="target_${escapeHtml(c)}" value="${current}"></td>
        </tr>`;
    }).join('');

    const yearOptions = [];
    for (let yy = now.getFullYear() - 2; yy <= now.getFullYear() + 1; yy++) {
      yearOptions.push(`<option value="${yy}"${yy === y ? ' selected' : ''}>${yy}</option>`);
    }
    const monthOptions = THAI_MONTHS.map((name, i) =>
      `<option value="${i + 1}"${i + 1 === m ? ' selected' : ''}>${name}</option>`
    ).join('');

    const saved = req.query.saved === '1';

    res.send(`<!DOCTYPE html>
<html>
<head>
  <title>Manage Targets</title>
  <meta charset="utf-8">
  <style>${reportSharedStyles()}</style>
</head>
<body>
  <div class="container">
    ${reportNav()}
    <h1>🎯 Manage Monthly Targets</h1>
    <div class="subtitle">ตั้งยอด target รายเดือนของแต่ละ category เพื่อใช้ใน MTD/YTD Report</div>

    ${saved ? '<div class="notice">✅ บันทึก target เรียบร้อย</div>' : ''}

    <div class="filter-section">
      <label for="month-filter">Month:</label>
      <select id="month-filter" onchange="applyFilter()">${monthOptions}</select>
      <label for="year-filter" style="margin-left: 20px;">Year:</label>
      <select id="year-filter" onchange="applyFilter()">${yearOptions}</select>
    </div>

    <form class="target-form" method="post" action="/targets">
      <input type="hidden" name="year" value="${y}">
      <input type="hidden" name="month" value="${m}">
      <table>
        <thead>
          <tr>
            <th>Category</th>
            <th>Target (${THAI_MONTHS[m - 1]} ${y})</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
      <button class="save-btn" type="submit">💾 บันทึก</button>
    </form>
  </div>

  <script>
    function applyFilter() {
      const month = document.getElementById('month-filter').value;
      const year = document.getElementById('year-filter').value;
      window.location.href = '/targets?month=' + month + '&year=' + year;
    }
  </script>
</body>
</html>`);
  } catch (err) {
    console.error('[TARGETS] Error:', err);
    res.status(500).send('Error loading targets: ' + escapeHtml(err.message));
  }
});

app.post('/targets', async (req, res) => {
  try {
    const y = parseInt(req.body.year, 10);
    const m = parseInt(req.body.month, 10);
    if (!y || !m || m < 1 || m > 12) {
      return res.status(400).send('Invalid year or month');
    }
    const entries = Object.entries(req.body)
      .filter(([k]) => k.startsWith('target_'))
      .map(([k, v]) => [k.slice('target_'.length), parseInt(v, 10)])
      .filter(([cat, val]) => cat && Number.isFinite(val) && val >= 0);

    for (const [cat, val] of entries) {
      await db.upsertTarget(y, m, cat, val);
    }
    console.log(`[TARGETS] Saved ${entries.length} targets for ${y}-${pad2(m)}`);
    res.redirect(`/targets?year=${y}&month=${m}&saved=1`);
  } catch (err) {
    console.error('[TARGETS] Save error:', err);
    res.status(500).send('Error saving targets: ' + escapeHtml(err.message));
  }
});

// Test page for uploading images
app.get('/test', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>Test OCR</title>
      <meta charset="utf-8">
      <style>
        body {
          font-family: Arial, sans-serif;
          max-width: 800px;
          margin: 50px auto;
          padding: 20px;
        }
        .upload-area {
          border: 2px dashed #ccc;
          border-radius: 10px;
          padding: 40px;
          text-align: center;
          cursor: pointer;
          transition: all 0.3s;
        }
        .upload-area:hover {
          border-color: #4CAF50;
          background-color: #f9f9f9;
        }
        .upload-area.dragging {
          border-color: #4CAF50;
          background-color: #e8f5e9;
        }
        input[type="file"] {
          display: none;
        }
        .result {
          margin-top: 30px;
          padding: 20px;
          background-color: #f5f5f5;
          border-radius: 5px;
          display: none;
        }
        .loading {
          display: none;
          margin-top: 20px;
          text-align: center;
        }
        .nav {
          margin-bottom: 20px;
        }
        .nav a {
          padding: 10px 20px;
          background-color: #4CAF50;
          color: white;
          text-decoration: none;
          margin-right: 10px;
          border-radius: 5px;
        }
      </style>
    </head>
    <body>
      ${reportNav()}

      <h1>🧪 Test OCR</h1>
      <p>Upload an Excel screenshot to test OCR extraction without using LINE bot</p>

      <div class="upload-area" id="uploadArea">
        <h2>📤 Drop image here or click to upload</h2>
        <p>Supports: PNG, JPG, JPEG</p>
        <input type="file" id="fileInput" accept="image/*">
      </div>

      <div class="loading" id="loading">
        <h3>⏳ Processing image...</h3>
      </div>

      <div class="result" id="result"></div>

      <script>
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const loading = document.getElementById('loading');
        const result = document.getElementById('result');

        uploadArea.addEventListener('click', () => fileInput.click());

        uploadArea.addEventListener('dragover', (e) => {
          e.preventDefault();
          uploadArea.classList.add('dragging');
        });

        uploadArea.addEventListener('dragleave', () => {
          uploadArea.classList.remove('dragging');
        });

        uploadArea.addEventListener('drop', (e) => {
          e.preventDefault();
          uploadArea.classList.remove('dragging');
          const file = e.dataTransfer.files[0];
          if (file && file.type.startsWith('image/')) {
            processImage(file);
          }
        });

        fileInput.addEventListener('change', (e) => {
          const file = e.target.files[0];
          if (file) {
            processImage(file);
          }
        });

        async function processImage(file) {
          loading.style.display = 'block';
          result.style.display = 'none';

          try {
            const response = await fetch('/test-ocr', {
              method: 'POST',
              headers: {
                'Content-Type': file.type
              },
              body: file
            });

            const data = await response.json();
            loading.style.display = 'none';

            if (data.success) {
              result.style.display = 'block';
              if (data.isExcelScreenshot) {
                // Whether the data actually reached the database is the point of this
                // page — detecting a spreadsheet but silently failing validation used
                // to look identical to a successful upload.
                const saved = data.recorded
                  ? \`<p style="color: #1b7a2f;"><strong>💾 บันทึกแล้ว:</strong> \${data.date} — \${data.categories.join(', ')}</p>\`
                  : \`<p style="color: #c62828;"><strong>⚠️ ไม่ได้บันทึกลงฐานข้อมูล:</strong> \${data.reason || 'ไม่พบสาเหตุ ดูที่หน้า Detection Logs'}</p>\`;
                const text = data.extractedText || '';
                result.innerHTML = \`
                  <h3>✅ Excel Screenshot Detected!</h3>
                  \${saved}
                  <p><strong>Extracted Text Length:</strong> \${text.length} characters</p>
                  <p><strong>Tables Found:</strong> \${data.tableData ? data.tableData.length : 0}</p>
                  <h4>Preview:</h4>
                  <pre style="max-height: 300px; overflow-y: auto; background: white; padding: 10px;">\${text.substring(0, 1000)}...</pre>
                  <p><a href="/latest-ocr" target="_blank">View Full Results</a> | <a href="/daily-report" target="_blank">Daily Report</a></p>
                \`;
              } else {
                result.innerHTML = '<h3>❌ Not detected as Excel screenshot</h3>';
              }
            } else {
              result.innerHTML = '<h3>❌ Error: ' + data.error + '</h3>';
            }
          } catch (error) {
            loading.style.display = 'none';
            result.style.display = 'block';
            result.innerHTML = '<h3>❌ Error: ' + error.message + '</h3>';
          }
        }
      </script>
    </body>
    </html>
  `);
});

// ---------- Missing-data watchdog ----------
//
// A LINE webhook that never arrives leaves no trace anywhere in this app — no log
// line, no detection_logs row — so a whole day of data can go missing silently
// (this is what happened on 09/08/2026). The watchdog closes that gap by checking
// the data itself rather than the delivery path, and alerting on any expected day
// that has no record at all.

const BANGKOK_TZ = 'Asia/Bangkok';

// Current date/hour in Bangkok, independent of the server's own timezone.
function bangkokNow() {
  const parts = Object.fromEntries(
    new Intl.DateTimeFormat('en-CA', {
      timeZone: BANGKOK_TZ,
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', hourCycle: 'h23'
    }).formatToParts(new Date()).map(p => [p.type, p.value])
  );
  return { date: `${parts.year}-${parts.month}-${parts.day}`, hour: parseInt(parts.hour, 10) % 24 };
}

// Shift a YYYY-MM-DD string by whole days. UTC arithmetic keeps DST and local
// timezone out of it — the string in, string out contract is what callers compare.
function shiftDate(isoDate, days) {
  const d = new Date(`${isoDate}T00:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

function isoToThaiDisplay(isoDate) {
  const [y, m, d] = isoDate.split('-');
  return `${d}/${m}/${y}`;
}

async function runWatchdog({ dryRun = false } = {}) {
  const { date: today, hour } = bangkokNow();

  // Today only counts as overdue once the plan is normally in (WATCHDOG_CUTOFF_HOUR),
  // otherwise every morning would report the current day as missing.
  const includeToday = hour >= WATCHDOG_CUTOFF_HOUR;
  const lastExpected = includeToday ? today : shiftDate(today, -1);
  const firstExpected = shiftDate(lastExpected, -(WATCHDOG_LOOKBACK_DAYS - 1));

  const recorded = new Set(await db.getRecordedDates(firstExpected, lastExpected));
  const expected = [];
  for (let d = firstExpected; d <= lastExpected; d = shiftDate(d, 1)) expected.push(d);
  const missing = expected.filter(d => !recorded.has(d));

  console.log(`[WATCHDOG] Window ${firstExpected}..${lastExpected} — ${recorded.size} day(s) recorded, ${missing.length} missing`);

  if (missing.length === 0) {
    return { window: { from: firstExpected, to: lastExpected }, missing: [], alerted: [], alreadyAlerted: [] };
  }

  // Report each missing day once. Days already alerted stay in the response so the
  // problem remains visible, but they do not re-notify the group every run.
  const toAlert = [];
  const alreadyAlerted = [];
  for (const date of missing) {
    if (await db.hasWatchdogAlert(date)) alreadyAlerted.push(date);
    else toAlert.push(date);
  }

  if (toAlert.length === 0 || dryRun) {
    console.log(`[WATCHDOG] ${dryRun ? 'Dry run' : 'Nothing new'} — missing: ${missing.join(', ')}`);
    return { window: { from: firstExpected, to: lastExpected }, missing, alerted: [], alreadyAlerted, dryRun };
  }

  const dateList = toAlert.map(isoToThaiDisplay).join('\n');
  const message =
    `⚠️ ไม่พบข้อมูล Plan ส่งสินค้า\n\n` +
    `วันที่ยังไม่มีข้อมูลในระบบ:\n${dateList}\n\n` +
    `กรุณาส่งรูปเข้ากลุ่มอีกครั้ง หรืออัปโหลดที่\n` +
    `https://shinsen.yushi-marketing.com/test`;

  const sent = await sendAlertToGroups(message);
  // Only mark as alerted when the push actually went out, so a LINE outage does not
  // silently swallow the one notification a missing day ever gets.
  if (sent) {
    for (const date of toAlert) await db.recordWatchdogAlert(date);
  }

  return { window: { from: firstExpected, to: lastExpected }, missing, alerted: sent ? toAlert : [], alreadyAlerted, sent };
}

// Hit this from an external scheduler (cron / uptime monitor) — it is the only part
// of the system that can notice data missing because the app was never reached.
// ?dry=1    — report without notifying
// ?strict=1 — answer 409 when a day is missing, so `curl -f` exits non-zero and the
//             scheduler that called it raises the alarm. This is the escape hatch for
//             setups that want no LINE alert at all: leave ALERT_GROUP_IDS empty and
//             let cron mail the failure instead.
app.get('/watchdog', async (req, res) => {
  try {
    // Drain first, and inside this request: an image still sitting in the queue is
    // not a missing day, and awaiting it here is what holds the process open long
    // enough for the OCR to finish.
    const drained = await drainQueue({ reason: 'watchdog', limit: 5 })
      .catch(error => ({ error: error.message }));
    const result = await runWatchdog({ dryRun: req.query.dry === '1' });
    const status = (req.query.strict === '1' && result.missing.length > 0) ? 409 : 200;
    res.status(status).json({ success: true, queue: drained, ...result });
  } catch (error) {
    console.error('[WATCHDOG] Run failed:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Queue worker endpoint. Point a per-minute cron at this — it is what makes an
// image survive the process being reaped mid-OCR. The work runs inside the request
// on purpose: Passenger keeps a process alive exactly as long as it owes a response.
app.get('/process-queue', async (req, res) => {
  try {
    const limit = Math.min(Math.max(parseInt(req.query.limit, 10) || 3, 1), 10);
    const result = await drainQueue({ reason: 'cron', limit });
    res.json({ success: true, ...result, queue: await db.getQueueStats() });
  } catch (error) {
    console.error('[QUEUE] Drain endpoint failed:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// The app is spawned per request on this host, so a boot-time run gives the watchdog
// coverage on ordinary traffic even without a scheduler. It cannot replace one: a
// total outage means no boot, and therefore no check.
if (process.env.WATCHDOG_ON_BOOT?.toLowerCase() !== 'false') {
  setTimeout(() => {
    // Queue first — the watchdog should not report a day missing while its image is
    // still waiting to be processed. This run may be cut short by the process being
    // reaped, which is fine: unfinished rows stay claimable.
    drainQueue({ reason: 'boot' })
      .catch(error => console.error('[QUEUE] Boot drain failed:', error.message))
      .then(() => runWatchdog())
      .catch(error => console.error('[WATCHDOG] Boot run failed:', error.message));
  }, 5000).unref?.();
}

// Passenger already restarts this process 35-50 times a day as a matter of course.
// A stray rejection killing it on top of that is pure downtime, and the queue makes
// a survivable process strictly better than a clean exit: staying up is what lets an
// in-flight OCR finish. Both handlers log loudly so nothing fails silently.
process.on('unhandledRejection', (reason) => {
  console.error('[FATAL] Unhandled promise rejection (staying up):', reason?.stack || reason);
});
process.on('uncaughtException', (error) => {
  console.error('[FATAL] Uncaught exception (staying up):', error?.stack || error);
});

app.get('/health', async (req, res) => {
  const health = {
    status: 'ok',
    time: new Date().toISOString(),
    uptimeSeconds: Math.round(process.uptime())
  };

  try {
    health.db = (await db.testConnection()) ? 'ok' : 'error';
    const { date } = bangkokNow();
    const recent = await db.getRecordedDates(shiftDate(date, -14), date);
    health.lastRecordedDate = recent.length ? isoToThaiDisplay(recent[recent.length - 1]) : null;
    // A growing queue means the drain cron is not running or not keeping up — the
    // one symptom that would otherwise look identical to "no images were sent".
    health.queue = await db.getQueueStats();
  } catch (error) {
    health.db = 'error';
    health.error = error.message;
  }

  if (health.db !== 'ok') {
    health.status = 'error';
    return res.status(503).json(health);
  }
  res.json(health);
});

// Network diagnostics — no SSH needed. Open in a browser and paste the JSON.
// Isolates whether the OpenAI "Premature close" is IPv6, egress block, proxy,
// request-size, or SDK-specific. Safe: never prints the API key.
app.get('/diag', async (req, res) => {
  const dnsp = require('dns').promises;
  const out = { node: process.version, time: new Date().toISOString() };

  out.proxyEnv = {
    HTTP_PROXY: process.env.HTTP_PROXY || process.env.http_proxy || null,
    HTTPS_PROXY: process.env.HTTPS_PROXY || process.env.https_proxy || null,
    NO_PROXY: process.env.NO_PROXY || process.env.no_proxy || null,
  };

  try {
    out.dns = {
      A: await dnsp.resolve4('api.openai.com').then(a => a.slice(0, 5)).catch(e => 'ERR ' + e.code),
      AAAA: await dnsp.resolve6('api.openai.com').then(a => a.slice(0, 5)).catch(e => 'ERR ' + e.code),
    };
  } catch (e) { out.dns = 'ERR ' + e.message; }

  // Raw HTTPS GET /v1/models with a forced IP family (isolates network from SDK)
  function rawGet(family) {
    return new Promise((resolve) => {
      const started = Date.now();
      const agent = new https.Agent({ family, keepAlive: false });
      const r = https.get('https://api.openai.com/v1/models', { agent, timeout: 15000 }, (resp) => {
        let bytes = 0;
        const ip = resp.socket ? resp.socket.remoteAddress : null;
        const status = resp.statusCode;
        resp.on('data', c => (bytes += c.length));
        resp.on('end', () => resolve({ ok: true, status, ip, bytes, ms: Date.now() - started }));
      });
      r.on('timeout', () => r.destroy(new Error('timeout 15s')));
      r.on('error', (e) => resolve({ ok: false, error: e.message, code: e.code, ms: Date.now() - started }));
    });
  }

  out.rawGet_IPv4 = await rawGet(4);
  out.rawGet_IPv6 = await rawGet(6);
  out.rawGet_default = await rawGet(undefined);

  // Raw POST /v1/chat/completions via Node https (bypasses the SDK / node-fetch).
  // If this works but the SDK fails -> it's an SDK/node-fetch issue, not the network.
  function rawPostChat(family, keepAlive) {
    return new Promise((resolve) => {
      const started = Date.now();
      const body = JSON.stringify({ model: 'gpt-4o', max_tokens: 5, messages: [{ role: 'user', content: 'say hi' }] });
      const agent = new https.Agent({ family, keepAlive });
      const r = https.request('https://api.openai.com/v1/chat/completions', {
        method: 'POST', agent, timeout: 20000,
        headers: {
          'Content-Type': 'application/json',
          'Content-Length': Buffer.byteLength(body),
          'Authorization': 'Bearer ' + (process.env.OPENAI_API_KEY || ''),
        },
      }, (resp) => {
        let bytes = 0, snippet = '';
        const ip = resp.socket ? resp.socket.remoteAddress : null;
        const status = resp.statusCode;
        resp.on('data', c => { bytes += c.length; if (snippet.length < 160) snippet += c.toString(); });
        resp.on('end', () => resolve({ ok: true, status, ip, bytes, ms: Date.now() - started, snippet: snippet.slice(0, 160) }));
      });
      r.on('timeout', () => r.destroy(new Error('timeout 20s')));
      r.on('error', (e) => resolve({ ok: false, error: e.message, code: e.code, ms: Date.now() - started }));
      r.write(body); r.end();
    });
  }

  out.rawPostChat_IPv4_noKeepAlive = await rawPostChat(4, false);
  out.rawPostChat_IPv4_keepAlive = await rawPostChat(4, true);

  // SDK via node-fetch v2 (current config: IPv4 + keep-alive agent)
  try {
    const t = Date.now();
    const r = await openai.chat.completions.create({
      model: 'gpt-4o', max_tokens: 5, messages: [{ role: 'user', content: 'say hi' }],
    });
    out.sdkChatTest_nodeFetch = { ok: true, ms: Date.now() - t, reply: r.choices[0].message.content };
  } catch (e) { out.sdkChatTest_nodeFetch = { ok: false, error: e.message, code: e.code }; }

  // SDK forced to use Node 22 native fetch (undici) instead of node-fetch v2
  try {
    const t = Date.now();
    const altClient = new OpenAI({
      apiKey: process.env.OPENAI_API_KEY, timeout: 20000, maxRetries: 0,
      fetch: (...a) => globalThis.fetch(...a),
    });
    const r = await altClient.chat.completions.create({
      model: 'gpt-4o', max_tokens: 5, messages: [{ role: 'user', content: 'say hi' }],
    });
    out.sdkChatTest_nativeFetch = { ok: true, ms: Date.now() - t, reply: r.choices[0].message.content };
  } catch (e) { out.sdkChatTest_nativeFetch = { ok: false, error: e.message, code: e.code }; }

  res.type('application/json').send(JSON.stringify(out, null, 2));
});

const PORT = process.env.PORT || 3000;
// Detection logs endpoint
app.get('/detection-logs', async (req, res) => {
  try {
    const logs = await loadDetectionLogs();

    let logsHTML = '';
    if (logs.length === 0) {
      logsHTML = '<tr><td colspan="6" class="no-data">No detection logs yet</td></tr>';
    } else {
      logs.forEach(log => {
        const timestamp = new Date(log.timestamp).toLocaleString('en-US', {
          year: 'numeric',
          month: '2-digit',
          day: '2-digit',
          hour: '2-digit',
          minute: '2-digit',
          second: '2-digit',
          hour12: false
        });

        const statusIcon = log.status === 'success' ? '✅' : log.status === 'failed' ? '❌' : '⚠️';
        const statusClass = log.status === 'success' ? 'status-success' : log.status === 'failed' ? 'status-failed' : 'status-error';
        const statusText = log.status === 'success' ? 'Success' : log.status === 'failed' ? 'Failed' : 'Error';

        const groupIdHTML = log.groupId
          ? `<code style="background-color: #4CAF50; color: white; padding: 4px 8px; border-radius: 3px; font-weight: bold;">${log.groupId}</code>`
          : '<span style="color: #999;">-</span>';

        const userIdHTML = log.userId
          ? `<code>${log.userId}</code>`
          : '<span style="color: #999;">-</span>';

        let detailsHTML = '';
        if (log.status === 'success') {
          detailsHTML = `<strong>Date:</strong> ${log.date || 'N/A'}<br>
                         <strong>Categories:</strong> ${log.categories ? log.categories.join(', ') : 'N/A'}<br>
                         <strong>Records Created:</strong> ${log.recordsCreated || 0}`;
        } else {
          detailsHTML = `<strong>Reason:</strong> ${log.reason || 'Unknown'}`;
        }

        logsHTML += `
          <tr>
            <td>${timestamp}</td>
            <td><code>${log.messageId || 'unknown'}</code></td>
            <td>${groupIdHTML}</td>
            <td>${userIdHTML}</td>
            <td class="${statusClass}">${statusIcon} ${statusText}</td>
            <td class="details">${detailsHTML}</td>
          </tr>
        `;
      });
    }

    res.send(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>Detection Logs</title>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
          }
          .container {
            max-width: 100%;
            margin: 0 auto;
            background-color: white;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          }
          h1 {
            text-align: center;
            color: #333;
            margin-bottom: 30px;
          }
          .nav {
            margin-bottom: 20px;
            text-align: center;
          }
          .nav a {
            padding: 10px 20px;
            background-color: #4CAF50;
            color: white;
            text-decoration: none;
            margin: 0 5px;
            border-radius: 5px;
            display: inline-block;
          }
          .nav a:hover {
            background-color: #45a049;
          }
          table {
            border-collapse: collapse;
            width: 100%;
            background-color: white;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
          }
          th, td {
            border: 1px solid #ddd;
            padding: 12px;
            text-align: left;
          }
          th {
            background-color: #2196F3;
            color: white;
            font-weight: bold;
          }
          tr:nth-child(even) {
            background-color: #f9f9f9;
          }
          tr:hover {
            background-color: #f1f1f1;
          }
          .status-success {
            color: #4CAF50;
            font-weight: bold;
          }
          .status-failed {
            color: #f44336;
            font-weight: bold;
          }
          .status-error {
            color: #ff9800;
            font-weight: bold;
          }
          .details {
            font-size: 14px;
            line-height: 1.6;
          }
          .no-data {
            padding: 40px;
            text-align: center;
            color: #666;
            font-style: italic;
          }
          code {
            background-color: #f5f5f5;
            padding: 2px 6px;
            border-radius: 3px;
            font-family: monospace;
            font-size: 12px;
          }
          .summary {
            display: flex;
            justify-content: space-around;
            margin-bottom: 30px;
            flex-wrap: wrap;
          }
          .summary-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 10px;
            min-width: 150px;
            text-align: center;
            margin: 10px;
          }
          .summary-number {
            font-size: 36px;
            font-weight: bold;
          }
          .summary-label {
            font-size: 14px;
            opacity: 0.9;
          }
        </style>
      </head>
      <body>
        <div class="container">
          ${reportNav()}

          <h1>📋 Image Detection Logs</h1>

          <div class="summary">
            <div class="summary-card" style="background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%);">
              <div class="summary-number">${logs.filter(l => l.status === 'success').length}</div>
              <div class="summary-label">Successful</div>
            </div>
            <div class="summary-card" style="background: linear-gradient(135deg, #f44336 0%, #d32f2f 100%);">
              <div class="summary-number">${logs.filter(l => l.status === 'failed').length}</div>
              <div class="summary-label">Failed</div>
            </div>
            <div class="summary-card" style="background: linear-gradient(135deg, #ff9800 0%, #f57c00 100%);">
              <div class="summary-number">${logs.filter(l => l.status === 'error').length}</div>
              <div class="summary-label">Errors</div>
            </div>
            <div class="summary-card">
              <div class="summary-number">${logs.length}</div>
              <div class="summary-label">Total Logs</div>
            </div>
          </div>

          <table>
            <tr>
              <th>Timestamp</th>
              <th>Message ID</th>
              <th>Group ID</th>
              <th>User ID</th>
              <th>Status</th>
              <th>Details</th>
            </tr>
            ${logsHTML}
          </table>
        </div>
      </body>
      </html>
    `);
  } catch (error) {
    console.error('Error loading detection logs:', error);
    res.status(500).send('Error loading detection logs');
  }
});

// API endpoint to send manual notification to LINE groups
app.post('/api/send-notification', async (req, res) => {
  try {
    const { message } = req.body;

    if (!message || message.trim() === '') {
      return res.status(400).json({ success: false, error: 'Message is required' });
    }

    if (NOTIFICATION_GROUP_IDS.length === 0) {
      return res.status(400).json({ success: false, error: 'No notification groups configured in .env' });
    }

    console.log('[MANUAL-NOTI] Sending manual notification to groups');
    console.log('[MANUAL-NOTI] Message:', message);
    console.log('[MANUAL-NOTI] Target groups:', NOTIFICATION_GROUP_IDS);

    const results = [];
    for (const groupId of NOTIFICATION_GROUP_IDS) {
      try {
        await client.pushMessage({
          to: groupId,
          messages: [{
            type: 'text',
            text: message,
          }],
        });
        console.log(`[MANUAL-NOTI] Successfully sent to group: ${groupId}`);
        results.push({ groupId, success: true });
      } catch (error) {
        console.error(`[MANUAL-NOTI] Failed to send to group ${groupId}:`, error.message);
        results.push({ groupId, success: false, error: error.message });
      }
    }

    const successCount = results.filter(r => r.success).length;
    res.json({
      success: true,
      message: `Sent to ${successCount}/${NOTIFICATION_GROUP_IDS.length} groups`,
      results
    });
  } catch (error) {
    console.error('[MANUAL-NOTI] Error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Page for sending manual notifications
app.get('/send-notification', (req, res) => {
  const groupCount = NOTIFICATION_GROUP_IDS.length;
  const groupIds = NOTIFICATION_GROUP_IDS.map(id => `<code>${id}</code>`).join('<br>');

  res.send(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>Send Notification</title>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <style>
        body {
          font-family: Arial, sans-serif;
          margin: 20px;
          background-color: #f5f5f5;
        }
        .container {
          max-width: 600px;
          margin: 0 auto;
          background-color: white;
          padding: 30px;
          border-radius: 10px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        h1 {
          text-align: center;
          color: #333;
          margin-bottom: 30px;
        }
        .nav {
          margin-bottom: 20px;
          text-align: center;
        }
        .nav a {
          padding: 10px 20px;
          background-color: #4CAF50;
          color: white;
          text-decoration: none;
          margin: 0 5px;
          border-radius: 5px;
          display: inline-block;
        }
        .nav a:hover {
          background-color: #45a049;
        }
        .info-box {
          background-color: #e3f2fd;
          border-left: 4px solid #2196F3;
          padding: 15px;
          margin-bottom: 20px;
          border-radius: 4px;
        }
        .info-box strong {
          color: #1976D2;
        }
        .form-group {
          margin-bottom: 20px;
        }
        label {
          display: block;
          margin-bottom: 8px;
          font-weight: bold;
          color: #333;
        }
        textarea {
          width: 100%;
          padding: 12px;
          border: 2px solid #ddd;
          border-radius: 5px;
          font-size: 16px;
          resize: vertical;
          min-height: 120px;
          box-sizing: border-box;
        }
        textarea:focus {
          border-color: #4CAF50;
          outline: none;
        }
        button {
          width: 100%;
          padding: 15px;
          background-color: #4CAF50;
          color: white;
          border: none;
          border-radius: 5px;
          font-size: 18px;
          cursor: pointer;
          transition: background-color 0.3s;
        }
        button:hover {
          background-color: #45a049;
        }
        button:disabled {
          background-color: #ccc;
          cursor: not-allowed;
        }
        .result {
          margin-top: 20px;
          padding: 15px;
          border-radius: 5px;
          display: none;
        }
        .result.success {
          background-color: #e8f5e9;
          border-left: 4px solid #4CAF50;
          color: #2e7d32;
        }
        .result.error {
          background-color: #ffebee;
          border-left: 4px solid #f44336;
          color: #c62828;
        }
        code {
          background-color: #f5f5f5;
          padding: 2px 6px;
          border-radius: 3px;
          font-family: monospace;
          font-size: 12px;
        }
      </style>
    </head>
    <body>
      <div class="container">
        ${reportNav()}

        <h1>📢 Send Notification to LINE Group</h1>

        <div class="info-box">
          <strong>Target Groups (${groupCount}):</strong><br>
          ${groupCount > 0 ? groupIds : '<span style="color: #f44336;">No groups configured in .env</span>'}
        </div>

        <form id="notificationForm">
          <div class="form-group">
            <label for="message">Message:</label>
            <textarea id="message" placeholder="Enter your notification message here..." required></textarea>
          </div>
          <button type="submit" id="sendBtn" ${groupCount === 0 ? 'disabled' : ''}>
            📤 Send Notification
          </button>
        </form>

        <div id="result" class="result"></div>
      </div>

      <script>
        document.getElementById('notificationForm').addEventListener('submit', async function(e) {
          e.preventDefault();

          const message = document.getElementById('message').value.trim();
          const sendBtn = document.getElementById('sendBtn');
          const result = document.getElementById('result');

          if (!message) {
            result.className = 'result error';
            result.textContent = 'Please enter a message';
            result.style.display = 'block';
            return;
          }

          sendBtn.disabled = true;
          sendBtn.textContent = '⏳ Sending...';
          result.style.display = 'none';

          try {
            const response = await fetch('/api/send-notification', {
              method: 'POST',
              headers: {
                'Content-Type': 'application/json'
              },
              body: JSON.stringify({ message })
            });

            const data = await response.json();

            if (data.success) {
              result.className = 'result success';
              result.innerHTML = '✅ ' + data.message;
            } else {
              result.className = 'result error';
              result.innerHTML = '❌ ' + data.error;
            }
          } catch (error) {
            result.className = 'result error';
            result.innerHTML = '❌ Error: ' + error.message;
          }

          result.style.display = 'block';
          sendBtn.disabled = false;
          sendBtn.textContent = '📤 Send Notification';
        });
      </script>
    </body>
    </html>
  `);
});

// Start the HTTP server unless explicitly imported for offline testing
// (set NODE_ENV=test before requiring this file to skip listen).
// NOTE: do NOT gate this on `require.main === module` — Phusion Passenger
// (Plesk production) loads this file via require(), so require.main is the
// Passenger loader, not this module. That makes the check false, listen()
// never runs, and Passenger fails with "A timeout occurred while spawning
// an application process".
if (process.env.NODE_ENV !== 'test') {
  app.listen(PORT, () => {
    console.log(`LINE Bot server is running on port ${PORT}`);
    console.log(`Webhook URL: http://localhost:${PORT}/webhook`);
  });
}

// Exported for offline testing/validation harnesses (no effect when run as server)
module.exports = {
  getRowProductValue,
  detectProductColumns,
  detectUnknownJuiceColumns,
  buildTablesFromResult,
  validateTableData,
  detectCrateColumns,
  deriveLayout,
  extractCDCTotals,
  extractYodruamTotals,
  preprocessTableData,
  CDC_NAME_MAPPING
};
