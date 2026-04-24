require('dotenv').config();
const express = require('express');
const line = require('@line/bot-sdk');
const OpenAI = require('openai');
const { DocumentAnalysisClient, AzureKeyCredential } = require('@azure/ai-form-recognizer');
const fs = require('fs').promises;
const fsSync = require('fs');
const path = require('path');
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

console.log('[CONFIG] Notification groups configured:', NOTIFICATION_GROUP_IDS.length);
console.log('[CONFIG] Notification group IDs:', NOTIFICATION_GROUP_IDS);
console.log('[CONFIG] Auto-notify:', AUTO_NOTIFY);

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

// Check if date already exists in records (uses MySQL)
async function isDateRecorded(date, category) {
  try {
    return await db.isDateRecorded(date, category);
  } catch (error) {
    console.error('[RECORD] Error checking if date recorded:', error);
    return false;
  }
}

// Product detection configuration - keywords to identify each product in column headers
const PRODUCT_DETECTION = {
  orange: {
    keywords: ['ส้ม', 'orange', 'น้ำส้ม'],
    dbCategory: 'orange'
  },
  yuzu: {
    keywords: ['ยูซุ', 'yuzu'],
    dbCategory: 'yuzu'
  },
  pop: {
    keywords: ['pop', 'shinsen pop', 'ป๊อป', 'มุกป๊อป'],
    dbCategory: 'pop'
  },
  mixed: {
    keywords: ['ผลไม้รวม', 'น้ำผลไม้รวม', 'mixed fruit'],
    dbCategory: 'mixed'
  },
  tomato: {
    keywords: ['tomato', 'มะเขือเทศ'],
    dbCategory: 'tomato'
  }
};

// Helper function to parse numbers from OCR
// Handles: commas, periods as thousand separators (OCR misreads), newlines/whitespace
function parseOCRNumber(str) {
  if (!str) return 0;
  let s = str.toString().trim();

  // Remove whitespace and newlines
  s = s.replace(/[\s]/g, '');

  // Check if period is used as thousand separator (e.g., "1.331" should be 1331)
  // Pattern: digit(s), period, exactly 3 digits, end of string
  if (/^\d+\.\d{3}$/.test(s)) {
    // Period is thousand separator, remove it
    s = s.replace(/\./g, '');
  }

  // Remove commas (thousand separators)
  s = s.replace(/,/g, '');

  return parseFloat(s) || 0;
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

// Detect product columns from table headers
function detectProductColumns(table) {
  const detectedProducts = {};

  console.log('[DETECT] Scanning table for product columns...');

  // Scan first 10 rows for header information
  for (let rowIdx = 0; rowIdx < Math.min(10, table.length); rowIdx++) {
    const row = table[rowIdx];
    if (!row) continue;

    // Check each column in this row
    for (let colIdx = 0; colIdx < row.length; colIdx++) {
      const cellText = row[colIdx] ? row[colIdx].toString().toLowerCase().trim() : '';
      if (!cellText) continue;

      // Check against each product's keywords
      for (const [productKey, productConfig] of Object.entries(PRODUCT_DETECTION)) {
        // Skip if already detected
        if (detectedProducts[productKey]) continue;

        // Check if any keyword matches
        const matched = productConfig.keywords.some(keyword =>
          cellText.includes(keyword.toLowerCase())
        );

        if (matched) {
          detectedProducts[productKey] = {
            column: colIdx,
            dbCategory: productConfig.dbCategory,
            foundIn: `Row ${rowIdx}, Col ${colIdx}: "${row[colIdx]}"`
          };
          console.log(`[DETECT] Found ${productKey} in column ${colIdx} (Row ${rowIdx}: "${row[colIdx]}")`);
        }
      }
    }
  }

  console.log(`[DETECT] Detected ${Object.keys(detectedProducts).length} products:`, Object.keys(detectedProducts));
  return detectedProducts;
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

  // Check all detected product columns for validation
  table.forEach((row, rowIndex) => {
    if (!row || row.length < 3) return;

    const c0Cell = row[0] ? row[0].toString().trim() : '';

    // Sum values from all product columns for validation
    productColumns.forEach(colIdx => {
      const value = row[colIdx] ? parseOCRNumber(row[colIdx]) : 0;

      if (c0Cell.includes('FC33') && c0Cell.includes('หาดใหญ่')) {
        hadyaiSum += value;
      }
      if (c0Cell.includes('FC07')) {
        phuketSum += value;
      }
    });
  });

  console.log(`[VALIDATION] FC33 หาดใหญ่ sum: ${hadyaiSum}, FC07 ภูเก็ต sum: ${phuketSum}`);

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
function extractCDCTotals(table, columnIndex, yodruamTotal = 0) {
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
      for (let col = 2; col <= 4; col++) {
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
      const value = parseOCRNumber(row[columnIndex]);
      if (value > 0) totalSum += value;
    }
  }

  console.log(`[CDC] Extracting CDC values, excluding totals row index: ${totalsRowIndex}`);

  // Extract CDC values row by row, tracking individual contributions for ratio correction
  const CRATE_TOTAL_COL = 6; // C6 = ตะกร้า รวม (total crates)
  const cdcRows = [];

  table.forEach((row, rowIndex) => {
    if (!row || row.length < 2) return;
    if (rowIndex === totalsRowIndex) return;

    CDC_NAMES.forEach(cdcName => {
      const fullCdcName = CDC_NAME_MAPPING[cdcName];
      let searchCell;

      if (fullCdcName === 'คลังบางบัวทอง') {
        searchCell = row[0] ? row[0].toString().trim() : '';
      } else if (fullCdcName.startsWith('คลัง')) {
        searchCell = row[1] ? row[1].toString().trim() : '';
      } else {
        searchCell = row[0] ? row[0].toString().trim() : '';
      }

      if (searchCell.includes(cdcName) && row[columnIndex]) {
        const value = parseOCRNumber(row[columnIndex]);
        const crateTotal = row[CRATE_TOTAL_COL] ? parseOCRNumber(row[CRATE_TOTAL_COL]) : 0;
        if (value > 0) {
          cdcRows.push({ rowIndex, cdcName: fullCdcName, value, crateTotal });
        }
      }
    });
  });

  // Ratio correction: fix truncated values where bottle/crate ratio is suspiciously low
  // Normal range: ~25-42 bottles per crate. Below 20 indicates a missing digit.
  // Only correct when crates >= 2 to avoid false positives on genuinely small orders
  const rawSum = cdcRows.reduce((s, r) => s + r.value, 0);
  const targetTotal = yodruamTotal > 0 ? yodruamTotal : totalSum;

  if (rawSum !== targetTotal && targetTotal > 0) {
    for (const entry of cdcRows) {
      if (entry.crateTotal < 2) continue;
      const ratio = entry.value / entry.crateTotal;
      if (ratio >= 20 || entry.value <= 1) continue;

      // Value appears truncated - try appending digit 0-9
      let bestCandidate = null;
      for (let d = 0; d <= 9; d++) {
        const candidate = entry.value * 10 + d;
        const candRatio = candidate / entry.crateTotal;
        if (candRatio >= 25 && candRatio <= 42) {
          bestCandidate = { val: candidate, ratio: candRatio, digit: d };
          break;
        }
      }

      if (bestCandidate) {
        console.log(`[RATIO-FIX] Row ${entry.rowIndex} (${entry.cdcName}): ${entry.value} → ${bestCandidate.val} (ratio ${ratio.toFixed(1)} → ${bestCandidate.ratio.toFixed(1)}, crates=${entry.crateTotal})`);
        entry.value = bestCandidate.val;
      }
    }

    const correctedSum = cdcRows.reduce((s, r) => s + r.value, 0);
    if (correctedSum !== rawSum) {
      console.log(`[RATIO-FIX] Corrected: ${rawSum} → ${correctedSum} (ยอดรวม=${targetTotal}, remaining diff=${targetTotal - correctedSum})`);
    }
  }

  // Build CDC totals from (possibly corrected) row values
  for (const entry of cdcRows) {
    cdcTotals[entry.cdcName] += entry.value;
  }

  return { cdcTotals, totalSum };
}

// Extract Khon Kaen Laos value for orange category
function extractKhonKaenLaos(table, columnIndex) {
  for (let i = table.length - 1; i >= 0; i--) {
    const row = table[i];
    if (!row || !row[0]) continue;

    const c0 = row[0].toString().trim();
    if (c0.includes('ขอนแก่น') && row[columnIndex]) {
      const value = parseOCRNumber(row[columnIndex]);
      console.log(`[LAOS] Found ขอนแก่น Laos value: ${value} in row ${i}`);
      return value;
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
async function recordDailyData(tableData, extractedText = '', rawResult = null) {
  if (!tableData || tableData.length === 0) {
    console.log('No table data to record');
    return { success: false, reason: 'No table data found' };
  }

  const table = tableData[0];

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
    const { cdcTotals, totalSum } = extractCDCTotals(table, columnIndex, yodruamValue);

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
    const khonKaenLaosValue = category === 'orange' ? extractKhonKaenLaos(table, columnIndex) : 0;

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

  if (results.length === 0) {
    return { success: false, reason: 'All detected products already recorded for this date' };
  }

  return { success: true, date: dateStr, results: results };
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
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.path}`);
  next();
});

// LINE webhook - must be before express.json() to preserve raw body for signature validation
app.post('/webhook', line.middleware(config), async (req, res) => {
  try {
    console.log('[WEBHOOK] Received webhook request');
    console.log('[WEBHOOK] Number of events:', req.body.events?.length || 0);
    const results = await Promise.all(req.body.events.map(handleEvent));
    res.json({ success: true });
  } catch (err) {
    console.error('[ERROR] Error handling webhook:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// Add JSON parsing middleware for other endpoints (after LINE webhook)
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

async function handleEvent(event) {
  console.log('[EVENT] Event type:', event.type);

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

          await client.replyMessage({
            replyToken: event.replyToken,
            messages: [{
              type: 'text',
              text: replyMessage,
            }],
          });

          console.log('Success message sent to user');

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
      console.error('Error processing image:', error);
      // Silent mode: don't reply on errors
    }
  } else {
    console.log('Non-image message received:', message.type);
  }

  return null;
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

    if (result.tables) {
      for (const table of result.tables) {
        const tableRows = [];
        const maxRow = Math.max(...table.cells.map(c => c.rowIndex)) + 1;
        const maxCol = Math.max(...table.cells.map(c => c.columnIndex)) + 1;

        for (let i = 0; i < maxRow; i++) {
          tableRows[i] = new Array(maxCol).fill('');
        }

        for (const cell of table.cells) {
          tableRows[cell.rowIndex][cell.columnIndex] = cell.content || '';
        }

        tableData.push(tableRows);
      }
    }

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
      recordResult = await recordDailyData(tableData, extractedText, result);
      if (recordResult && recordResult.success) {
        console.log('Daily data recorded successfully');
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

  // Initialize totals for all CDC locations
  const cdcTotals = {};
  cdcOrder.forEach(cdc => {
    cdcTotals[cdc] = 0;
  });

  // Iterate through table rows and sum values based on CDC name
  table.forEach((row, rowIndex) => {
    if (!row || row.length < 2) return;

    // Find matching CDC name
    cdcNames.forEach(cdcName => {
      const fullCdcName = cdcNameMapping[cdcName];

      // Determine which column to search based on CDC name
      let searchColumnIndex;
      let searchCell;

      if (fullCdcName.startsWith('คลัง')) {
        // Search in C1 for locations starting with "คลัง"
        searchColumnIndex = 1;
        searchCell = row[1] ? row[1].toString().trim() : '';
      } else {
        // Search in C0 for other locations
        searchColumnIndex = 0;
        searchCell = row[0] ? row[0].toString().trim() : '';
      }

      // Check if this row matches the CDC name
      if (searchCell.includes(cdcName)) {
        // Get value from the specified column (C2 for orange, C3 for yuzu)
        if (row[columnIndex]) {
          const value = parseOCRNumber(row[columnIndex]);

          if (value > 0) {
            cdcTotals[fullCdcName] += value;
          }
        }
      }
    });
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
      <div class="nav">
        <a href="/latest-ocr">Original Table</a>
        <a href="/transformed-data">Transformed Data</a>
        <a href="/daily-report">Daily Report</a>
      </div>

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
      <div class="nav">
        <a href="/latest-ocr">Original Table</a>
        <a href="/transformed-data">Transformed Data</a>
        <a href="/daily-report">Daily Report</a>
      </div>

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
    const categoryData = {
      orange: { records: records.orange || [], name: 'Orange', thaiName: 'น้ำส้ม', icon: '🍊', bgColor: '#fff3e0', borderColor: '#ff9800' },
      yuzu: { records: records.yuzu || [], name: 'Yuzu', thaiName: 'ยูซุ', icon: '🍋', bgColor: '#fffde7', borderColor: '#fdd835' },
      pop: { records: records.pop || [], name: 'Shinsen Pop', thaiName: 'Shinsen Pop', icon: '🍹', bgColor: '#e1f5fe', borderColor: '#03a9f4' },
      mixed: { records: records.mixed || [], name: 'Mixed Fruit', thaiName: 'น้ำผลไม้รวม', icon: '🍇', bgColor: '#f3e5f5', borderColor: '#9c27b0' },
      tomato: { records: records.tomato || [], name: 'Tomato Yuzu', thaiName: 'Tomato Yuzu', icon: '🍅', bgColor: '#ffebee', borderColor: '#f44336' }
    };

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
          <div class="nav">
            <a href="/latest-ocr">Original Table</a>
            <a href="/transformed-data">Transformed Data</a>
            <a href="/daily-report">Daily Report</a>
            <a href="/detection-logs">Detection Logs</a>
            <a href="/send-notification" style="background-color: #2196F3;">Send Notification</a>
          </div>

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
      const extractedText = await performOCR(imageBuffer);

      res.json({
        success: true,
        isExcelScreenshot: true,
        extractedText: extractedText,
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

// ---------- MTD / YTD Report ----------

const THAI_MONTHS = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];

function pad2(n) { return String(n).padStart(2, '0'); }
function ymd(y, m, d) { return `${y}-${pad2(m)}-${pad2(d)}`; }
function formatInt(n) { return (n == null || Number.isNaN(n)) ? '-' : Math.round(n).toLocaleString('en-US'); }
function formatThaiDate(y, m, d) { return `${d} ${THAI_MONTHS[m - 1]} ${y}`; }

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

const CATEGORY_ICONS = { orange: '🍊', yuzu: '🍋', pop: '🍹', mixed: '🍇', tomato: '🍅' };
const CATEGORY_COLORS = { orange: '#ff9800', yuzu: '#f9a825', pop: '#03a9f4', mixed: '#9c27b0', tomato: '#f44336' };

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
  const icon = CATEGORY_ICONS[category] || '';
  const color = CATEGORY_COLORS[category] || '#4CAF50';
  return `
    <tr>
      <td style="text-align:left;font-weight:600;color:${color};white-space:nowrap;">${icon} ${escapeHtml(name)}</td>
      ${buildProgressCell(volume, target, yoy)}
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
    <div class="nav">
      <a href="/daily-report">Daily Report</a>
      <a href="/mtd-report">MTD/YTD Report</a>
      <a href="/targets">Manage Targets</a>
      <a href="/detection-logs">Detection Logs</a>
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
    const mtdTotalGrowth = mtdTotalPy > 0 ? ((mtdTotalVol - mtdTotalPy) / mtdTotalPy) * 100 : null;
    const ytdTotalGrowth = ytdTotalPy > 0 ? ((ytdTotalVol - ytdTotalPy) / ytdTotalPy) * 100 : null;

    const growthBadge = (pct) => {
      if (pct === null) return '<span style="opacity:0.8;">N/A</span>';
      const sign = pct >= 0 ? '+' : '';
      return `${sign}${pct.toFixed(1)}%`;
    };

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

    ${hasAny ? `
    <div class="summary">
      <div class="summary-card">
        <div class="summary-number">${formatInt(mtdTotalVol)}</div>
        <div class="summary-label">Volume MTD · Growth ${growthBadge(mtdTotalGrowth)}</div>
      </div>
      <div class="summary-card" style="background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);">
        <div class="summary-number">${formatInt(ytdTotalVol)}</div>
        <div class="summary-label">Volume YTD · Growth ${growthBadge(ytdTotalGrowth)}</div>
      </div>
    </div>
    ` : ''}

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
        <tbody>${mtdRows || emptyRow}</tbody>
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
        <tbody>${ytdRows || emptyRow}</tbody>
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
      const icon = CATEGORY_ICONS[c] || '';
      const color = CATEGORY_COLORS[c] || '#4CAF50';
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
      <div class="nav">
        <a href="/latest-ocr">Original Table</a>
        <a href="/transformed-data">Transformed Data</a>
        <a href="/daily-report">Daily Report</a>
        <a href="/detection-logs">Detection Logs</a>
        <a href="/test">Test OCR</a>
      </div>

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
                result.innerHTML = \`
                  <h3>✅ Excel Screenshot Detected!</h3>
                  <p><strong>Extracted Text Length:</strong> \${data.extractedText.length} characters</p>
                  <p><strong>Tables Found:</strong> \${data.tableData ? data.tableData.length : 0}</p>
                  <h4>Preview:</h4>
                  <pre style="max-height: 300px; overflow-y: auto; background: white; padding: 10px;">\${data.extractedText.substring(0, 1000)}...</pre>
                  <p><a href="/latest-ocr" target="_blank">View Full Results</a></p>
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

app.get('/health', (req, res) => {
  res.json({ status: 'ok' });
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
          <div class="nav">
            <a href="/latest-ocr">Original Table</a>
            <a href="/transformed-data">Transformed Data</a>
            <a href="/daily-report">Daily Report</a>
            <a href="/detection-logs">Detection Logs</a>
            <a href="/send-notification" style="background-color: #2196F3;">Send Notification</a>
          </div>

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
        <div class="nav">
          <a href="/daily-report">Daily Report</a>
          <a href="/detection-logs">Detection Logs</a>
          <a href="/send-notification">Send Notification</a>
        </div>

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

app.listen(PORT, () => {
  console.log(`LINE Bot server is running on port ${PORT}`);
  console.log(`Webhook URL: http://localhost:${PORT}/webhook`);
});
