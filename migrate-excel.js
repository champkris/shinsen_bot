#!/usr/bin/env node
/**
 * Migration script: Excel to MySQL
 * Imports historical daily sales data from Excel file
 *
 * Excel structure:
 * - Multiple sheets (one per product: orange, yuzu/pop)
 * - Each sheet contains monthly tables
 * - Tables have date, CDC locations, and totals
 */

require('dotenv').config();
const XLSX = require('xlsx');
const path = require('path');
const db = require('./db');

const EXCEL_FILE = path.join(__dirname, 'docs', 'จำนวนสั่งชินเซน 7-11 2024 2025 2026.xlsx');

// Month names in Thai that might appear in the Excel
const THAI_MONTHS = [
  'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
];

const ENGLISH_MONTHS = [
  'january', 'february', 'march', 'april', 'may', 'june',
  'july', 'august', 'september', 'october', 'november', 'december'
];

// CDC name variations that might appear in Excel
const CDC_VARIATIONS = {
  'บางบัวทอง': ['บางบัวทอง', 'คลังบางบัวทอง', 'BBT'],
  'นครราชสีมา': ['นครราชสีมา', 'โคราช', 'NMA'],
  'นครสวรรค์': ['นครสวรรค์', 'NSW'],
  'ชลบุรี': ['ชลบุรี', 'CBR'],
  'มหาชัย': ['มหาชัย', 'คลังมหาชัย', 'MHC'],
  'สุวรรณภูมิ': ['สุวรรณภูมิ', 'คลังสุวรรณภูมิ', 'SVB'],
  'หาดใหญ่': ['หาดใหญ่', 'HDY'],
  'ภูเก็ต': ['ภูเก็ต', 'PKT'],
  'เชียงใหม่': ['เชียงใหม่', 'CNX'],
  'สุราษฎร์': ['สุราษฎร์', 'สุราษฎร์ธานี', 'SRT'],
  'ขอนแก่น': ['ขอนแก่น', 'KKN']
};

// Mapping from short CDC names to full Thai names used in db.js
const CDC_NAME_TO_FULL = {
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

function analyzeWorkbook(workbook) {
  console.log('\n=== Analyzing Excel File ===\n');
  console.log('Sheet names:', workbook.SheetNames);

  for (const sheetName of workbook.SheetNames) {
    console.log(`\n--- Sheet: ${sheetName} ---`);
    const sheet = workbook.Sheets[sheetName];
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
    console.log(`Range: A1 to ${XLSX.utils.encode_col(range.e.c)}${range.e.r + 1}`);

    // Get first few rows to understand structure
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    console.log(`Total rows: ${data.length}`);

    // Show first 10 rows
    console.log('First 10 rows:');
    for (let i = 0; i < Math.min(10, data.length); i++) {
      const row = data[i].slice(0, 8).map(c => String(c).substring(0, 15));
      console.log(`  Row ${i}: ${JSON.stringify(row)}`);
    }
  }
}

function detectCategory(sheetName) {
  const name = sheetName.toLowerCase();
  const nameThai = sheetName;

  // Orange (ส้ม)
  if (name.includes('orange') || nameThai.includes('ส้ม')) {
    return 'orange';
  }

  // Shinsen Pop - check before yuzu since "pop" shouldn't match yuzu
  if (name.includes('pop') || name.includes('shinsen pop')) {
    return 'pop';
  }

  // Tomato Yuzu - check before general yuzu
  if (name.includes('tomato') || nameThai.includes('มะเขือเท')) {
    return 'tomato';
  }

  // Yuzu (ยูซุ)
  if (name.includes('yuzu') || nameThai.includes('ยูซุ')) {
    return 'yuzu';
  }

  // Mixed fruit juice (น้ำผลไม้รวม)
  if (name.includes('mixed') || nameThai.includes('ผลไม้รวม') || nameThai.includes('น้ำผลไม้รวม')) {
    return 'mixed';
  }

  return null;
}

function findCdcName(cellValue) {
  if (!cellValue) return null;
  const value = String(cellValue).trim();

  for (const [cdcKey, variations] of Object.entries(CDC_VARIATIONS)) {
    for (const variation of variations) {
      if (value.includes(variation)) {
        return CDC_NAME_TO_FULL[cdcKey];
      }
    }
  }
  return null;
}

function parseExcelDate(value) {
  if (!value) return null;

  // If it's an Excel serial date number
  if (typeof value === 'number') {
    const date = XLSX.SSF.parse_date_code(value);
    if (date) {
      const day = String(date.d).padStart(2, '0');
      const month = String(date.m).padStart(2, '0');
      const year = date.y;
      return `${day}/${month}/${year}`;
    }
  }

  // If it's a string date
  const strValue = String(value).trim();

  // Try DD/MM/YYYY format
  const match1 = strValue.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (match1) {
    return `${match1[1].padStart(2, '0')}/${match1[2].padStart(2, '0')}/${match1[3]}`;
  }

  // Try YYYY-MM-DD format
  const match2 = strValue.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match2) {
    return `${match2[3]}/${match2[2]}/${match2[1]}`;
  }

  return null;
}

function parseNumericValue(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return Math.round(value);
  const num = parseFloat(String(value).replace(/,/g, ''));
  return isNaN(num) ? 0 : Math.round(num);
}

async function parseAndMigrateSheet(sheet, sheetName, category, dryRun = false) {
  console.log(`\nProcessing sheet: ${sheetName} (${category})`);

  const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  const records = [];

  // Look for header row with CDC names
  let headerRowIndex = -1;
  let cdcColumns = {}; // Maps column index to CDC name
  let dateColumnIndex = -1;
  let totalColumnIndex = -1;

  // Scan for header row (contains multiple CDC names)
  for (let rowIdx = 0; rowIdx < Math.min(50, data.length); rowIdx++) {
    const row = data[rowIdx];
    let cdcCount = 0;
    const tempCdcColumns = {};

    for (let colIdx = 0; colIdx < row.length; colIdx++) {
      const cellValue = String(row[colIdx]).trim();

      // Check for date column
      if (cellValue.includes('วันที่') || cellValue.toLowerCase().includes('date')) {
        dateColumnIndex = colIdx;
      }

      // Check for total column
      if (cellValue.includes('รวม') || cellValue.toLowerCase().includes('total') || cellValue.includes('ยอดรวม')) {
        if (!cellValue.includes('CDC')) { // Avoid matching "รวม CDC"
          totalColumnIndex = colIdx;
        }
      }

      // Check for CDC name
      const cdcName = findCdcName(cellValue);
      if (cdcName) {
        tempCdcColumns[colIdx] = cdcName;
        cdcCount++;
      }
    }

    // If we found multiple CDC columns, this is likely the header row
    if (cdcCount >= 5) {
      headerRowIndex = rowIdx;
      cdcColumns = tempCdcColumns;
      console.log(`  Found header row at index ${rowIdx}`);
      console.log(`  CDC columns found:`, Object.values(cdcColumns));
      console.log(`  Date column: ${dateColumnIndex}, Total column: ${totalColumnIndex}`);
      break;
    }
  }

  if (headerRowIndex === -1) {
    console.log(`  Warning: Could not find header row with CDC columns`);
    return { processed: 0, skipped: 0 };
  }

  // Parse data rows
  let processed = 0;
  let skipped = 0;

  for (let rowIdx = headerRowIndex + 1; rowIdx < data.length; rowIdx++) {
    const row = data[rowIdx];
    if (!row || row.length === 0) continue;

    // Try to parse date from first few columns
    let dateStr = null;
    for (let colIdx = 0; colIdx <= Math.max(dateColumnIndex, 2); colIdx++) {
      dateStr = parseExcelDate(row[colIdx]);
      if (dateStr) break;
    }

    if (!dateStr) {
      // Try the designated date column
      if (dateColumnIndex >= 0) {
        dateStr = parseExcelDate(row[dateColumnIndex]);
      }
    }

    if (!dateStr) continue; // Skip non-data rows

    // Build CDC totals
    const cdcTotals = {};
    for (const cdcName of Object.values(CDC_NAME_TO_FULL)) {
      cdcTotals[cdcName] = 0;
    }

    for (const [colIdx, cdcName] of Object.entries(cdcColumns)) {
      const value = parseNumericValue(row[parseInt(colIdx)]);
      cdcTotals[cdcName] = value;
    }

    // Get total sum
    let totalSum = 0;
    if (totalColumnIndex >= 0 && row[totalColumnIndex]) {
      totalSum = parseNumericValue(row[totalColumnIndex]);
    } else {
      // Calculate from CDC values
      totalSum = Object.values(cdcTotals).reduce((a, b) => a + b, 0);
    }

    // Skip if all values are 0
    if (totalSum === 0 && Object.values(cdcTotals).every(v => v === 0)) {
      skipped++;
      continue;
    }

    const record = {
      date: dateStr,
      timestamp: new Date().toISOString(),
      fc33HadyaiSum: cdcTotals['หาดใหญ่'] || 0,
      totalSum: totalSum,
      cdcTotals: cdcTotals
    };

    // Add Laos/Cambodia fields for orange category
    if (category === 'orange') {
      record.khonKaenLaos = 0; // Historical data may not have this
      record.khonKaenCambodia = 0;
    }

    records.push(record);

    if (!dryRun) {
      try {
        // Check if already exists
        const exists = await db.isDateRecorded(dateStr, category);
        if (!exists) {
          await db.saveDailyRecord(record, category);
          processed++;
        } else {
          skipped++;
        }
      } catch (error) {
        console.error(`  Error saving record for ${dateStr}:`, error.message);
        skipped++;
      }
    } else {
      processed++;
    }

    // Progress indicator
    if (processed % 50 === 0) {
      process.stdout.write(`\r  Processed: ${processed}, Skipped: ${skipped}`);
    }
  }

  console.log(`\r  Processed: ${processed}, Skipped: ${skipped}`);
  return { processed, skipped, records: dryRun ? records : undefined };
}

async function main() {
  const args = process.argv.slice(2);
  const dryRun = args.includes('--dry-run');
  const analyze = args.includes('--analyze');

  console.log('=================================');
  console.log('Excel to MySQL Migration Script');
  console.log('=================================');

  if (dryRun) {
    console.log('\n** DRY RUN MODE - No data will be saved **\n');
  }

  // Load Excel file
  let workbook;
  try {
    workbook = XLSX.readFile(EXCEL_FILE);
    console.log(`Loaded Excel file: ${EXCEL_FILE}`);
  } catch (error) {
    console.error(`Failed to load Excel file: ${error.message}`);
    process.exit(1);
  }

  // Analyze mode - just show structure
  if (analyze) {
    analyzeWorkbook(workbook);
    process.exit(0);
  }

  // Test database connection (unless dry run)
  if (!dryRun) {
    const connected = await db.testConnection();
    if (!connected) {
      console.error('\nFailed to connect to MySQL database.');
      console.error('Please ensure:');
      console.error('1. MySQL server is running');
      console.error('2. Database "shinsen_bot" exists (run schema.sql first)');
      console.error('3. .env file has correct MYSQL_* credentials');
      process.exit(1);
    }
  }

  // Process each sheet
  const results = {
    orange: { processed: 0, skipped: 0 },
    yuzu: { processed: 0, skipped: 0 },
    pop: { processed: 0, skipped: 0 },
    mixed: { processed: 0, skipped: 0 },
    tomato: { processed: 0, skipped: 0 }
  };

  for (const sheetName of workbook.SheetNames) {
    const category = detectCategory(sheetName);

    if (!category) {
      console.log(`\nSkipping sheet "${sheetName}" - cannot determine category`);
      continue;
    }

    const sheet = workbook.Sheets[sheetName];
    const result = await parseAndMigrateSheet(sheet, sheetName, category, dryRun);

    results[category].processed += result.processed;
    results[category].skipped += result.skipped;
  }

  // Summary
  console.log('\n=================================');
  console.log('Migration Summary');
  console.log('=================================');
  let totalProcessed = 0;
  for (const [category, data] of Object.entries(results)) {
    if (data.processed > 0 || data.skipped > 0) {
      console.log(`${category}: ${data.processed} imported, ${data.skipped} skipped`);
      totalProcessed += data.processed;
    }
  }
  console.log(`Total: ${totalProcessed} imported`);

  if (!dryRun) {
    await db.pool.end();
  }

  console.log('\nDone!');
}

main().catch(error => {
  console.error('Fatal error:', error);
  process.exit(1);
});
