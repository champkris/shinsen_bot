require('dotenv').config();
const express = require('express');
const line = require('@line/bot-sdk');
const OpenAI = require('openai');
const { DocumentAnalysisClient, AzureKeyCredential } = require('@azure/ai-form-recognizer');
const fs = require('fs').promises;
const path = require('path');

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

const DAILY_RECORDS_FILE = path.join(__dirname, 'daily_records.json');
const DETECTION_LOGS_FILE = path.join(__dirname, 'detection_logs.json');

// Notification group IDs (comma-separated in .env)
const NOTIFICATION_GROUP_IDS = process.env.NOTIFICATION_GROUP_IDS
  ? process.env.NOTIFICATION_GROUP_IDS.split(',').map(id => id.trim()).filter(id => id)
  : [];

console.log('[CONFIG] Notification groups configured:', NOTIFICATION_GROUP_IDS.length);
console.log('[CONFIG] Notification group IDs:', NOTIFICATION_GROUP_IDS);

// Load detection logs from file
async function loadDetectionLogs() {
  try {
    const data = await fs.readFile(DETECTION_LOGS_FILE, 'utf8');
    return JSON.parse(data);
  } catch (error) {
    if (error.code === 'ENOENT') {
      return [];
    }
    throw error;
  }
}

// Save detection log
async function saveDetectionLog(logEntry) {
  try {
    const logs = await loadDetectionLogs();
    logs.unshift(logEntry); // Add to beginning for chronological order

    // Keep only last 100 logs
    const trimmedLogs = logs.slice(0, 100);

    await fs.writeFile(DETECTION_LOGS_FILE, JSON.stringify(trimmedLogs, null, 2), 'utf8');
    console.log('[LOG] Detection log saved:', logEntry.status);
  } catch (error) {
    console.error('[LOG] Error saving detection log:', error);
  }
}

// Load daily records from file
async function loadDailyRecords() {
  try {
    const data = await fs.readFile(DAILY_RECORDS_FILE, 'utf8');
    return JSON.parse(data);
  } catch (error) {
    if (error.code === 'ENOENT') {
      return { orange: [], yuzu: [] };
    }
    throw error;
  }
}

// Save daily records to file
async function saveDailyRecords(records) {
  await fs.writeFile(DAILY_RECORDS_FILE, JSON.stringify(records, null, 2), 'utf8');
}

// Send notification to configured groups about data update
async function sendNotificationToGroups(date, categories) {
  console.log(`[NOTIFICATION] sendNotificationToGroups called with date: ${date}, categories:`, categories);
  console.log(`[NOTIFICATION] NOTIFICATION_GROUP_IDS:`, NOTIFICATION_GROUP_IDS);

  if (NOTIFICATION_GROUP_IDS.length === 0) {
    console.log('[NOTIFICATION] No notification groups configured');
    return;
  }

  const message = `Report for ${date} has been recorded\n\nCategories: ${categories.join(', ')}\n\nView report: https://shinsen.yushi-marketing.com/daily-report`;
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

// Check if date already exists in records
function isDateRecorded(records, date, category) {
  return records[category].some(record => record.date === date);
}

// Record daily data if R38C2 is not 0
async function recordDailyData(tableData, extractedText = '') {
  if (!tableData || tableData.length === 0) {
    console.log('No table data to record');
    return { success: false, reason: 'No table data found' };
  }

  const table = tableData[0];

  // Check if FC33 หาดใหญ่ OR FC07 ภูเก็ต C2 sum is not 0
  let hadyaiSum = 0;
  let phuketSum = 0;
  table.forEach((row, rowIndex) => {
    if (!row || row.length < 3) return;

    // Check C0 for FC33 หาดใหญ่ and FC07 ภูเก็ต
    const c0Cell = row[0] ? row[0].toString().trim() : '';

    if (c0Cell.includes('FC33') && c0Cell.includes('หาดใหญ่')) {
      const c2Value = row[2] ? row[2].toString().replace(/,/g, '') : '0';
      const value = parseFloat(c2Value) || 0;
      hadyaiSum += value;
      console.log(`[VALIDATION] Row ${rowIndex}: FC33 หาดใหญ่ found in C0, C2 value: ${value}`);
    }

    if (c0Cell.includes('FC07')) {
      const c2Value = row[2] ? row[2].toString().replace(/,/g, '') : '0';
      const value = parseFloat(c2Value) || 0;
      phuketSum += value;
      console.log(`[VALIDATION] Row ${rowIndex}: FC07 found in C0, C2 value: ${value}`);
    }
  });

  console.log(`[VALIDATION] FC33 หาดใหญ่ total sum in C2: ${hadyaiSum}`);
  console.log(`[VALIDATION] FC07 ภูเก็ต total sum in C2: ${phuketSum}`);

  // Pass validation if EITHER FC33 หาดใหญ่ OR FC07 ภูเก็ต has non-zero sum
  if (hadyaiSum === 0 && phuketSum === 0) {
    console.log('[VALIDATION] Both FC33 หาดใหญ่ and FC07 ภูเก็ต C2 sums are 0, not recording');
    return { success: false, reason: 'Both FC33 หาดใหญ่ and FC07 ภูเก็ต C2 sums are 0 (validation failed)' };
  }

  console.log(`[VALIDATION] Passed - FC33 หาดใหญ่: ${hadyaiSum}, FC07 ภูเก็ต: ${phuketSum}`);

  // Extract date from table or extracted text
  let dateStr = null;

  // First, try to find date in the extracted text
  if (extractedText) {
    console.log('[DATE] Searching for date in extracted text...');
    const textDateMatch = extractedText.match(/(?:วันที่\s*)?(\d{1,2}\/\d{1,2}\/\d{4})/);
    if (textDateMatch) {
      dateStr = textDateMatch[1];
      console.log(`[DATE] Found date in extracted text: ${dateStr}`);
    }
  }

  // If not found in text, search in table
  if (!dateStr) {
    console.log('[DATE] Date not in extracted text, searching table...');
    for (let i = 0; i < Math.min(5, table.length); i++) {
      const row = table[i];
      console.log(`[DATE] Row ${i}:`, row ? row.slice(0, 5) : 'undefined');
    }

    // Search entire table for date pattern (วันที่ DD/MM/YYYY)
    for (let i = 0; i < table.length; i++) {
      const row = table[i];
      if (!row) continue;

      // Check all columns in this row
      for (let col = 0; col < row.length; col++) {
        const cellText = row[col] ? row[col].toString() : '';

        // Look for date patterns like "วันที่ 18/10/2025" or "18/10/2025"
        const dateMatch = cellText.match(/(?:วันที่\s*)?(\d{1,2}\/\d{1,2}\/\d{4})/);
        if (dateMatch) {
          dateStr = dateMatch[1]; // Extract just the date part (DD/MM/YYYY)
          console.log(`[DATE] Found date: ${dateStr} in row ${i}, column ${col}, cell: "${cellText}"`);
          break;
        }
      }

      if (dateStr) break;
    }
  }

  if (!dateStr) {
    console.log('[DATE] No date found in extracted text or table');
    return { success: false, reason: 'No date found in table or extracted text' };
  }

  console.log(`[DATE] Using date: ${dateStr}`);

  const records = await loadDailyRecords();

  // Calculate CDC totals for both orange and yuzu
  const cdcConfig = {
    orange: {
      column: 2, // Column index for orange data (C2)
    },
    yuzu: {
      column: 3, // Column index for yuzu data (C3)
    }
  };

  // CDC names to track
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

  // Map short names to full names for consistency
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

  const results = [];

  // Loop through both categories (orange and yuzu)
  for (const category of ['orange', 'yuzu']) {
    // Check if this date is already recorded for this category
    if (isDateRecorded(records, dateStr, category)) {
      console.log(`Date ${dateStr} already recorded for ${category}`);
      continue;
    }

    const columnIndex = cdcConfig[category].column;
    const cdcTotals = {};
    let totalSum = 0;

    // Initialize all CDC totals to 0
    Object.values(cdcNameMapping).forEach(cdcFullName => {
      cdcTotals[cdcFullName] = 0;
    });

    console.log(`[RECORD] Calculating CDC totals for ${category} using column ${columnIndex}`);

    // First, find the total sum from the table
    // Look for a row containing "รวม" or "Total" in C0 or C1
    console.log(`[TOTAL] Searching for total sum in ${table.length} rows for ${category}...`);

    for (let i = 0; i < table.length; i++) {
      const row = table[i];
      if (!row || row.length < 2) continue;

      const c0 = row[0] ? row[0].toString().trim() : '';
      const c1 = row[1] ? row[1].toString().trim() : '';

      // Log last few rows to debug
      if (i >= table.length - 5) {
        console.log(`[TOTAL] Row ${i}: C0="${c0}", C1="${c1}", C2="${row[2] || ''}", C3="${row[3] || ''}"`);
      }

      // Check if this row contains total indicator
      // Check for various total indicators
      const totalIndicators = ['ยอดรวม', 'รวม', 'total', 'grand total', 'sum'];
      const foundTotal = totalIndicators.some(indicator =>
        c0.toLowerCase().includes(indicator.toLowerCase()) || c1.toLowerCase().includes(indicator.toLowerCase())
      );

      if (foundTotal) {
        // Extract the total sum from the appropriate column
        if (row[columnIndex]) {
          const cellValue = row[columnIndex].toString().replace(/,/g, '');
          totalSum = parseFloat(cellValue) || 0;
          console.log(`[TOTAL] Found total sum for ${category}: ${totalSum} in row ${i} (C0="${c0}", C1="${c1}")`);
          break;
        }
      }
    }

    // If no total row found, calculate total by summing all values in the column
    if (totalSum === 0) {
      console.log(`[TOTAL] No ยอดรวม row found, calculating total by summing all values in column ${columnIndex}`);
      for (let i = 0; i < table.length; i++) {
        const row = table[i];
        if (!row || !row[columnIndex]) continue;

        const cellValue = row[columnIndex].toString().replace(/,/g, '').trim();
        const value = parseFloat(cellValue) || 0;

        if (value > 0) {
          totalSum += value;
        }
      }
      console.log(`[TOTAL] Calculated total sum for ${category}: ${totalSum}`);
    }

    // Iterate through table rows and sum values based on CDC name
    table.forEach((row, rowIndex) => {
      if (!row || row.length < 2) return;

      // Find matching CDC name
      cdcNames.forEach(cdcName => {
        const fullCdcName = cdcNameMapping[cdcName];

        // Determine which column to search based on CDC name
        // Exception: "คลังบางบัวทอง" searches in C0 (looks for "FC01 บางบัวทอง" or similar)
        // Other "คลัง" prefixed names search in C1
        // All other locations search in C0
        let searchColumnIndex;
        let searchCell;

        if (fullCdcName === 'คลังบางบัวทอง') {
          // Exception: บางบัวทอง searches in C0 for "FC01 บางบัวทอง" or similar
          searchColumnIndex = 0;
          searchCell = row[0] ? row[0].toString().trim() : '';
        } else if (fullCdcName.startsWith('คลัง')) {
          // Search in C1 for other locations starting with "คลัง" (มหาชัย, สุวรรณภูมิ)
          searchColumnIndex = 1;
          searchCell = row[1] ? row[1].toString().trim() : '';
        } else {
          // Search in C0 for other locations
          searchColumnIndex = 0;
          searchCell = row[0] ? row[0].toString().trim() : '';
        }

        // Check if this row matches the CDC name
        if (searchCell.includes(cdcName)) {
          // Get value from the appropriate column (orange=C2, yuzu=C3)
          if (row[columnIndex]) {
            const cellValue = row[columnIndex].toString().replace(/,/g, '');
            const value = parseFloat(cellValue) || 0;

            if (value > 0) {
              cdcTotals[fullCdcName] += value;
              console.log(`[RECORD] Row ${rowIndex}: C${searchColumnIndex}="${searchCell}" -> ${fullCdcName} += ${value} (total: ${cdcTotals[fullCdcName]})`);
            }
          }
        }
      });
    });

    // For Orange category only: Extract ขอนแก่น Laos value from the last ขอนแก่น row
    let khonKaenLaosValue = 0;
    if (category === 'orange') {
      // Find the last row containing ขอนแก่น in C0
      for (let i = table.length - 1; i >= 0; i--) {
        const row = table[i];
        if (!row || !row[0]) continue;

        const c0 = row[0].toString().trim();
        if (c0.includes('ขอนแก่น')) {
          // Extract value from C2 (column index 2)
          if (row[2]) {
            const cellValue = row[2].toString().replace(/,/g, '');
            khonKaenLaosValue = parseFloat(cellValue) || 0;
            console.log(`[LAOS] Found last ขอนแก่น row at index ${i}, C2 value: ${khonKaenLaosValue}`);
            break;
          }
        }
      }
    }

    // Extract all relevant data
    const dailyRecord = {
      date: dateStr,
      timestamp: new Date().toISOString(),
      fc33HadyaiSum: hadyaiSum, // Store FC33 หาดใหญ่ validation value
      totalSum: totalSum, // Store total sum from source table
      cdcTotals: cdcTotals, // Store CDC totals
      khonKaenLaos: category === 'orange' ? khonKaenLaosValue : undefined, // ขอนแก่น Laos (Orange only)
      khonKaenCambodia: category === 'orange' ? 0 : undefined // ขอนแก่น Cambodia (Orange only, always 0 for now)
    };

    records[category].push(dailyRecord);
    console.log(`Recorded daily data for ${dateStr} in ${category} category`);
    results.push({ category, record: dailyRecord });
  }

  // Save all records at once
  await saveDailyRecords(records);

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
          const replyMessage = `Report for ${recordResult.date} has been recorded\n\nView report: https://shinsen.yushi-marketing.com/daily-report`;

          await client.replyMessage({
            replyToken: event.replyToken,
            messages: [{
              type: 'text',
              text: replyMessage,
            }],
          });

          console.log('Success message sent to user');

          // Send notifications to configured groups
          const categories = recordResult.results.map(r => r.category);
          await sendNotificationToGroups(recordResult.date, categories);
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
      recordResult = await recordDailyData(tableData, extractedText);
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
          const cellValue = row[columnIndex].toString().replace(/,/g, '');
          const value = parseFloat(cellValue) || 0;

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
    const records = await loadDailyRecords();

    // Get current date for default filter
    const now = new Date();
    const currentMonth = now.getMonth() + 1; // 1-12
    const currentYear = now.getFullYear();

    // Get filter parameters from query string (default to current month/year)
    const selectedMonth = parseInt(req.query.month) || currentMonth;
    const selectedYear = parseInt(req.query.year) || currentYear;

    // Filter records by selected month and year
    const filterByMonthYear = (recordsList) => {
      if (!recordsList || recordsList.length === 0) return [];

      return recordsList.filter(record => {
        // Parse date in DD/MM/YYYY format
        const dateParts = record.date.split('/');
        if (dateParts.length !== 3) return false;

        const recordMonth = parseInt(dateParts[1]);
        const recordYear = parseInt(dateParts[2]);

        return recordMonth === selectedMonth && recordYear === selectedYear;
      });
    };

    const filteredOrange = filterByMonthYear(records.orange);
    const filteredYuzu = filterByMonthYear(records.yuzu);

    const orangeTableHTML = generateDailyRecordsTable(filteredOrange, 'Orange');
    const yuzuTableHTML = generateDailyRecordsTable(filteredYuzu, 'Yuzu');

    res.send(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>Daily Records Report</title>
        <meta charset="utf-8">
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
            <div class="summary-card">
              <div class="summary-number">${filteredOrange.length}</div>
              <div class="summary-label">Orange Records</div>
            </div>
            <div class="summary-card">
              <div class="summary-number">${filteredYuzu.length}</div>
              <div class="summary-label">Yuzu Records</div>
            </div>
            <div class="summary-card">
              <div class="summary-number">${filteredOrange.length + filteredYuzu.length}</div>
              <div class="summary-label">Total Records</div>
            </div>
          </div>

          <div class="category-section orange-section">
            <div class="category-title orange-title">
              🍊 Orange Records
            </div>
            ${orangeTableHTML}
          </div>

          <div class="category-section yuzu-section">
            <div class="category-title yuzu-title">
              🍋 Yuzu Records
            </div>
            ${yuzuTableHTML}
          </div>
        </div>

        <script>
          function applyFilter() {
            const month = document.getElementById('month-filter').value;
            const year = document.getElementById('year-filter').value;
            window.location.href = '/daily-report?month=' + month + '&year=' + year;
          }
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
      const formattedLaos = laosValue.toLocaleString('en-US');
      html += `<td style="background-color: #FFE0B2; font-weight: bold;">${formattedLaos}</td>`;

      // ขอนแก่น Cambodia column (always blank/0)
      const cambodiaValue = record.khonKaenCambodia || 0;
      const formattedCambodia = cambodiaValue.toLocaleString('en-US');
      html += `<td style="background-color: #FFE0B2;">${formattedCambodia}</td>`;
    }

    // Total sum column
    const totalValue = record.totalSum || 0;
    const formattedTotal = totalValue.toLocaleString('en-US');
    html += `<td style="background-color: #c8e6c9; font-weight: bold;">${formattedTotal}</td>`;

    // Recorded timestamp
    html += `<td style="background-color: #f3e5f5; font-size: 0.9em;">${recordedDate}</td>`;
    html += '</tr>';
  });

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

app.listen(PORT, () => {
  console.log(`LINE Bot server is running on port ${PORT}`);
  console.log(`Webhook URL: http://localhost:${PORT}/webhook`);
});
