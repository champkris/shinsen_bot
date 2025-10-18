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

// Check if date already exists in records
function isDateRecorded(records, date, category) {
  return records[category].some(record => record.date === date);
}

// Record daily data if R38C2 is not 0
async function recordDailyData(tableData) {
  if (!tableData || tableData.length === 0) {
    console.log('No table data to record');
    return null;
  }

  const table = tableData[0];

  // Check if R38C2 exists and is not 0
  if (!table[38] || !table[38][2]) {
    console.log('R38C2 does not exist');
    return null;
  }

  const r38c2Value = parseFloat(table[38][2].toString().replace(/,/g, '')) || 0;

  if (r38c2Value === 0) {
    console.log('R38C2 is 0, not recording');
    return null;
  }

  // Extract date from top of table (usually in first few rows)
  let dateStr = null;

  // Search first 5 rows for date pattern (วันที่ DD/MM/YYYY)
  for (let i = 0; i < Math.min(5, table.length); i++) {
    const row = table[i];
    if (!row) continue;

    // Check all columns in this row
    for (let col = 0; col < row.length; col++) {
      const cellText = row[col] ? row[col].toString() : '';

      // Look for date patterns like "วันที่ 18/10/2025" or "18/10/2025"
      const dateMatch = cellText.match(/(?:วันที่\s*)?(\d{1,2}\/\d{1,2}\/\d{4})/);
      if (dateMatch) {
        dateStr = dateMatch[1]; // Extract just the date part (DD/MM/YYYY)
        console.log(`[DATE] Found date: ${dateStr} in row ${i}, column ${col}`);
        break;
      }
    }

    if (dateStr) break;
  }

  if (!dateStr) {
    console.log('[DATE] No date found in top of table');
    return null;
  }

  console.log(`[DATE] Using date: ${dateStr}`);

  // Determine category (orange or yuzu) - you can modify this logic based on actual data
  // For now, we'll check if there's a category indicator in the table
  const category = detectCategory(table);

  const records = await loadDailyRecords();

  // Check if this date is already recorded for this category
  if (isDateRecorded(records, dateStr, category)) {
    console.log(`Date ${dateStr} already recorded for ${category}`);
    return null;
  }

  // Calculate CDC totals by searching for CDC names in the table (dynamic approach)
  const cdcConfig = {
    orange: {
      column: 2, // Column index for orange data (0-indexed)
    },
    yuzu: {
      column: 3, // Column index for yuzu data (0-indexed)
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

  const columnIndex = cdcConfig[category].column;
  const cdcTotals = {};

  // Initialize all CDC totals to 0
  Object.values(cdcNameMapping).forEach(cdcFullName => {
    cdcTotals[cdcFullName] = 0;
  });

  console.log(`[RECORD] Calculating CDC totals for ${category} using column ${columnIndex}`);

  // Iterate through table rows and sum values based on CDC name
  table.forEach((row, rowIndex) => {
    if (!row || row.length < 2) return;

    // Find matching CDC name
    cdcNames.forEach(cdcName => {
      const fullCdcName = cdcNameMapping[cdcName];

      // Determine which column to search based on CDC name
      // If CDC name starts with "คลัง", search in C1 (column 1)
      // Otherwise, search in C0 (column 0)
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

  // Extract all relevant data from row 38
  const dailyRecord = {
    date: dateStr,
    timestamp: new Date().toISOString(),
    r38c2: r38c2Value,
    cdcTotals: cdcTotals, // Store CDC totals
    rowData: table[38] // Store entire row 38 data for reference
  };

  records[category].push(dailyRecord);
  await saveDailyRecords(records);

  console.log(`Recorded daily data for ${dateStr} in ${category} category`);
  return { category, record: dailyRecord };
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
        const extractedText = await performOCR(imageBuffer);
        console.log('[OCR] Extraction completed, text length:', extractedText.length);

        const replyMessage = `This is an Excel screenshot!\n\nExtracted text:\n${extractedText.substring(0, 4000)}`;

        await client.replyMessage({
          replyToken: event.replyToken,
          messages: [{
            type: 'text',
            text: replyMessage,
          }],
        });

        console.log('OCR result sent to user');
      } else {
        await client.replyMessage({
          replyToken: event.replyToken,
          messages: [{
            type: 'text',
            text: 'This is an image, but not an Excel screenshot.',
          }],
        });
      }

      console.log('Detection result:', { isExcelScreenshot });
    } catch (error) {
      console.error('Error processing image:', error);

      await client.replyMessage({
        replyToken: event.replyToken,
        messages: [{
          type: 'text',
          text: 'Sorry, I encountered an error processing your image.',
        }],
      });
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

async function performOCR(imageBuffer) {
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
      rawResult: result
    };

    console.log('OCR completed. Text length:', extractedText.length);
    console.log('Tables found:', tableData.length);

    // Record daily data if conditions are met
    try {
      const recordResult = await recordDailyData(tableData);
      if (recordResult) {
        console.log('Daily data recorded successfully');
      }
    } catch (error) {
      console.error('Error recording daily data:', error);
    }

    return extractedText;
  } catch (error) {
    console.error('Error in OCR:', error);
    throw error;
  }
}

function transformTableData(tableData) {
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
  const columnIndex = 2; // Orange column (C2)

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
        // Get value from column C2
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

function generateTransformedTable(transformedData) {
  if (!transformedData || transformedData.length === 0) {
    return '<p>No transformed data available</p>';
  }

  let html = '<style>table { border-collapse: collapse; width: 100%; margin: 20px 0; } th, td { border: 1px solid #ddd; padding: 8px; text-align: left; } th { background-color: #2196F3; color: white; } tr:nth-child(even) { background-color: #f2f2f2; }</style>';

  html += '<table>';
  html += '<tr><th>CDC</th><th>ขวดน้ำส้ม</th></tr>';

  transformedData.forEach(item => {
    html += `<tr><td>${item.cdc}</td><td>${item.value}</td></tr>`;
  });

  html += '</table>';

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

  const transformedData = transformTableData(latestOCRResult.tableData);
  const transformedHTML = generateTransformedTable(transformedData);

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
        <strong>CDC locations:</strong> ${transformedData ? transformedData.length : 0}
      </div>

      ${transformedHTML}
    </body>
    </html>
  `);
});

app.get('/daily-report', async (req, res) => {
  try {
    const records = await loadDailyRecords();

    const orangeTableHTML = generateDailyRecordsTable(records.orange, 'Orange');
    const yuzuTableHTML = generateDailyRecordsTable(records.yuzu, 'Yuzu');

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
            max-width: 1200px;
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
          .action-buttons {
            margin-bottom: 20px;
            text-align: center;
          }
          .btn-clear {
            padding: 12px 24px;
            background-color: #f44336;
            color: white;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-size: 16px;
            transition: background-color 0.3s;
          }
          .btn-clear:hover {
            background-color: #d32f2f;
          }
          .btn-clear:disabled {
            background-color: #ccc;
            cursor: not-allowed;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="nav">
            <a href="/latest-ocr">Original Table</a>
            <a href="/transformed-data">Transformed Data</a>
            <a href="/daily-report">Daily Report</a>
          </div>

          <h1>📊 Daily Records Report</h1>

          <div class="action-buttons">
            <button class="btn-clear" onclick="clearAllRecords()">🗑️ Clear All Records</button>
          </div>

          <div class="summary">
            <div class="summary-card">
              <div class="summary-number">${records.orange.length}</div>
              <div class="summary-label">Orange Records</div>
            </div>
            <div class="summary-card">
              <div class="summary-number">${records.yuzu.length}</div>
              <div class="summary-label">Yuzu Records</div>
            </div>
            <div class="summary-card">
              <div class="summary-number">${records.orange.length + records.yuzu.length}</div>
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
          async function clearAllRecords() {
            if (!confirm('Are you sure you want to clear all records? This action cannot be undone.')) {
              return;
            }

            const button = document.querySelector('.btn-clear');
            button.disabled = true;
            button.textContent = 'Clearing...';

            try {
              const response = await fetch('/clear-records', {
                method: 'POST',
                headers: {
                  'Content-Type': 'application/json'
                }
              });

              const result = await response.json();

              if (result.success) {
                alert('All records cleared successfully!');
                window.location.reload();
              } else {
                alert('Failed to clear records: ' + result.error);
                button.disabled = false;
                button.textContent = '🗑️ Clear All Records';
              }
            } catch (error) {
              alert('Error clearing records: ' + error.message);
              button.disabled = false;
              button.textContent = '🗑️ Clear All Records';
            }
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

  // Column headers matching CDC locations (in proper order)
  const cdcColumns = [
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
  html += '<th style="background-color: #9c27b0;">Recorded At</th></tr>';

  // Sort by date (most recent first)
  const sortedRecords = [...records].sort((a, b) => {
    return new Date(b.timestamp) - new Date(a.timestamp);
  });

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

    // Recorded timestamp
    html += `<td style="background-color: #f3e5f5; font-size: 0.9em;">${recordedDate}</td>`;
    html += '</tr>';
  });

  html += '</table></div>';
  return html;
}

app.post('/clear-records', async (req, res) => {
  try {
    const emptyRecords = { orange: [], yuzu: [] };
    await saveDailyRecords(emptyRecords);
    console.log('[CLEAR] All records cleared successfully');
    res.json({ success: true, message: 'All records cleared successfully' });
  } catch (error) {
    console.error('[ERROR] Error clearing records:', error);
    res.status(500).json({ success: false, error: 'Failed to clear records' });
  }
});

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
app.listen(PORT, () => {
  console.log(`LINE Bot server is running on port ${PORT}`);
  console.log(`Webhook URL: http://localhost:${PORT}/webhook`);
});
