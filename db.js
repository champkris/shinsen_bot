require('dotenv').config();
const mysql = require('mysql2/promise');

// Create connection pool
const pool = mysql.createPool({
  host: process.env.MYSQL_HOST || 'localhost',
  user: process.env.MYSQL_USER,
  password: process.env.MYSQL_PASSWORD,
  database: process.env.MYSQL_DATABASE || 'shinsen_bot',
  waitForConnections: true,
  connectionLimit: 10,
  queueLimit: 0
});

// CDC name mapping from Thai to database column names
const cdcColumnMapping = {
  'คลังบางบัวทอง': 'cdc_bangbuathong',
  'นครราชสีมา': 'cdc_nakhonratchasima',
  'นครสวรรค์': 'cdc_nakhonsawan',
  'ชลบุรี': 'cdc_chonburi',
  'คลังมหาชัย': 'cdc_mahachai',
  'คลังสุวรรณภูมิ': 'cdc_suvarnabhumi',
  'หาดใหญ่': 'cdc_hadyai',
  'ภูเก็ต': 'cdc_phuket',
  'เชียงใหม่': 'cdc_chiangmai',
  'สุราษฎร์': 'cdc_surat',
  'ขอนแก่น': 'cdc_khonkaen'
};

// Reverse mapping: database column to Thai name
const columnToCdcMapping = Object.fromEntries(
  Object.entries(cdcColumnMapping).map(([k, v]) => [v, k])
);

// Convert DD/MM/YYYY to MySQL DATE format (YYYY-MM-DD)
function toMySQLDate(dateStr) {
  if (!dateStr) return null;
  // If already in YYYY-MM-DD format
  if (dateStr.match(/^\d{4}-\d{2}-\d{2}$/)) {
    return dateStr;
  }
  // Convert from DD/MM/YYYY
  const parts = dateStr.split('/');
  if (parts.length !== 3) return null;
  return `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
}

// Convert MySQL DATE to DD/MM/YYYY format
function toDisplayDate(mysqlDate) {
  if (!mysqlDate) return null;
  const date = new Date(mysqlDate);
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
}

// All supported product categories
const CATEGORIES = ['orange', 'yuzu', 'pop', 'mixed', 'tomato'];

// Category display names (Thai)
const CATEGORY_NAMES = {
  orange: 'น้ำส้ม',
  yuzu: 'ยูซุ',
  pop: 'Shinsen Pop',
  mixed: 'น้ำผลไม้รวม',
  tomato: 'Tomato Yuzu'
};

// Convert record from database format to application format
function dbToAppRecord(row) {
  const cdcTotals = {};
  for (const [thaiName, columnName] of Object.entries(cdcColumnMapping)) {
    cdcTotals[thaiName] = row[columnName] || 0;
  }

  const record = {
    date: toDisplayDate(row.date),
    timestamp: row.timestamp ? row.timestamp.toISOString() : null,
    fc33HadyaiSum: row.fc33_hadyai_sum || 0,
    totalSum: row.total_sum || 0,
    cdcTotals: cdcTotals
  };

  // Add khonKaenLaos and khonKaenCambodia for orange category
  if (row.category === 'orange') {
    record.khonKaenLaos = row.khon_kaen_laos || 0;
    record.khonKaenCambodia = row.khon_kaen_cambodia || 0;
  }

  return record;
}

// Convert record from application format to database format
function appToDbRecord(record, category) {
  const dbRecord = {
    date: toMySQLDate(record.date),
    category: category,
    timestamp: record.timestamp ? new Date(record.timestamp) : new Date(),
    fc33_hadyai_sum: record.fc33HadyaiSum || 0,
    total_sum: record.totalSum || 0
  };

  // Map CDC totals
  if (record.cdcTotals) {
    for (const [thaiName, columnName] of Object.entries(cdcColumnMapping)) {
      dbRecord[columnName] = record.cdcTotals[thaiName] || 0;
    }
  }

  // Add khonKaenLaos and khonKaenCambodia for orange category
  if (category === 'orange') {
    dbRecord.khon_kaen_laos = record.khonKaenLaos || null;
    dbRecord.khon_kaen_cambodia = record.khonKaenCambodia || null;
  }

  return dbRecord;
}

// Get a single daily record by date and category
async function getDailyRecord(date, category) {
  const mysqlDate = toMySQLDate(date);
  const [rows] = await pool.execute(
    'SELECT * FROM daily_records WHERE date = ? AND category = ?',
    [mysqlDate, category]
  );
  return rows.length > 0 ? dbToAppRecord(rows[0]) : null;
}

// Check if a date is already recorded for a category
async function isDateRecorded(date, category) {
  const mysqlDate = toMySQLDate(date);
  const [rows] = await pool.execute(
    'SELECT 1 FROM daily_records WHERE date = ? AND category = ?',
    [mysqlDate, category]
  );
  return rows.length > 0;
}

// Save a daily record (insert or update)
async function saveDailyRecord(record, category) {
  const dbRecord = appToDbRecord(record, category);

  const columns = Object.keys(dbRecord);
  const values = Object.values(dbRecord);
  const placeholders = columns.map(() => '?').join(', ');
  const updateClause = columns.map(col => `${col} = VALUES(${col})`).join(', ');

  const sql = `
    INSERT INTO daily_records (${columns.join(', ')})
    VALUES (${placeholders})
    ON DUPLICATE KEY UPDATE ${updateClause}
  `;

  const [result] = await pool.execute(sql, values);
  return result;
}

// Get daily records by month and year
async function getDailyRecordsByMonth(month, year) {
  const startDate = `${year}-${String(month).padStart(2, '0')}-01`;
  const endDate = `${year}-${String(month).padStart(2, '0')}-31`;

  const [rows] = await pool.execute(
    `SELECT * FROM daily_records
     WHERE date >= ? AND date <= ?
     ORDER BY date ASC, category ASC`,
    [startDate, endDate]
  );

  // Initialize all categories
  const records = {};
  for (const cat of CATEGORIES) {
    records[cat] = [];
  }

  for (const row of rows) {
    const record = dbToAppRecord(row);
    if (records[row.category]) {
      records[row.category].push(record);
    }
  }

  return records;
}

// Load all daily records (for backwards compatibility)
async function loadDailyRecords() {
  const [rows] = await pool.execute(
    'SELECT * FROM daily_records ORDER BY date ASC'
  );

  // Initialize all categories
  const records = {};
  for (const cat of CATEGORIES) {
    records[cat] = [];
  }

  for (const row of rows) {
    const record = dbToAppRecord(row);
    if (records[row.category]) {
      records[row.category].push(record);
    }
  }

  return records;
}

// Save detection log
async function saveDetectionLog(logEntry) {
  const sql = `
    INSERT INTO detection_logs
    (timestamp, message_id, group_id, user_id, status, date, categories, records_created, reason)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
  `;

  const values = [
    logEntry.timestamp ? new Date(logEntry.timestamp) : new Date(),
    logEntry.messageId || null,
    logEntry.groupId || null,
    logEntry.userId || null,
    logEntry.status,
    logEntry.date ? toMySQLDate(logEntry.date) : null,
    logEntry.categories ? JSON.stringify(logEntry.categories) : null,
    logEntry.recordsCreated || 0,
    logEntry.reason || null
  ];

  const [result] = await pool.execute(sql, values);
  return result;
}

// Get detection logs (most recent first)
async function getDetectionLogs(limit = 100) {
  const [rows] = await pool.execute(
    'SELECT * FROM detection_logs ORDER BY timestamp DESC LIMIT ?',
    [limit]
  );

  return rows.map(row => ({
    timestamp: row.timestamp ? row.timestamp.toISOString() : null,
    messageId: row.message_id,
    groupId: row.group_id,
    userId: row.user_id,
    status: row.status,
    date: row.date ? toDisplayDate(row.date) : null,
    categories: row.categories ? JSON.parse(row.categories) : null,
    recordsCreated: row.records_created,
    reason: row.reason
  }));
}

// Test database connection
async function testConnection() {
  try {
    const connection = await pool.getConnection();
    console.log('[DB] MySQL connection successful');
    connection.release();
    return true;
  } catch (error) {
    console.error('[DB] MySQL connection failed:', error.message);
    return false;
  }
}

module.exports = {
  pool,
  getDailyRecord,
  isDateRecorded,
  saveDailyRecord,
  getDailyRecordsByMonth,
  loadDailyRecords,
  saveDetectionLog,
  getDetectionLogs,
  testConnection,
  toMySQLDate,
  toDisplayDate,
  cdcColumnMapping,
  columnToCdcMapping,
  CATEGORIES,
  CATEGORY_NAMES
};
