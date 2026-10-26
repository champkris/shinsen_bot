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

// The queue table is created on boot rather than by a deploy step. Deploys here are
// "upload files, press Restart App" in the Plesk panel with no shell in the loop, so a
// migration that has to be remembered is a migration that gets skipped — and skipping
// this one silently turns every incoming image back into a lost screenshot.
// CREATE TABLE IF NOT EXISTS is a no-op once the table exists, which is the normal
// case for all 35-50 of this process's daily boots.
const PENDING_EVENTS_DDL = `
  CREATE TABLE IF NOT EXISTS pending_events (
    message_id VARCHAR(255) NOT NULL PRIMARY KEY,
    event_json MEDIUMTEXT NOT NULL,
    status ENUM('pending', 'processing', 'done', 'failed') NOT NULL DEFAULT 'pending',
    attempts INT NOT NULL DEFAULT 0,
    claimed_by VARCHAR(64) DEFAULT NULL,
    claimed_at DATETIME DEFAULT NULL,
    last_error TEXT DEFAULT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX idx_status_created (status, created_at)
  )
`;

// Resolves either way: a DB user without CREATE rights must still get a working app
// (the webhook falls back to inline processing), not a boot that hangs on schema setup.
const schemaReady = pool.query(PENDING_EVENTS_DDL)
  .then(() => { console.log('[DB] pending_events ready'); return true; })
  .catch(error => {
    console.error('[DB] Could not ensure pending_events table:', error.message);
    return false;
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
const CATEGORIES = ['orange', 'yuzu', 'pop', 'mixed', 'tomato', 'pineapple', 'watermelon'];

// Category display names (Thai)
const CATEGORY_NAMES = {
  orange: 'น้ำส้ม',
  yuzu: 'ยูซุ',
  pop: 'Shinsen Pop',
  mixed: 'น้ำผลไม้รวม',
  tomato: 'Tomato Yuzu',
  pineapple: 'น้ำสับปะรด',
  watermelon: 'น้ำแตงโม'
};

// daily_records.category is an ENUM, so a category added to CATEGORIES is rejected by
// MySQL until the column learns it. Widened on boot for the same reason pending_events is
// created on boot: a deploy has no shell step, so a hand-run ALTER is one that gets skipped.
// Checked against information_schema first so the ALTER only runs once, not on every boot.
const categoryEnumReady = pool.query(
  `SELECT COLUMN_TYPE AS type FROM information_schema.COLUMNS
   WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'daily_records' AND COLUMN_NAME = 'category'`
)
  .then(async ([rows]) => {
    if (!rows.length) return false;
    const type = String(rows[0].type);
    const missing = CATEGORIES.filter(c => !type.includes(`'${c}'`));
    if (!missing.length) return true;
    const enumList = CATEGORIES.map(c => `'${c}'`).join(', ');
    await pool.query(`ALTER TABLE daily_records MODIFY category ENUM(${enumList}) NOT NULL`);
    console.log(`[DB] daily_records.category ENUM extended with: ${missing.join(', ')}`);
    return true;
  })
  .catch(error => {
    console.error('[DB] Could not ensure daily_records.category ENUM:', error.message);
    return false;
  });

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
  await categoryEnumReady;
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
    categories: row.categories ? (typeof row.categories === 'string' ? JSON.parse(row.categories) : row.categories) : null,
    recordsCreated: row.records_created,
    reason: row.reason
  }));
}

// Sum total_sum by category within an inclusive date range.
// Returns an object like { orange: 12345, yuzu: 6789 } — categories with zero are omitted.
async function getAggregateByCategory(startDate, endDate) {
  const [rows] = await pool.execute(
    `SELECT category, SUM(total_sum) AS volume
     FROM daily_records
     WHERE date BETWEEN ? AND ?
     GROUP BY category`,
    [startDate, endDate]
  );
  const result = {};
  for (const row of rows) {
    result[row.category] = Number(row.volume) || 0;
  }
  return result;
}

// Latest day-of-month with any data in the given month, or null if none.
async function getLatestDayWithData(year, month) {
  const [rows] = await pool.execute(
    `SELECT MAX(DAY(date)) AS day FROM daily_records
     WHERE YEAR(date) = ? AND MONTH(date) = ?`,
    [year, month]
  );
  return rows[0]?.day ? Number(rows[0].day) : null;
}

// Distinct categories ever recorded or ever given a target.
async function getKnownCategories() {
  const [rows] = await pool.execute(
    `SELECT category FROM daily_records
     UNION
     SELECT category FROM monthly_targets`
  );
  return rows.map(r => r.category);
}

// Targets for a single month as { category: target_value }.
async function getTargets(year, month) {
  const [rows] = await pool.execute(
    `SELECT category, target_value FROM monthly_targets
     WHERE year = ? AND month = ?`,
    [year, month]
  );
  const result = {};
  for (const row of rows) {
    result[row.category] = Number(row.target_value) || 0;
  }
  return result;
}

// Sum of monthly targets from January through the given month (inclusive), per category.
async function getTargetsYTD(year, throughMonth) {
  const [rows] = await pool.execute(
    `SELECT category, SUM(target_value) AS total
     FROM monthly_targets
     WHERE year = ? AND month BETWEEN 1 AND ?
     GROUP BY category`,
    [year, throughMonth]
  );
  const result = {};
  for (const row of rows) {
    result[row.category] = Number(row.total) || 0;
  }
  return result;
}

// Insert or update a single monthly target.
async function upsertTarget(year, month, category, value) {
  await pool.execute(
    `INSERT INTO monthly_targets (year, month, category, target_value)
     VALUES (?, ?, ?, ?)
     ON DUPLICATE KEY UPDATE target_value = VALUES(target_value)`,
    [year, month, category, value]
  );
}

// Has this LINE message already been through the pipeline? Used to make webhook
// handling idempotent so LINE's webhook redelivery cannot double-process an image.
// 'error' rows are deliberately excluded: they mean the run blew up before reaching
// a verdict, so the image still needs processing. Counting them would let one
// transient LINE/Azure blip permanently mark a screenshot as handled — the retries
// would all short-circuit here and the data would never arrive.
async function hasProcessedMessage(messageId) {
  if (!messageId) return false;
  const [rows] = await pool.execute(
    "SELECT 1 FROM detection_logs WHERE message_id = ? AND status IN ('success', 'failed') LIMIT 1",
    [messageId]
  );
  return rows.length > 0;
}

// Dates (YYYY-MM-DD strings) in an inclusive range that have at least one recorded
// category. DATE_FORMAT keeps this a string end-to-end — a JS Date round-trip would
// shift the day across timezones.
async function getRecordedDates(startDate, endDate) {
  const [rows] = await pool.execute(
    `SELECT DISTINCT DATE_FORMAT(date, '%Y-%m-%d') AS d
     FROM daily_records
     WHERE date BETWEEN ? AND ?
     ORDER BY d ASC`,
    [startDate, endDate]
  );
  return rows.map(row => row.d);
}

// Watchdog alert bookkeeping — one row per (date, kind) so a missing day is
// reported once, not on every watchdog run.
async function hasWatchdogAlert(date, kind = 'missing_data') {
  const [rows] = await pool.execute(
    'SELECT 1 FROM watchdog_alerts WHERE alert_date = ? AND kind = ? LIMIT 1',
    [toMySQLDate(date), kind]
  );
  return rows.length > 0;
}

async function recordWatchdogAlert(date, kind = 'missing_data') {
  const [result] = await pool.execute(
    `INSERT INTO watchdog_alerts (alert_date, kind, sent_at)
     VALUES (?, ?, NOW())
     ON DUPLICATE KEY UPDATE sent_at = VALUES(sent_at)`,
    [toMySQLDate(date), kind]
  );
  return result;
}

// ---------------------------------------------------------------------------
// Durable event queue
//
// Passenger reaps the app process as soon as no HTTP request is in flight, so
// anything still running after the webhook responds can be killed mid-run — that
// is how a screenshot that GPT-4 Vision had already accepted disappeared without
// even a detection_logs row. Every image event is written here inside the webhook
// request (a ~10ms INSERT that always beats the teardown) and processed later,
// so a killed run leaves a claimable row behind instead of nothing.
// ---------------------------------------------------------------------------

// Queue an event, leaving any existing row untouched: a LINE redelivery must not
// resurrect a row already marked done or reset its attempt count. INSERT IGNORE
// rather than ON DUPLICATE KEY UPDATE because MariaDB reports affectedRows = 1 for
// the no-op update branch, which makes "was this new?" unanswerable there.
// Returns true only when the row was actually inserted.
async function enqueueEvent(messageId, event) {
  if (!messageId) return false;
  // The very first webhook after a deploy can land before the DDL finishes.
  await schemaReady;
  const [result] = await pool.execute(
    'INSERT IGNORE INTO pending_events (message_id, event_json) VALUES (?, ?)',
    [messageId, JSON.stringify(event)]
  );
  return result.affectedRows === 1;
}

// Take ownership of up to `limit` events. Rows stuck in 'processing' past
// staleSeconds are reclaimed — that is the crash path: the process died holding
// the claim, so nobody will ever finish or release it.
async function claimPendingEvents(claimToken, { limit = 3, staleSeconds = 180, maxAttempts = 5 } = {}) {
  await schemaReady;
  await pool.execute(
    `UPDATE pending_events
        SET status = 'processing', claimed_by = ?, claimed_at = NOW(), attempts = attempts + 1
      WHERE (status = 'pending'
             OR (status = 'processing' AND claimed_at < NOW() - INTERVAL ? SECOND))
        AND attempts < ?
      ORDER BY created_at ASC
      LIMIT ?`,
    [claimToken, staleSeconds, maxAttempts, limit]
  );

  const [rows] = await pool.execute(
    `SELECT message_id, event_json, attempts
       FROM pending_events
      WHERE claimed_by = ? AND status = 'processing'
      ORDER BY created_at ASC`,
    [claimToken]
  );

  return rows.map(row => ({
    messageId: row.message_id,
    event: JSON.parse(row.event_json),
    attempts: row.attempts
  }));
}

async function completeEvent(messageId) {
  await pool.execute(
    `UPDATE pending_events
        SET status = 'done', claimed_by = NULL, claimed_at = NULL, last_error = NULL
      WHERE message_id = ?`,
    [messageId]
  );
}

// Hand the row back for a later attempt, or bury it once it has burned through
// maxAttempts so a permanently broken event cannot loop forever.
async function releaseEvent(messageId, errorMessage, maxAttempts = 5) {
  const [result] = await pool.execute(
    `UPDATE pending_events
        SET status = IF(attempts >= ?, 'failed', 'pending'),
            claimed_by = NULL, claimed_at = NULL, last_error = ?
      WHERE message_id = ?`,
    [maxAttempts, errorMessage ? String(errorMessage).slice(0, 1000) : null, messageId]
  );
  if (result.affectedRows === 0) return null;
  const [rows] = await pool.execute(
    'SELECT status, attempts FROM pending_events WHERE message_id = ?',
    [messageId]
  );
  return rows[0] || null;
}

// Rows still owing work, oldest first — what /health and the queue page report.
async function getQueueStats() {
  const stats = { pending: 0, processing: 0, done: 0, failed: 0, oldestUnfinished: null };
  // Reported on /health, so a missing table must degrade to a note on the queue —
  // not a 503 that says the whole database is down.
  if (!(await schemaReady)) return { ...stats, error: 'pending_events table unavailable' };

  const [rows] = await pool.execute(
    `SELECT status, COUNT(*) AS n, MIN(created_at) AS oldest
       FROM pending_events
      GROUP BY status`
  );
  for (const row of rows) {
    stats[row.status] = Number(row.n) || 0;
    if ((row.status === 'pending' || row.status === 'processing') && row.oldest) {
      const oldest = new Date(row.oldest).toISOString();
      if (!stats.oldestUnfinished || oldest < stats.oldestUnfinished) stats.oldestUnfinished = oldest;
    }
  }
  return stats;
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
  getAggregateByCategory,
  getLatestDayWithData,
  getKnownCategories,
  getTargets,
  getTargetsYTD,
  upsertTarget,
  hasProcessedMessage,
  getRecordedDates,
  hasWatchdogAlert,
  recordWatchdogAlert,
  enqueueEvent,
  claimPendingEvents,
  completeEvent,
  releaseEvent,
  getQueueStats,
  testConnection,
  toMySQLDate,
  toDisplayDate,
  cdcColumnMapping,
  columnToCdcMapping,
  CATEGORIES,
  CATEGORY_NAMES
};
