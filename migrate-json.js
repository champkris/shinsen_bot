#!/usr/bin/env node
/**
 * Migration script: JSON to MySQL
 * Migrates existing daily_records.json and detection_logs.json to MySQL database
 */

require('dotenv').config();
const fs = require('fs').promises;
const path = require('path');
const db = require('./db');

const DAILY_RECORDS_FILE = path.join(__dirname, 'daily_records.json');
const DETECTION_LOGS_FILE = path.join(__dirname, 'detection_logs.json');

async function migrateDailyRecords() {
  console.log('\n=== Migrating Daily Records ===');

  try {
    // Check if file exists
    const data = await fs.readFile(DAILY_RECORDS_FILE, 'utf8');
    const records = JSON.parse(data);

    let orangeCount = 0;
    let yuzuCount = 0;

    // Migrate orange records
    if (records.orange && records.orange.length > 0) {
      console.log(`Found ${records.orange.length} orange records`);
      for (const record of records.orange) {
        try {
          await db.saveDailyRecord(record, 'orange');
          orangeCount++;
          process.stdout.write(`\rOrange: ${orangeCount}/${records.orange.length}`);
        } catch (error) {
          console.error(`\nError migrating orange record ${record.date}:`, error.message);
        }
      }
      console.log(`\nOrange migration complete: ${orangeCount} records`);
    }

    // Migrate yuzu records (check both 'yuzu' and 'pop' keys for backwards compatibility)
    const yuzuRecords = records.yuzu || records.pop || [];
    if (yuzuRecords.length > 0) {
      console.log(`Found ${yuzuRecords.length} yuzu records`);
      for (const record of yuzuRecords) {
        try {
          await db.saveDailyRecord(record, 'yuzu');
          yuzuCount++;
          process.stdout.write(`\rYuzu: ${yuzuCount}/${yuzuRecords.length}`);
        } catch (error) {
          console.error(`\nError migrating yuzu record ${record.date}:`, error.message);
        }
      }
      console.log(`\nYuzu migration complete: ${yuzuCount} records`);
    }

    console.log(`\nTotal daily records migrated: ${orangeCount + yuzuCount}`);
    return { orange: orangeCount, yuzu: yuzuCount };

  } catch (error) {
    if (error.code === 'ENOENT') {
      console.log('No daily_records.json file found, skipping...');
      return { orange: 0, yuzu: 0 };
    }
    throw error;
  }
}

async function migrateDetectionLogs() {
  console.log('\n=== Migrating Detection Logs ===');

  try {
    const data = await fs.readFile(DETECTION_LOGS_FILE, 'utf8');
    const logs = JSON.parse(data);

    if (logs.length === 0) {
      console.log('No detection logs found');
      return 0;
    }

    console.log(`Found ${logs.length} detection logs`);
    let count = 0;

    for (const log of logs) {
      try {
        await db.saveDetectionLog(log);
        count++;
        process.stdout.write(`\rLogs: ${count}/${logs.length}`);
      } catch (error) {
        console.error(`\nError migrating log:`, error.message);
      }
    }

    console.log(`\nDetection logs migration complete: ${count} records`);
    return count;

  } catch (error) {
    if (error.code === 'ENOENT') {
      console.log('No detection_logs.json file found, skipping...');
      return 0;
    }
    throw error;
  }
}

async function verifyMigration() {
  console.log('\n=== Verifying Migration ===');

  // Load from MySQL
  const records = await db.loadDailyRecords();
  const logs = await db.getDetectionLogs(1000);

  console.log(`MySQL daily_records: Orange=${records.orange.length}, Yuzu=${records.yuzu.length}`);
  console.log(`MySQL detection_logs: ${logs.length}`);

  // Load from JSON for comparison
  try {
    const jsonData = await fs.readFile(DAILY_RECORDS_FILE, 'utf8');
    const jsonRecords = JSON.parse(jsonData);
    const jsonYuzu = jsonRecords.yuzu || jsonRecords.pop || [];
    console.log(`JSON daily_records: Orange=${jsonRecords.orange?.length || 0}, Yuzu=${jsonYuzu.length}`);
  } catch (e) {
    // File might not exist
  }

  try {
    const jsonLogs = JSON.parse(await fs.readFile(DETECTION_LOGS_FILE, 'utf8'));
    console.log(`JSON detection_logs: ${jsonLogs.length}`);
  } catch (e) {
    // File might not exist
  }
}

async function main() {
  console.log('=================================');
  console.log('JSON to MySQL Migration Script');
  console.log('=================================');

  // Test database connection
  const connected = await db.testConnection();
  if (!connected) {
    console.error('\nFailed to connect to MySQL database.');
    console.error('Please ensure:');
    console.error('1. MySQL server is running');
    console.error('2. Database "shinsen_bot" exists (run schema.sql first)');
    console.error('3. .env file has correct MYSQL_* credentials');
    process.exit(1);
  }

  try {
    // Run migrations
    await migrateDailyRecords();
    await migrateDetectionLogs();

    // Verify
    await verifyMigration();

    console.log('\n=================================');
    console.log('Migration completed successfully!');
    console.log('=================================');

  } catch (error) {
    console.error('\nMigration failed:', error);
    process.exit(1);
  } finally {
    await db.pool.end();
  }
}

main();
