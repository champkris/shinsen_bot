-- Shinsen Bot MySQL Schema
-- Run this file to create the database and tables

CREATE DATABASE IF NOT EXISTS shinsen_bot;
USE shinsen_bot;

-- Table: daily_records
-- Stores daily sales records for all product categories
-- Categories: orange, yuzu, pop (Shinsen Pop), mixed (น้ำผลไม้รวม), tomato (Tomato Yuzu), pineapple (น้ำสับปะรด)
CREATE TABLE IF NOT EXISTS daily_records (
  id INT AUTO_INCREMENT PRIMARY KEY,
  date DATE NOT NULL,
  category ENUM('orange', 'yuzu', 'pop', 'mixed', 'tomato', 'pineapple') NOT NULL,
  timestamp DATETIME NOT NULL,
  fc33_hadyai_sum INT DEFAULT 0,
  total_sum INT DEFAULT 0,
  cdc_bangbuathong INT DEFAULT 0,
  cdc_nakhonratchasima INT DEFAULT 0,
  cdc_nakhonsawan INT DEFAULT 0,
  cdc_chonburi INT DEFAULT 0,
  cdc_mahachai INT DEFAULT 0,
  cdc_suvarnabhumi INT DEFAULT 0,
  cdc_hadyai INT DEFAULT 0,
  cdc_phuket INT DEFAULT 0,
  cdc_chiangmai INT DEFAULT 0,
  cdc_surat INT DEFAULT 0,
  cdc_khonkaen INT DEFAULT 0,
  khon_kaen_laos INT DEFAULT NULL,
  khon_kaen_cambodia INT DEFAULT NULL,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  UNIQUE KEY unique_date_category (date, category),
  INDEX idx_date (date),
  INDEX idx_category (category),
  INDEX idx_date_category (date, category)
);

-- Table: monthly_targets
-- Stores sales targets per category per month
CREATE TABLE IF NOT EXISTS monthly_targets (
  id INT AUTO_INCREMENT PRIMARY KEY,
  year INT NOT NULL,
  month INT NOT NULL,
  category VARCHAR(50) NOT NULL,
  target_value BIGINT NOT NULL DEFAULT 0,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  UNIQUE KEY unique_target (year, month, category),
  INDEX idx_year_month (year, month)
);

-- Table: watchdog_alerts
-- One row per day already reported missing, so the watchdog alerts once per date
-- instead of on every run
CREATE TABLE IF NOT EXISTS watchdog_alerts (
  alert_date DATE NOT NULL,
  kind VARCHAR(32) NOT NULL DEFAULT 'missing_data',
  sent_at DATETIME NOT NULL,
  PRIMARY KEY (alert_date, kind)
);

-- Table: pending_events
-- Durable work queue for LINE image events. The webhook writes a row here and
-- returns 200 immediately; the actual OCR run happens later. Passenger tears the
-- app process down once no request is in flight, so any work left running after
-- the response can be killed mid-flight — a row that survives in this table is
-- what makes that recoverable instead of a silently lost screenshot.
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
);

-- Table: detection_logs
-- Stores logs from image detection attempts
CREATE TABLE IF NOT EXISTS detection_logs (
  id INT AUTO_INCREMENT PRIMARY KEY,
  timestamp DATETIME NOT NULL,
  message_id VARCHAR(255),
  group_id VARCHAR(255),
  user_id VARCHAR(255),
  status ENUM('success', 'failed', 'error') NOT NULL,
  date DATE DEFAULT NULL,
  categories JSON DEFAULT NULL,
  records_created INT DEFAULT 0,
  reason TEXT DEFAULT NULL,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  INDEX idx_timestamp (timestamp),
  INDEX idx_status (status),
  INDEX idx_group_id (group_id)
);
