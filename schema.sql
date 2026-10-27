-- Shinsen Bot MySQL Schema
-- Run this file to create the database and tables

CREATE DATABASE IF NOT EXISTS shinsen_bot;
USE shinsen_bot;

-- Table: daily_records
-- Stores daily sales records for all product categories
-- Categories: orange, yuzu, pop (Shinsen Pop), mixed (น้ำผลไม้รวม), tomato (Tomato Yuzu), pineapple (น้ำสับปะรด), watermelon (น้ำแตงโม)
CREATE TABLE IF NOT EXISTS daily_records (
  id INT AUTO_INCREMENT PRIMARY KEY,
  date DATE NOT NULL,
  category ENUM('orange', 'yuzu', 'pop', 'mixed', 'tomato', 'pineapple', 'watermelon') NOT NULL,
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

-- Table: product_categories
-- Juices added automatically when a plan carried a column no keyword matched (or named on
-- the juice page). Merged into the built-in categories at runtime; keywords is a JSON
-- array matched against headers.
CREATE TABLE IF NOT EXISTS product_categories (
  category VARCHAR(50) NOT NULL PRIMARY KEY,
  thai_name VARCHAR(100) NOT NULL,
  english_name VARCHAR(100) DEFAULT NULL,
  color VARCHAR(7) DEFAULT NULL,
  icon VARCHAR(16) DEFAULT NULL,
  keywords TEXT NOT NULL,
  created_by VARCHAR(64) DEFAULT NULL,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) DEFAULT CHARSET=utf8mb4;

-- Table: pending_juices
-- Every juice column found on a plan with no matching keyword. Auto-added ones are kept
-- as history (status 'added'); one whose name couldn't be read waits as 'pending' until
-- it is named or ignored on the juice page. context_json keeps the OCR table so it is
-- recorded then without resending the image.
CREATE TABLE IF NOT EXISTS pending_juices (
  id INT AUTO_INCREMENT PRIMARY KEY,
  plan_date DATE NOT NULL,
  header VARCHAR(255) NOT NULL,
  keyword VARCHAR(100) NOT NULL,
  english_name VARCHAR(100) DEFAULT NULL,
  color VARCHAR(7) DEFAULT NULL,
  total INT NOT NULL DEFAULT 0,
  context_json MEDIUMTEXT NOT NULL,
  status ENUM('pending', 'added', 'ignored') NOT NULL DEFAULT 'pending',
  category VARCHAR(50) DEFAULT NULL,
  resolved_by VARCHAR(64) DEFAULT NULL,
  resolved_at DATETIME DEFAULT NULL,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  UNIQUE KEY uniq_plan_keyword (plan_date, keyword)
) DEFAULT CHARSET=utf8mb4;

-- Table: juice_display
-- Display settings for the built-in juices, set on the juice page: a name, English name,
-- colour or icon here replaces the one in code (NULL keeps it). Keywords are not affected.
CREATE TABLE IF NOT EXISTS juice_display (
  category VARCHAR(50) NOT NULL PRIMARY KEY,
  thai_name VARCHAR(100) DEFAULT NULL,
  english_name VARCHAR(100) DEFAULT NULL,
  color VARCHAR(7) DEFAULT NULL,
  icon VARCHAR(16) DEFAULT NULL
) DEFAULT CHARSET=utf8mb4;

-- Table: juice_aliases
-- Extra spellings left by the juice page: merging a duplicate juice points its keywords at
-- the juice it was merged into; deleting a juice points them at NULL, so a column with
-- that name is never auto-added again.
CREATE TABLE IF NOT EXISTS juice_aliases (
  keyword VARCHAR(100) NOT NULL PRIMARY KEY,
  category VARCHAR(50) DEFAULT NULL
) DEFAULT CHARSET=utf8mb4;

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
