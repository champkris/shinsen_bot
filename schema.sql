-- Shinsen Bot MySQL Schema
-- Run this file to create the database and tables

CREATE DATABASE IF NOT EXISTS shinsen_bot;
USE shinsen_bot;

-- Table: daily_records
-- Stores daily sales records for all product categories
-- Categories: orange, yuzu, pop (Shinsen Pop), mixed (น้ำผลไม้รวม), tomato (Tomato Yuzu)
CREATE TABLE IF NOT EXISTS daily_records (
  id INT AUTO_INCREMENT PRIMARY KEY,
  date DATE NOT NULL,
  category ENUM('orange', 'yuzu', 'pop', 'mixed', 'tomato') NOT NULL,
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
