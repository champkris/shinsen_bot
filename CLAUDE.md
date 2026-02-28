# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
npm start        # Production: node server.js
npm run dev      # Development: nodemon with auto-reload
```

No test framework is configured. Manual testing is done via the `/test` web endpoint (image upload) and `/detection-logs` page.

Database setup: `mysql < schema.sql`

Migration scripts: `node migrate-json.js` (JSON→MySQL), `node migrate-excel.js` (Excel→MySQL)

## Architecture

This is a LINE bot that detects Excel screenshots, extracts sales data via OCR, and stores it in MySQL. The entire application lives in two main files.

### server.js (~2400 lines)

The monolithic application file containing:
- **LINE webhook handler** (`POST /webhook`): Receives image messages, downloads from LINE, sends to GPT-4o Vision for Excel detection, then to Azure OCR for data extraction
- **OCR processing pipeline**: Azure Document Intelligence extracts tables → preprocesses merged cells (fill-down) → dynamically detects product columns from headers → extracts CDC warehouse totals → saves to MySQL
- **Web dashboard routes**: `/daily-report`, `/latest-ocr`, `/transformed-data`, `/detection-logs`, `/send-notification`, `/test` — all rendered as inline HTML (no template engine)

### db.js (~280 lines)

Database abstraction layer:
- MySQL connection pool (mysql2/promise)
- CDC name mapping: Thai warehouse names ↔ database column names
- Date format conversion: DD/MM/YYYY ↔ YYYY-MM-DD
- CRUD operations for `daily_records` and `detection_logs`
- Exports `CATEGORIES` array and `CATEGORY_NAMES` object used throughout

### Data Flow

1. Image received via LINE webhook
2. GPT-4o Vision determines if image is an Excel screenshot
3. Azure Document Intelligence OCR extracts table data
4. Table preprocessed: merged cells filled down, product columns detected from headers by keyword matching
5. CDC (warehouse) totals extracted per product category
6. Records saved to `daily_records` table (upsert on date+category)
7. Notification sent to configured LINE groups

### Key Concepts

- **CDC**: Warehouse/distribution center locations across Thailand (11 locations). Mapped between Thai short names (e.g., `บางบัวทอง`) and database columns (e.g., `cdc_bangbuathong`)
- **Categories**: 5 product types — orange, yuzu, pop (Shinsen Pop), mixed (น้ำผลไม้รวม), tomato (Tomato Yuzu)
- **Dynamic product detection**: Product columns are identified from Excel header rows using keyword matching, not hardcoded positions
- **Totals row exclusion**: The totals row at the bottom of tables is excluded from CDC calculations to avoid double-counting

### Environment Variables

See `.env.example`. Requires: LINE credentials, OpenAI API key, Azure Document Intelligence endpoint/key, MySQL connection details, and optional `NOTIFICATION_GROUP_IDS` (comma-separated LINE group IDs).

### Logging Convention

Console logs use `[TAG]` prefix pattern: `[CONFIG]`, `[DB]`, `[LOG]`, `[RECORD]`, `[EXTRACT]`, `[OCR]`, etc.
