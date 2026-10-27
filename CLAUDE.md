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
- **Operational routes**: `/health` (DB status + last recorded date, 503 when the DB is down), `/watchdog` (missing-data check, meant to be called by an external scheduler), `/diag` (OpenAI connectivity probe)

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
- **Categories**: 7 product types — orange, yuzu, pop (Shinsen Pop), mixed (น้ำผลไม้รวม), tomato (Tomato Yuzu), pineapple (น้ำสับปะรด), watermelon (น้ำแตงโม)
- **Dynamic product detection**: Product columns are identified from Excel header rows using keyword matching, not hardcoded positions
- **New juices (auto-added)**: A juice column with bottles but no matching keyword is saved to `pending_juices` (with its OCR table). If its header name is usable it is added straight away as a `juice_<id>` category in `product_categories` — English label machine-translated via OpenAI (5s timeout, falls back to the Thai name), dashboard colour sampled from the column's cells in the plan image (`sampleColumnColor`, via `jpeg-js` and Azure cell polygons; green when unreadable), `daily_records.category` ENUM widened — and recorded; the dev alert only reports it. An unreadable name stays `pending` until named on the juice page. `db.refreshCategories()` merges these into `CATEGORIES`/`CATEGORY_NAMES` in place
- **Juice page** (`/juices/<token>`): linked only from dev alerts, no login — the token is an HMAC of `LINE_CHANNEL_SECRET`, a wrong one gets 404, and it is masked in request logs. Lists every juice. For the built-in ones it only changes how they look on the dashboards — Thai/English name, colour, icon — stored in `juice_display` as overrides of the defaults in code (`juiceIcon`/`juiceColor`/`db.CATEGORY_NAMES` read the DB first), with a reset; their keywords stay in code. For auto-added ones it edits names, colour and icon (icons come from a fixed fruit-emoji list, `JUICE_ICON_CHOICES`), names pending columns, merges a duplicate into another juice (its keywords become `juice_aliases` of the target; refused when both have a record on the same day), or deletes a junk juice with its records (its keywords become aliases to NULL, never auto-added again)
- **Totals row exclusion**: The totals row at the bottom of tables is excluded from CDC calculations to avoid double-counting
- **Missing-data watchdog**: A webhook LINE never delivers leaves no trace in this app — no log, no `detection_logs` row — so a whole day can go missing silently. `runWatchdog()` checks `daily_records` for expected days instead of trusting the delivery path, and alerts `ALERT_GROUP_IDS` once per missing date (dedup via the `watchdog_alerts` table). It runs on boot and via `GET /watchdog`; an external scheduler is still required, since a total outage means the app never boots
- **Webhook idempotency**: `/webhook` acknowledges LINE before processing (a full OCR run takes ~15s, well past LINE's timeout). `handleEvent` skips images whose message ID already has a `detection_logs` row, so LINE webhook redelivery cannot double-process a screenshot

### Environment Variables

See `.env.example`. Requires: LINE credentials, OpenAI API key, Azure Document Intelligence endpoint/key, MySQL connection details, and optional `NOTIFICATION_GROUP_IDS` (comma-separated LINE group IDs).

### Logging Convention

Console logs use `[TAG]` prefix pattern: `[CONFIG]`, `[DB]`, `[LOG]`, `[RECORD]`, `[EXTRACT]`, `[OCR]`, etc.
