# AGENTS.md — Tabula Loom

## 1. Project Overview
**Tabula Loom** (`tabula-loom`) is a multi-sheet `.xlsx` spreadsheet editor and dynamic field configuration platform built with **React + TypeScript + Vite**, **SheetJS (`xlsx`)**, **Express**, and **PostgreSQL + Railway Storage**.

Each worksheet inside an uploaded `.xlsx` file has its own independent, dynamic field schema (`fields: FieldConfig[]`) and row records (`rows: RowRecord[]`).

---

## 2. Tech Stack & Package Manager
- **Package Manager**: **Yarn (`yarn@1.22.21`)**. Always use `yarn` (`yarn install`, `yarn dev`, `yarn build`, `yarn start`). **Never** commit `package-lock.json` so Railway's Nixpacks builder cleanly detects Yarn.
- **Frontend**: React 19, TypeScript, Vite (`src/`).
- **Spreadsheet Engine**: SheetJS (`xlsx`) using lossless 2D matrix extraction (`header: 1`) in `src/xlsxUtils.ts`.
- **Backend Server**: Express + TypeScript (`server/index.ts`), executed via `tsx`.
- **Database**: PostgreSQL (`pg` pool in `server/db.ts`) using a Relational + `JSONB` hybrid schema.
- **Object / File Storage**: Railway S3-compatible Buckets (`@aws-sdk/client-s3`) or Railway Volume / local disk fallback (`server/storage.ts`).

---

## 3. Directory Structure
```text
tabula-loom/
├── index.html            # Root HTML with Geist, Geist Mono, and Newsreader fonts
├── package.json          # Yarn scripts & dependencies
├── vite.config.ts        # Vite config (port 5175, proxies /api -> localhost:3001)
├── .env.example          # Template for DATABASE_URL and Railway Storage variables
├── src/
│   ├── main.tsx          # React entry point
│   ├── App.tsx           # Dual-mode workspace (Spreadsheet Grid + Fields Configuration)
│   ├── types.ts          # TypeScript interfaces for WorkbookProject, SheetData, FieldConfig
│   ├── xlsxUtils.ts      # Lossless 2D matrix parser, header detector, validator, exporter
│   ├── sampleData.ts     # Blank initial workbook state (EMPTY_WORKBOOK)
│   ├── Icons.tsx         # Clean SVG icons and field type glyphs
│   └── index.css         # Utilitarian minimalist UI theme (#F7F6F3 / #FFFFFF, 1px borders)
└── server/
    ├── index.ts          # Express API (/api/health, /api/workbooks) + static dist/ server
    ├── db.ts             # PostgreSQL connection pool and automatic DDL initializer
    ├── storage.ts        # S3 Bucket / Railway Volume adapter for raw .xlsx binaries
    └── initDb.ts         # CLI helper (`yarn db:init`) to initialize PostgreSQL tables
```

---

## 4. Core Domain Rules (MUST PRESERVE)

### 4.1 Dynamic Per-Sheet Schemas
- Different sheets in the same `.xlsx` workbook rarely share the same columns.
- Never hardcode spreadsheet columns in SQL tables. Each `sheets` row in PostgreSQL stores its own `fields JSONB` array, and each `sheet_rows` record stores cell values keyed by `field.id` in `cells JSONB`.

### 4.2 Headerless Sheets vs. Header Rows (`headerRowIndex`)
- `headerRowIndex === -1`: The sheet has **no header row**. Columns must be labeled with Excel alphabet letters (`A`, `B`, `C`, ...) and machine keys (`col_a`, `col_b`, ...). **Row 1 must remain data row 1** and must never be swallowed as a header.
- `headerRowIndex >= 0`: 0-based row index in `rawMatrix` acting as the column headers.
- Duplicate header names in the same sheet (e.g., two `"Amount"` columns) or blank header cells must never overwrite each other; always parse sheets with `XLSX.utils.sheet_to_json(ws, { header: 1 })`.

### 4.3 Hidden Excel Worksheets (`hiddenState`)
- `parseXlsxFile` inspects `workbook.Workbook?.Sheets?.[i]?.Hidden` (`0` = visible, `1` = hidden, `2` = very hidden).
- Hidden sheets must remain hidden by default in the tab bar, and the UI must display the notification banner informing the user which sheets are hidden with a toggle to reveal them.

### 4.4 UI & Typography Guidelines
- **Minimalist Editorial Aesthetic**: Warm paper background (`#F7F6F3`), crisp white surfaces (`#FFFFFF`), `1px solid #EAEAEA` borders, and muted spot pastels.
- **Fonts**: `Geist` (UI & body), `Geist Mono` (machine keys, coordinates, numbers, formulas), `Newsreader` (headings). Avoid heavy drop shadows, neon gradients, or emojis in the UI.
- **No Dummy Data**: The default initial state in `src/sampleData.ts` (`EMPTY_WORKBOOK`) and new sheets created by the user must start completely blank (`A`, `B`, `C` columns with an empty row).

---

## 5. Development & Deployment Commands

### Local Development
```bash
# Install dependencies
yarn install

# Run both Express API (:3001) and Vite Frontend (:5175) concurrently
yarn dev

# Or run frontend only
yarn dev:web
```

### Production Build & Railway Deployment
Railway automatically runs `yarn build` followed by `yarn start`:
```bash
# Type-check and build Vite frontend into dist/
yarn build

# Start production Express server (serves both /api/* and static dist/)
yarn start
```

### Environment Variables (Railway)
- `DATABASE_URL`: PostgreSQL connection string (automatically injected when linking a Railway PostgreSQL service).
- `PORT`: Automatically injected by Railway in production.
- **Optional Storage**:
  - Railway S3 Bucket: `BUCKET`, `ENDPOINT`, `REGION`, `ACCESS_KEY_ID`, `SECRET_ACCESS_KEY`.
  - Railway Volume: `RAILWAY_VOLUME_MOUNT_PATH` (falls back to `./storage/workbooks` locally).
