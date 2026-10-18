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
├── index.html            # Root HTML with Inter, Poppins, and Geist Mono fonts
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

### 4.4 No Dummy Data
- The default initial state in `src/sampleData.ts` (`EMPTY_WORKBOOK`) and new sheets created by the user must start completely blank (`A`, `B`, `C` columns with an empty row).

---

## 5. UI, Typography & Styling Protocol (STRICT)

All AI agents and developers modifying the frontend (`src/index.css`, `src/App.tsx`, `src/Icons.tsx`) MUST follow this **Utilitarian Editorial Minimalist** design system:

### 5.1 Typography Hierarchy
- **Primary UI & Cell Text (`--font-sans`)**: `'Inter', -apple-system, BlinkMacSystemFont, 'Helvetica Neue', sans-serif`.
  - Base size: `13.5px`, `line-height: 1.5`.
  - Avoid Roboto and Open Sans (Inter is the approved sans).
- **Headings (`--font-serif`, now Poppins)**: `'Poppins', -apple-system, BlinkMacSystemFont, 'Helvetica Neue', sans-serif`.
  - Used for brand title (`.brand-title`), inspector titles (`.inspector-title`), dialog headers (`.dialog-title`), and schema view headings (`.schema-hero-title`).
  - Use light tracking (`letter-spacing: -0.01em`) and `font-weight: 500`.
- **Technical / Tabular Data (`--font-mono`)**: `'Geist Mono', 'SF Mono', Menlo, monospace`.
  - MUST be used for: machine keys (`field.key`), Excel coordinates (`A1`, `B2`), row/column indices, numeric/currency/date cells, formulas, and `<kbd>` shortcuts.

### 5.2 Color Palette (Warm Monochrome + Muted Spot Pastels)
- **Background & Surfaces**:
  - Canvas background (`--bg-canvas`): `#F7F6F3` (warm bone/paper).
  - Primary card/table surface (`--bg-surface`): `#FFFFFF`.
  - Subtle toolbar/header surface (`--bg-subtle`): `#FBFBFA`.
  - Hover surface (`--bg-hover`): `#F2F1EE`.
  - Selected row/column tint (`--bg-selected`): `#F0F6FA`.
- **Text Colors**:
  - Never use pure `#000000` for body text. Use `--text-primary: #111111` for headings/labels, `--text-body: #2F3437` for cell content, and `--text-muted: #787774` for secondary metadata.
- **Borders**:
  - All structural dividers, table cells, cards, and inputs MUST use `1px solid var(--border-hairline)` (`#EAEAEA`).
- **Spot Pastels (Status Pills & Type Badges Only)**:
  - Blue (`--pastel-blue-bg: #E1F3FE`, `--pastel-blue-fg: #1F6C9F`)
  - Green (`--pastel-green-bg: #EDF3EC`, `--pastel-green-fg: #346538`)
  - Yellow (`--pastel-yellow-bg: #FBF3DB`, `--pastel-yellow-fg: #956400`)
  - Red (`--pastel-red-bg: #FDEBEC`, `--pastel-red-fg: #9F2F2D`)

### 5.3 Layout & Anti-Overlap Rules (CRITICAL)
- **No Overlapping Flex Items in Toolbars**:
  - Top navigation bars (`.top-navbar`, `.sheet-toolbar`, `.formula-bar`) must keep interactive controls (`button`, `.mode-switcher`, `.status-pill`, `.top-actions`) at `flex-shrink: 0` and `white-space: nowrap` so button labels never wrap onto two lines or collide with adjacent pills.
  - Text inputs inside toolbars (such as `.workbook-title-input`) must use flexible sizing (`flex: 0 1 190px; min-width: 110px; overflow: hidden; text-overflow: ellipsis`) so long filenames or high validation counts (`119 VALIDATION ISSUES`) never push or overlap neighboring controls.
- **Sticky Grid Headers & Row Numbers**:
  - `.sheet-table thead th` must remain `position: sticky; top: 0; z-index: 10`.
  - `.th-row-index` and `.td-row-index` must remain `position: sticky; left: 0` so row numbers stay pinned during horizontal scrolling.
- **Banned Visual Elements**:
  - **NO emojis** anywhere in the UI, buttons, badges, or headings. Use clean 1.5px stroke SVG icons in `src/Icons.tsx` or monospace type glyphs (`Aa`, `#`, `Cal`, `Opt`, `T/F`, `fx`, `Ref`).
  - **NO heavy drop shadows**, neon gradients, or 3D glassmorphism. Primary buttons are flat solid `#111111` with `#FFFFFF` text and `4px–5px` border radius.

---

## 6. Git Commit Message Convention (STRICT)

All commits MUST follow the **Conventional Commits** specification:

```text
<type>(<scope>): <imperative summary under 72 chars>

- <what changed and why>
- <key architectural, UI, or schema details>
```

### Allowed `<type>` Prefixes
- `feat`: A new feature (e.g., `feat(parser): support headerless sheets with A, B, C columns`)
- `fix`: A bug or UI layout fix (e.g., `fix(ui): prevent top navbar validation badge from overlapping mode switcher`)
- `refactor`: Code restructuring that neither fixes a bug nor adds a feature
- `chore`: Build scripts, dependencies, or tooling updates (e.g., `chore(deps): configure yarn lockfile and Railway scripts`)
- `docs`: Documentation updates (`AGENTS.md`, `ENVIRONMENT_SETUP.md`, comments)
- `perf`: Performance improvements

### Commit Rules
- Use the **imperative mood** in the subject line (`add`, `fix`, `prevent`, not `added` or `fixes`).
- Common scopes: `ui`, `parser`, `schema`, `db`, `storage`, `api`, `deps`.
- **Do NOT** append AI attribution trailers (`Co-Authored-By:` or `Generated with...`) unless explicitly requested.

---

## 7. Development & Deployment Commands

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
