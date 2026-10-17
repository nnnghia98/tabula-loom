# Environment & Railway Deployment Guide — Tabula Loom

This guide walks the engineering and DevOps team through setting up **Tabula Loom** locally, configuring **PostgreSQL**, and deploying the full stack (Frontend + API + PostgreSQL + Storage) on **Railway**.

---

## 1. Prerequisites (Local Development)

- **Node.js**: `v20+` or `v22+` (LTS recommended)
- **Package Manager**: **Yarn** (`v1.22.x`). Do **not** use `npm install` or commit a `package-lock.json` file, as Railway relies on `yarn.lock` to auto-detect Yarn.
- **PostgreSQL**: `v14+` (local instance, Docker container, or a Railway PostgreSQL connection string)

---

## 2. Local Environment Setup

### Step 1: Install Dependencies
```bash
cd /path/to/tabula-loom
yarn install
```

### Step 2: Create `.env` File
Copy the template from `.env.example`:
```bash
cp .env.example .env
```

Edit `.env` with your local or remote PostgreSQL credentials:
```dotenv
# Required for database persistence
DATABASE_URL=postgresql://postgres:postgres@localhost:5432/tabula_loom

# Optional: S3-compatible Bucket credentials (for storing original .xlsx files in cloud object storage)
BUCKET=
ENDPOINT=
REGION=auto
ACCESS_KEY_ID=
SECRET_ACCESS_KEY=

# Optional: Mounted volume path (if not using S3 Bucket, defaults to ./storage/workbooks locally)
RAILWAY_VOLUME_MOUNT_PATH=
```

### Step 3: Initialize Database Schema
Tabula Loom automatically creates its tables on server startup (`server/db.ts`), or you can run the schema initializer explicitly:
```bash
yarn db:init
```

### Step 4: Start the Development Servers
```bash
# Starts both the Express API (:3001) and Vite Frontend (:5175) concurrently
yarn dev
```
- **Frontend UI**: `http://localhost:5175`
- **Backend API Healthcheck**: `http://localhost:3001/api/health`

*(Vite automatically proxies `/api/*` requests from `:5175` to `:3001` in development.)*

---

## 3. PostgreSQL Database Architecture

Because uploaded `.xlsx` workbooks contain multiple sheets where each sheet has a completely different column structure, Tabula Loom uses a **Relational + `JSONB` hybrid schema** (defined in `server/db.ts`):

| Table | Purpose | Key Columns |
| :--- | :--- | :--- |
| **`workbooks`** | Top-level `.xlsx` file metadata | `id` (PK), `file_name`, `storage_key` (path/URL to raw `.xlsx` binary), `last_modified`, `created_at` |
| **`sheets`** | Worksheet tabs & per-sheet dynamic schema | `id` (PK), `workbook_id` (FK -> `workbooks.id` CASCADE), `name`, `sheet_order`, `hidden_state` (`0` visible, `1` hidden, `2` very hidden), `header_row_index` (`-1` headerless `A,B,C`, `>=0` header row), `fields` (`JSONB` array of `FieldConfig`), `raw_matrix` (`JSONB`) |
| **`sheet_rows`** | Row records for each worksheet | `id` (PK), `sheet_id` (FK -> `sheets.id` CASCADE), `row_order`, `excel_row_number`, `cells` (`JSONB` object mapping `field.id -> value` with a `GIN` index) |

### SSL Behavior
In `server/db.ts`, SSL (`rejectUnauthorized: false`) is automatically enabled whenever `DATABASE_URL` points to `*.railway.app`, `*.rlwy.net`, or when `PGSSLMODE=require` is set.

---

## 4. Deploying on Railway (Step-by-Step)

Tabula Loom is architected to deploy on Railway as a **single unified web service** connected to a **Railway PostgreSQL database** and **Railway Storage (Bucket or Volume)**.

### Step 1: Create the Railway Project & Add PostgreSQL
1. Log in to [Railway](https://railway.app/) and click **New Project**.
2. Select **Deploy from GitHub repo** and choose the `tabula-loom` repository (or add **PostgreSQL** first via **+ New** → **Database** → **Add PostgreSQL**).
3. In your **App Service** settings, go to the **Variables** tab and add a variable reference to your Railway Postgres service:
   ```dotenv
   DATABASE_URL=${{Postgres.DATABASE_URL}}
   ```

### Step 2: Configure Railway Storage for Original `.xlsx` Files
When a user uploads an `.xlsx` file and clicks **Save to DB**, Tabula Loom saves both the parsed JSONB sheets in PostgreSQL and the raw `.xlsx` binary in storage (`server/storage.ts`). You can choose either option on Railway:

#### Option A: Railway S3-Compatible Bucket *(Recommended)*
1. In your Railway project canvas, click **+ New** → **Bucket** to create a Railway Object Storage bucket.
2. In your **App Service** → **Variables** tab, link the Bucket credentials:
   ```dotenv
   BUCKET=${{Bucket.BUCKET}}
   ENDPOINT=${{Bucket.ENDPOINT}}
   REGION=${{Bucket.REGION}}
   ACCESS_KEY_ID=${{Bucket.ACCESS_KEY_ID}}
   SECRET_ACCESS_KEY=${{Bucket.SECRET_ACCESS_KEY}}
   ```
3. When these variables are present, `server/storage.ts` automatically switches `storageMode` to `"railway-s3-bucket"`.

#### Option B: Railway Volume
1. If you prefer a mounted filesystem volume instead of S3, click **+ New** → **Volume**, attach it to your App Service, and set its mount path to `/data`.
2. Railway automatically injects `RAILWAY_VOLUME_MOUNT_PATH=/data`, and `server/storage.ts` will store uploaded `.xlsx` binaries in `/data/workbooks/`.

### Step 3: Build & Start Commands on Railway
Railway (Nixpacks / Railpack) automatically reads `package.json` and `yarn.lock`:
- **Install Command**: `yarn install`
- **Build Command**: `yarn build` *(compiles TypeScript & bundles the React app into `dist/`)*
- **Start Command**: `yarn start` *(runs `tsx server/index.ts`, which initializes the PostgreSQL tables, serves `/api/*`, and serves the static frontend from `dist/` on `0.0.0.0:$PORT`)*

### Step 4: Verify Deployment
Once deployed, open your Railway public domain:
1. Check `https://<your-railway-domain>/api/health` — it should return:
   ```json
   {
     "ok": true,
     "dbConnected": true,
     "dbError": null,
     "storageMode": "railway-s3-bucket"
   }
   ```
2. Open `https://<your-railway-domain>/` in the browser — the **PostgreSQL** indicator in the top navigation bar will show a green status dot (`PostgreSQL Connected`).
