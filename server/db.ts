import pg from 'pg';
import dotenv from 'dotenv';

dotenv.config();

const { Pool } = pg;

export const DATABASE_URL = process.env.DATABASE_URL || '';

export const pool = DATABASE_URL
  ? new Pool({
      connectionString: DATABASE_URL,
      ssl:
        DATABASE_URL.includes('railway.app') ||
        DATABASE_URL.includes('rlwy.net') ||
        process.env.PGSSLMODE === 'require'
          ? { rejectUnauthorized: false }
          : undefined,
    })
  : null;

export const SCHEMA_SQL = `
CREATE TABLE IF NOT EXISTS workbooks (
  id TEXT PRIMARY KEY,
  file_name TEXT NOT NULL,
  storage_key TEXT,
  last_modified TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS sheets (
  id TEXT PRIMARY KEY,
  workbook_id TEXT NOT NULL REFERENCES workbooks(id) ON DELETE CASCADE,
  name TEXT NOT NULL,
  description TEXT,
  sheet_order INT NOT NULL DEFAULT 0,
  hidden_state SMALLINT NOT NULL DEFAULT 0,
  header_row_index INT NOT NULL DEFAULT 0,
  fields JSONB NOT NULL DEFAULT '[]'::jsonb,
  raw_matrix JSONB,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_sheets_workbook_id ON sheets(workbook_id, sheet_order);

CREATE TABLE IF NOT EXISTS sheet_rows (
  id TEXT PRIMARY KEY,
  sheet_id TEXT NOT NULL REFERENCES sheets(id) ON DELETE CASCADE,
  row_order INT NOT NULL DEFAULT 0,
  excel_row_number INT,
  cells JSONB NOT NULL DEFAULT '{}'::jsonb
);

CREATE INDEX IF NOT EXISTS idx_sheet_rows_sheet_id ON sheet_rows(sheet_id, row_order);
CREATE INDEX IF NOT EXISTS idx_sheet_rows_cells_gin ON sheet_rows USING GIN (cells);
`;

export async function ensureDatabaseSchema(): Promise<{ connected: boolean; error?: string }> {
  if (!pool) {
    return {
      connected: false,
      error: 'DATABASE_URL is not configured yet. Add DATABASE_URL in .env or Railway variables.',
    };
  }
  try {
    await pool.query(SCHEMA_SQL);
    return { connected: true };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.error('[Tabula Loom DB] Schema initialization error:', message);
    return { connected: false, error: message };
  }
}
