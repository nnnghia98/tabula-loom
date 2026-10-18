import express from 'express';
import cors from 'cors';
import path from 'node:path';
import fs from 'node:fs';
import { pool, ensureDatabaseSchema } from './db.js';
import {
  getStorageMode,
  saveOriginalXlsxBinary,
  readOriginalXlsxBinary,
  deleteOriginalXlsxBinary,
} from './storage.js';
import type { WorkbookProject, SheetData, RowRecord, FieldConfig, CellValue } from '../src/types.js';

const app = express();
const PORT = Number(process.env.PORT) || 3001;

app.use(cors());
app.use(express.json({ limit: '50mb' }));

let dbReadyState: { connected: boolean; error?: string } = { connected: false };

// Initialize schema on startup
const databaseReady = ensureDatabaseSchema().then((status) => {
  dbReadyState = status;
  if (status.connected) {
    console.log('[Tabula Loom API] PostgreSQL connected and dynamic schema tables verified.');
  } else {
    console.warn('[Tabula Loom API] Running without PostgreSQL connection:', status.error);
  }
});

app.use('/api', async (_req, _res, next) => {
  await databaseReady;
  next();
});

// 1. Health & Database Status
app.get('/api/health', async (_req, res) => {
  if (!dbReadyState.connected && pool) {
    dbReadyState = await ensureDatabaseSchema();
  }
  res.json({
    ok: true,
    dbConnected: dbReadyState.connected,
    dbError: dbReadyState.error || null,
    storageMode: getStorageMode(),
  });
});

// 2. List all saved workbooks in PostgreSQL
app.get('/api/workbooks', async (_req, res) => {
  if (!pool || !dbReadyState.connected) {
    res.json({ dbConnected: false, workbooks: [] });
    return;
  }

  try {
    const result = await pool.query(`
      SELECT
        w.id,
        w.file_name AS "fileName",
        w.storage_key AS "storageKey",
        w.last_modified AS "lastModified",
        w.created_at AS "createdAt",
        COUNT(DISTINCT s.id)::int AS "sheetCount"
      FROM workbooks w
      LEFT JOIN sheets s ON s.workbook_id = w.id
      GROUP BY w.id
      ORDER BY w.last_modified DESC
    `);
    res.json({ dbConnected: true, workbooks: result.rows });
  } catch (err) {
    res.status(500).json({
      error: err instanceof Error ? err.message : 'Failed to list workbooks',
    });
  }
});

// 3. Load a single workbook with all dynamic sheets, field schemas, and rows
app.get('/api/workbooks/:id', async (req, res) => {
  if (!pool || !dbReadyState.connected) {
    res.status(503).json({ error: 'PostgreSQL is not connected' });
    return;
  }

  const workbookId = req.params.id;
  try {
    const wbRes = await pool.query(
      `SELECT id, file_name, storage_key, last_modified FROM workbooks WHERE id = $1`,
      [workbookId]
    );
    if (wbRes.rowCount === 0) {
      res.status(404).json({ error: 'Workbook not found' });
      return;
    }

    const wbRow = wbRes.rows[0];
    const sheetsRes = await pool.query(
      `SELECT id, name, description, sheet_order, hidden_state, header_row_index, fields, raw_matrix
       FROM sheets
       WHERE workbook_id = $1
       ORDER BY sheet_order ASC`,
      [workbookId]
    );

    const sheets: SheetData[] = [];
    for (const sRow of sheetsRes.rows) {
      const rowsRes = await pool.query(
        `SELECT id, excel_row_number, cells
         FROM sheet_rows
         WHERE sheet_id = $1
         ORDER BY row_order ASC`,
        [sRow.id]
      );

      const rows: RowRecord[] = rowsRes.rows.map((r) => ({
        id: r.id,
        excelRowNumber: r.excel_row_number ?? undefined,
        cells: (r.cells || {}) as Record<string, CellValue>,
      }));

      sheets.push({
        id: sRow.id,
        name: sRow.name,
        description: sRow.description || undefined,
        hiddenState: (sRow.hidden_state || 0) as 0 | 1 | 2,
        headerRowIndex: sRow.header_row_index ?? 0,
        fields: (sRow.fields || []) as FieldConfig[],
        rawMatrix: (sRow.raw_matrix || undefined) as CellValue[][] | undefined,
        rows,
      });
    }

    const project: WorkbookProject & { id: string; hasOriginalBinary: boolean } = {
      id: wbRow.id,
      fileName: wbRow.file_name,
      lastModified: new Date(wbRow.last_modified).toISOString(),
      hasOriginalBinary: Boolean(wbRow.storage_key),
      sheets,
    };

    res.json(project);
  } catch (err) {
    res.status(500).json({
      error: err instanceof Error ? err.message : 'Failed to load workbook',
    });
  }
});

// 4. Create or Upsert an entire WorkbookProject (dynamic sheets + JSONB fields + rows + optional raw base64 .xlsx binary)
app.post('/api/workbooks', async (req, res) => {
  if (!pool || !dbReadyState.connected) {
    res.status(503).json({
      error: 'PostgreSQL is not connected. Set DATABASE_URL in .env or Railway.',
    });
    return;
  }

  const body = req.body as {
    id?: string;
    fileName: string;
    sheets: SheetData[];
    rawFileBase64?: string;
  };

  if (!body || !body.fileName || !Array.isArray(body.sheets)) {
    res.status(400).json({ error: 'Invalid workbook payload' });
    return;
  }

  const workbookId = body.id || `wb_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 6)}`;
  const client = await pool.connect();

  try {
    await client.query('BEGIN');

    let storageKey: string | null = null;
    if (body.rawFileBase64) {
      const binaryBuffer = Buffer.from(body.rawFileBase64, 'base64');
      storageKey = await saveOriginalXlsxBinary(workbookId, body.fileName, binaryBuffer);
    }

    await client.query(
      `INSERT INTO workbooks (id, file_name, storage_key, last_modified)
       VALUES ($1, $2, COALESCE($3, NULL), NOW())
       ON CONFLICT (id) DO UPDATE SET
         file_name = EXCLUDED.file_name,
         storage_key = COALESCE(EXCLUDED.storage_key, workbooks.storage_key),
         last_modified = NOW()`,
      [workbookId, body.fileName, storageKey]
    );

    // Replace sheets and rows cleanly inside the transaction so dynamic schema changes stay in sync
    await client.query(`DELETE FROM sheets WHERE workbook_id = $1`, [workbookId]);

    for (let sIdx = 0; sIdx < body.sheets.length; sIdx++) {
      const sheet = body.sheets[sIdx];
      // Ensure sheet ID is globally scoped to workbookId to avoid PK collisions across multiple saved workbooks
      const scopedSheetId = sheet.id.startsWith(`${workbookId}_`)
        ? sheet.id
        : `${workbookId}_${sheet.id}`;

      await client.query(
        `INSERT INTO sheets (
           id, workbook_id, name, description, sheet_order, hidden_state, header_row_index, fields, raw_matrix, updated_at
         ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8::jsonb, $9::jsonb, NOW())`,
        [
          scopedSheetId,
          workbookId,
          sheet.name,
          sheet.description || null,
          sIdx,
          sheet.hiddenState || 0,
          sheet.headerRowIndex ?? 0,
          JSON.stringify(sheet.fields || []),
          sheet.rawMatrix ? JSON.stringify(sheet.rawMatrix) : null,
        ]
      );

      for (let rIdx = 0; rIdx < sheet.rows.length; rIdx++) {
        const row = sheet.rows[rIdx];
        const scopedRowId = row.id.startsWith(`${scopedSheetId}_`)
          ? row.id
          : `${scopedSheetId}_${row.id}`;

        await client.query(
          `INSERT INTO sheet_rows (id, sheet_id, row_order, excel_row_number, cells)
           VALUES ($1, $2, $3, $4, $5::jsonb)`,
          [
            scopedRowId,
            scopedSheetId,
            rIdx,
            row.excelRowNumber ?? rIdx + 1,
            JSON.stringify(row.cells || {}),
          ]
        );
      }
    }

    await client.query('COMMIT');
    res.json({
      ok: true,
      id: workbookId,
      fileName: body.fileName,
      lastModified: new Date().toISOString(),
    });
  } catch (err) {
    await client.query('ROLLBACK');
    console.error('[Tabula Loom API] Failed to save workbook:', err);
    res.status(500).json({
      error: err instanceof Error ? err.message : 'Failed to save workbook to PostgreSQL',
    });
  } finally {
    client.release();
  }
});

// 5. Download original uploaded .xlsx binary from Railway Storage / disk
app.get('/api/workbooks/:id/original-file', async (req, res) => {
  if (!pool || !dbReadyState.connected) {
    res.status(503).json({ error: 'PostgreSQL is not connected' });
    return;
  }
  const wbRes = await pool.query(`SELECT file_name, storage_key FROM workbooks WHERE id = $1`, [
    req.params.id,
  ]);
  if (wbRes.rowCount === 0 || !wbRes.rows[0].storage_key) {
    res.status(404).json({ error: 'Original .xlsx file not found in storage' });
    return;
  }

  const stored = await readOriginalXlsxBinary(wbRes.rows[0].storage_key);
  if (!stored) {
    res.status(404).json({ error: 'Stored binary could not be read' });
    return;
  }

  res.setHeader(
    'Content-Type',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );
  res.setHeader('Content-Disposition', `attachment; filename="${wbRes.rows[0].file_name}"`);
  res.send(stored.buffer);
});

// 6. Delete a saved workbook from PostgreSQL & Storage
app.delete('/api/workbooks/:id', async (req, res) => {
  if (!pool || !dbReadyState.connected) {
    res.status(503).json({ error: 'PostgreSQL is not connected' });
    return;
  }
  try {
    const wbRes = await pool.query(`SELECT storage_key FROM workbooks WHERE id = $1`, [
      req.params.id,
    ]);
    if (wbRes.rowCount && wbRes.rows[0].storage_key) {
      await deleteOriginalXlsxBinary(wbRes.rows[0].storage_key);
    }
    await pool.query(`DELETE FROM workbooks WHERE id = $1`, [req.params.id]);
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({
      error: err instanceof Error ? err.message : 'Failed to delete workbook',
    });
  }
});

// Serve static Vite build in production (for single-service Railway deployment)
const distDir = path.resolve(process.cwd(), 'dist');
if (fs.existsSync(distDir)) {
  app.use(express.static(distDir));
  app.get('*', (req, res, next) => {
    if (req.path.startsWith('/api')) {
      next();
      return;
    }
    res.sendFile(path.join(distDir, 'index.html'));
  });
}

if (!process.env.VERCEL) {
  app.listen(PORT, '0.0.0.0', () => {
    console.log(`[Tabula Loom Server] Listening on http://0.0.0.0:${PORT}`);
  });
}

export default app;
