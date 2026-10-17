import { ensureDatabaseSchema, pool } from './db.js';

async function run() {
  const status = await ensureDatabaseSchema();
  if (status.connected) {
    console.log('✅ PostgreSQL schema initialized successfully.');
  } else {
    console.error('❌ Could not initialize PostgreSQL schema:', status.error);
    process.exitCode = 1;
  }
  if (pool) {
    await pool.end();
  }
}

run();
