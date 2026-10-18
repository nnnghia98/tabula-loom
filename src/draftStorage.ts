import type { WorkbookProject } from './types';

/**
 * Local unsaved-work draft, stored in IndexedDB (large quota, unlike localStorage).
 * The explicit Save button still persists to PostgreSQL; the draft is cleared on Save.
 */
export interface WorkbookDraft {
  project: WorkbookProject;
  rawUploadedBase64: string | null;
  activeSheetId: string;
  savedAt: number;
}

const DB_NAME = 'tabula-loom';
const STORE = 'drafts';
const KEY = 'current';

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function withStore<T>(
  mode: IDBTransactionMode,
  run: (store: IDBObjectStore) => IDBRequest<T>
): Promise<T> {
  const db = await openDb();
  try {
    return await new Promise<T>((resolve, reject) => {
      const req = run(db.transaction(STORE, mode).objectStore(STORE));
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  } finally {
    db.close();
  }
}

export async function loadDraft(): Promise<WorkbookDraft | null> {
  try {
    return ((await withStore('readonly', (s) => s.get(KEY))) as WorkbookDraft) ?? null;
  } catch {
    return null;
  }
}

export async function saveDraft(draft: WorkbookDraft): Promise<boolean> {
  try {
    await withStore('readwrite', (s) => s.put(draft, KEY));
    return true;
  } catch {
    return false;
  }
}

export async function clearDraft(): Promise<void> {
  try {
    await withStore('readwrite', (s) => s.delete(KEY));
  } catch {
    /* ignore */
  }
}
