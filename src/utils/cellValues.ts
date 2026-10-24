import type { CellValue, FieldConfig, SheetData } from '../types';

/** Coerce a raw user-entered value to the storage type of the field. */
export function normalizeCellValue(field: FieldConfig | undefined, rawValue: CellValue): CellValue {
  if (field?.type === 'number' && typeof rawValue === 'string' && rawValue.trim() !== '') {
    const parsedNum = Number(rawValue.replace(/[$,%\s]/g, ''));
    return Number.isNaN(parsedNum) ? rawValue : parsedNum;
  }
  if (field?.type === 'boolean') {
    if (typeof rawValue === 'string') return rawValue.toLowerCase() === 'true';
  }
  return rawValue;
}

/** Initial cell values for a new row, based on each field's default value. */
export function buildDefaultRowCells(fields: FieldConfig[]): Record<string, CellValue> {
  const cells: Record<string, CellValue> = {};
  fields.forEach((f) => {
    if (f.type === 'boolean') {
      cells[f.id] = f.defaultValue === 'true';
    } else if (f.type === 'number' && f.defaultValue !== '') {
      cells[f.id] = Number(f.defaultValue) || 0;
    } else {
      cells[f.id] = f.defaultValue || '';
    }
  });
  return cells;
}

/** Heuristic: is this column an auto-numbered "No./STT/ID" column? */
export function isAutoIdField(sheet: SheetData, field: FieldConfig, colIdx: number): boolean {
  if (/^(stt|no\.?|#|id)$/i.test(field.label.trim())) return true;
  if (colIdx !== 0 || sheet.rows.length < 2) return false;
  const nonEmpty = sheet.rows
    .map((r) => r.cells[field.id])
    .filter((v) => v !== null && v !== undefined && String(v).trim() !== '');
  return nonEmpty.length >= 2 && nonEmpty.every((v, idx) => Number(String(v).trim()) === idx + 1);
}
