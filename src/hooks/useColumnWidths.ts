import { useEffect, useState } from 'react';
import type { FieldConfig, SheetData } from '../types';
import { isAutoIdField } from '../utils/cellValues';

const COL_WIDTHS_STORAGE_KEY = 'tabula-loom:column-widths';

/** Persisted per-column width overrides (localStorage), keyed by `${sheetId}:${fieldId}`. */
export function useColumnWidths(activeSheet: SheetData) {
  const [colWidths, setColWidths] = useState<Record<string, number>>(() => {
    try {
      return JSON.parse(localStorage.getItem(COL_WIDTHS_STORAGE_KEY) || '{}');
    } catch {
      return {};
    }
  });

  useEffect(() => {
    try {
      localStorage.setItem(COL_WIDTHS_STORAGE_KEY, JSON.stringify(colWidths));
    } catch {
      /* storage unavailable */
    }
  }, [colWidths]);

  const getColWidth = (field: FieldConfig, colIdx = 0): number => {
    const manualWidth = colWidths[`${activeSheet.id}:${field.id}`];
    if (manualWidth !== undefined) return manualWidth;
    if (isAutoIdField(activeSheet, field, colIdx)) return 64;
    return field.width;
  };

  const startColumnResize = (e: React.MouseEvent, field: FieldConfig, colIdx = 0) => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startWidth = getColWidth(field, colIdx);
    const key = `${activeSheet.id}:${field.id}`;
    const onMove = (ev: MouseEvent) => {
      const next = Math.max(48, Math.round(startWidth + ev.clientX - startX));
      setColWidths((prev) => ({ ...prev, [key]: next }));
    };
    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      document.body.style.userSelect = '';
      document.body.style.cursor = '';
    };
    document.body.style.userSelect = 'none';
    document.body.style.cursor = 'col-resize';
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  };

  return { getColWidth, startColumnResize };
}
