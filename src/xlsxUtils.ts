import * as XLSX from 'xlsx';
import {
  WorkbookProject,
  SheetData,
  SheetHiddenState,
  FieldConfig,
  RowRecord,
  FieldDataType,
  CellValue,
  CellValidationError,
} from './types';

export function toMachineKey(label: string): string {
  return (
    label
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '_')
      .replace(/^_+|_+$/g, '') || 'field'
  );
}

export function getExcelColumnLetter(colIndex: number): string {
  let temp = colIndex;
  let letter = '';
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

function isPureNumericOrCurrency(val: string): boolean {
  const cleaned = val.replace(/[$,%\s]/g, '');
  return cleaned !== '' && !Number.isNaN(Number(cleaned));
}

function isDateLike(val: string): boolean {
  return /^\d{4}-\d{2}-\d{2}$|^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(val);
}

function isCodeOrDataToken(val: string): boolean {
  // Matches IDs like INV-9001, 10024, SKU_01, emails, URLs, formulas, booleans
  const s = val.trim();
  if (s.startsWith('=')) return true;
  if (/^(true|false|yes|no)$/i.test(s)) return true;
  if (/^[A-Z]{2,5}-\d{2,}$/i.test(s)) return true;
  if (s.includes('@') || /^https?:\/\//i.test(s)) return true;
  if (s.length > 42) return true; // Header labels are rarely sentences > 42 chars
  return false;
}

export function inferFieldType(values: CellValue[], header: string): {
  type: FieldDataType;
  selectOptions?: string[];
  numberFormat?: 'raw' | 'integer' | 'decimal' | 'currency' | 'percent';
} {
  const nonEmpty = values.filter((v) => v !== null && v !== undefined && String(v).trim() !== '');
  if (nonEmpty.length === 0) return { type: 'text' };

  const lowerHeader = header.toLowerCase();

  // Check if formula
  if (nonEmpty.some((v) => typeof v === 'string' && v.trim().startsWith('='))) {
    return { type: 'formula' };
  }

  // Check boolean
  const boolSet = new Set(['true', 'false', 'yes', 'no']);
  if (
    nonEmpty.every(
      (v) => typeof v === 'boolean' || boolSet.has(String(v).trim().toLowerCase())
    )
  ) {
    return { type: 'boolean' };
  }

  // Check number
  const allNumbers = nonEmpty.every((v) => {
    if (typeof v === 'number') return !Number.isNaN(v);
    return isPureNumericOrCurrency(String(v));
  });

  if (allNumbers) {
    if (
      lowerHeader.includes('price') ||
      lowerHeader.includes('cost') ||
      lowerHeader.includes('budget') ||
      lowerHeader.includes('amount') ||
      lowerHeader.includes('revenue') ||
      lowerHeader.includes('total') ||
      lowerHeader.includes('fee')
    ) {
      return { type: 'number', numberFormat: 'currency' };
    }
    if (lowerHeader.includes('rate') || lowerHeader.includes('margin') || lowerHeader.includes('pct')) {
      return { type: 'number', numberFormat: 'percent' };
    }
    const hasFloats = nonEmpty.some((v) => {
      const cleaned = String(v).replace(/[$,%\s]/g, '');
      return Number(cleaned) % 1 !== 0;
    });
    return { type: 'number', numberFormat: hasFloats ? 'decimal' : 'integer' };
  }

  // Check date
  if (
    lowerHeader.includes('date') ||
    lowerHeader.includes('created') ||
    lowerHeader.includes('deadline') ||
    lowerHeader.includes('eta') ||
    nonEmpty.every((v) => isDateLike(String(v).trim()))
  ) {
    return { type: 'date' };
  }

  // Check select / enum (if low cardinality relative to row count)
  const uniqueStrings = Array.from(new Set(nonEmpty.map((v) => String(v).trim())));
  if (
    uniqueStrings.length >= 2 &&
    uniqueStrings.length <= 7 &&
    (nonEmpty.length >= 4 ||
      lowerHeader.includes('status') ||
      lowerHeader.includes('stage') ||
      lowerHeader.includes('category') ||
      lowerHeader.includes('tier') ||
      lowerHeader.includes('priority') ||
      lowerHeader.includes('region') ||
      lowerHeader.includes('department') ||
      lowerHeader.includes('type'))
  ) {
    return { type: 'select', selectOptions: uniqueStrings };
  }

  return { type: 'text' };
}

/**
 * Automatically detects the most likely header row index (0-based) OR returns -1 if the sheet
 * has no header row (so columns should be named A, B, C... and Row 1 preserved as data).
 */
export function detectBestHeaderRowIndex(matrix: CellValue[][]): number {
  const nonEmptyRows = matrix.filter((r) =>
    r.some((c) => c !== null && c !== undefined && String(c).trim() !== '')
  );

  // Empty sheet or single-row sheet -> treat as headerless so we don't swallow the only row as header
  if (nonEmptyRows.length <= 1) return -1;

  // Determine typical table width
  const colCounts = nonEmptyRows.map(
    (r) => r.filter((c) => c !== null && c !== undefined && String(c).trim() !== '').length
  );
  const maxPopulatedCols = Math.max(...colCounts, 0);

  // If the entire sheet is just 1 column of values/notes, don't force Row 1 to be a header unless it's clearly a label above numbers
  const maxScan = Math.min(8, matrix.length);
  let bestRowIdx = -1;
  let bestScore = 0;

  for (let r = 0; r < maxScan; r++) {
    const row = matrix[r] || [];
    const nonEmptyCells = row
      .map((c) => (c !== null && c !== undefined ? String(c).trim() : ''))
      .filter(Boolean);

    if (nonEmptyCells.length === 0) continue;

    // A single-cell banner on a multi-column sheet is a title, not column headers
    if (nonEmptyCells.length === 1 && maxPopulatedCols >= 2) {
      continue;
    }

    // Count how many cells in this row are actual data values (numbers, dates, codes, booleans)
    const dataLikeCount = nonEmptyCells.filter(
      (val) => isPureNumericOrCurrency(val) || isDateLike(val) || isCodeOrDataToken(val)
    ).length;

    // If ANY cell in a candidate header row is a pure number/date, or >=30% are data tokens, it is a data row, NOT a header row!
    const hasNumericOrDate = nonEmptyCells.some(
      (val) => isPureNumericOrCurrency(val) || isDateLike(val)
    );
    if (hasNumericOrDate || dataLikeCount / nonEmptyCells.length >= 0.3) {
      continue;
    }

    // Check if subsequent rows have different type patterns (e.g., header row is all text labels, next row has numbers/dates/codes)
    const nextRow = matrix.slice(r + 1).find((nr) =>
      nr.some((c) => c !== null && c !== undefined && String(c).trim() !== '')
    );
    let contrastBonus = 0;
    if (nextRow) {
      const nextNonEmpty = nextRow
        .map((c) => (c !== null && c !== undefined ? String(c).trim() : ''))
        .filter(Boolean);
      const nextHasData = nextNonEmpty.some(
        (val) => isPureNumericOrCurrency(val) || isDateLike(val) || isCodeOrDataToken(val)
      );
      if (nextHasData) {
        contrastBonus = 4;
      }
    }

    // Check if cells look like clean header labels (short text <= 32 chars)
    const cleanLabelCount = nonEmptyCells.filter((val) => val.length <= 32).length;
    const coverageRatio = nonEmptyCells.length / Math.max(1, maxPopulatedCols);

    // Require at least 50% column coverage and clean labels
    if (coverageRatio < 0.5) continue;

    const score =
      cleanLabelCount * 2 +
      contrastBonus +
      coverageRatio * 3 -
      r * 0.3;

    if (score > bestScore && score >= 4.5) {
      bestScore = score;
      bestRowIdx = r;
    }
  }

  return bestRowIdx;
}

/**
 * Builds dynamic FieldConfig[] and RowRecord[] from a 2D matrix given a specific `headerRowIndex`.
 * If `headerRowIndex === -1`, no row is consumed as headers; columns are named A, B, C...
 * and all non-empty rows starting from row 0 are kept as data rows.
 */
export function buildSheetFromMatrix(
  sheetId: string,
  sheetName: string,
  rawMatrix: CellValue[][],
  headerRowIndex: number,
  hiddenState: SheetHiddenState = 0
): SheetData {
  // Determine max column width across all non-empty rows
  let maxCols = 0;
  rawMatrix.forEach((row) => {
    for (let c = row.length - 1; c >= 0; c--) {
      if (row[c] !== null && row[c] !== undefined && String(row[c]).trim() !== '') {
        if (c + 1 > maxCols) maxCols = c + 1;
        break;
      }
    }
  });

  if (maxCols === 0) maxCols = 3;

  const isHeaderless = headerRowIndex < 0 || headerRowIndex >= rawMatrix.length;
  const effectiveHeaderIdx = isHeaderless ? -1 : headerRowIndex;
  const headerRow = isHeaderless ? [] : rawMatrix[effectiveHeaderIdx] || [];

  // Build indexed data rows so we preserve the original 1-based Excel row number
  const rawDataSlice = isHeaderless
    ? rawMatrix.map((row, idx) => ({ row, excelRowNumber: idx + 1 }))
    : rawMatrix
        .slice(effectiveHeaderIdx + 1)
        .map((row, idx) => ({ row, excelRowNumber: effectiveHeaderIdx + idx + 2 }));

  const nonEmptyDataEntries = rawDataSlice.filter(({ row }) =>
    row.some((cell) => cell !== null && cell !== undefined && String(cell).trim() !== '')
  );

  const usedKeys = new Map<string, number>();
  const usedLabels = new Map<string, number>();

  const fields: FieldConfig[] = [];
  const activeColIndices: number[] = [];

  for (let colIdx = 0; colIdx < maxCols; colIdx++) {
    const colLetter = getExcelColumnLetter(colIdx);
    const rawHeaderCell = isHeaderless ? '' : headerRow[colIdx];
    const rawHeaderStr =
      rawHeaderCell !== null && rawHeaderCell !== undefined
        ? String(rawHeaderCell).trim()
        : '';

    const colValues: CellValue[] = nonEmptyDataEntries.map(({ row }) => {
      const v = row[colIdx];
      return v === undefined || v === null ? '' : v;
    });

    const hasAnyData =
      rawHeaderStr !== '' ||
      colValues.some((v) => v !== null && v !== undefined && String(v).trim() !== '');

    if (!hasAnyData && rawMatrix.length > 0) continue;

    activeColIndices.push(colIdx);

    // If headerless (or blank header cell), use the Excel column alphabet A, B, C...
    const baseLabel = isHeaderless ? colLetter : rawHeaderStr || colLetter;
    const labelCount = usedLabels.get(baseLabel.toLowerCase()) || 0;
    usedLabels.set(baseLabel.toLowerCase(), labelCount + 1);
    const finalLabel = labelCount === 0 ? baseLabel : `${baseLabel} (${labelCount + 1})`;

    const baseKey = isHeaderless
      ? `col_${colLetter.toLowerCase()}`
      : toMachineKey(rawHeaderStr || `col_${colLetter.toLowerCase()}`);
    const keyCount = usedKeys.get(baseKey) || 0;
    usedKeys.set(baseKey, keyCount + 1);
    const finalKey = keyCount === 0 ? baseKey : `${baseKey}_${keyCount + 1}`;

    const inferred = inferFieldType(colValues, isHeaderless ? '' : finalLabel);
    const fieldId = `${sheetId}_c${colIdx}_${finalKey}`;

    fields.push({
      id: fieldId,
      key: finalKey,
      label: finalLabel,
      sourceHeader: isHeaderless
        ? `(No Header — Col ${colLetter})`
        : rawHeaderStr || `(Blank — Col ${colLetter})`,
      columnLetter: colLetter,
      type: inferred.type,
      visible: true,
      width: Math.max(130, Math.min(260, finalLabel.length * 10 + 65)),
      defaultValue: '',
      numberFormat: inferred.numberFormat || 'raw',
      currencySymbol: '$',
      selectOptions: inferred.selectOptions || ['Active', 'Pending', 'Archived'],
      formulaExpression: '',
      validation: {
        required: false,
        unique: false,
      },
    });
  }

  // Fallback if completely empty sheet
  if (fields.length === 0) {
    ['A', 'B', 'C'].forEach((letter, i) => {
      activeColIndices.push(i);
      fields.push({
        id: `${sheetId}_c${i}_col_${letter.toLowerCase()}`,
        key: `col_${letter.toLowerCase()}`,
        label: letter,
        sourceHeader: `(No Header — Col ${letter})`,
        columnLetter: letter,
        type: 'text',
        visible: true,
        width: 140,
        defaultValue: '',
        validation: {},
      });
    });
  }

  // Build RowRecord[] mapped to field.id by column index
  const rows: RowRecord[] = nonEmptyDataEntries.map(({ row: rawRow, excelRowNumber }, rIdx) => {
    const cells: Record<string, CellValue> = {};
    fields.forEach((field, fIdx) => {
      const origColIdx = activeColIndices[fIdx] ?? fIdx;
      const rawVal = rawRow[origColIdx];

      if (rawVal === undefined || rawVal === null || String(rawVal).trim() === '') {
        cells[field.id] = '';
      } else if (field.type === 'number') {
        const cleaned = String(rawVal).replace(/[$,%\s]/g, '');
        const num = Number(cleaned);
        cells[field.id] = Number.isNaN(num) ? String(rawVal) : num;
      } else if (field.type === 'boolean') {
        const s = String(rawVal).trim().toLowerCase();
        cells[field.id] = s === 'true' || s === 'yes' || s === '1';
      } else {
        cells[field.id] = String(rawVal);
      }
    });

    return {
      id: `${sheetId}_r${rIdx}_${Math.random().toString(36).slice(2, 6)}`,
      excelRowNumber,
      cells,
    };
  });

  const headerDesc = isHeaderless
    ? 'No header row (Columns A, B, C...)'
    : `Header at Excel Row ${effectiveHeaderIdx + 1}`;

  return {
    id: sheetId,
    name: sheetName,
    description: `${fields.length} columns, ${rows.length} rows • ${headerDesc}`,
    hiddenState,
    headerRowIndex: effectiveHeaderIdx,
    rawMatrix,
    fields,
    rows,
  };
}

/**
 * Cross-sheet lookup auto-detector: links columns across sheets when keys match.
 */
export function autoDetectCrossSheetLookups(sheets: SheetData[]): SheetData[] {
  if (sheets.length <= 1) return sheets;

  return sheets.map((sheet) => {
    const updatedFields = sheet.fields.map((field, fIdx) => {
      if (fIdx === 0 || field.key.startsWith('col_')) return field;

      for (const otherSheet of sheets) {
        if (otherSheet.id === sheet.id || otherSheet.fields.length === 0) continue;
        const pkField = otherSheet.fields[0];
        if (pkField.key.startsWith('col_')) continue;

        const displayField =
          otherSheet.fields.find((f) => f.key.includes('name') || f.key.includes('title')) ||
          otherSheet.fields[1] ||
          pkField;

        const keyMatches =
          field.key === pkField.key ||
          (field.key.endsWith('_id') && pkField.key === field.key) ||
          (field.key.endsWith('_code') && pkField.key === field.key);

        if (keyMatches) {
          return {
            ...field,
            type: 'lookup' as FieldDataType,
            lookup: {
              targetSheetId: otherSheet.id,
              targetFieldId: pkField.id,
              displayFieldId: displayField?.id,
            },
          };
        }
      }
      return field;
    });

    return {
      ...sheet,
      fields: updatedFields,
    };
  });
}

export function parseXlsxFile(buffer: ArrayBuffer, fileName: string): WorkbookProject {
  const workbook = XLSX.read(buffer, { type: 'array', cellFormula: true, cellDates: true });
  const wbSheetMeta = workbook.Workbook?.Sheets || [];

  const initialSheets: SheetData[] = workbook.SheetNames.map((sheetName, sheetIdx) => {
    const ws = workbook.Sheets[sheetName];
    // Read hidden status from workbook metadata (0 = visible, 1 = hidden, 2 = very hidden)
    const rawHidden = wbSheetMeta[sheetIdx]?.Hidden;
    const hiddenState: SheetHiddenState =
      rawHidden === 1 ? 1 : rawHidden === 2 ? 2 : 0;

    const rawAoA = XLSX.utils.sheet_to_json<CellValue[]>(ws, {
      header: 1,
      defval: '',
      raw: false,
      blankrows: true,
    });

    const cleanMatrix: CellValue[][] = rawAoA.map((row) =>
      Array.isArray(row) ? row.map((cell) => (cell === undefined || cell === null ? '' : cell)) : []
    );

    const bestHeaderRowIdx = detectBestHeaderRowIndex(cleanMatrix);
    const sheetId = `sheet_${sheetIdx}_${toMachineKey(sheetName)}`;

    return buildSheetFromMatrix(sheetId, sheetName, cleanMatrix, bestHeaderRowIdx, hiddenState);
  });

  const sheetsWithLookups = autoDetectCrossSheetLookups(initialSheets);

  return {
    fileName,
    lastModified: new Date().toISOString(),
    sheets: sheetsWithLookups,
  };
}

export function evaluateRowFormula(
  expression: string | undefined,
  row: RowRecord,
  fields: FieldConfig[],
  rawCellVal: CellValue
): string | number {
  const expr =
    (expression && expression.trim()) ||
    (typeof rawCellVal === 'string' && rawCellVal.startsWith('=') ? rawCellVal.slice(1) : '');
  if (!expr) return rawCellVal !== null && rawCellVal !== undefined ? String(rawCellVal) : '';

  try {
    const resolved = expr.replace(/\{([^}]+)\}/g, (_, token: string) => {
      const cleanToken = token.trim().toLowerCase();
      const targetField = fields.find(
        (f) => f.key.toLowerCase() === cleanToken || f.label.toLowerCase() === cleanToken
      );
      if (!targetField) return '0';
      const val = row.cells[targetField.id];
      const num = Number(String(val ?? '0').replace(/[$,%\s]/g, ''));
      return Number.isNaN(num) ? '0' : String(num);
    });

    if (!/^[0-9+\-*/().\s]+$/.test(resolved)) {
      return resolved;
    }
    const result = Function(`"use strict"; return (${resolved})`)();
    if (typeof result === 'number' && Number.isFinite(result)) {
      return Math.round(result * 100) / 100;
    }
    return String(result);
  } catch {
    return '#ERR';
  }
}

export function formatCellValue(
  value: CellValue,
  field: FieldConfig,
  row: RowRecord,
  allFields: FieldConfig[]
): string {
  if (field.type === 'formula') {
    const evaluated = evaluateRowFormula(field.formulaExpression, row, allFields, value);
    if (typeof evaluated === 'number') {
      if (field.numberFormat === 'currency') {
        return `${field.currencySymbol || '$'}${evaluated.toLocaleString('en-US', {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        })}`;
      }
      if (field.numberFormat === 'percent') {
        return `${evaluated.toFixed(1)}%`;
      }
      return evaluated.toLocaleString('en-US');
    }
    return String(evaluated);
  }

  if (value === null || value === undefined || value === '') return '';

  if (field.type === 'number') {
    const num = typeof value === 'number' ? value : Number(String(value).replace(/[$,%\s]/g, ''));
    if (Number.isNaN(num)) return String(value);
    switch (field.numberFormat) {
      case 'currency':
        return `${field.currencySymbol || '$'}${num.toLocaleString('en-US', {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        })}`;
      case 'integer':
        return Math.round(num).toLocaleString('en-US');
      case 'decimal':
        return num.toLocaleString('en-US', {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        });
      case 'percent':
        return `${num.toFixed(1)}%`;
      default:
        return String(num);
    }
  }

  if (field.type === 'boolean') {
    const isTrue =
      value === true ||
      String(value).toLowerCase() === 'true' ||
      String(value).toLowerCase() === 'yes' ||
      String(value) === '1';
    return isTrue ? 'True' : 'False';
  }

  return String(value);
}

export function validateSheet(
  sheet: SheetData,
  allSheets: SheetData[]
): Record<string, CellValidationError> {
  const errors: Record<string, CellValidationError> = {};

  sheet.fields.forEach((field) => {
    const seenValues = new Map<string, string[]>();

    let validLookupValues: Set<string> | null = null;
    if (field.type === 'lookup' && field.lookup) {
      const targetSheet = allSheets.find((s) => s.id === field.lookup?.targetSheetId);
      if (targetSheet) {
        validLookupValues = new Set(
          targetSheet.rows
            .map((r) => String(r.cells[field.lookup!.targetFieldId] ?? '').trim())
            .filter(Boolean)
        );
      }
    }

    sheet.rows.forEach((row) => {
      const rawVal = row.cells[field.id];
      const strVal = rawVal === null || rawVal === undefined ? '' : String(rawVal).trim();
      const cellKey = `${row.id}:${field.id}`;

      if (field.type === 'formula') return;

      if (field.validation.required && strVal === '') {
        errors[cellKey] = {
          rowId: row.id,
          fieldId: field.id,
          message: `"${field.label}" is required`,
        };
        return;
      }

      if (strVal === '') return;

      if (field.type === 'number') {
        const num = Number(strVal.replace(/[$,%\s]/g, ''));
        if (Number.isNaN(num)) {
          errors[cellKey] = {
            rowId: row.id,
            fieldId: field.id,
            message: `Expected numeric value`,
          };
          return;
        }
        if (field.validation.min !== undefined && num < field.validation.min) {
          errors[cellKey] = {
            rowId: row.id,
            fieldId: field.id,
            message: `Minimum allowed is ${field.validation.min}`,
          };
          return;
        }
        if (field.validation.max !== undefined && num > field.validation.max) {
          errors[cellKey] = {
            rowId: row.id,
            fieldId: field.id,
            message: `Maximum allowed is ${field.validation.max}`,
          };
          return;
        }
      }

      if (field.type === 'select' && field.selectOptions && field.selectOptions.length > 0) {
        if (!field.selectOptions.includes(strVal)) {
          errors[cellKey] = {
            rowId: row.id,
            fieldId: field.id,
            message: `Value "${strVal}" is not in allowed options (${field.selectOptions.join(', ')})`,
          };
          return;
        }
      }

      if (field.type === 'lookup' && validLookupValues) {
        if (!validLookupValues.has(strVal)) {
          errors[cellKey] = {
            rowId: row.id,
            fieldId: field.id,
            message: `Unresolved foreign key "${strVal}" in target sheet`,
          };
          return;
        }
      }

      if (field.validation.regex) {
        try {
          const re = new RegExp(field.validation.regex);
          if (!re.test(strVal)) {
            errors[cellKey] = {
              rowId: row.id,
              fieldId: field.id,
              message: `Does not match pattern ${field.validation.regex}`,
            };
            return;
          }
        } catch {
          // ignore invalid regex while typing
        }
      }

      if (field.validation.unique) {
        const list = seenValues.get(strVal) || [];
        list.push(row.id);
        seenValues.set(strVal, list);
      }
    });

    if (field.validation.unique) {
      seenValues.forEach((rowIds, val) => {
        if (rowIds.length > 1) {
          rowIds.forEach((rId) => {
            errors[`${rId}:${field.id}`] = {
              rowId: rId,
              fieldId: field.id,
              message: `Duplicate value "${val}" violates unique constraint`,
            };
          });
        }
      });
    }
  });

  return errors;
}

export function exportWorkbookToXlsx(
  project: WorkbookProject,
  options: {
    useMachineKeys?: boolean;
    includeHiddenFields?: boolean;
    includeHiddenSheets?: boolean;
    includeSchemaSheet?: boolean;
  } = {}
): void {
  const wb = XLSX.utils.book_new();
  const sheetsMeta: { name: string; Hidden?: SheetHiddenState }[] = [];

  project.sheets.forEach((sheet) => {
    if (sheet.hiddenState && !options.includeHiddenSheets) {
      return;
    }

    const activeFields = options.includeHiddenFields
      ? sheet.fields
      : sheet.fields.filter((f) => f.visible);

    const dataRows = sheet.rows.map((row) =>
      activeFields.map((field) => {
        if (field.type === 'formula') {
          return evaluateRowFormula(
            field.formulaExpression,
            row,
            sheet.fields,
            row.cells[field.id]
          );
        }
        return row.cells[field.id] ?? field.defaultValue ?? '';
      })
    );

    // If the sheet has no headers (headerRowIndex === -1) and user didn't force machine keys, don't inject an artificial header row!
    const isHeaderless = sheet.headerRowIndex === -1 && !options.useMachineKeys;
    const headerRow = activeFields.map((f) => (options.useMachineKeys ? f.key : f.label));
    const aoa = isHeaderless ? dataRows : [headerRow, ...dataRows];

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = activeFields.map((f) => ({ wch: Math.max(12, Math.round(f.width / 8)) }));

    const safeName = sheet.name.slice(0, 31);
    XLSX.utils.book_append_sheet(wb, ws, safeName);
    sheetsMeta.push({ name: safeName, Hidden: sheet.hiddenState || 0 });
  });

  if (options.includeSchemaSheet !== false) {
    const schemaRows: (string | boolean)[][] = [
      [
        'Sheet Name',
        'Sheet Visibility',
        'Header Mode',
        'Excel Col',
        'Field Label',
        'Machine Key',
        'Source Header',
        'Data Type',
        'Visible',
        'Required',
        'Unique',
        'Default Value',
        'Format / Options / Reference',
      ],
    ];

    project.sheets.forEach((sheet) => {
      if (sheet.hiddenState && !options.includeHiddenSheets) return;
      const visText =
        sheet.hiddenState === 2
          ? 'VeryHidden'
          : sheet.hiddenState === 1
          ? 'Hidden'
          : 'Visible';
      const headerMode =
        sheet.headerRowIndex === -1
          ? 'No Header (A, B, C...)'
          : `Row ${sheet.headerRowIndex + 1}`;

      sheet.fields.forEach((f, idx) => {
        let extra = '';
        if (f.type === 'select') extra = `Options: ${(f.selectOptions || []).join(', ')}`;
        else if (f.type === 'number') extra = `Format: ${f.numberFormat || 'raw'}`;
        else if (f.type === 'formula') extra = `Formula: ${f.formulaExpression || ''}`;
        else if (f.type === 'lookup' && f.lookup) {
          const targetS = project.sheets.find((s) => s.id === f.lookup?.targetSheetId);
          const targetF = targetS?.fields.find((tf) => tf.id === f.lookup?.targetFieldId);
          extra = `Lookup -> ${targetS?.name || '?'}.${targetF?.key || '?'}`;
        }

        schemaRows.push([
          sheet.name,
          visText,
          headerMode,
          f.columnLetter || getExcelColumnLetter(idx),
          f.label,
          f.key,
          f.sourceHeader,
          f.type,
          f.visible,
          Boolean(f.validation.required),
          Boolean(f.validation.unique),
          f.defaultValue || '',
          extra,
        ]);
      });
    });

    const schemaWs = XLSX.utils.aoa_to_sheet(schemaRows);
    schemaWs['!cols'] = [
      { wch: 18 },
      { wch: 14 },
      { wch: 18 },
      { wch: 10 },
      { wch: 20 },
      { wch: 18 },
      { wch: 22 },
      { wch: 12 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 },
      { wch: 15 },
      { wch: 36 },
    ];
    XLSX.utils.book_append_sheet(wb, schemaWs, '_field_configs');
    sheetsMeta.push({ name: '_field_configs', Hidden: 0 });
  }

  wb.Workbook = { Sheets: sheetsMeta };

  const cleanName = project.fileName.replace(/\.xlsx$/i, '');
  XLSX.writeFile(wb, `${cleanName}_configured.xlsx`);
}

export function exportSchemaAsJson(project: WorkbookProject): void {
  const payload = {
    fileName: project.fileName,
    exportedAt: new Date().toISOString(),
    sheets: project.sheets.map((s) => ({
      id: s.id,
      name: s.name,
      hiddenState: s.hiddenState || 0,
      headerRowIndex: s.headerRowIndex,
      description: s.description,
      fields: s.fields,
    })),
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${project.fileName.replace(/\.xlsx$/i, '')}_schema.json`;
  a.click();
  URL.revokeObjectURL(url);
}
