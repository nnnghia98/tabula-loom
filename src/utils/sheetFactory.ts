import type { FieldConfig, SheetData } from '../types';
import { getExcelColumnLetter } from '../xlsxUtils';

/** A new, blank headerless column (A, B, C ...). */
export function createBlankField(index: number, isHeaderless: boolean): FieldConfig {
  const colLetter = getExcelColumnLetter(index);
  const label = isHeaderless ? colLetter : `Column ${colLetter}`;
  return {
    id: `fld_${Date.now().toString(36)}`,
    key: `col_${colLetter.toLowerCase()}`,
    label,
    sourceHeader: isHeaderless ? `(No Header — Col ${colLetter})` : label,
    columnLetter: colLetter,
    type: 'text',
    visible: true,
    width: 170,
    defaultValue: '',
    selectOptions: ['Option A', 'Option B'],
    validation: {},
  };
}

/** A new blank sheet: headerless columns A, B, C and one empty row. */
export function createBlankSheet(name: string): SheetData {
  const id = `sheet_${Date.now().toString(36)}`;
  const letters = ['A', 'B', 'C'];
  const fields: FieldConfig[] = letters.map((letter, i) => ({
    id: `${id}_c${i}_col_${letter.toLowerCase()}`,
    key: `col_${letter.toLowerCase()}`,
    label: letter,
    sourceHeader: `(No Header — Col ${letter})`,
    columnLetter: letter,
    type: 'text',
    visible: true,
    width: 170,
    defaultValue: '',
    validation: {},
  }));
  return {
    id,
    name,
    description: '',
    hiddenState: 0,
    headerRowIndex: -1,
    fields,
    rows: [
      {
        id: `row_${Date.now().toString(36)}`,
        excelRowNumber: 1,
        cells: Object.fromEntries(fields.map((f) => [f.id, ''])),
      },
    ],
  };
}
