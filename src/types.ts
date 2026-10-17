export type FieldDataType =
  | 'text'
  | 'number'
  | 'date'
  | 'select'
  | 'boolean'
  | 'formula'
  | 'lookup';

export type NumberFormatType = 'raw' | 'integer' | 'decimal' | 'currency' | 'percent';

export interface FieldValidationRule {
  required?: boolean;
  unique?: boolean;
  min?: number;
  max?: number;
  regex?: string;
}

export interface LookupReference {
  targetSheetId: string;
  targetFieldId: string;
  displayFieldId?: string;
}

export interface FieldConfig {
  id: string;
  key: string;             // Machine / export key (e.g. "unit_price" or "col_a")
  label: string;           // Display header name (e.g. "Unit Price" or "A")
  sourceHeader: string;    // Original header from imported .xlsx
  columnLetter?: string;   // Original Excel column letter (A, B, C...)
  type: FieldDataType;
  visible: boolean;
  width: number;
  defaultValue: string;
  description?: string;
  numberFormat?: NumberFormatType;
  currencySymbol?: string;
  selectOptions?: string[];
  formulaExpression?: string; // e.g. "{qty} * {unit_price}" or Excel-style
  lookup?: LookupReference;
  validation: FieldValidationRule;
}

export type CellValue = string | number | boolean | null;

export interface RowRecord {
  id: string;
  excelRowNumber?: number; // Original 1-based Excel row number
  cells: Record<string, CellValue>;
}

export type SheetHiddenState = 0 | 1 | 2; // 0 = Visible, 1 = Hidden, 2 = VeryHidden

export interface SheetData {
  id: string;
  name: string;
  description?: string;
  hiddenState?: SheetHiddenState; // From Excel workbook.Workbook.Sheets[i].Hidden
  headerRowIndex: number;         // -1 means "No Header Row (Use A, B, C...)"; >= 0 is 0-based row in rawMatrix
  rawMatrix?: CellValue[][];      // Preserved 2D array of the imported sheet
  fields: FieldConfig[];
  rows: RowRecord[];
}

export interface WorkbookProject {
  id?: string;
  fileName: string;
  lastModified: string;
  hasOriginalBinary?: boolean;
  sheets: SheetData[];
}

export interface SavedWorkbookSummary {
  id: string;
  fileName: string;
  storageKey: string | null;
  lastModified: string;
  createdAt: string;
  sheetCount: number;
}

export interface CellValidationError {
  rowId: string;
  fieldId: string;
  message: string;
}
