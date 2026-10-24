import type { SheetData, FieldConfig, CellValue, RowRecord } from '../types';
import type { SheetErrors } from '../types';
import type { LookupOption } from '../types';
import { formatCellValue, getExcelColumnLetter } from '../xlsxUtils';
import { IconTrash, IconLink, IconWarning, FieldTypeGlyph } from '../Icons';

interface SheetGridProps {
  activeSheet: SheetData;
  activeSheetErrors: SheetErrors;
  editingCell: { rowId: string; fieldId: string } | null;
  filteredRows: RowRecord[];
  getColWidth: (field: FieldConfig, colIdx?: number) => number;
  getLookupOptions: (field: FieldConfig) => LookupOption[];
  handleDeleteRow: (rowId: string) => void;
  isHeaderlessSheet: boolean;
  selectedCell: { rowId: string; fieldId: string } | null;
  selectedFieldId: string;
  setEditingCell: React.Dispatch<React.SetStateAction<{ rowId: string; fieldId: string } | null>>;
  setIsInspectorOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setSelectedCell: React.Dispatch<React.SetStateAction<{ rowId: string; fieldId: string } | null>>;
  setSelectedFieldId: React.Dispatch<React.SetStateAction<string>>;
  startColumnResize: (e: React.MouseEvent, field: FieldConfig, colIdx?: number) => void;
  updateCellValue: (rowId: string, fieldId: string, rawValue: CellValue) => void;
  visibleFields: FieldConfig[];
}

export function SheetGrid({ activeSheet, activeSheetErrors, editingCell, filteredRows, getColWidth, getLookupOptions, handleDeleteRow, isHeaderlessSheet, selectedCell, selectedFieldId, setEditingCell, setIsInspectorOpen, setSelectedCell, setSelectedFieldId, startColumnResize, updateCellValue, visibleFields }: SheetGridProps) {
  return (
    <div className="grid-scroll-container">
      <table
        className="sheet-table"
        style={{ width: 44 + visibleFields.reduce((s, f, idx) => s + getColWidth(f, idx), 0) }}
      >
        <colgroup>
          <col style={{ width: 44 }} />
          {visibleFields.map((f, idx) => (
            <col key={f.id} style={{ width: getColWidth(f, idx) }} />
          ))}
        </colgroup>
        <thead>
          {/* Classic Excel Row 1: Column Alphabet Letters (A, B, C, D...) */}
          <tr className="excel-letter-row">
            <th className="th-row-index" />
            {visibleFields.map((field, colIdx) => {
              const isSelectedCol = field.id === selectedFieldId;
              const colLetter = field.columnLetter || getExcelColumnLetter(colIdx);
              return (
                <th
                  key={`letter_${field.id}`}
                  className={isSelectedCol ? 'selected-column-letter' : ''}
                  onClick={() => setSelectedFieldId(field.id)}
                  onDoubleClick={() => {
                    setSelectedFieldId(field.id);
                    setIsInspectorOpen(true);
                  }}
                  title="Click to select column, double-click to configure column settings"
                >
                  {colLetter}
                  <div
                    className="col-resize-handle"
                    title="Drag to resize column"
                    onMouseDown={(e) => startColumnResize(e, field, colIdx)}
                    onClick={(e) => e.stopPropagation()}
                    onDoubleClick={(e) => e.stopPropagation()}
                  />
                </th>
              );
            })}
          </tr>

          {/* Row 2: Clean Human-Readable Column Names (only shown when sheet has headers or custom column names) */}
          {!isHeaderlessSheet && (
            <tr className="column-title-row">
              <th className="th-row-index">#</th>
              {visibleFields.map((field, colIdx) => {
                const isSelectedCol = field.id === selectedFieldId;
                return (
                  <th
                    key={field.id}
                    className={isSelectedCol ? 'selected-column' : ''}
                    onClick={() => setSelectedFieldId(field.id)}
                    onDoubleClick={() => {
                      setSelectedFieldId(field.id);
                      setIsInspectorOpen(true);
                    }}
                    title="Double-click to open Column Settings"
                  >
                    <div className="th-header-single-line">
                      <div className="th-label-group">
                        <span className="th-label">{field.label}</span>
                        {field.validation.required && (
                          <span className="req-dot" title="Required column">
                            *
                          </span>
                        )}
                      </div>
                      <FieldTypeGlyph type={field.type} />
                    </div>
                    <div
                      className="col-resize-handle"
                      title="Drag to resize column"
                      onMouseDown={(e) => startColumnResize(e, field, colIdx)}
                      onClick={(e) => e.stopPropagation()}
                      onDoubleClick={(e) => e.stopPropagation()}
                    />
                  </th>
                );
              })}
            </tr>
          )}
        </thead>
        <tbody>
          {filteredRows.map((row, rowIndex) => {
            const displayRowNumber = row.excelRowNumber ?? rowIndex + 1;
            return (
              <tr key={row.id}>
                <td className="td-row-index" title={`Excel Row ${displayRowNumber}`}>
                  <div className="row-index-inner">
                    <span>{displayRowNumber}</span>
                    <button
                      type="button"
                      className="row-delete-btn"
                      title="Delete this row"
                      onClick={() => handleDeleteRow(row.id)}
                    >
                      <IconTrash size={13} />
                    </button>
                  </div>
                </td>
                {visibleFields.map((field) => {
                  const cellKey = `${row.id}:${field.id}`;
                  const cellError = activeSheetErrors[cellKey];
                  const isSelected =
                    selectedCell?.rowId === row.id && selectedCell?.fieldId === field.id;
                  const isEditing =
                    editingCell?.rowId === row.id && editingCell?.fieldId === field.id;
                  const rawVal = row.cells[field.id];
                  const formatted = formatCellValue(
                    rawVal,
                    field,
                    row,
                    activeSheet.fields
                  );

                  const lookupChoices =
                    field.type === 'lookup' ? getLookupOptions(field) : [];
                  const matchedLookup =
                    field.type === 'lookup'
                      ? lookupChoices.find((c) => c.keyVal === String(rawVal ?? ''))
                      : undefined;

                  return (
                    <td
                      key={field.id}
                      className={[
                        'grid-cell',
                        isSelected ? 'cell-selected' : '',
                        cellError ? 'cell-error' : '',
                        field.type === 'formula' ? 'cell-formula' : '',
                      ]
                        .filter(Boolean)
                        .join(' ')}
                      onClick={() => {
                        setSelectedCell({ rowId: row.id, fieldId: field.id });
                        setSelectedFieldId(field.id);
                      }}
                      onDoubleClick={() => {
                        if (field.type !== 'formula') {
                          setEditingCell({ rowId: row.id, fieldId: field.id });
                        }
                      }}
                    >
                      {isEditing ? (
                        field.type === 'select' ? (
                          <select
                            autoFocus
                            className="cell-inline-select"
                            value={String(rawVal ?? '')}
                            onChange={(e) => {
                              updateCellValue(row.id, field.id, e.target.value);
                              setEditingCell(null);
                            }}
                            onBlur={() => setEditingCell(null)}
                          >
                            <option value="">— Select —</option>
                            {(field.selectOptions || []).map((opt) => (
                              <option key={opt} value={opt}>
                                {opt}
                              </option>
                            ))}
                          </select>
                        ) : field.type === 'lookup' ? (
                          <select
                            autoFocus
                            className="cell-inline-select"
                            value={String(rawVal ?? '')}
                            onChange={(e) => {
                              updateCellValue(row.id, field.id, e.target.value);
                              setEditingCell(null);
                            }}
                            onBlur={() => setEditingCell(null)}
                          >
                            <option value="">— Select Linked Value —</option>
                            {lookupChoices.map((item) => (
                              <option key={item.keyVal} value={item.keyVal}>
                                {item.keyVal}
                                {item.displayVal ? ` — ${item.displayVal}` : ''}
                              </option>
                            ))}
                          </select>
                        ) : field.type === 'boolean' ? (
                          <select
                            autoFocus
                            className="cell-inline-select"
                            value={String(Boolean(rawVal))}
                            onChange={(e) => {
                              updateCellValue(
                                row.id,
                                field.id,
                                e.target.value === 'true'
                              );
                              setEditingCell(null);
                            }}
                            onBlur={() => setEditingCell(null)}
                          >
                            <option value="true">Yes / True</option>
                            <option value="false">No / False</option>
                          </select>
                        ) : (
                          <input
                            autoFocus
                            type={field.type === 'date' ? 'date' : 'text'}
                            className={`cell-inline-input ${
                              field.type === 'number' || field.type === 'date'
                                ? 'mono'
                                : ''
                            }`}
                            defaultValue={String(rawVal ?? '')}
                            onBlur={(e) => {
                              updateCellValue(row.id, field.id, e.target.value);
                              setEditingCell(null);
                            }}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') {
                                updateCellValue(
                                  row.id,
                                  field.id,
                                  (e.target as HTMLInputElement).value
                                );
                                setEditingCell(null);
                              } else if (e.key === 'Escape') {
                                setEditingCell(null);
                              }
                            }}
                          />
                        )
                      ) : (
                        <div
                          className={[
                            'cell-display',
                            field.type === 'number' || field.type === 'formula'
                              ? 'align-right'
                              : '',
                            field.type === 'date' ? 'mono' : '',
                          ]
                            .filter(Boolean)
                            .join(' ')}
                          title={
                            cellError
                              ? cellError.message
                              : 'Double-click to edit cell'
                          }
                        >
                          {field.type === 'select' && formatted ? (
                            <span className="cell-tag">{formatted}</span>
                          ) : field.type === 'lookup' && formatted ? (
                            <span
                              style={{
                                display: 'inline-flex',
                                alignItems: 'center',
                                minWidth: 0,
                                overflow: 'hidden',
                              }}
                            >
                              <span className="cell-lookup-chip">
                                <IconLink size={12} />
                                {formatted}
                              </span>
                              {matchedLookup?.displayVal && (
                                <span className="cell-lookup-label">
                                  {matchedLookup.displayVal}
                                </span>
                              )}
                            </span>
                          ) : field.type === 'boolean' ? (
                            <input
                              type="checkbox"
                              checked={
                                rawVal === true ||
                                String(rawVal).toLowerCase() === 'true'
                              }
                              onChange={(e) =>
                                updateCellValue(row.id, field.id, e.target.checked)
                              }
                              style={{
                                cursor: 'pointer',
                                accentColor: '#111111',
                                width: '16px',
                                height: '16px',
                              }}
                            />
                          ) : (
                            <span
                              style={{
                                overflow: 'hidden',
                                textOverflow: 'ellipsis',
                                width: '100%',
                              }}
                            >
                              {formatted}
                            </span>
                          )}

                          {cellError && (
                            <span
                              className="cell-error-indicator"
                              title={cellError.message}
                            >
                              <IconWarning size={15} />
                            </span>
                          )}
                        </div>
                      )}
                    </td>
                  );
                })}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}
