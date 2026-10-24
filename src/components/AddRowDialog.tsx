import type { SheetData, FieldConfig, CellValue } from '../types';
import type { LookupOption } from '../types';
import { formatCellValue, getExcelColumnLetter } from '../xlsxUtils';
import { IconPlus, IconWarning, IconCheck, FieldTypeGlyph } from '../Icons';

interface AddRowDialogProps {
  activeSheet: SheetData;
  addRowDialogRef: React.RefObject<HTMLDialogElement | null>;
  getLookupOptions: (field: FieldConfig) => LookupOption[];
  handleAddBlankRow: () => void;
  handleSubmitAddRowForm: (e: React.FormEvent, keepOpen?: boolean) => void;
  newRowDraft: Record<string, CellValue>;
  newRowErrors: Record<string, string>;
  setNewRowDraft: React.Dispatch<React.SetStateAction<Record<string, CellValue>>>;
  setNewRowErrors: React.Dispatch<React.SetStateAction<Record<string, string>>>;
}

export function AddRowDialog({ activeSheet, addRowDialogRef, getLookupOptions, handleAddBlankRow, handleSubmitAddRowForm, newRowDraft, newRowErrors, setNewRowDraft, setNewRowErrors }: AddRowDialogProps) {
  return (
    <dialog
      ref={addRowDialogRef}
      closedby="any"
      className="editorial-dialog"
      style={{ maxWidth: '640px' }}
    >
      <form onSubmit={(e) => handleSubmitAddRowForm(e, false)}>
        <div className="dialog-header">
          <div>
            <h3 className="dialog-title">Add Row to "{activeSheet.name}"</h3>
            <div style={{ fontSize: '13px', color: 'var(--text-muted)', marginTop: '2px' }}>
              Row #{(activeSheet.rows[activeSheet.rows.length - 1]?.excelRowNumber || activeSheet.rows.length) + 1} •{' '}
              {activeSheet.fields.length} configured {activeSheet.fields.length === 1 ? 'field' : 'fields'}
            </div>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <button
              type="button"
              className="btn btn-sm"
              onClick={handleAddBlankRow}
              title="Insert a blank row immediately without filling out the form"
            >
              Insert Blank Row
            </button>
            <button
              type="button"
              className="btn btn-ghost btn-sm"
              onClick={() => addRowDialogRef.current?.close()}
            >
              Close
            </button>
          </div>
        </div>

        <div
          className="dialog-body"
          style={{
            maxHeight: '62vh',
            overflowY: 'auto',
            display: 'grid',
            gridTemplateColumns: activeSheet.fields.length > 4 ? '1fr 1fr' : '1fr',
            gap: '14px 16px',
          }}
        >
          {activeSheet.fields.map((field, idx) => {
            const colLetter = field.columnLetter || getExcelColumnLetter(idx);
            const rawVal = newRowDraft[field.id] ?? '';
            const fieldErr = newRowErrors[field.id];
            const lookupChoices = field.type === 'lookup' ? getLookupOptions(field) : [];
            const formulaPreview =
              field.type === 'formula'
                ? formatCellValue(
                    '',
                    field,
                    { id: '__preview__', cells: newRowDraft },
                    activeSheet.fields
                  )
                : '';

            const setDraftField = (val: CellValue) => {
              setNewRowDraft((prev) => ({ ...prev, [field.id]: val }));
              if (newRowErrors[field.id]) {
                setNewRowErrors((prev) => {
                  const next = { ...prev };
                  delete next[field.id];
                  return next;
                });
              }
            };

            return (
              <div key={field.id} className="form-field">
                <label className="form-label">
                  <span
                    style={{
                      display: 'inline-flex',
                      alignItems: 'center',
                      gap: '6px',
                      minWidth: 0,
                    }}
                  >
                    <span
                      style={{
                        fontFamily: 'var(--font-mono)',
                        fontSize: '11.5px',
                        color: 'var(--text-muted)',
                        backgroundColor: 'var(--bg-canvas)',
                        border: '1px solid var(--border-hairline)',
                        borderRadius: '4px',
                        padding: '1px 5px',
                      }}
                    >
                      {colLetter}
                    </span>
                    <span
                      style={{
                        overflow: 'hidden',
                        textOverflow: 'ellipsis',
                        whiteSpace: 'nowrap',
                      }}
                    >
                      {field.label === colLetter ? `Column ${colLetter}` : field.label}
                    </span>
                    {field.validation.required && (
                      <span className="req-dot" title="Required field">
                        *
                      </span>
                    )}
                  </span>
                  <FieldTypeGlyph type={field.type} />
                </label>

                {field.type === 'select' ? (
                  <select
                    className="select-control"
                    value={String(rawVal ?? '')}
                    onChange={(e) => setDraftField(e.target.value)}
                  >
                    <option value="">— Select option —</option>
                    {(field.selectOptions || []).map((opt) => (
                      <option key={opt} value={opt}>
                        {opt}
                      </option>
                    ))}
                  </select>
                ) : field.type === 'lookup' ? (
                  <select
                    className="select-control"
                    value={String(rawVal ?? '')}
                    onChange={(e) => setDraftField(e.target.value)}
                  >
                    <option value="">— Select linked record —</option>
                    {lookupChoices.map((item) => (
                      <option key={item.keyVal} value={item.keyVal}>
                        {item.keyVal}
                        {item.displayVal ? ` — ${item.displayVal}` : ''}
                      </option>
                    ))}
                  </select>
                ) : field.type === 'boolean' ? (
                  <label className="checkbox-row" style={{ height: '38px' }}>
                    <span>
                      {rawVal === true || String(rawVal).toLowerCase() === 'true'
                        ? 'Yes / True'
                        : 'No / False'}
                    </span>
                    <input
                      type="checkbox"
                      checked={rawVal === true || String(rawVal).toLowerCase() === 'true'}
                      onChange={(e) => setDraftField(e.target.checked)}
                      style={{ accentColor: '#111111', width: '16px', height: '16px' }}
                    />
                  </label>
                ) : field.type === 'formula' ? (
                  <input
                    type="text"
                    className="input-control mono"
                    value={formulaPreview || '0'}
                    readOnly
                    disabled
                    title={`Calculated automatically: ${field.formulaExpression || 'Formula'}`}
                    style={{
                      backgroundColor: '#FCFBF7',
                      color: 'var(--text-muted)',
                      cursor: 'not-allowed',
                    }}
                  />
                ) : (
                  <input
                    type={
                      field.type === 'date'
                        ? 'date'
                        : field.type === 'number'
                        ? 'number'
                        : 'text'
                    }
                    step={field.type === 'number' ? 'any' : undefined}
                    className={`input-control ${
                      field.type === 'number' || field.type === 'date' ? 'mono' : ''
                    }`}
                    placeholder={
                      field.type === 'number'
                        ? field.numberFormat === 'currency'
                          ? `${field.currencySymbol || '$'}0.00`
                          : 'Enter number...'
                        : `Enter ${field.label}...`
                    }
                    value={String(rawVal ?? '')}
                    onChange={(e) => setDraftField(e.target.value)}
                    autoFocus={idx === 0}
                  />
                )}

                {fieldErr ? (
                  <div
                    style={{
                      fontSize: '12.5px',
                      color: 'var(--pastel-red-fg)',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '4px',
                    }}
                  >
                    <IconWarning size={13} /> {fieldErr}
                  </div>
                ) : field.description ? (
                  <div className="form-hint">{field.description}</div>
                ) : null}
              </div>
            );
          })}
        </div>

        <div className="dialog-footer">
          <button
            type="button"
            className="btn"
            onClick={() => addRowDialogRef.current?.close()}
          >
            Cancel
          </button>
          <button
            type="button"
            className="btn"
            onClick={(e) => handleSubmitAddRowForm(e, true)}
            title="Insert this row and reset the form to add another row"
          >
            <IconPlus size={14} /> Add &amp; Add Another
          </button>
          <button type="submit" className="btn btn-primary">
            <IconCheck size={14} /> Add Row
          </button>
        </div>
      </form>
    </dialog>
  );
}
