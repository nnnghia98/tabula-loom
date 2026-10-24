import type { WorkbookProject, SheetData, FieldConfig, FieldDataType, NumberFormatType } from '../types';
import { IconTrash } from '../Icons';

interface FieldInspectorProps {
  activeSheet: SheetData;
  handleDeleteField: (fieldId: string) => void;
  handleMoveField: (fieldId: string, direction: -1 | 1) => void;
  project: WorkbookProject;
  selectedField: FieldConfig;
  setIsInspectorOpen: React.Dispatch<React.SetStateAction<boolean>>;
  updateFieldConfig: (fieldId: string, patch: Partial<FieldConfig>) => void;
}

export function FieldInspector({ activeSheet, handleDeleteField, handleMoveField, project, selectedField, setIsInspectorOpen, updateFieldConfig }: FieldInspectorProps) {
  return (
    <aside className="field-inspector" aria-label="Column Settings">
      <div className="inspector-header">
        <div>
          <div className="section-heading">
            Column {selectedField.columnLetter || '—'} Settings
          </div>
          <h2 className="inspector-title">
            {selectedField.label === selectedField.columnLetter
              ? `Column ${selectedField.columnLetter}`
              : selectedField.label}
          </h2>
        </div>
        <button
          type="button"
          className="btn btn-sm"
          onClick={() => setIsInspectorOpen(false)}
        >
          Close
        </button>
      </div>

      {/* Essential Settings in Plain Language */}
      <div className="inspector-section">
        <div className="form-field">
          <label className="form-label">Column Name</label>
          <input
            type="text"
            className="input-control"
            value={selectedField.label}
            onChange={(e) =>
              updateFieldConfig(selectedField.id, {
                label: e.target.value,
              })
            }
          />
        </div>

        <div className="form-field">
          <label className="form-label">Column Data Type</label>
          <select
            className="select-control"
            value={selectedField.type}
            onChange={(e) => {
              const nextType = e.target.value as FieldDataType;
              const patch: Partial<FieldConfig> = { type: nextType };
              if (nextType === 'lookup' && !selectedField.lookup) {
                const otherSheet =
                  project.sheets.find((s) => s.id !== activeSheet.id) ||
                  project.sheets[0];
                patch.lookup = {
                  targetSheetId: otherSheet.id,
                  targetFieldId: otherSheet.fields[0]?.id || '',
                  displayFieldId: otherSheet.fields[1]?.id || otherSheet.fields[0]?.id,
                };
              }
              updateFieldConfig(selectedField.id, patch);
            }}
          >
            <option value="text">Text</option>
            <option value="number">Number / Currency</option>
            <option value="date">Date</option>
            <option value="select">Dropdown List (Options)</option>
            <option value="boolean">Checkbox (Yes / No)</option>
            <option value="formula">Formula Calculation</option>
            <option value="lookup">Link to Another Sheet</option>
          </select>
        </div>

        {(selectedField.type === 'number' || selectedField.type === 'formula') && (
          <div className="form-row-2">
            <div className="form-field">
              <label className="form-label">Number Style</label>
              <select
                className="select-control"
                value={selectedField.numberFormat || 'raw'}
                onChange={(e) =>
                  updateFieldConfig(selectedField.id, {
                    numberFormat: e.target.value as NumberFormatType,
                  })
                }
              >
                <option value="raw">Plain Number</option>
                <option value="integer">Whole Number (1,000)</option>
                <option value="decimal">Two Decimals (1,000.00)</option>
                <option value="currency">Currency ($1,000.00)</option>
                <option value="percent">Percent (10.0%)</option>
              </select>
            </div>
            <div className="form-field">
              <label className="form-label">Symbol</label>
              <input
                type="text"
                className="input-control mono"
                value={selectedField.currencySymbol || '$'}
                onChange={(e) =>
                  updateFieldConfig(selectedField.id, {
                    currencySymbol: e.target.value,
                  })
                }
              />
            </div>
          </div>
        )}

        {selectedField.type === 'select' && (
          <div className="form-field">
            <label className="form-label">Dropdown Choices (comma-separated)</label>
            <textarea
              className="textarea-control"
              value={(selectedField.selectOptions || []).join(', ')}
              onChange={(e) =>
                updateFieldConfig(selectedField.id, {
                  selectOptions: e.target.value
                    .split(',')
                    .map((s) => s.trim())
                    .filter(Boolean),
                })
              }
            />
          </div>
        )}

        {selectedField.type === 'formula' && (
          <div className="form-field">
            <label className="form-label">Formula</label>
            <input
              type="text"
              className="input-control mono"
              placeholder="e.g. {col_a} * {col_b}"
              value={selectedField.formulaExpression || ''}
              onChange={(e) =>
                updateFieldConfig(selectedField.id, {
                  formulaExpression: e.target.value,
                })
              }
            />
          </div>
        )}

        {selectedField.type === 'lookup' && (
          <>
            <div className="form-field">
              <label className="form-label">Link to Sheet</label>
              <select
                className="select-control"
                value={selectedField.lookup?.targetSheetId || ''}
                onChange={(e) => {
                  const targetSheet = project.sheets.find(
                    (s) => s.id === e.target.value
                  );
                  if (!targetSheet) return;
                  updateFieldConfig(selectedField.id, {
                    lookup: {
                      targetSheetId: targetSheet.id,
                      targetFieldId: targetSheet.fields[0]?.id || '',
                      displayFieldId:
                        targetSheet.fields[1]?.id || targetSheet.fields[0]?.id,
                    },
                  });
                }}
              >
                {project.sheets.map((s) => (
                  <option key={s.id} value={s.id}>
                    {s.name}
                  </option>
                ))}
              </select>
            </div>

            {(() => {
              const targetSheet = project.sheets.find(
                (s) => s.id === selectedField.lookup?.targetSheetId
              );
              if (!targetSheet) return null;
              return (
                <div className="form-row-2">
                  <div className="form-field">
                    <label className="form-label">Match Column</label>
                    <select
                      className="select-control"
                      value={selectedField.lookup?.targetFieldId || ''}
                      onChange={(e) =>
                        updateFieldConfig(selectedField.id, {
                          lookup: {
                            ...selectedField.lookup!,
                            targetFieldId: e.target.value,
                          },
                        })
                      }
                    >
                      {targetSheet.fields.map((tf) => (
                        <option key={tf.id} value={tf.id}>
                          {tf.label}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div className="form-field">
                    <label className="form-label">Display Column</label>
                    <select
                      className="select-control"
                      value={selectedField.lookup?.displayFieldId || ''}
                      onChange={(e) =>
                        updateFieldConfig(selectedField.id, {
                          lookup: {
                            ...selectedField.lookup!,
                            displayFieldId: e.target.value,
                          },
                        })
                      }
                    >
                      {targetSheet.fields.map((tf) => (
                        <option key={tf.id} value={tf.id}>
                          {tf.label}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
              );
            })()}
          </>
        )}
      </div>

      {/* Simple Validation Checkboxes */}
      <div className="inspector-section">
        <div className="section-heading">Data Rules</div>

        <div className="form-row-2">
          <label className="checkbox-row">
            <span>Required</span>
            <input
              type="checkbox"
              checked={Boolean(selectedField.validation.required)}
              onChange={(e) =>
                updateFieldConfig(selectedField.id, {
                  validation: {
                    ...selectedField.validation,
                    required: e.target.checked,
                  },
                })
              }
            />
          </label>
          <label className="checkbox-row">
            <span>No Duplicates</span>
            <input
              type="checkbox"
              checked={Boolean(selectedField.validation.unique)}
              onChange={(e) =>
                updateFieldConfig(selectedField.id, {
                  validation: {
                    ...selectedField.validation,
                    unique: e.target.checked,
                  },
                })
              }
            />
          </label>
        </div>

        {selectedField.type === 'number' && (
          <div className="form-row-2">
            <div className="form-field">
              <label className="form-label">Minimum</label>
              <input
                type="number"
                className="input-control mono"
                placeholder="No min"
                value={selectedField.validation.min ?? ''}
                onChange={(e) =>
                  updateFieldConfig(selectedField.id, {
                    validation: {
                      ...selectedField.validation,
                      min:
                        e.target.value === '' ? undefined : Number(e.target.value),
                    },
                  })
                }
              />
            </div>
            <div className="form-field">
              <label className="form-label">Maximum</label>
              <input
                type="number"
                className="input-control mono"
                placeholder="No max"
                value={selectedField.validation.max ?? ''}
                onChange={(e) =>
                  updateFieldConfig(selectedField.id, {
                    validation: {
                      ...selectedField.validation,
                      max:
                        e.target.value === '' ? undefined : Number(e.target.value),
                    },
                  })
                }
              />
            </div>
          </div>
        )}

        <div className="form-field">
          <label className="form-label">Default Value for New Rows</label>
          <input
            type="text"
            className="input-control"
            placeholder="Blank"
            value={selectedField.defaultValue}
            onChange={(e) =>
              updateFieldConfig(selectedField.id, { defaultValue: e.target.value })
            }
          />
        </div>
      </div>

      {/* Column Position & Visibility */}
      <div className="inspector-section" style={{ borderBottom: 'none' }}>
        <div className="section-heading">Column Layout</div>
        <div className="form-row-2">
          <button
            type="button"
            className="btn"
            onClick={() => handleMoveField(selectedField.id, -1)}
          >
            ← Move Left
          </button>
          <button
            type="button"
            className="btn"
            onClick={() => handleMoveField(selectedField.id, 1)}
          >
            Move Right →
          </button>
        </div>

        <div className="form-row-2">
          <label className="checkbox-row">
            <span>Show Column</span>
            <input
              type="checkbox"
              checked={selectedField.visible}
              onChange={(e) =>
                updateFieldConfig(selectedField.id, { visible: e.target.checked })
              }
            />
          </label>
          <div className="form-field">
            <input
              type="number"
              className="input-control mono"
              min={100}
              max={450}
              value={selectedField.width}
              onChange={(e) =>
                updateFieldConfig(selectedField.id, {
                  width: Number(e.target.value) || 160,
                })
              }
              title="Column width in pixels"
            />
          </div>
        </div>

        <button
          type="button"
          className="btn btn-danger"
          style={{ marginTop: '8px' }}
          disabled={activeSheet.fields.length <= 1}
          onClick={() => handleDeleteField(selectedField.id)}
        >
          <IconTrash size={14} /> Delete Column
        </button>
      </div>
    </aside>
  );
}
