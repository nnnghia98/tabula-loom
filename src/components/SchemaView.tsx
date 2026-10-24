import type { WorkbookProject, SheetData, FieldConfig, FieldDataType, NumberFormatType } from '../types';
import type { WorkspaceMode } from '../types';
import { toMachineKey, getExcelColumnLetter } from '../xlsxUtils';
import { IconPlus, IconTrash, IconEye, IconTable } from '../Icons';

interface SchemaViewProps {
  activeSheet: SheetData;
  handleAddField: () => void;
  handleDeleteField: (fieldId: string) => void;
  handleHeaderRowChange: (newHeaderRowIdx: number) => void;
  project: WorkbookProject;
  selectedFieldId: string;
  setSelectedFieldId: React.Dispatch<React.SetStateAction<string>>;
  setShowHiddenColumns: React.Dispatch<React.SetStateAction<boolean>>;
  setWorkspaceMode: React.Dispatch<React.SetStateAction<WorkspaceMode>>;
  showHiddenColumns: boolean;
  updateFieldConfig: (fieldId: string, patch: Partial<FieldConfig>) => void;
}

export function SchemaView({ activeSheet, handleAddField, handleDeleteField, handleHeaderRowChange, project, selectedFieldId, setSelectedFieldId, setShowHiddenColumns, setWorkspaceMode, showHiddenColumns, updateFieldConfig }: SchemaViewProps) {
  return (
    <main className="schema-view-container">
      <div className="schema-view-inner">
        <div className="schema-hero-card">
          <div>
            <div className="section-heading">Sheet Column Settings</div>
            <h1 className="schema-hero-title">{activeSheet.name}</h1>
            <p style={{ color: 'var(--text-muted)', fontSize: '15px' }}>
              Configure column names, data types, dropdown lists, and rules for this sheet.
            </p>
          </div>
          <div style={{ display: 'flex', gap: '10px', alignItems: 'center', flexWrap: 'wrap' }}>
            <div
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: '6px',
                height: '38px',
                padding: '0 12px',
                backgroundColor: 'var(--bg-subtle)',
                border: '1px solid var(--border-hairline)',
                borderRadius: '6px',
                fontSize: '13.5px',
                color: 'var(--text-body)',
              }}
              title="Choose which Excel row contains the column headers for this sheet"
            >
              <span>Header Row:</span>
              <select
                value={activeSheet.headerRowIndex}
                onChange={(e) => handleHeaderRowChange(Number(e.target.value))}
                style={{
                  border: 'none',
                  background: 'transparent',
                  fontSize: '13.5px',
                  fontWeight: 600,
                  color: 'var(--text-primary)',
                  cursor: 'pointer',
                  paddingRight: '22px',
                }}
              >
                <option value={-1}>No Header (Use A, B, C)</option>
                {(activeSheet.rawMatrix || []).slice(0, 10).map((r, idx) => {
                  const preview = r
                    .filter((c) => String(c ?? '').trim() !== '')
                    .slice(0, 2)
                    .join(', ');
                  return (
                    <option key={idx} value={idx}>
                      Row {idx + 1} {preview ? `(${preview.slice(0, 18)})` : ''}
                    </option>
                  );
                })}
                {!activeSheet.rawMatrix && <option value={0}>Row 1</option>}
              </select>
            </div>
            <button
              type="button"
              className="btn"
              onClick={() => setShowHiddenColumns((prev) => !prev)}
            >
              <IconEye visible={showHiddenColumns} size={15} />
              {showHiddenColumns ? 'Showing Hidden Columns' : 'Show Hidden Columns'}
            </button>
            <button type="button" className="btn" onClick={handleAddField}>
              <IconPlus size={15} /> Add Column
            </button>
            <button
              type="button"
              className="btn btn-primary"
              onClick={() => setWorkspaceMode('grid')}
            >
              <IconTable size={15} /> Back to Spreadsheet
            </button>
          </div>
        </div>

        <div className="schema-table-card">
          <table className="schema-config-table">
            <thead>
              <tr>
                <th style={{ width: '64px' }}>Show</th>
                <th style={{ width: '64px' }}>Col</th>
                <th style={{ width: '240px' }}>Column Name</th>
                <th style={{ width: '200px' }}>Data Type</th>
                <th>Options / Formula / Linked Sheet</th>
                <th style={{ width: '150px' }}>Default Value</th>
                <th style={{ width: '160px' }}>Rules</th>
                <th style={{ width: '64px' }} />
              </tr>
            </thead>
            <tbody>
              {activeSheet.fields.map((field, idx) => (
                <tr
                  key={field.id}
                  className={field.id === selectedFieldId ? 'active-schema-row' : ''}
                  onClick={() => setSelectedFieldId(field.id)}
                >
                  <td>
                    <button
                      type="button"
                      className="btn btn-ghost btn-sm"
                      onClick={() =>
                        updateFieldConfig(field.id, { visible: !field.visible })
                      }
                      title={field.visible ? 'Visible in sheet' : 'Hidden in sheet'}
                    >
                      <IconEye visible={field.visible} size={15} />
                    </button>
                  </td>
                  <td
                    style={{
                      fontFamily: 'var(--font-mono)',
                      fontSize: '14px',
                      fontWeight: 600,
                      color: 'var(--text-muted)',
                    }}
                  >
                    {field.columnLetter || getExcelColumnLetter(idx)}
                  </td>
                  <td>
                    <input
                      type="text"
                      className="input-control"
                      value={field.label}
                      onChange={(e) =>
                        updateFieldConfig(field.id, {
                          label: e.target.value,
                          key: toMachineKey(e.target.value),
                        })
                      }
                    />
                  </td>
                  <td>
                    <select
                      className="select-control"
                      value={field.type}
                      onChange={(e) =>
                        updateFieldConfig(field.id, {
                          type: e.target.value as FieldDataType,
                        })
                      }
                    >
                      <option value="text">Text</option>
                      <option value="number">Number / Currency</option>
                      <option value="date">Date</option>
                      <option value="select">Dropdown List</option>
                      <option value="boolean">Checkbox (Yes/No)</option>
                      <option value="formula">Formula</option>
                      <option value="lookup">Link to Sheet</option>
                    </select>
                  </td>
                  <td>
                    {field.type === 'select' && (
                      <input
                        type="text"
                        className="input-control"
                        placeholder="Option 1, Option 2, Option 3"
                        value={(field.selectOptions || []).join(', ')}
                        onChange={(e) =>
                          updateFieldConfig(field.id, {
                            selectOptions: e.target.value
                              .split(',')
                              .map((s) => s.trim())
                              .filter(Boolean),
                          })
                        }
                      />
                    )}
                    {field.type === 'formula' && (
                      <input
                        type="text"
                        className="input-control mono"
                        placeholder="{col_a} * {col_b}"
                        value={field.formulaExpression || ''}
                        onChange={(e) =>
                          updateFieldConfig(field.id, {
                            formulaExpression: e.target.value,
                          })
                        }
                      />
                    )}
                    {field.type === 'number' && (
                      <select
                        className="select-control"
                        value={field.numberFormat || 'raw'}
                        onChange={(e) =>
                          updateFieldConfig(field.id, {
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
                    )}
                    {field.type === 'lookup' && (
                      <div style={{ display: 'flex', gap: '8px' }}>
                        <select
                          className="select-control"
                          value={field.lookup?.targetSheetId || ''}
                          onChange={(e) => {
                            const ts = project.sheets.find(
                              (s) => s.id === e.target.value
                            );
                            if (!ts) return;
                            updateFieldConfig(field.id, {
                              lookup: {
                                targetSheetId: ts.id,
                                targetFieldId: ts.fields[0]?.id || '',
                                displayFieldId:
                                  ts.fields[1]?.id || ts.fields[0]?.id,
                              },
                            });
                          }}
                        >
                          {project.sheets.map((s) => (
                            <option key={s.id} value={s.id}>
                              Sheet: {s.name}
                            </option>
                          ))}
                        </select>
                        <select
                          className="select-control"
                          value={field.lookup?.targetFieldId || ''}
                          onChange={(e) =>
                            updateFieldConfig(field.id, {
                              lookup: {
                                ...(field.lookup || {
                                  targetSheetId: project.sheets[0].id,
                                }),
                                targetFieldId: e.target.value,
                              },
                            })
                          }
                        >
                          {(
                            project.sheets.find(
                              (s) => s.id === field.lookup?.targetSheetId
                            )?.fields || []
                          ).map((tf) => (
                            <option key={tf.id} value={tf.id}>
                              Column: {tf.label}
                            </option>
                          ))}
                        </select>
                      </div>
                    )}
                    {(field.type === 'text' ||
                      field.type === 'date' ||
                      field.type === 'boolean') && (
                      <span className="form-hint">Standard {field.type} column</span>
                    )}
                  </td>
                  <td>
                    <input
                      type="text"
                      className="input-control"
                      placeholder="Blank"
                      value={field.defaultValue}
                      onChange={(e) =>
                        updateFieldConfig(field.id, {
                          defaultValue: e.target.value,
                        })
                      }
                    />
                  </td>
                  <td>
                    <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
                      <label
                        style={{
                          display: 'inline-flex',
                          alignItems: 'center',
                          gap: '5px',
                          fontSize: '13.5px',
                          cursor: 'pointer',
                        }}
                      >
                        <input
                          type="checkbox"
                          checked={Boolean(field.validation.required)}
                          onChange={(e) =>
                            updateFieldConfig(field.id, {
                              validation: {
                                ...field.validation,
                                required: e.target.checked,
                              },
                            })
                          }
                        />
                        Required
                      </label>
                      <label
                        style={{
                          display: 'inline-flex',
                          alignItems: 'center',
                          gap: '5px',
                          fontSize: '13.5px',
                          cursor: 'pointer',
                        }}
                      >
                        <input
                          type="checkbox"
                          checked={Boolean(field.validation.unique)}
                          onChange={(e) =>
                            updateFieldConfig(field.id, {
                              validation: {
                                ...field.validation,
                                unique: e.target.checked,
                              },
                            })
                          }
                        />
                        Unique
                      </label>
                    </div>
                  </td>
                  <td>
                    <button
                      type="button"
                      className="btn btn-ghost btn-sm"
                      disabled={activeSheet.fields.length <= 1}
                      onClick={() => handleDeleteField(field.id)}
                      title="Delete column"
                    >
                      <IconTrash size={15} />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </main>
  );
}
