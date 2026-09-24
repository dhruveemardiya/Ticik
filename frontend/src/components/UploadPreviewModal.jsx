import React, { useState } from 'react';
import {
  FileSpreadsheet,
  Layers,
  Users,
  CheckCircle2,
  X,
  AlertCircle,
  Loader2,
  FileCheck,
  ShieldCheck,
  Database
} from 'lucide-react';

export default function UploadPreviewModal({
  isOpen,
  previewData,
  onClose,
  onSaveConfirmed,
  isReplacing = false,
}) {
  const [activeFileIdx, setActiveFileIdx] = useState(0);
  const [activeSheetIdx, setActiveSheetIdx] = useState(0);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);

  if (!isOpen || !previewData) return null;

  const {
    filenames = [],
    filename = 'Excel Workbooks',
    total_files = 1,
    sheet_count = 0,
    customer_count = 0,
    total_records = 0,
    files = [],
    sheets = []
  } = previewData;

  // Determine current active file and active sheet
  const activeFile = files && files.length > 0 ? files[activeFileIdx] || files[0] : null;
  const currentSheets = activeFile ? activeFile.sheets || [] : sheets;
  const currentSheet = currentSheets[activeSheetIdx] || currentSheets[0];

  const handleSave = async () => {
    setSaving(true);
    setError(null);

    try {
      let payload;
      if (files && files.length > 0) {
        payload = {
          files: files.map((f) => ({
            temp_filename: f.temp_filename,
            filename: f.filename,
          })),
        };
      } else {
        payload = {
          temp_filename: previewData.temp_filename,
          filename: previewData.filename,
        };
      }

      const res = await fetch('/api/customers/save/', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Failed to commit Excel files');
      }

      onSaveConfirmed(data.data);
    } catch (err) {
      setError(err.message || 'Error saving Excel files');
      setSaving(false);
    }
  };

  return (
    <div className="modal-backdrop animate-fade-in" onClick={!saving ? onClose : undefined}>
      <div className="modal-card preview-modal-card glass-panel" onClick={(e) => e.stopPropagation()}>
        {/* Modal Header */}
        <div className="preview-modal-header">
          <div className="preview-header-info">
            <div className="preview-file-icon">
              <FileSpreadsheet size={26} className="text-emerald" />
            </div>
            <div>
              <div className="preview-title-row">
                <h3 className="preview-modal-title">
                  {total_files > 1 ? `${total_files} Excel Files Preview & Merge` : 'Excel Workbook Preview'}
                </h3>
                <span className="badge badge-emerald">Ready to Save & Merge</span>
              </div>
              <div className="preview-filenames-tags">
                {(filenames.length > 0 ? filenames : [filename]).map((name, i) => (
                  <span key={i} className="preview-filename-pill" title={name}>
                    <FileCheck size={12} className="text-cyan mr-1 flex-shrink-0" />
                    <span className="truncate">{name}</span>
                  </span>
                ))}
              </div>
            </div>
          </div>
          {!saving && (
            <button
              type="button"
              className="btn-close-modal"
              onClick={onClose}
              title="Cancel and keep current data"
            >
              <X size={20} />
            </button>
          )}
        </div>

        {/* Safety Banner */}
        <div className="safety-banner">
          <ShieldCheck size={16} className="text-emerald flex-shrink-0" />
          <span>
            {total_files > 1
              ? `All ${total_files} files will be permanently saved as separate source workbooks. Customer records will be indexed and grouped without overwriting.`
              : 'Your customer data will be safely stored and indexed for instant name-only search.'}
          </span>
        </div>

        {error && (
          <div className="alert-error mb-3">
            <AlertCircle size={18} className="alert-icon" />
            <div className="alert-content">{error}</div>
          </div>
        )}

        {/* High-Level Summary Stats */}
        <div className="preview-stats-bar">
          <div className="preview-stat-item">
            <span className="preview-stat-label">Total Files</span>
            <span className="preview-stat-val text-primary font-bold">
              <FileSpreadsheet size={15} className="text-cyan inline mr-1" />
              {total_files} {total_files === 1 ? 'File' : 'Files'}
            </span>
          </div>
          <div className="preview-stat-divider" />
          <div className="preview-stat-item">
            <span className="preview-stat-label">Total Sheets</span>
            <span className="preview-stat-val">
              <Layers size={15} className="text-cyan inline mr-1" />
              {sheet_count} Sheets
            </span>
          </div>
          <div className="preview-stat-divider" />
          <div className="preview-stat-item">
            <span className="preview-stat-label">Total Records</span>
            <span className="preview-stat-val">
              <Database size={15} className="text-muted inline mr-1" />
              {total_records.toLocaleString()} Rows
            </span>
          </div>
          <div className="preview-stat-divider" />
          <div className="preview-stat-item">
            <span className="preview-stat-label">Merged Customers</span>
            <span className="preview-stat-val text-emerald font-bold">
              <Users size={15} className="inline mr-1" />
              {customer_count.toLocaleString()} Unique
            </span>
          </div>
          <div className="preview-stat-divider" />
          <div className="preview-stat-item">
            <span className="preview-stat-label">Status</span>
            <span className="preview-stat-val text-emerald font-medium">
              <CheckCircle2 size={15} className="inline mr-1" />
              Validated
            </span>
          </div>
        </div>

        {/* File Tabs (if multi-file) */}
        {files && files.length > 1 && (
          <div className="file-tabs-bar">
            <span className="text-xs font-semibold text-secondary mr-2">Select File to Inspect:</span>
            <div className="file-tabs-row">
              {files.map((f, fIdx) => (
                <button
                  key={f.file_id || fIdx}
                  type="button"
                  className={`file-tab-btn ${fIdx === activeFileIdx ? 'file-tab-active' : ''}`}
                  onClick={() => {
                    setActiveFileIdx(fIdx);
                    setActiveSheetIdx(0);
                  }}
                >
                  <FileSpreadsheet size={14} className="mr-1.5" />
                  <span className="tab-file-title">{f.filename}</span>
                  <span className="tab-file-count">
                    ({f.sheet_count} Sheets • {f.total_records} rows)
                  </span>
                </button>
              ))}
            </div>
          </div>
        )}

        {/* Sheet Summary Chips */}
        <div className="sheet-summary-section">
          <div className="flex items-center justify-between mb-2">
            <h4 className="section-label-sm">
              {activeFile ? `${activeFile.filename} Sheets (${currentSheets.length})` : `All Sheets (${currentSheets.length})`}
            </h4>
            <span className="text-xs text-secondary">Click any sheet to preview rows</span>
          </div>

          <div className="sheet-summary-chips">
            {currentSheets.map((s, idx) => (
              <button
                key={s.name || idx}
                type="button"
                className={`sheet-summary-chip ${idx === activeSheetIdx ? 'chip-active' : ''}`}
                onClick={() => setActiveSheetIdx(idx)}
              >
                <span className="chip-name">{s.name}</span>
                <span className="chip-count">
                  {s.records_count} {s.records_count === 1 ? 'record' : 'records'}
                </span>
                {s.sheet_route && s.sheet_route !== 'Route information not available' && (
                  <span className="chip-route">{s.sheet_route}</span>
                )}
              </button>
            ))}
          </div>
        </div>

        {/* Active Sheet Sample Data Preview */}
        {currentSheet && (
          <div className="preview-sheet-details">
            <div className="preview-sheet-info-row">
              <span className="font-semibold text-primary">
                {activeFile ? `${activeFile.filename} → ` : ''}{currentSheet.name}
              </span>
              <span className="text-secondary text-xs">
                Headers detected: {currentSheet.headers?.length || 0} columns • Total{' '}
                {currentSheet.records_count} records
              </span>
            </div>

            {currentSheet.preview_rows && currentSheet.preview_rows.length > 0 ? (
              <div className="preview-table-container">
                <table className="preview-table">
                  <thead>
                    <tr>
                      <th className="th-num">#</th>
                      {currentSheet.headers?.map((h) => (
                        <th key={h}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {currentSheet.preview_rows.map((row) => (
                      <tr key={row.record_id || row.row_number}>
                        <td className="td-num">Row {row.row_number}</td>
                        {currentSheet.headers?.map((h) => (
                          <td key={h}>{row.data?.[h] || '—'}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="preview-empty-sheet">
                This sheet has 0 records (empty or header-only).
              </div>
            )}
          </div>
        )}

        {/* Action Buttons */}
        <div className="modal-actions preview-modal-actions">
          <button
            type="button"
            className="btn btn-secondary"
            onClick={onClose}
            disabled={saving}
          >
            Cancel
          </button>
          <button
            type="button"
            className="btn btn-primary"
            onClick={handleSave}
            disabled={saving}
            id="btn-save-replace-workbook"
          >
            {saving ? (
              <>
                <Loader2 size={16} className="spin-icon mr-2" />
                Saving to Permanent Storage...
              </>
            ) : (
              <>
                <CheckCircle2 size={16} className="mr-2" />
                {total_files > 1 ? `Save & Merge All ${total_files} Excel Files` : 'Save Excel Workbook'}
              </>
            )}
          </button>
        </div>
      </div>
    </div>
  );
}
