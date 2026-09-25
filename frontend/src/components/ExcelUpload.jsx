import React, { useState, useRef } from 'react';
import {
  UploadCloud,
  FileSpreadsheet,
  AlertCircle,
  Loader2,
  CheckCircle2,
  X,
  FileCheck,
  Layers,
  ArrowRight,
  Sparkles,
  RefreshCw
} from 'lucide-react';
import UploadPreviewModal from './UploadPreviewModal';

export default function ExcelUpload({ onSavedSuccessfully }) {
  // 3 distinct file slots
  const [fileSlot1, setFileSlot1] = useState(null);
  const [fileSlot2, setFileSlot2] = useState(null);
  const [fileSlot3, setFileSlot3] = useState(null);

  const [dragActive, setDragActive] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [previewData, setPreviewData] = useState(null);
  const [error, setError] = useState(null);

  const multiFileInputRef = useRef(null);
  const slot1InputRef = useRef(null);
  const slot2InputRef = useRef(null);
  const slot3InputRef = useRef(null);

  const formatFileSize = (bytes) => {
    if (!bytes && bytes !== 0) return '';
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
  };

  const handleDrag = (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === 'dragenter' || e.type === 'dragover') {
      setDragActive(true);
    } else if (e.type === 'dragleave') {
      setDragActive(false);
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      distributeFiles(Array.from(e.dataTransfer.files));
    }
  };

  const handleMultiFileInputChange = (e) => {
    if (e.target.files && e.target.files.length > 0) {
      distributeFiles(Array.from(e.target.files));
    }
  };

  // Helper: Distribute dropped/selected files into Slot 1, Slot 2, Slot 3
  const distributeFiles = (files) => {
    const validFiles = files.filter(
      (f) => f.name.endsWith('.xlsx') || f.name.endsWith('.xls')
    );

    if (validFiles.length === 0) {
      setError('Please select valid Excel files (.xlsx or .xls).');
      return;
    }

    setError(null);

    // If 3 or more files dropped, assign first 3 to slots 1, 2, 3
    if (validFiles.length >= 3) {
      setFileSlot1(validFiles[0]);
      setFileSlot2(validFiles[1]);
      setFileSlot3(validFiles[2]);
    } else if (validFiles.length === 2) {
      if (!fileSlot1) {
        setFileSlot1(validFiles[0]);
        setFileSlot2(validFiles[1]);
      } else if (!fileSlot2) {
        setFileSlot2(validFiles[0]);
        setFileSlot3(validFiles[1]);
      } else {
        setFileSlot1(validFiles[0]);
        setFileSlot2(validFiles[1]);
      }
    } else if (validFiles.length === 1) {
      // Put in first empty slot
      if (!fileSlot1) setFileSlot1(validFiles[0]);
      else if (!fileSlot2) setFileSlot2(validFiles[0]);
      else if (!fileSlot3) setFileSlot3(validFiles[0]);
      else setFileSlot1(validFiles[0]);
    }

    if (multiFileInputRef.current) {
      multiFileInputRef.current.value = '';
    }
  };

  const handleSingleSlotChange = (e, slotNumber) => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
        setError('Please select a valid Excel file (.xlsx or .xls).');
        return;
      }
      setError(null);
      if (slotNumber === 1) setFileSlot1(file);
      if (slotNumber === 2) setFileSlot2(file);
      if (slotNumber === 3) setFileSlot3(file);
    }
    e.target.value = '';
  };

  const selectedCount = [fileSlot1, fileSlot2, fileSlot3].filter(Boolean).length;

  const handleUploadAll = async () => {
    const filesToUpload = [fileSlot1, fileSlot2, fileSlot3].filter(Boolean);
    if (filesToUpload.length === 0) {
      setError('Please select at least 1 Excel file (3 files recommended).');
      return;
    }

    setError(null);
    setUploading(true);

    try {
      const formData = new FormData();
      filesToUpload.forEach((f) => {
        formData.append('files', f);
      });

      const res = await fetch('/api/customers/upload/', {
        method: 'POST',
        body: formData,
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Failed to parse Excel files');
      }

      setPreviewData(data);
    } catch (err) {
      setError(err.message || 'Error uploading Excel files');
    } finally {
      setUploading(false);
    }
  };

  const handleCancelPreview = async () => {
    if (previewData) {
      const tempFilenames = [];
      if (previewData.temp_filename) tempFilenames.push(previewData.temp_filename);
      if (previewData.files) {
        previewData.files.forEach((f) => {
          if (f.temp_filename) tempFilenames.push(f.temp_filename);
        });
      }
      try {
        await fetch('/api/customers/cancel_upload/', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ temp_filenames: tempFilenames }),
        });
      } catch (e) {
        // silent
      }
    }
    setPreviewData(null);
  };

  const handleSaveConfirmed = (savedMetadata) => {
    setPreviewData(null);
    if (onSavedSuccessfully) {
      onSavedSuccessfully(savedMetadata);
    }
  };

  const handleResetSlots = () => {
    setFileSlot1(null);
    setFileSlot2(null);
    setFileSlot3(null);
    setError(null);
  };

  return (
    <div className="upload-container animate-fade-in">
      <div className="upload-header text-center">
        <div className="setup-badge">
          <Sparkles size={14} className="text-sky mr-1.5" />
          <span>3 Excel Workbooks Multi-Source Setup</span>
        </div>
        <h1 className="setup-title">Upload Exactly 3 Excel Files</h1>
        <p className="setup-subtitle">
          Upload 3 Excel workbooks at once. Data from ALL files and ALL sheets will be parsed,
          merged, and indexed together without overwriting. Customers appearing in multiple files
          will be grouped under the same customer name.
        </p>
      </div>

      {error && (
        <div className="alert-error animate-fade-in mb-4">
          <AlertCircle size={20} className="alert-icon flex-shrink-0" />
          <div className="alert-content">
            <strong>Upload Error:</strong> {error}
          </div>
        </div>
      )}

      {/* Main Multi-File Dropzone */}
      <div
        className={`dropzone glass-panel ${dragActive ? 'dropzone-active' : ''}`}
        onDragEnter={handleDrag}
        onDragLeave={handleDrag}
        onDragOver={handleDrag}
        onDrop={handleDrop}
        onClick={() => multiFileInputRef.current?.click()}
      >
        <input
          ref={multiFileInputRef}
          type="file"
          accept=".xlsx, .xls"
          multiple
          onChange={handleMultiFileInputChange}
          style={{ display: 'none' }}
        />

        <div className="dropzone-content">
          <div className="dropzone-icon-wrapper">
            {uploading ? (
              <Loader2 size={44} className="spin-icon text-sky" />
            ) : (
              <UploadCloud size={44} className="text-sky" />
            )}
          </div>

          <h3 className="dropzone-title">
            {uploading ? 'Analyzing All Sheets Across 3 Files...' : 'Select or Drop 3 Excel Files Here'}
          </h3>

          <p className="dropzone-desc">
            Drag & drop all 3 files at once, or browse files to auto-fill Slots 1, 2, and 3 below.
          </p>

          <div className="flex items-center justify-center gap-3 mt-4">
            <button
              type="button"
              className="btn btn-primary btn-md"
              disabled={uploading}
              onClick={(e) => {
                e.stopPropagation();
                multiFileInputRef.current?.click();
              }}
              id="btn-browse-3-files"
            >
              <FileSpreadsheet size={16} className="mr-2" /> Browse 3 Excel Files
            </button>

            {selectedCount > 0 && (
              <button
                type="button"
                className="btn btn-secondary btn-md"
                disabled={uploading}
                onClick={(e) => {
                  e.stopPropagation();
                  handleResetSlots();
                }}
              >
                <RefreshCw size={14} className="mr-1.5" /> Clear All
              </button>
            )}
          </div>
        </div>
      </div>

      {/* Hidden file inputs for individual slots */}
      <input
        ref={slot1InputRef}
        type="file"
        accept=".xlsx, .xls"
        onChange={(e) => handleSingleSlotChange(e, 1)}
        style={{ display: 'none' }}
      />
      <input
        ref={slot2InputRef}
        type="file"
        accept=".xlsx, .xls"
        onChange={(e) => handleSingleSlotChange(e, 2)}
        style={{ display: 'none' }}
      />
      <input
        ref={slot3InputRef}
        type="file"
        accept=".xlsx, .xls"
        onChange={(e) => handleSingleSlotChange(e, 3)}
        style={{ display: 'none' }}
      />

      {/* 3 Explicit File Slot Cards */}
      <div className="file-slots-grid">
        {/* Slot 1 */}
        <div
          className={`file-slot-card glass-panel ${fileSlot1 ? 'slot-filled' : 'slot-empty'}`}
          onClick={() => !fileSlot1 && slot1InputRef.current?.click()}
        >
          <div className="slot-badge-header">
            <span className="slot-number-badge">FILE 01</span>
            {fileSlot1 ? (
              <span className="slot-status-tag tag-ready">
                <CheckCircle2 size={13} className="mr-1 text-sky" /> Selected
              </span>
            ) : (
              <span className="slot-status-tag tag-waiting">Waiting</span>
            )}
          </div>

          <div className="slot-body">
            <div className="slot-icon-box">
              <FileSpreadsheet size={24} className={fileSlot1 ? 'text-sky' : 'text-muted'} />
            </div>
            <div className="slot-info">
              {fileSlot1 ? (
                <>
                  <h4 className="slot-filename" title={fileSlot1.name}>
                    {fileSlot1.name}
                  </h4>
                  <p className="slot-meta">{formatFileSize(fileSlot1.size)} • Excel Workbook</p>
                </>
              ) : (
                <>
                  <h4 className="slot-filename-placeholder">Select File 1</h4>
                  <p className="slot-meta-placeholder">e.g. TICKET DATA 1.xlsx</p>
                </>
              )}
            </div>
          </div>

          <div className="slot-footer">
            {fileSlot1 ? (
              <div className="slot-actions">
                <button
                  type="button"
                  className="btn btn-secondary btn-xs"
                  onClick={(e) => {
                    e.stopPropagation();
                    slot1InputRef.current?.click();
                  }}
                  title="Change file"
                >
                  Change
                </button>
                <button
                  type="button"
                  className="btn btn-danger-soft btn-xs"
                  onClick={(e) => {
                    e.stopPropagation();
                    setFileSlot1(null);
                  }}
                  title="Remove file"
                >
                  <X size={12} className="mr-1" /> Remove
                </button>
              </div>
            ) : (
              <button
                type="button"
                className="btn btn-secondary btn-xs w-full"
                onClick={(e) => {
                  e.stopPropagation();
                  slot1InputRef.current?.click();
                }}
              >
                Browse File 1
              </button>
            )}
          </div>
        </div>

        {/* Slot 2 */}
        <div
          className={`file-slot-card glass-panel ${fileSlot2 ? 'slot-filled' : 'slot-empty'}`}
          onClick={() => !fileSlot2 && slot2InputRef.current?.click()}
        >
          <div className="slot-badge-header">
            <span className="slot-number-badge">FILE 02</span>
            {fileSlot2 ? (
              <span className="slot-status-tag tag-ready">
                <CheckCircle2 size={13} className="mr-1 text-sky" /> Selected
              </span>
            ) : (
              <span className="slot-status-tag tag-waiting">Waiting</span>
            )}
          </div>

          <div className="slot-body">
            <div className="slot-icon-box">
              <FileSpreadsheet size={24} className={fileSlot2 ? 'text-sky' : 'text-muted'} />
            </div>
            <div className="slot-info">
              {fileSlot2 ? (
                <>
                  <h4 className="slot-filename" title={fileSlot2.name}>
                    {fileSlot2.name}
                  </h4>
                  <p className="slot-meta">{formatFileSize(fileSlot2.size)} • Excel Workbook</p>
                </>
              ) : (
                <>
                  <h4 className="slot-filename-placeholder">Select File 2</h4>
                  <p className="slot-meta-placeholder">e.g. TICKET ONEWAY.xlsx</p>
                </>
              )}
            </div>
          </div>

          <div className="slot-footer">
            {fileSlot2 ? (
              <div className="slot-actions">
                <button
                  type="button"
                  className="btn btn-secondary btn-xs"
                  onClick={(e) => {
                    e.stopPropagation();
                    slot2InputRef.current?.click();
                  }}
                  title="Change file"
                >
                  Change
                </button>
                <button
                  type="button"
                  className="btn btn-danger-soft btn-xs"
                  onClick={(e) => {
                    e.stopPropagation();
                    setFileSlot2(null);
                  }}
                  title="Remove file"
                >
                  <X size={12} className="mr-1" /> Remove
                </button>
              </div>
            ) : (
              <button
                type="button"
                className="btn btn-secondary btn-xs w-full"
                onClick={(e) => {
                  e.stopPropagation();
                  slot2InputRef.current?.click();
                }}
              >
                Browse File 2
              </button>
            )}
          </div>
        </div>

        {/* Slot 3 */}
        <div
          className={`file-slot-card glass-panel ${fileSlot3 ? 'slot-filled' : 'slot-empty'}`}
          onClick={() => !fileSlot3 && slot3InputRef.current?.click()}
        >
          <div className="slot-badge-header">
            <span className="slot-number-badge">FILE 03</span>
            {fileSlot3 ? (
              <span className="slot-status-tag tag-ready">
                <CheckCircle2 size={13} className="mr-1 text-sky" /> Selected
              </span>
            ) : (
              <span className="slot-status-tag tag-waiting">Waiting</span>
            )}
          </div>

          <div className="slot-body">
            <div className="slot-icon-box">
              <FileSpreadsheet size={24} className={fileSlot3 ? 'text-sky' : 'text-muted'} />
            </div>
            <div className="slot-info">
              {fileSlot3 ? (
                <>
                  <h4 className="slot-filename" title={fileSlot3.name}>
                    {fileSlot3.name}
                  </h4>
                  <p className="slot-meta">{formatFileSize(fileSlot3.size)} • Excel Workbook</p>
                </>
              ) : (
                <>
                  <h4 className="slot-filename-placeholder">Select File 3</h4>
                  <p className="slot-meta-placeholder">e.g. Train HRDW-AHM ON 3RD OCT (2).xlsx</p>
                </>
              )}
            </div>
          </div>

          <div className="slot-footer">
            {fileSlot3 ? (
              <div className="slot-actions">
                <button
                  type="button"
                  className="btn btn-secondary btn-xs"
                  onClick={(e) => {
                    e.stopPropagation();
                    slot3InputRef.current?.click();
                  }}
                  title="Change file"
                >
                  Change
                </button>
                <button
                  type="button"
                  className="btn btn-danger-soft btn-xs"
                  onClick={(e) => {
                    e.stopPropagation();
                    setFileSlot3(null);
                  }}
                  title="Remove file"
                >
                  <X size={12} className="mr-1" /> Remove
                </button>
              </div>
            ) : (
              <button
                type="button"
                className="btn btn-secondary btn-xs w-full"
                onClick={(e) => {
                  e.stopPropagation();
                  slot3InputRef.current?.click();
                }}
              >
                Browse File 3
              </button>
            )}
          </div>
        </div>
      </div>

      {/* Upload Submission Bar */}
      <div className="upload-submit-bar glass-panel mt-6">
        <div className="submit-bar-left">
          <div className="status-indicator-badge">
            <span
              className={`status-circle ${
                selectedCount === 3
                  ? 'status-circle-success'
                  : selectedCount > 0
                  ? 'status-circle-warning'
                  : 'status-circle-muted'
              }`}
            />
            <span className="font-semibold text-primary">
              {selectedCount === 3
                ? 'All 3 Files Selected'
                : `${selectedCount} of 3 Files Selected`}
            </span>
          </div>
          <span className="submit-bar-hint text-xs text-secondary">
            {selectedCount === 3
              ? 'Ready to parse all sheets and merge customer records together.'
              : 'Please select all 3 Excel files to begin unified multi-source processing.'}
          </span>
        </div>

        <button
          type="button"
          className="btn btn-primary btn-lg"
          disabled={selectedCount === 0 || uploading}
          onClick={handleUploadAll}
          id="btn-submit-3-files"
        >
          {uploading ? (
            <>
              <Loader2 size={18} className="spin-icon mr-2" />
              Parsing All Sheets & Files...
            </>
          ) : (
            <>
              <CheckCircle2 size={18} className="mr-2" />
              Upload & Merge {selectedCount > 0 ? `${selectedCount} Excel Files` : '3 Files'}
            </>
          )}
        </button>
      </div>

      {/* Upload Preview Modal */}
      <UploadPreviewModal
        isOpen={!!previewData}
        previewData={previewData}
        onClose={handleCancelPreview}
        onSaveConfirmed={handleSaveConfirmed}
        isReplacing={false}
      />
    </div>
  );
}
