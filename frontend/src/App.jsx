import React, { useState, useEffect, useCallback, useRef } from 'react';
import ExcelUpload from './components/ExcelUpload';
import CustomerSearch from './components/CustomerSearch';
import CustomerListPanel from './components/CustomerListPanel';
import CustomerDetails from './components/CustomerDetails';
import CustomerDirectoryModal from './components/CustomerDirectoryModal';
import UploadPreviewModal from './components/UploadPreviewModal';
import SearchResultsTable from './components/SearchResultsTable';
import DeleteConfirmModal from './components/DeleteConfirmModal';
import {
  FileSpreadsheet,
  Users,
  Upload,
  AlertTriangle,
  RotateCw,
  FileCheck,
  Loader2,
  Trash2,
  X
} from 'lucide-react';
import './App.css';

export default function App() {
  // System & Connection State
  const [systemStatus, setSystemStatus] = useState(null);
  const [loadingStatus, setLoadingStatus] = useState(true);
  const [backendOffline, setBackendOffline] = useState(false);
  const [retryingConnection, setRetryingConnection] = useState(false);

  // Selected Unique Customer State (Requirements 7, 25, 38, 39)
  // Initially null: NO customer is auto-selected! Right panel remains clean and empty on initial load.
  const [selectedCustomerName, setSelectedCustomerName] = useState(null);
  const [selectedCustomerData, setSelectedCustomerData] = useState(null);
  const [loadingRecord, setLoadingRecord] = useState(false);

  // Customer list state (Unique customer names)
  const [customerList, setCustomerList] = useState([]);
  const [loadingList, setLoadingList] = useState(false);
  const [sheetFilter, setSheetFilter] = useState('');

  // Modals & Upload State
  const [showDirectoryModal, setShowDirectoryModal] = useState(false);
  const [showDeleteAllModal, setShowDeleteAllModal] = useState(false);
  const [showUploadModal, setShowUploadModal] = useState(false);
  const [uploadPreviewData, setUploadPreviewData] = useState(null);
  const [isUploading, setIsUploading] = useState(false);
  const [uploadError, setUploadError] = useState(null);

  // Name-Only Search Suggestions State (Requirements 4, 5, 6, 23, 24)
  const [searchQuery, setSearchQuery] = useState('');
  const [searchSuggestions, setSearchSuggestions] = useState([]);
  const [searchLoading, setSearchLoading] = useState(false);
  const searchDebounceRef = useRef(null);

  const fileInputRef = useRef(null);

  // 1. Fetch system status from Django
  const fetchStatus = useCallback(async (isRetry = false) => {
    if (isRetry) setRetryingConnection(true);
    try {
      const res = await fetch('/api/customers/status/');
      if (res.ok) {
        const data = await res.json();
        setSystemStatus(data);
        setBackendOffline(false);
        setRetryingConnection(false);
        return data;
      } else {
        setBackendOffline(true);
      }
    } catch (err) {
      console.warn('Backend connection check failed:', err.message);
      setBackendOffline(true);
    } finally {
      if (isRetry) setRetryingConnection(false);
    }
    return null;
  }, []);

  // 2. Fetch unique customers list for left panel (NO AUTO SELECT FIRST)
  const fetchCustomersList = useCallback(async (sheet = '') => {
    setLoadingList(true);
    try {
      const params = new URLSearchParams();
      if (sheet) params.append('sheet', sheet);

      const res = await fetch(`/api/customers/list/?${params.toString()}`);
      if (res.ok) {
        const data = await res.json();
        const list = data.customers || [];
        setCustomerList(list);
      }
    } catch (err) {
      console.error('Failed to load customer list:', err);
    } finally {
      setLoadingList(false);
    }
  }, []);

  // 3. Initial check on mount: NEVER auto-select first customer
  useEffect(() => {
    let isMounted = true;
    const init = async () => {
      setLoadingStatus(true);
      const st = await fetchStatus();
      if (isMounted && st && st.available) {
        // Load list but DO NOT auto-select
        await fetchCustomersList('');
      }
      if (isMounted) setLoadingStatus(false);
    };
    init();
    return () => {
      isMounted = false;
    };
  }, [fetchStatus, fetchCustomersList]);

  // When sheet filter changes: reset customer selection so right panel shows empty state (Requirement 39)
  const handleSheetFilterChange = (sheet) => {
    setSheetFilter(sheet);
    setSelectedCustomerName(null);
    setSelectedCustomerData(null);
    fetchCustomersList(sheet);
  };

  // 4. Fetch all records for the selected unique customer name (Requirements 8, 9, 25, 36)
  const handleSelectCustomer = async (customerName) => {
    if (!customerName) return;
    setSelectedCustomerName(customerName);
    setLoadingRecord(true);

    try {
      const res = await fetch(`/api/customers/records/?name=${encodeURIComponent(customerName)}`);
      if (res.ok) {
        const data = await res.json();
        setSelectedCustomerData(data);
      } else {
        setSelectedCustomerData(null);
      }
    } catch (err) {
      console.error('Failed to fetch customer records:', err);
      setSelectedCustomerData(null);
    } finally {
      setLoadingRecord(false);
    }
  };

  // 5. When an individual Excel record is deleted (Action 1: Delete Record)
  const handleRecordDeleted = async (deletedRecordId) => {
    await fetchStatus();
    await fetchCustomersList(sheetFilter);

    if (selectedCustomerName) {
      try {
        const res = await fetch(`/api/customers/records/?name=${encodeURIComponent(selectedCustomerName)}`);
        if (res.ok) {
          const data = await res.json();
          if (data && data.records && data.records.length > 0) {
            setSelectedCustomerData(data);
          } else {
            setSelectedCustomerName(null);
            setSelectedCustomerData(null);
          }
        } else {
          setSelectedCustomerName(null);
          setSelectedCustomerData(null);
        }
      } catch (e) {
        setSelectedCustomerName(null);
        setSelectedCustomerData(null);
      }
    }
  };

  // 5b. When an entire customer is deleted across all sheets (Action 2: Delete Customer)
  const handleCustomerEntirelyDeleted = async (deletedCustomerName) => {
    setSelectedCustomerName(null);
    setSelectedCustomerData(null);
    await fetchStatus();
    await fetchCustomersList(sheetFilter);
  };

  // 5c. When all customers are deleted across all sheets (Action 3: Delete All Customers)
  const handleAllCustomersDeleted = async () => {
    setSelectedCustomerName(null);
    setSelectedCustomerData(null);
    setSheetFilter('');
    setShowDeleteAllModal(false);
    await fetchStatus();
    await fetchCustomersList('');
  };

  // 6. Name-Only Search Query Handler: Fetches unique suggestions
  const executeSearchSuggestions = useCallback(async (queryText) => {
    if (!queryText) {
      setSearchSuggestions([]);
      setSearchLoading(false);
      return;
    }

    setSearchLoading(true);
    try {
      const res = await fetch(`/api/customers/search/?name=${encodeURIComponent(queryText)}`);
      if (res.ok) {
        const data = await res.json();
        setSearchSuggestions(data.results || []);
      }
    } catch (err) {
      console.error('Search request failed:', err);
    } finally {
      setSearchLoading(false);
    }
  }, []);

  const handleSearchQueryChange = (newQuery) => {
    setSearchQuery(newQuery);
    const trimmed = newQuery.trim();

    if (searchDebounceRef.current) {
      clearTimeout(searchDebounceRef.current);
    }

    if (!trimmed) {
      setSearchSuggestions([]);
      setSearchLoading(false);
      return;
    }

    setSearchLoading(true);
    searchDebounceRef.current = setTimeout(() => {
      executeSearchSuggestions(trimmed);
    }, 150);
  };

  const handleClearSearch = () => {
    if (searchDebounceRef.current) {
      clearTimeout(searchDebounceRef.current);
    }
    setSearchQuery('');
    setSearchSuggestions([]);
    setSearchLoading(false);
  };

  const handleSelectFromSearch = async (customerName) => {
    await handleSelectCustomer(customerName);
    setSearchSuggestions([]);
  };

  // 7. Handle file selection from Top-Right [ Upload Excel ] button
  const handleTopRightUploadClick = () => {
    setUploadError(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
      fileInputRef.current.click();
    }
  };

  const handleFileInputChange = async (e) => {
    if (!e.target.files || !e.target.files[0]) return;
    const file = e.target.files[0];

    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
      setUploadError('Please select a valid Excel workbook (.xlsx or .xls)');
      return;
    }

    setUploadError(null);
    setIsUploading(true);

    try {
      const formData = new FormData();
      formData.append('file', file);

      const res = await fetch('/api/customers/upload/', {
        method: 'POST',
        body: formData,
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Failed to parse Excel workbook');
      }

      setUploadPreviewData(data);
    } catch (err) {
      setUploadError(
        err.message || 'Unable to read this Excel file. Current customer data is still safe.'
      );
    } finally {
      setIsUploading(false);
      if (fileInputRef.current) {
        fileInputRef.current.value = '';
      }
    }
  };

  // 8. Cancel Upload in Preview Modal
  const handleCancelUploadPreview = async () => {
    if (uploadPreviewData?.temp_filename) {
      try {
        await fetch('/api/customers/cancel_upload/', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ temp_filename: uploadPreviewData.temp_filename }),
        });
      } catch (e) {
        // silent
      }
    }
    setUploadPreviewData(null);
  };

  // 9. Confirm Save / Replace in Preview Modal
  const handleSaveConfirmed = async (savedData) => {
    setUploadPreviewData(null);
    setShowUploadModal(false);
    setUploadError(null);
    await fetchStatus();
    setSheetFilter('');
    handleClearSearch();
    await fetchCustomersList('', true);
  };

  // Manual retry connection button handler
  const handleRetryConnection = async () => {
    const st = await fetchStatus(true);
    if (st && st.available) {
      await fetchCustomersList('', true);
    }
  };

  // --- RENDER LOGIC ---

  // 1. Initial Loading Screen
  if (loadingStatus) {
    return (
      <div className="app-loading-screen">
        <div className="status-dot status-dot-active mb-3" />
        <p className="text-secondary font-medium">Connecting to Customer Data System...</p>
      </div>
    );
  }

  // 2. BACKEND_OFFLINE
  if (backendOffline && (!systemStatus || !systemStatus.available)) {
    return (
      <div className="offline-screen animate-fade-in">
        <div className="offline-card glass-panel">
          <div className="offline-icon-wrap">
            <AlertTriangle size={36} className="text-amber" />
          </div>
          <h2 className="offline-title">Backend Connection Required</h2>
          <p className="offline-message">Django backend is not running.</p>
          <div className="offline-instruction">
            <span>Please start the Django server on:</span>
            <code className="offline-url">http://127.0.0.1:8000</code>
          </div>
          <button
            type="button"
            className="btn btn-primary btn-md mt-4"
            onClick={handleRetryConnection}
            disabled={retryingConnection}
            id="btn-retry-backend-connection"
          >
            {retryingConnection ? (
              <>
                <Loader2 size={16} className="spin-icon mr-2" /> Connecting...
              </>
            ) : (
              <>
                <RotateCw size={16} className="mr-2" /> Retry Connection
              </>
            )}
          </button>
        </div>
      </div>
    );
  }

  const isDataReady = systemStatus && systemStatus.available;
  const sheetCount = systemStatus?.sheet_count || systemStatus?.sheets?.length || 1;
  const customerCount = systemStatus?.customer_count || 0;
  const isSearchActive = searchQuery.trim().length > 0;

  return (
    <div className="app-layout">
      {/* Hidden file input for [ Upload Excel ] button */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx, .xls"
        onChange={handleFileInputChange}
        style={{ display: 'none' }}
      />

      {/* Backend connection warning banner if temporarily disconnected while data is loaded */}
      {backendOffline && isDataReady && (
        <div className="connection-alert-banner">
          <AlertTriangle size={15} className="mr-2 flex-shrink-0" />
          <span>Django backend is currently offline. Existing data remains visible.</span>
          <button
            type="button"
            className="btn btn-secondary btn-xs ml-auto"
            onClick={handleRetryConnection}
            disabled={retryingConnection}
          >
            {retryingConnection ? (
              <Loader2 size={12} className="spin-icon mr-1" />
            ) : (
              <RotateCw size={12} className="mr-1" />
            )}
            Retry Connection
          </button>
        </div>
      )}

      {/* Floating Upload Error Banner */}
      {uploadError && (
        <div className="upload-floating-alert animate-fade-in">
          <div className="upload-alert-text">
            <strong>Upload Error:</strong> {uploadError}
          </div>
          <button
            type="button"
            className="btn-close-alert"
            onClick={() => setUploadError(null)}
          >
            <X size={16} />
          </button>
        </div>
      )}

      {/* Top Navbar */}
      <header className="app-navbar glass-panel">
        <div className="navbar-container">
          {/* Brand Left */}
          <div className="navbar-brand">
            <div className="brand-logo">
              <FileSpreadsheet size={22} className="text-sky" />
            </div>
            <div className="brand-text">
              <h1 className="brand-title">Customer Manager</h1>
            </div>
          </div>

          {/* Top-Right Excel Status Area — Redesigned ONE-LINE */}
          <div className="navbar-excel-bar">
            {isDataReady ? (
              <div className="excel-status-line">
                {/* Total Files Badge */}
                <div
                  className="excel-pill excel-pill-filename cursor-pointer"
                  title={
                    systemStatus.filenames && systemStatus.filenames.length > 0
                      ? systemStatus.filenames.join('\n')
                      : systemStatus.filename
                  }
                  onClick={() => setShowUploadModal(true)}
                >
                  <FileCheck size={14} className="text-cyan flex-shrink-0" />
                  <span className="file-name-truncate font-semibold">
                    {systemStatus.total_files > 1
                      ? `${systemStatus.total_files} Excel Files`
                      : systemStatus.filename}
                  </span>
                </div>

                <span className="excel-separator">•</span>

                {/* Sheets Count */}
                <span className="excel-meta-text">
                  {sheetCount} {sheetCount === 1 ? 'Sheet' : 'Sheets'}
                </span>

                <span className="excel-separator">•</span>

                {/* Total Records Count */}
                <span className="excel-meta-text">
                  {(systemStatus.total_records || 0).toLocaleString()} Records
                </span>

                <span className="excel-separator">•</span>

                {/* Clickable Customer Count */}
                <button
                  type="button"
                  className="excel-pill excel-pill-clickable"
                  onClick={() => setShowDirectoryModal(true)}
                  title="Click to view full customer directory"
                  id="btn-open-customers-directory"
                >
                  <Users size={14} className="text-sky flex-shrink-0" />
                  <span>
                    <strong>{customerCount.toLocaleString()} Customers</strong>
                  </span>
                </button>

                <span className="excel-separator">•</span>

                {/* Saved & Ready Status */}
                <div className="excel-pill excel-pill-status">
                  <span className="status-dot status-dot-active" />
                  <span className="text-sky font-semibold">Saved & Ready</span>
                </div>

                {/* Upload 3 Files Button */}
                <button
                  type="button"
                  className="btn btn-primary btn-sm btn-upload-topright"
                  onClick={() => setShowUploadModal(true)}
                  title="Upload 3 Excel workbooks"
                  id="btn-upload-excel"
                >
                  <Upload size={14} className="mr-1.5" /> Upload 3 Files
                </button>

                {/* Delete All Customers Button */}
                <button
                  type="button"
                  className="btn btn-danger-soft btn-sm btn-delete-all-top"
                  onClick={() => setShowDeleteAllModal(true)}
                  disabled={isUploading}
                  title="Permanently delete all customer records across all sheets"
                  id="btn-delete-all-customers"
                >
                  <Trash2 size={13} className="mr-1 text-rose" /> Delete All Customers
                </button>
              </div>
            ) : (
              <div className="excel-status-line">
                <span className="badge badge-neutral">No Excel Files Configured</span>
                <button
                  type="button"
                  className="btn btn-primary btn-sm btn-upload-topright ml-3"
                  onClick={() => setShowUploadModal(true)}
                  id="btn-upload-excel"
                >
                  <Upload size={14} className="mr-1.5" /> Upload 3 Files
                </button>
              </div>
            )}
          </div>
        </div>
      </header>

      {/* Main Container */}
      <main className="main-content">
        {!isDataReady ? (
          /* Step 1: Initial upload dropzone */
          <ExcelUpload onSavedSuccessfully={handleSaveConfirmed} />
        ) : (
          /* Step 2: Customer Management Workspace */
          <div className="customer-workspace animate-fade-in">
            {/* NAME-ONLY Search Bar with Unique Suggestions Dropdown (Requirements 4, 5, 6, 23, 24) */}
            <CustomerSearch
              query={searchQuery}
              onQueryChange={handleSearchQueryChange}
              suggestions={searchSuggestions}
              loading={searchLoading}
              onSelectCustomer={handleSelectFromSearch}
            />

            {/* Standard Two-Column Desktop Responsive Workspace (Requirements 7, 8, 9, 31, 38, 39, 40) */}
            <div className="two-column-workspace">
              {/* Left Column: Unique Customer Cards List with Sheet Filter */}
              <div className="workspace-left-col">
                <CustomerListPanel
                  customers={customerList}
                  selectedCustomerName={selectedCustomerName}
                  onSelectCustomer={handleSelectCustomer}
                  loading={loadingList}
                  sheetFilter={sheetFilter}
                  onSheetFilterChange={handleSheetFilterChange}
                  sheets={systemStatus?.sheets || []}
                />
              </div>

              {/* Right Column: Customer Details (Starts completely EMPTY on load; shows all records when customer selected) */}
              <div className="workspace-right-col">
                <CustomerDetails
                  customerData={selectedCustomerData}
                  loading={loadingRecord}
                  onRecordDeleted={handleRecordDeleted}
                  onCustomerDeleted={handleCustomerEntirelyDeleted}
                />
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Upload Preview Modal */}
      <UploadPreviewModal
        isOpen={!!uploadPreviewData}
        previewData={uploadPreviewData}
        onClose={handleCancelUploadPreview}
        onSaveConfirmed={handleSaveConfirmed}
        isReplacing={isDataReady}
      />

      {/* Full Customer Directory Modal */}
      <CustomerDirectoryModal
        isOpen={showDirectoryModal}
        onClose={() => setShowDirectoryModal(false)}
        onSelectCustomer={handleSelectCustomer}
        sheets={systemStatus?.sheets || []}
        totalCustomers={customerCount}
      />

      {/* Delete All Customers Modal (Requirement 4: Delete All Customers) */}
      <DeleteConfirmModal
        isOpen={showDeleteAllModal}
        onClose={() => setShowDeleteAllModal(false)}
        mode="all"
        systemStatus={systemStatus}
        onDeleted={handleAllCustomersDeleted}
      />

      {/* 3-File Upload Modal (when opened from top navbar) */}
      {showUploadModal && (
        <div className="modal-backdrop animate-fade-in" onClick={() => setShowUploadModal(false)}>
          <div className="modal-card upload-modal-dialog glass-panel" onClick={(e) => e.stopPropagation()}>
            <div className="modal-dialog-header">
              <div className="flex items-center gap-2">
                <FileSpreadsheet size={22} className="text-sky" />
                <h3 className="modal-dialog-title">Upload Exactly 3 Excel Files</h3>
              </div>
              <button
                type="button"
                className="btn-close-modal"
                onClick={() => setShowUploadModal(false)}
              >
                <X size={20} />
              </button>
            </div>
            <div className="modal-dialog-body">
              <ExcelUpload
                onSavedSuccessfully={(savedData) => {
                  setShowUploadModal(false);
                  handleSaveConfirmed(savedData);
                }}
              />
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
