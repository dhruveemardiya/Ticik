import React, { useState, useEffect } from 'react';
import { X, Search, Users, User, Train, Navigation, Layers, Hash, ChevronRight } from 'lucide-react';

export default function CustomerDirectoryModal({
  isOpen,
  onClose,
  onSelectCustomer,
  sheets = [],
  totalCustomers = 0
}) {
  const [customers, setCustomers] = useState([]);
  const [loading, setLoading] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [activeSheet, setActiveSheet] = useState('');

  useEffect(() => {
    if (!isOpen) return;

    const fetchCustomers = async () => {
      setLoading(true);
      try {
        const params = new URLSearchParams();
        if (activeSheet) params.append('sheet', activeSheet);
        if (searchQuery.trim()) params.append('search', searchQuery.trim());

        const res = await fetch(`/api/customers/list/?${params.toString()}`);
        if (res.ok) {
          const data = await res.json();
          setCustomers(data.customers || []);
        }
      } catch (err) {
        console.error('Failed to load customers:', err);
      } finally {
        setLoading(false);
      }
    };

    const timer = setTimeout(fetchCustomers, 120);
    return () => clearTimeout(timer);
  }, [isOpen, activeSheet, searchQuery]);

  if (!isOpen) return null;

  const visibleSheets = sheets.filter(
    (s) => !s.name?.toLowerCase().includes('missing data')
  );

  return (
    <div className="modal-backdrop animate-fade-in" onClick={onClose}>
      <div className="directory-modal glass-panel" onClick={(e) => e.stopPropagation()}>
        {/* Header */}
        <div className="directory-header">
          <div className="directory-title-wrap">
            <div className="directory-icon-badge">
              <Users size={22} className="text-sky" />
            </div>
            <div>
              <h2 className="directory-title">Customer Directory</h2>
              <p className="directory-subtitle">
                {totalCustomers.toLocaleString()} total customers across {visibleSheets.length} sheet
                {visibleSheets.length === 1 ? '' : 's'} in saved Excel workbook
              </p>
            </div>
          </div>
          <button type="button" className="btn-close-modal" onClick={onClose} title="Close directory">
            <X size={20} />
          </button>
        </div>

        {/* Filter Toolbar */}
        <div className="directory-toolbar">
          <div className="directory-search-input-box">
            <Search size={18} className="text-muted mr-2 flex-shrink-0" />
            <input
              type="text"
              placeholder="Filter by customer name across all sheets..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="directory-search-input"
              autoFocus
            />
            {searchQuery && (
              <button
                type="button"
                className="search-clear-btn"
                onClick={() => setSearchQuery('')}
              >
                <X size={16} />
              </button>
            )}
          </div>

          {/* Sheet Selector Pills */}
          <div className="sheet-pills-row">
            <button
              type="button"
              className={`sheet-pill ${activeSheet === '' ? 'sheet-pill-active' : ''}`}
              onClick={() => setActiveSheet('')}
            >
              All Sheets ({totalCustomers.toLocaleString()})
            </button>
            {visibleSheets.map((s) => (
              <button
                key={s.name}
                type="button"
                className={`sheet-pill ${activeSheet === s.name ? 'sheet-pill-active' : ''}`}
                onClick={() => setActiveSheet(s.name)}
              >
                {s.name} ({s.records_count})
              </button>
            ))}
          </div>
        </div>

        {/* Customers List Body */}
        <div className="directory-body">
          {loading ? (
            <div className="directory-loading">
              <div className="status-dot status-dot-active mb-2" />
              <p className="text-secondary text-sm">Loading customer directory...</p>
            </div>
          ) : customers.length === 0 ? (
            <div className="directory-empty">
              <User size={36} className="text-dim mb-2" />
              <p className="font-semibold text-primary">No customers found</p>
              <span className="text-dim text-sm">
                {searchQuery ? `No matches found for "${searchQuery}"` : 'This sheet has no customer records'}
              </span>
            </div>
          ) : (
            <div className="directory-grid">
              {customers.map((c, idx) => {
                return (
                  <div
                    key={c.normalized_name || c.name}
                    className="directory-card"
                    onClick={() => {
                      onSelectCustomer(c.name);
                      onClose();
                    }}
                  >
                    <div className="directory-card-index">
                      {String(idx + 1).padStart(2, '0')}
                    </div>

                    <div className="directory-card-info">
                      <h4 className="directory-card-name">{c.name}</h4>
                      <div className="directory-card-meta">
                        <span className="meta-tag">
                          {c.record_count} {c.record_count === 1 ? 'record' : 'records'}
                        </span>
                      </div>
                    </div>

                    <div className="directory-card-action">
                      <ChevronRight size={18} className="text-dim" />
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
