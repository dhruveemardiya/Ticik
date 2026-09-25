import React, { useState } from 'react';
import {
  Table,
  Layers,
  Hash,
  X,
  ChevronDown,
  ChevronRight,
  UserCheck,
  UserX,
  Loader2,
  Train,
  Navigation,
  ExternalLink
} from 'lucide-react';

export default function SearchResultsTable({
  query = '',
  count = 0,
  columns = [],
  results = [],
  loading = false,
  onClearSearch,
  onSelectRecord,
}) {
  const [expandedRows, setExpandedRows] = useState({});

  const toggleRowExpand = (recordId) => {
    setExpandedRows((prev) => ({
      ...prev,
      [recordId]: !prev[recordId],
    }));
  };

  return (
    <div className="search-results-workspace animate-fade-in">
      {/* Top Results Bar */}
      <div className="search-results-header glass-panel">
        <div className="search-results-title-group">
          <h2 className="search-results-heading">Search Results</h2>
          <span className="results-count-pill">
            <strong>{count.toLocaleString()}</strong> Matching {count === 1 ? 'Record' : 'Records'}
            {query && (
              <>
                {' '}
                for <span className="query-highlight">"{query}"</span>
              </>
            )}
          </span>
          <span className="search-scope-tag">Name-only search across all sheets</span>
        </div>

        <button
          type="button"
          className="btn btn-secondary btn-sm btn-clear-search"
          onClick={onClearSearch}
          title="Return to standard customer list"
          id="btn-clear-search-results"
        >
          <X size={15} className="mr-1.5" /> Clear Search
        </button>
      </div>

      {/* Main Table Container */}
      <div className="search-table-panel glass-panel">
        {loading ? (
          <div className="table-loading-state">
            <Loader2 size={36} className="spin-icon text-sky mb-3" />
            <p className="font-semibold text-primary">Searching customer name across all sheets...</p>
            <span className="text-secondary text-xs">Querying "{query}"</span>
          </div>
        ) : results.length === 0 ? (
          <div className="table-empty-state">
            <div className="empty-icon-circle">
              <UserX size={36} className="text-muted" />
            </div>
            <h3 className="empty-state-title">No customers found with this name</h3>
            <p className="empty-state-desc">
              No records in the workbook have a Customer Name containing "{query}".
              <br />
              <span className="text-dim text-xs">
                Search is strictly restricted to the customer Name column per system specification.
              </span>
            </p>
            <button
              type="button"
              className="btn btn-secondary btn-sm mt-3"
              onClick={onClearSearch}
            >
              <X size={14} className="mr-1" /> Reset Search
            </button>
          </div>
        ) : (
          <div className="table-scroll-wrapper">
            <table className="excel-data-table">
              <thead>
                <tr>
                  <th className="th-sticky-index">#</th>
                  <th className="th-sticky-meta">Sheet</th>
                  <th className="th-sticky-meta">Row</th>
                  {columns.map((col) => (
                    <th key={col.key} title={col.key} className="th-data-col">
                      {col.label}
                    </th>
                  ))}
                  <th className="th-action-col">Actions</th>
                </tr>
              </thead>
              <tbody>
                {results.map((rec, idx) => {
                  const isExpanded = !!expandedRows[rec.record_id];
                  const hasRoute =
                    rec.train_route && rec.train_route !== 'Route information not available';

                  return (
                    <React.Fragment key={rec.record_id}>
                      <tr
                        className={`table-row-item ${isExpanded ? 'row-expanded' : ''}`}
                        onClick={() => toggleRowExpand(rec.record_id)}
                      >
                        {/* Index */}
                        <td className="td-index">
                          <span className="row-index-num">
                            {String(idx + 1).padStart(2, '0')}
                          </span>
                        </td>

                        {/* Sheet */}
                        <td className="td-sheet">
                          <span className="sheet-badge" title={rec.sheet_name}>
                            <Layers size={11} className="inline mr-1 text-cyan" />
                            {rec.sheet_name}
                          </span>
                        </td>

                        {/* Excel Row */}
                        <td className="td-row-num">
                          <span className="row-number-badge">
                            <Hash size={10} className="inline mr-0.5 text-muted" />
                            {rec.row_number}
                          </span>
                        </td>

                        {/* All Excel Data Columns in Exact Original Order */}
                        {columns.map((col) => {
                          const val = rec.row_data?.[col.key];
                          const isNameCol =
                            col.label === 'Customer Name' || col.key.toUpperCase() === 'NAME';
                          const isAmountCol =
                            col.label === 'Amount' || col.key.toUpperCase() === 'RS.';

                          return (
                            <td
                              key={col.key}
                              className={`td-cell ${isNameCol ? 'cell-customer-name' : ''} ${
                                isAmountCol ? 'cell-amount' : ''
                              }`}
                            >
                              {val !== undefined && val !== '' && val !== null ? (
                                <span className="cell-val-text">{val}</span>
                              ) : (
                                <span className="cell-empty">—</span>
                              )}
                            </td>
                          );
                        })}

                        {/* Row Action: Expand / Inspect */}
                        <td
                          className="td-action"
                          onClick={(e) => {
                            e.stopPropagation();
                            if (onSelectRecord) onSelectRecord(rec.record_id);
                          }}
                        >
                          <button
                            type="button"
                            className="btn btn-secondary btn-xs btn-inspect-row"
                            title="Open full detail view for this row"
                          >
                            <ExternalLink size={12} className="mr-1" /> View
                          </button>
                        </td>
                      </tr>

                      {/* Expandable Inline Row Detail (Req 16) */}
                      {isExpanded && (
                        <tr className="row-expanded-detail">
                          <td colSpan={columns.length + 4} className="expanded-content-cell">
                            <div className="expanded-record-card">
                              <div className="expanded-header">
                                <div className="expanded-title-line">
                                  <UserCheck size={18} className="text-sky mr-2" />
                                  <strong className="text-primary text-base">
                                    {rec.name}
                                  </strong>
                                  <span className="badge badge-cyan ml-2">
                                    {rec.sheet_name} • Row {rec.row_number}
                                  </span>
                                </div>

                                <div className="expanded-meta-tags">
                                  {rec.train_no && (
                                    <span className="meta-tag">
                                      <Train size={12} className="text-sky mr-1 inline" />
                                      Train: <strong>{rec.train_no}</strong>
                                    </span>
                                  )}
                                  {hasRoute && (
                                    <span className="meta-tag">
                                      <Navigation size={12} className="text-cyan mr-1 inline" />
                                      Route: <strong>{rec.train_route}</strong>
                                    </span>
                                  )}
                                </div>
                              </div>

                              <div className="expanded-fields-grid">
                                {rec.fields?.map((f) => (
                                  <div key={f.key} className="expanded-field-box">
                                    <span className="expanded-field-label">{f.label}</span>
                                    <span className="expanded-field-val">
                                      {f.display_value || f.value || '—'}
                                    </span>
                                  </div>
                                ))}
                              </div>
                            </div>
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
