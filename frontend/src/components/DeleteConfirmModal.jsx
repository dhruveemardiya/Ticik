import React, { useState } from 'react';
import { AlertTriangle, Trash2, Loader2, Train, Layers, Hash, Users, User } from 'lucide-react';

export default function DeleteConfirmModal({
  isOpen,
  onClose,
  mode = 'record', // 'record' | 'customer' | 'all'
  record = null,
  customer = null,
  systemStatus = null,
  onDeleted
}) {
  const [deleting, setDeleting] = useState(false);
  const [error, setError] = useState(null);

  if (!isOpen) return null;

  const handleDelete = async () => {
    setDeleting(true);
    setError(null);
    try {
      let endpoint = '';
      if (mode === 'record') {
        if (!record || !record.record_id) throw new Error('No record specified to delete');
        endpoint = `/api/customers/record/?record_id=${encodeURIComponent(record.record_id)}`;
      } else if (mode === 'customer') {
        if (!customer || !customer.name) throw new Error('No customer specified to delete');
        endpoint = `/api/customers/customer/?name=${encodeURIComponent(customer.name)}`;
      } else if (mode === 'all') {
        endpoint = '/api/customers/all/';
      }

      const res = await fetch(endpoint, {
        method: 'DELETE'
      });
      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Failed to complete delete action');
      }

      if (onDeleted) {
        onDeleted({
          mode,
          record,
          customer,
          data
        });
      }
      onClose();
    } catch (err) {
      setError(err.message || 'Error executing delete');
    } finally {
      setDeleting(false);
    }
  };

  return (
    <div className="modal-backdrop animate-fade-in" onClick={onClose}>
      <div className="modal-card glass-panel" onClick={(e) => e.stopPropagation()}>
        <div className="modal-icon-wrap delete-icon-wrap">
          <AlertTriangle size={32} className="text-rose" />
        </div>

        {mode === 'record' && (
          <>
            <h3 className="modal-title">Delete this Excel Record?</h3>
            {record && (
              <div className="delete-target-card">
                <div className="delete-target-row mb-1">
                  <span className="text-secondary text-xs">Customer:</span>
                  <strong className="delete-target-name ml-2">{record.name}</strong>
                </div>
                <div className="delete-target-meta">
                  <span className="meta-tag">
                    <Layers size={12} className="text-cyan" /> Sheet: {record.sheet_name}
                  </span>
                  <span className="meta-tag">
                    <Hash size={12} className="text-muted" /> Excel Row: {record.row_number}
                  </span>
                  {record.train_no && (
                    <span className="meta-tag">
                      <Train size={12} className="text-sky" /> Train: {record.train_no}
                    </span>
                  )}
                </div>
              </div>
            )}
            <p className="modal-desc delete-warning-text">
              This will <strong>permanently remove this one Excel row</strong> from the saved Excel workbook.
              All other records and sheets will remain untouched.
            </p>
          </>
        )}

        {mode === 'customer' && (
          <>
            <h3 className="modal-title">Delete Customer and All Records?</h3>
            {customer && (
              <div className="delete-target-card">
                <div className="delete-target-row mb-1">
                  <User size={14} className="text-rose mr-1.5 flex-shrink-0" />
                  <strong className="delete-target-name">{customer.name}</strong>
                </div>
                <div className="delete-target-meta mt-1">
                  <span className="meta-tag">
                    <Layers size={12} className="text-cyan" /> {customer.record_count || customer.records?.length || 0} Records across all sheets
                  </span>
                </div>
              </div>
            )}
            <p className="modal-desc delete-warning-text">
              This will <strong>permanently delete all {customer?.record_count || 'matching'} physical Excel records</strong> for this customer across all sheets. This action cannot be undone.
            </p>
          </>
        )}

        {mode === 'all' && (
          <>
            <h3 className="modal-title">Delete ALL Customers from Excel?</h3>
            <div className="delete-target-card">
              <div className="delete-target-row mb-1">
                <Users size={16} className="text-rose mr-1.5 flex-shrink-0" />
                <strong className="delete-target-name">Entire Excel Workbook Database</strong>
              </div>
              <div className="delete-target-meta mt-1">
                <span className="meta-tag">
                  {systemStatus?.customer_count || 0} Unique Customers
                </span>
                <span className="meta-tag">
                  {systemStatus?.total_records || 0} Total Records
                </span>
                <span className="meta-tag">
                  {systemStatus?.sheets?.length || 0} Sheets
                </span>
              </div>
            </div>
            <p className="modal-desc delete-warning-text">
              This will <strong>permanently wipe ALL customer records</strong> from the Excel workbook on disk. All sheet columns and structure will be preserved. This action is irreversible.
            </p>
          </>
        )}

        {error && (
          <div className="alert-error text-sm mb-3">
            {error}
          </div>
        )}

        <div className="modal-actions">
          <button
            type="button"
            className="btn btn-secondary"
            onClick={onClose}
            disabled={deleting}
          >
            Cancel
          </button>
          <button
            type="button"
            className="btn btn-outline-danger"
            onClick={handleDelete}
            disabled={deleting}
            id={
              mode === 'record'
                ? 'btn-confirm-delete-record'
                : mode === 'customer'
                ? 'btn-confirm-delete-customer'
                : 'btn-confirm-delete-all'
            }
          >
            {deleting ? (
              <>
                <Loader2 size={16} className="spin-icon mr-1" /> Deleting from Excel...
              </>
            ) : (
              <>
                <Trash2 size={16} className="mr-1" />
                {mode === 'record' && 'Delete Record'}
                {mode === 'customer' && 'Delete Customer'}
                {mode === 'all' && 'Delete All Customers'}
              </>
            )}
          </button>
        </div>
      </div>
    </div>
  );
}
