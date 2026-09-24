import React, { useState } from 'react';
import { AlertTriangle, Loader2 } from 'lucide-react';

export default function ReplaceModal({ isOpen, onClose, onConfirm }) {
  const [replacing, setReplacing] = useState(false);

  if (!isOpen) return null;

  const handleConfirm = async () => {
    setReplacing(true);
    try {
      await onConfirm();
    } finally {
      setReplacing(false);
    }
  };

  return (
    <div className="modal-backdrop animate-fade-in" onClick={onClose}>
      <div className="modal-card glass-panel" onClick={(e) => e.stopPropagation()}>
        <div className="modal-icon-wrap">
          <AlertTriangle size={32} className="text-amber" />
        </div>

        <h3 className="modal-title">Replace Customer Data?</h3>
        <p className="modal-desc">
          Existing customer data will be replaced. You will need to upload a new Excel file.
          <br /><br />
          <strong>Are you sure?</strong>
        </p>

        <div className="modal-actions">
          <button
            type="button"
            className="btn btn-secondary"
            onClick={onClose}
            disabled={replacing}
          >
            Cancel
          </button>
          <button
            type="button"
            className="btn btn-outline-danger"
            onClick={handleConfirm}
            disabled={replacing}
            id="btn-confirm-replace"
          >
            {replacing ? (
              <>
                <Loader2 size={16} className="spin-icon" /> Replacing...
              </>
            ) : (
              'Replace'
            )}
          </button>
        </div>
      </div>
    </div>
  );
}
