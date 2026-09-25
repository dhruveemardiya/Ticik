import React from 'react';
import { Users, User, ChevronRight, ArrowLeft } from 'lucide-react';

export default function CustomerListPanel({
  customers = [],
  selectedCustomerName,
  onSelectCustomer,
  loading = false,
  onBackMobile
}) {
  return (
    <div className="customer-list-panel customer-list-panel-dark">
      {/* Panel Header */}
      <div className="panel-header">
        <div className="panel-title-wrap">
          {onBackMobile && (
            <button
              type="button"
              className="btn btn-secondary btn-xs btn-mobile-back"
              onClick={onBackMobile}
              title="Back to Details"
              aria-label="Back"
            >
              <ArrowLeft size={14} className="text-sky mr-1" />
              <span>Back</span>
            </button>
          )}
          <Users size={18} className="text-sky" />
          <h3 className="panel-title">Customer Directory</h3>
          <span className="badge badge-neutral ml-auto">
            {customers.length.toLocaleString()} {customers.length === 1 ? 'Customer' : 'Customers'}
          </span>
        </div>
      </div>

      {/* Unique Customer Cards List */}
      <div className="panel-list-container">
        {loading ? (
          <div className="panel-loading">
            <div className="status-dot status-dot-active mb-2" />
            <span className="text-secondary text-xs">Loading customer directory...</span>
          </div>
        ) : customers.length === 0 ? (
          <div className="panel-empty">
            <User size={28} className="text-dim mb-2" />
            <p className="text-secondary text-sm">No customers in this view</p>
          </div>
        ) : (
          <div className="panel-cards-stack">
            {customers.map((c, idx) => {
              const isSelected = selectedCustomerName &&
                selectedCustomerName.trim().toLowerCase() === c.name.trim().toLowerCase();

              return (
                <div
                  key={c.normalized_name || c.name}
                  className={`panel-customer-card ${isSelected ? 'panel-card-selected' : ''}`}
                  onClick={() => {
                    onSelectCustomer(c.name);
                    if (onBackMobile) {
                      onBackMobile();
                    }
                  }}
                  title={`Click to view all ${c.record_count} Excel records for ${c.name}`}
                >
                  <div className="panel-card-header">
                    <span className="panel-card-index">{String(idx + 1).padStart(2, '0')}</span>
                    <h5 className="panel-card-name">{c.name}</h5>
                    <div className="panel-card-arrow">
                      <ChevronRight size={16} className={isSelected ? 'text-sky' : 'text-dim'} />
                    </div>
                  </div>

                  <div className="panel-card-body">
                    <div className="panel-card-count-badge">
                      <span className="text-secondary text-xs">
                        {c.record_count} {c.record_count === 1 ? 'record' : 'records'}
                      </span>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}
