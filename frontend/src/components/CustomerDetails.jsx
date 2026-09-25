import React, { useState, useMemo } from 'react';
import {
  Layers,
  Hash,
  Trash2,
  Copy,
  Check,
  UserCheck,
  Search,
  FileSpreadsheet,
  User,
  Ticket,
  ArrowLeft
} from 'lucide-react';
import DeleteConfirmModal from './DeleteConfirmModal';

export default function CustomerDetails({
  customerData,
  loading = false,
  onRecordDeleted,
  onCustomerDeleted
}) {
  const [recordToDelete, setRecordToDelete] = useState(null);
  const [customerToDelete, setCustomerToDelete] = useState(null);
  const [copiedRecordId, setCopiedRecordId] = useState(null);
  const [copiedAll, setCopiedAll] = useState(false);

  // Helper: Format Time (handles 11.25 -> 11:25, 4.5 -> 04:30, or '26/9/2026- 4:50' -> 04:50)
  const formatTime = (val) => {
    if (!val) return '';
    const s = String(val).trim();
    if (['DEP', 'DEP (T)', 'DEPARTURE', 'ARR', 'ARR (T)', 'ARRIVAL', '—', '-', '--', 'N/A'].includes(s.toUpperCase())) {
      return '';
    }
    // Handle decimal times like 11.25 -> 11:25, 4.5 -> 04:30, 5.05 -> 05:05
    if (/^\d{1,2}\.\d{1,2}$/.test(s)) {
      const [h, m] = s.split('.');
      const hours = h.padStart(2, '0');
      const mins = m.length === 1 ? (parseInt(m, 10) * 6).toString().padStart(2, '0') : m.padEnd(2, '0');
      return `${hours}:${mins}`;
    }
    // Extract time from string like '26/9/2026- 4:50' or '2026-09-26 12:20:00'
    const matchTime = s.match(/(?:^|[^\d])(\d{1,2}:\d{2})(?::\d{2})?/);
    if (matchTime) {
      return matchTime[1];
    }
    return s;
  };

  // Helper: Format Date of Journey
  const formatDOJ = (val) => {
    if (!val) return '';
    const s = String(val).trim();
    if (['DOJ', 'DATE OF JOURNEY', 'TRAVEL DATE', '—', '-', '--', 'N/A'].includes(s.toUpperCase())) {
      return '';
    }
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
      try {
        const parts = s.split(' ')[0].split('-');
        const d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
        if (!isNaN(d.getTime())) {
          return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
        }
      } catch (e) {}
    }
    return s;
  };

  // Core Data Processing:
  // COMMON INFO (show once): Customer Name, Age, Sex/Gender, Aadhaar No.
  // RECORD INFO (every record): DOJ, DEP, ARR, TICKET NO., SEAT NO/STATUS
  // HIDE: Price / RS. / Amount, SC, Ticket Serial No.
  const { commonFields, recordsWithSpecificFields } = useMemo(() => {
    if (!customerData || !customerData.records || customerData.records.length === 0) {
      return { commonFields: [], recordsWithSpecificFields: [] };
    }

    const records = customerData.records;
    const customerName = (customerData.name || '').trim();

    const isAgeCol = (k, l) => k === 'AGE' || l === 'AGE';
    const isGenderCol = (k, l) => k === 'SEX' || k === 'GENDER' || l === 'SEX' || l === 'GENDER';
    const isAadhaarCol = (k, l) =>
      k.includes('ADHAAR') || k.includes('AADHAAR') || k.includes('AADHAR') ||
      l.includes('ADHAAR') || l.includes('AADHAAR') || l.includes('AADHAR');

    // 1. Extract Common Age
    let commonAge = '';
    for (const rec of records) {
      for (const f of rec.fields || []) {
        const k = String(f.key || '').trim().toUpperCase();
        const l = String(f.label || '').trim().toUpperCase();
        if (isAgeCol(k, l)) {
          const val = String(f.display_value || f.value || '').trim();
          if (val && !['—', '---', '--', '-', 'N/A', 'AGE'].includes(val.toUpperCase())) {
            commonAge = val;
            break;
          }
        }
      }
      if (commonAge) break;
    }

    // 2. Extract Common Gender/Sex
    let commonGender = '';
    for (const rec of records) {
      for (const f of rec.fields || []) {
        const k = String(f.key || '').trim().toUpperCase();
        const l = String(f.label || '').trim().toUpperCase();
        if (isGenderCol(k, l)) {
          const val = String(f.display_value || f.value || '').trim();
          if (val && !['—', '---', '--', '-', 'N/A', 'SEX', 'GENDER'].includes(val.toUpperCase())) {
            commonGender = val;
            break;
          }
        }
      }
      if (commonGender) break;
    }

    // 3. Extract Common Aadhaar No.
    let commonAadhaar = '';
    for (const rec of records) {
      for (const f of rec.fields || []) {
        const k = String(f.key || '').trim().toUpperCase();
        const l = String(f.label || '').trim().toUpperCase();
        if (isAadhaarCol(k, l)) {
          const val = String(f.display_value || f.value || '').trim();
          if (val && !['—', '---', '--', '-', 'N/A'].includes(val.toUpperCase()) && !isAadhaarCol(val, val)) {
            commonAadhaar = val;
            break;
          }
        }
      }
      if (commonAadhaar) break;
    }

    // COMMON INFO — show only once: Customer Name, Age, Sex/Gender, Aadhaar No.
    const commonList = [
      { key: 'CUSTOMER_NAME', label: 'Customer Name', value: customerName || '—' },
      { key: 'AGE', label: 'Age', value: commonAge || '—' },
      { key: 'GENDER', label: 'Sex/Gender', value: commonGender || '—' },
      { key: 'AADHAAR', label: 'Aadhaar No.', value: commonAadhaar || '—' }
    ];

    // Helper: Resolve DEP (Departure Time + Location)
    const resolveDep = (rec) => {
      let depTimeRaw = '';
      for (const f of rec.fields || []) {
        const k = String(f.key || '').trim().toUpperCase();
        const l = String(f.label || '').trim().toUpperCase();
        if (['DEP (T)', 'DEP', 'DEPARTURE'].includes(k) || ['DEPARTURE', 'DEP'].includes(l)) {
          const v = String(f.display_value || f.value || '').trim();
          if (v && !['DEP (T)', 'DEPARTURE', 'DEP', '—', '-', '--', 'N/A'].includes(v.toUpperCase())) {
            depTimeRaw = v;
            break;
          }
        }
      }

      const depTime = formatTime(depTimeRaw);
      let depLoc = '';

      if (rec.departure && !rec.departure.toLowerCase().includes('information not available')) {
        depLoc = rec.departure.trim();
      } else if (rec.train_route && rec.train_route.includes('→')) {
        const p0 = rec.train_route.split('→')[0].trim();
        if (p0 && !p0.toUpperCase().includes('DEP')) {
          depLoc = p0;
        }
      }

      // If depTime is non-numeric station text
      if (depTime && !/\d/.test(depTime)) {
        depLoc = depTime;
        return depLoc;
      }

      if (depTime && depLoc && !depTime.toLowerCase().includes(depLoc.toLowerCase())) {
        return `${depTime} • ${depLoc}`;
      }
      return depTime || depLoc || '—';
    };

    // Helper: Resolve ARR (Arrival Time + Location)
    const resolveArr = (rec) => {
      let arrTimeRaw = '';
      for (const f of rec.fields || []) {
        const k = String(f.key || '').trim().toUpperCase();
        const l = String(f.label || '').trim().toUpperCase();
        if (['ARR (T)', 'ARR', 'ARRIVAL'].includes(k) || ['ARRIVAL', 'ARR'].includes(l)) {
          const v = String(f.display_value || f.value || '').trim();
          if (v && !['ARR (T)', 'ARRIVAL', 'ARR', '—', '-', '--', 'N/A'].includes(v.toUpperCase())) {
            arrTimeRaw = v;
            break;
          }
        }
      }

      const arrTime = formatTime(arrTimeRaw);
      let arrLoc = '';

      if (rec.arrival && !rec.arrival.toLowerCase().includes('information not available')) {
        arrLoc = rec.arrival.trim();
      } else if (rec.train_route && rec.train_route.includes('→')) {
        const p1 = rec.train_route.split('→')[1].trim();
        if (p1 && !p1.toUpperCase().includes('ARR')) {
          arrLoc = p1;
        }
      }

      // If arrTime is non-numeric station text
      if (arrTime && !/\d/.test(arrTime)) {
        arrLoc = arrTime;
        return arrLoc;
      }

      if (arrTime && arrLoc && !arrTime.toLowerCase().includes(arrLoc.toLowerCase())) {
        return `${arrTime} • ${arrLoc}`;
      }
      return arrTime || arrLoc || '—';
    };

    // Helper: Resolve TICKET NO.
    const resolveTicketNo = (rec) => {
      const candidates = [];
      for (const f of rec.fields || []) {
        const k = String(f.key || '').trim().toUpperCase();
        const l = String(f.label || '').trim().toUpperCase();
        if (
          k.includes('TICKET') || k.includes('TCKT') || k.includes('TKT') || k.includes('PNR') ||
          l.includes('TICKET') || l.includes('PNR')
        ) {
          // Exclude Serial Number
          if (k.includes('SR') || l.includes('SERIAL') || l.includes('SR')) {
            continue;
          }
          const v = String(f.display_value || f.value || '').trim();
          if (v && v !== '—' && !v.toUpperCase().includes('TICKET') && !v.toUpperCase().includes('TKT')) {
            candidates.push(v);
          }
        }
      }

      // Prefer real 5-12 digit Ticket/PNR number (e.g. 15917481) over small counts like 2
      const longTicket = candidates.find((c) => /^\d{5,12}$/.test(c.replace(/\s+/g, '')));
      if (longTicket) return longTicket;

      if (candidates.length > 0) return candidates[candidates.length - 1];
      return '—';
    };

    // Helper: Resolve SEAT NO/STATUS
    const resolveSeatStatus = (rec) => {
      for (const f of rec.fields || []) {
        const k = String(f.key || '').trim().toUpperCase();
        const l = String(f.label || '').trim().toUpperCase();
        if (k.includes('SEAT') || k.includes('STATUS') || l.includes('SEAT') || l.includes('STATUS')) {
          const v = String(f.display_value || f.value || '').trim();
          if (v && v !== '—' && !v.toUpperCase().includes('SEAT') && !v.toUpperCase().includes('STATUS')) {
            return v;
          }
        }
      }
      return '—';
    };

    // Helper: Resolve DOJ
    const resolveDOJ = (rec) => {
      for (const f of rec.fields || []) {
        const k = String(f.key || '').trim().toUpperCase();
        const l = String(f.label || '').trim().toUpperCase();
        if (k.includes('DOJ') || k.includes('DATE') || l.includes('DATE') || l.includes('DOJ')) {
          const v = String(f.display_value || f.value || '').trim();
          if (v && v !== '—' && !['DOJ', 'DATE OF JOURNEY', 'TRAVEL DATE'].includes(v.toUpperCase())) {
            return formatDOJ(v);
          }
        }
      }
      if (rec.doj) return formatDOJ(rec.doj);
      return '—';
    };

    // RECORD INFO — For every Excel record show ONLY:
    // 1. DOJ — Date of Journey
    // 2. DEP — Departure Time + Location
    // 3. ARR — Arrival Time + Location
    // 4. TICKET NO. — Ticket Number
    // 5. SEAT NO/STATUS — Seat + Status
    // Hide completely: Price/RS/Amount, SC, Ticket Serial No.
    const recordsSpecific = records.map((rec) => {
      const doj = resolveDOJ(rec);
      const dep = resolveDep(rec);
      const arr = resolveArr(rec);
      const ticketNo = resolveTicketNo(rec);
      const seatStatus = resolveSeatStatus(rec);

      const cleanRecordFields = [
        { key: 'DOJ', label: 'DOJ', sublabel: 'Date of Journey', value: doj },
        { key: 'DEP', label: 'DEP', sublabel: 'Departure Time + Location', value: dep },
        { key: 'ARR', label: 'ARR', sublabel: 'Arrival Time + Location', value: arr },
        { key: 'TICKET_NO', label: 'TICKET NO.', sublabel: 'Ticket Number', value: ticketNo },
        { key: 'SEAT_STATUS', label: 'SEAT NO/STATUS', sublabel: 'Seat + Status', value: seatStatus }
      ];

      return {
        ...rec,
        cleanRecordFields
      };
    });

    return {
      commonFields: commonList,
      recordsWithSpecificFields: recordsSpecific
    };
  }, [customerData]);

  // 1. Loading State
  if (loading) {
    return (
      <div className="details-container details-container-light animate-fade-in text-center p-8">
        <div className="details-loading">
          <div className="status-dot status-dot-active mb-3" />
          <p className="text-secondary font-medium">Loading customer records from Excel...</p>
        </div>
      </div>
    );
  }

  // 2. Clean Empty State (When NO customer is selected)
  if (!customerData || !customerData.records || customerData.records.length === 0) {
    return (
      <div className="details-empty-state details-empty-state-light animate-fade-in text-center">
        <div className="empty-state-card-visual">
          <div className="empty-state-ticket-badge">
            <Ticket size={38} className="empty-state-ticket-icon" />
          </div>
          <div className="empty-state-pulse-ring" />
        </div>
        <h3 className="empty-state-title">Select a Customer</h3>
        <p className="empty-state-text">
          Choose a customer from the list to view their records.
        </p>
        <div className="empty-state-hint-pill">
          <ArrowLeft size={13} className="text-sky mr-1.5 flex-shrink-0 animate-bounce-x" />
          <span>Select from Customer Directory on the left</span>
        </div>
      </div>
    );
  }

  const { name, record_count } = customerData;

  const handleCopyRecord = (rec) => {
    const lines = [
      `Customer: ${name}`,
      `Sheet: ${rec.sheet_name}, Excel Row: ${rec.row_number}`,
      `--- COMMON INFO ---`,
      ...commonFields.map((f) => `${f.label}: ${f.value}`),
      `--- RECORD INFO ---`,
      ...(rec.cleanRecordFields || []).map((f) => `${f.label}: ${f.value}`)
    ];

    const textToCopy = lines.join('\n');
    navigator.clipboard.writeText(textToCopy).then(() => {
      setCopiedRecordId(rec.record_id);
      setTimeout(() => setCopiedRecordId(null), 2000);
    });
  };

  const handleCopyAll = () => {
    const lines = [
      `CUSTOMER: ${name} (${record_count} Records)`,
      '--- COMMON INFO ---',
      ...commonFields.map((f) => `${f.label}: ${f.value}`),
      '--- RECORD INFO ---'
    ];
    recordsWithSpecificFields.forEach((rec, idx) => {
      lines.push(`[Record ${idx + 1}] Sheet: ${rec.sheet_name}, Row: ${rec.row_number}`);
      (rec.cleanRecordFields || []).forEach((f) => {
        lines.push(`  ${f.label}: ${f.value}`);
      });
    });

    navigator.clipboard.writeText(lines.join('\n')).then(() => {
      setCopiedAll(true);
      setTimeout(() => setCopiedAll(false), 2000);
    });
  };

  return (
    <>
      <div className="details-container details-container-light animate-details-enter">
        {/* Main Customer Header (Requirements 8, 14) */}
        <div className="details-header">
          <div className="details-header-main">
            <span className="details-kicker">CUSTOMER DETAILS</span>
            <h2 className="details-customer-name">{name}</h2>
          </div>
          <div className="details-header-right">
            <div className="details-record-count-badge">
              <FileSpreadsheet size={15} className="mr-1.5 text-sky inline flex-shrink-0" />
              <span className="font-semibold text-sky">
                {record_count} {record_count === 1 ? 'Excel Record' : 'Excel Records'}
              </span>
            </div>
            <button
              type="button"
              className="btn btn-secondary btn-xs ml-2"
              onClick={handleCopyAll}
              title="Copy all customer records to clipboard"
            >
              {copiedAll ? (
                <>
                  <Check size={12} className="text-sky mr-1" /> Copied All
                </>
              ) : (
                <>
                  <Copy size={12} className="mr-1" /> Copy All
                </>
              )}
            </button>
            <button
              type="button"
              className="btn btn-danger-soft btn-xs ml-2"
              onClick={() => setCustomerToDelete(customerData)}
              title={`Permanently delete all ${record_count} records for ${name} across all sheets`}
              id="btn-delete-customer"
            >
              <Trash2 size={12} className="mr-1 text-rose" /> Delete Customer
            </button>
          </div>
        </div>

        {/* Section 1: COMMON INFO — Show only once (Customer Name, Age, Sex/Gender, Aadhaar No.) */}
        {commonFields.length > 0 && (
          <div className="customer-info-section">
            <div className="section-label-bar">
              <div className="section-title-wrap">
                <User size={14} className="text-sky mr-1.5 flex-shrink-0" />
                <span className="section-title">COMMON INFO</span>
              </div>
              <span className="section-subtitle">
                Customer identity across all records
              </span>
            </div>

            <div className="customer-info-card">
              <div className="customer-info-grid">
                {commonFields.map((field) => (
                  <div key={field.key} className="info-cell">
                    <span className="info-cell-label">{field.label}</span>
                    <span className="info-cell-value">{field.value}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* Section 2: RECORD INFO — Only DOJ, DEP, ARR, TICKET NO., SEAT NO/STATUS */}
        <div className="excel-records-section">
          <div className="section-label-bar">
            <div className="section-title-wrap">
              <FileSpreadsheet size={14} className="text-sky mr-1.5 flex-shrink-0" />
              <span className="section-title">RECORD INFO</span>
            </div>
            <span className="section-subtitle">
              {recordsWithSpecificFields.length} individual Excel rows
            </span>
          </div>

          <div className="details-records-stack">
            {recordsWithSpecificFields.map((rec, idx) => {
              const isCopied = copiedRecordId === rec.record_id;

              return (
                <div key={rec.record_id} className="record-card">
                  {/* Record Header */}
                  <div className="record-card-header">
                    <div className="record-header-left">
                      <span className="record-badge-index">
                        RECORD {String(idx + 1).padStart(2, '0')}
                      </span>
                      {rec.file_name && (
                        <span className="record-source-tag record-file-tag" title={`Source File: ${rec.file_name}`}>
                          <FileSpreadsheet size={12} className="inline mr-1 text-sky" /> {rec.file_name}
                        </span>
                      )}
                      <span className="record-source-tag">
                        <Layers size={12} className="inline mr-1 text-sky" /> {rec.sheet_name}
                      </span>
                      <span className="record-source-tag">
                        <Hash size={12} className="inline mr-1 text-muted" /> Row {rec.row_number}
                      </span>
                      {rec.train_no && (
                        <span className="record-source-tag font-semibold text-sky">
                          Train {rec.train_no}
                        </span>
                      )}
                    </div>

                    <div className="record-header-actions">
                      <button
                        type="button"
                        className="btn btn-secondary btn-xs"
                        onClick={() => handleCopyRecord(rec)}
                        title="Copy this record's details to clipboard"
                      >
                        {isCopied ? (
                          <>
                            <Check size={12} className="text-sky mr-1" /> Copied
                          </>
                        ) : (
                          <>
                            <Copy size={12} className="mr-1" /> Copy Record
                          </>
                        )}
                      </button>
                      <button
                        type="button"
                        className="btn btn-danger-soft btn-xs"
                        onClick={() => setRecordToDelete(rec)}
                        title={`Delete Row ${rec.row_number} in ${rec.sheet_name}`}
                      >
                        <Trash2 size={12} className="mr-1 text-rose" /> Delete Record
                      </button>
                    </div>
                  </div>

                  {/* Clean 5-Field Grid: DOJ, DEP, ARR, TICKET NO., SEAT NO/STATUS */}
                  <div className="record-fields-grid record-fields-grid-5">
                    {(rec.cleanRecordFields || []).map((field) => (
                      <div key={field.key} className="record-field-cell">
                        <span className="field-label font-bold text-cyan">{field.label}</span>
                        <span className="field-sublabel text-xs text-muted block mb-0.5">{field.sublabel}</span>
                        <div className="field-value font-medium text-primary">
                          {field.value && field.value !== '—' ? (
                            <span className="value-text">{field.value}</span>
                          ) : (
                            <span className="value-empty">—</span>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* 1. Delete Record Modal (Requirement 4: Delete Record) */}
      <DeleteConfirmModal
        isOpen={!!recordToDelete}
        onClose={() => setRecordToDelete(null)}
        mode="record"
        record={recordToDelete}
        onDeleted={(payload) => {
          setRecordToDelete(null);
          if (onRecordDeleted) {
            onRecordDeleted(payload.record?.record_id, payload.data?.data?.new_customer_count);
          }
        }}
      />

      {/* 2. Delete Customer Modal (Requirement 4: Delete Customer) */}
      <DeleteConfirmModal
        isOpen={!!customerToDelete}
        onClose={() => setCustomerToDelete(null)}
        mode="customer"
        customer={customerToDelete}
        onDeleted={(payload) => {
          const deletedName = customerToDelete?.name;
          setCustomerToDelete(null);
          if (onCustomerDeleted) {
            onCustomerDeleted(deletedName, payload.data?.data?.new_customer_count);
          }
        }}
      />
    </>
  );
}
