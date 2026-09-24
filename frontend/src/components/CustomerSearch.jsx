import React, { useState, useEffect, useRef } from 'react';
import { Search, X, Loader2, User } from 'lucide-react';

export default function CustomerSearch({
  query = '',
  onQueryChange,
  suggestions = [],
  loading = false,
  onSelectCustomer,
}) {
  const [isOpen, setIsOpen] = useState(false);
  const [activeIndex, setActiveIndex] = useState(-1);
  const wrapperRef = useRef(null);
  const inputRef = useRef(null);

  // Close dropdown on click outside
  useEffect(() => {
    const handleClickOutside = (e) => {
      if (wrapperRef.current && !wrapperRef.current.contains(e.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // When query changes and has text, open dropdown
  useEffect(() => {
    if (query.trim().length > 0) {
      setIsOpen(true);
      setActiveIndex(-1);
    } else {
      setIsOpen(false);
      setActiveIndex(-1);
    }
  }, [query]);

  const handleClear = () => {
    if (onQueryChange) onQueryChange('');
    setIsOpen(false);
    setActiveIndex(-1);
    if (inputRef.current) inputRef.current.focus();
  };

  const handleSelect = (customerName) => {
    setIsOpen(false);
    setActiveIndex(-1);
    if (onSelectCustomer) {
      onSelectCustomer(customerName);
    }
  };

  const handleKeyDown = (e) => {
    if (!isOpen || suggestions.length === 0) return;

    if (e.key === 'ArrowDown') {
      e.preventDefault();
      setActiveIndex((prev) => (prev + 1 < suggestions.length ? prev + 1 : 0));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setActiveIndex((prev) => (prev > 0 ? prev - 1 : suggestions.length - 1));
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (activeIndex >= 0 && activeIndex < suggestions.length) {
        handleSelect(suggestions[activeIndex].name);
      } else if (suggestions.length > 0) {
        handleSelect(suggestions[0].name);
      }
    } else if (e.key === 'Escape') {
      setIsOpen(false);
    }
  };

  return (
    <div className="search-section" ref={wrapperRef}>
      <div className="search-bar-wrapper">
        <div className="search-input-box">
          <Search size={18} className="search-icon text-muted flex-shrink-0" />
          <input
            ref={inputRef}
            id="customer-search-input"
            type="text"
            className="search-input"
            placeholder="🔍 Search customer name across all sheets..."
            value={query}
            onChange={(e) => onQueryChange && onQueryChange(e.target.value)}
            onFocus={() => {
              if (query.trim().length > 0 && suggestions.length > 0) {
                setIsOpen(true);
              }
            }}
            onKeyDown={handleKeyDown}
            autoComplete="off"
            title="Search is restricted to customer Name only"
          />

          <div className="search-input-actions">
            {loading && <Loader2 size={16} className="spin-icon text-emerald" />}
            {query && !loading && (
              <button
                type="button"
                className="search-clear-btn"
                onClick={handleClear}
                title="Clear search and return to customer list"
                id="btn-search-clear-x"
              >
                <X size={16} />
              </button>
            )}
          </div>
        </div>

        {/* Unique Name Suggestions Dropdown (Requirements 5, 6, 23, 24) */}
        {isOpen && query.trim().length > 0 && (
          <div className="suggestions-dropdown glass-panel animate-fade-in">
            <div className="suggestions-header">
              <span>Search Results</span>
              <span>
                {suggestions.length} {suggestions.length === 1 ? 'name' : 'names'}
              </span>
            </div>

            {loading ? (
              <div className="p-4 text-center text-muted text-xs">
                <Loader2 size={14} className="spin-icon inline mr-2 text-emerald" />
                Searching unique customer names...
              </div>
            ) : suggestions.length === 0 ? (
              <div className="suggestions-empty">
                <p className="text-secondary text-sm">No customers matching "{query}"</p>
              </div>
            ) : (
              suggestions.map((item, idx) => (
                <div
                  key={item.normalized_name || item.name}
                  className={`suggestion-item ${idx === activeIndex ? 'suggestion-highlighted' : ''}`}
                  onClick={() => handleSelect(item.name)}
                  onMouseEnter={() => setActiveIndex(idx)}
                >
                  <div className="suggestion-avatar">
                    <User size={15} />
                  </div>
                  <div className="suggestion-info">
                    <div className="suggestion-main-line">
                      <span className="suggestion-name">{item.name}</span>
                    </div>
                    <div className="suggestion-meta-line">
                      <span className="suggestion-sub-tag text-emerald">
                        {item.record_count} {item.record_count === 1 ? 'record' : 'records'}
                      </span>
                    </div>
                  </div>
                </div>
              ))
            )}
          </div>
        )}
      </div>
    </div>
  );
}
