import React from 'react';
import { AlertCircle, RefreshCw } from 'lucide-react';

export default class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error) {
    return { hasError: true, error };
  }

  componentDidCatch(error, errorInfo) {
    console.error('ErrorBoundary caught an error:', error, errorInfo);
  }

  handleReset = () => {
    this.setState({ hasError: false, error: null });
  };

  render() {
    if (this.state.hasError) {
      return (
        <div className="details-container details-container-light text-center p-8 animate-fade-in">
          <div className="empty-state-card-visual mb-3">
            <AlertCircle size={38} className="text-rose" />
          </div>
          <h3 className="empty-state-title text-rose">Something went wrong</h3>
          <p className="empty-state-text text-secondary mb-4">
            An unexpected error occurred while rendering customer details.
          </p>
          <button
            type="button"
            className="btn btn-secondary btn-sm"
            onClick={this.handleReset}
          >
            <RefreshCw size={14} className="mr-1.5" /> Try Again
          </button>
        </div>
      );
    }

    return this.props.children;
  }
}
