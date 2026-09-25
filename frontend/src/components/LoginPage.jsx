import React, { useState } from 'react';
import {
  Train,
  Lock,
  Eye,
  EyeOff,
  ArrowRight,
  AlertCircle,
  Loader2,
  ShieldCheck
} from 'lucide-react';

export default function LoginPage({ onLoginSuccess }) {
  // Login form state
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [loginLoading, setLoginLoading] = useState(false);
  const [loginError, setLoginError] = useState('');

  // Handle Login submission
  const handleLoginSubmit = async (e) => {
    e.preventDefault();
    if (!password) {
      setLoginError('Please enter your password.');
      return;
    }

    setLoginLoading(true);
    setLoginError('');

    try {
      const res = await fetch('/api/customers/auth/login/', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ password })
      });

      const data = await res.json();
      if (res.ok && data.success) {
        if (data.token) {
          localStorage.setItem('train_auth_token', data.token);
        }
        onLoginSuccess();
      } else {
        setLoginError(data.error || 'Incorrect password. Please try again.');
      }
    } catch (err) {
      setLoginError('Network connection error. Please try again.');
    } finally {
      setLoginLoading(false);
    }
  };

  return (
    <div className="login-viewport">
      {/* Background layer with subtle train visual */}
      <div className="login-bg-media" />
      <div className="login-bg-overlay" />

      {/* Login Card */}
      <div className="login-card-container animate-fade-in">
        <div className="login-card glass-panel">
          {/* Card Header */}
          <div className="login-card-header">
            <div className="login-logo-glow">
              <div className="login-logo-badge">
                <Train size={28} className="text-sky" />
              </div>
            </div>
            <h1 className="login-title">Rail Passenger Portal</h1>
            <p className="login-subtitle">
              Enter password to access customer & journey records
            </p>
          </div>

          {/* LOGIN FORM */}
          <form onSubmit={handleLoginSubmit} className="login-form">
            {loginError && (
              <div className="login-alert login-alert-error animate-shake">
                <AlertCircle size={16} className="flex-shrink-0" />
                <span>{loginError}</span>
              </div>
            )}

            <div className="login-field-group">
              <label className="login-field-label" htmlFor="login-password-input">
                Password
              </label>
              <div className="login-input-wrapper">
                <Lock size={16} className="login-input-icon text-muted" />
                <input
                  id="login-password-input"
                  type={showPassword ? 'text' : 'password'}
                  className="login-input"
                  placeholder="Enter system password"
                  value={password}
                  onChange={(e) => {
                    setPassword(e.target.value);
                    if (loginError) setLoginError('');
                  }}
                  autoFocus
                  autoComplete="current-password"
                />
                <button
                  type="button"
                  className="login-eye-btn"
                  onClick={() => setShowPassword(!showPassword)}
                  tabIndex={-1}
                  title={showPassword ? 'Hide password' : 'Show password'}
                >
                  {showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
                </button>
              </div>
            </div>

            <button
              type="submit"
              className="btn btn-primary login-submit-btn"
              disabled={loginLoading}
              id="btn-login-submit"
            >
              {loginLoading ? (
                <>
                  <Loader2 size={16} className="animate-spin mr-2" />
                  <span>Verifying...</span>
                </>
              ) : (
                <>
                  <span>Enter Portal</span>
                  <ArrowRight size={16} className="ml-2" />
                </>
              )}
            </button>
          </form>

          <div className="login-card-foot-badge">
            <span className="login-security-tag">
              <ShieldCheck size={12} className="inline mr-1 text-sky" /> Protected Rail Terminal
            </span>
          </div>
        </div>
      </div>
    </div>
  );
}
