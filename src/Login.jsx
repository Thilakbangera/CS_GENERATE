import { useState } from 'react';

const VALID_EMAIL    = 'patent.admin@gmail.com';
const VALID_PASSWORD = 'Drafters@334';

export default function Login({ onLogin }) {
  const [email, setEmail]       = useState('');
  const [password, setPassword] = useState('');
  const [showPwd, setShowPwd]   = useState(false);
  const [error, setError]       = useState('');
  const [loading, setLoading]   = useState(false);
  const [shake, setShake]       = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError('');
    setLoading(true);

    // Simulate a brief loading feel
    await new Promise(r => setTimeout(r, 600));

    if (email.trim().toLowerCase() === VALID_EMAIL && password === VALID_PASSWORD) {
      sessionStorage.setItem('patent_auth', '1');
      onLogin();
    } else {
      setLoading(false);
      setError('Invalid email or password. Please try again.');
      setShake(true);
      setTimeout(() => setShake(false), 600);
    }
  };

  return (
    <div className="login-shell">
      {/* Left decorative panel */}
      <div className="login-left">
        <div className="login-left-inner">
          <div className="login-brand-icon">
            <span className="material-icons-round" style={{ fontSize: 36, color: '#fff' }}>description</span>
          </div>
          <div className="login-brand-title">Patent CS Generator</div>
          <div className="login-brand-sub">Complete Specification Drafting · IPO India Form 2</div>

          <div className="login-features">
            {[
              { icon: 'bolt',         text: 'AI-powered instant drafting'       },
              { icon: 'gavel',        text: 'IPO India Form 2 compliant'        },
              { icon: 'format_quote', text: 'Article usage rules enforced'      },
              { icon: 'download',     text: 'Export to Word (.docx) instantly'  },
              { icon: 'api',          text: 'Gemini, Groq & Claude support'     },
            ].map(f => (
              <div className="login-feature-item" key={f.text}>
                <span className="material-icons-round login-feature-icon">{f.icon}</span>
                <span>{f.text}</span>
              </div>
            ))}
          </div>

          <div className="login-left-badge">Form 2 · IPO India</div>
        </div>

        {/* Decorative blobs */}
        <div className="login-blob login-blob-1" />
        <div className="login-blob login-blob-2" />
        <div className="login-blob login-blob-3" />
      </div>

      {/* Right login form */}
      <div className="login-right">
        <div className={`login-card${shake ? ' login-shake' : ''}`}>
          {/* Logo for mobile */}
          <div className="login-card-logo">
            <div className="login-card-logo-icon">
              <span className="material-icons-round" style={{ fontSize: 22, color: '#fff' }}>description</span>
            </div>
            <div>
              <div className="login-card-logo-text">Patent CS Generator</div>
              <div className="login-card-logo-sub">IPO India Form 2</div>
            </div>
          </div>

          <div className="login-card-heading">Welcome back</div>
          <div className="login-card-subheading">Sign in to access the specification generator</div>

          <form onSubmit={handleSubmit} autoComplete="off">
            {/* Email */}
            <div className="login-field">
              <label className="login-label">
                <span className="material-icons-round login-label-icon">email</span>
                Email address
              </label>
              <div className="login-input-wrap">
                <input
                  id="login-email"
                  type="email"
                  className="login-input"
                  placeholder="patent.admin@gmail.com"
                  value={email}
                  onChange={e => { setEmail(e.target.value); setError(''); }}
                  required
                  autoFocus
                />
              </div>
            </div>

            {/* Password */}
            <div className="login-field">
              <label className="login-label">
                <span className="material-icons-round login-label-icon">lock</span>
                Password
              </label>
              <div className="login-input-wrap login-input-wrap-pwd">
                <input
                  id="login-password"
                  type={showPwd ? 'text' : 'password'}
                  className="login-input"
                  placeholder="Enter your password"
                  value={password}
                  onChange={e => { setPassword(e.target.value); setError(''); }}
                  required
                />
                <button
                  type="button"
                  className="login-pwd-toggle"
                  onClick={() => setShowPwd(v => !v)}
                  tabIndex={-1}
                  title={showPwd ? 'Hide password' : 'Show password'}
                >
                  <span className="material-icons-round" style={{ fontSize: 18 }}>
                    {showPwd ? 'visibility_off' : 'visibility'}
                  </span>
                </button>
              </div>
            </div>

            {/* Error */}
            {error && (
              <div className="login-error">
                <span className="material-icons-round" style={{ fontSize: 16 }}>error_outline</span>
                {error}
              </div>
            )}

            {/* Submit */}
            <button
              id="login-submit-btn"
              type="submit"
              className={`login-btn${loading ? ' login-btn-loading' : ''}`}
              disabled={loading}
            >
              {loading ? (
                <>
                  <span className="login-spinner" />
                  Signing in…
                </>
              ) : (
                <>
                  <span className="material-icons-round" style={{ fontSize: 18 }}>login</span>
                  Sign in
                </>
              )}
            </button>
          </form>

          <div className="login-footer">
            Patent CS Generator &nbsp;·&nbsp; v2.0 &nbsp;·&nbsp; IPO India Form 2
          </div>
        </div>
      </div>
    </div>
  );
}
