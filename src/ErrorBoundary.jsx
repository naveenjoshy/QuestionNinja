import React from 'react';

export class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error) {
    return { hasError: true, error };
  }

  componentDidCatch(error, errorInfo) {
    console.error('QuestionNinja caught a rendering error:', error, errorInfo);
  }

  handleResetDraft = () => {
    try {
      localStorage.removeItem('question_ninja_draft');
    } catch {
      // ignore
    }
    window.location.reload();
  };

  handleReload = () => {
    window.location.reload();
  };

  render() {
    if (this.state.hasError) {
      return (
        <div style={{
          minHeight: '100vh',
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          justifyContent: 'center',
          backgroundColor: '#0f172a',
          color: '#f8fafc',
          padding: '24px',
          fontFamily: 'Inter, system-ui, -apple-system, sans-serif'
        }}>
          <div style={{
            maxWidth: '520px',
            width: '100%',
            backgroundColor: '#1e293b',
            borderRadius: '16px',
            border: '1px solid rgba(239, 68, 68, 0.3)',
            padding: '32px',
            boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.5), 0 8px 10px -6px rgba(0, 0, 0, 0.5)',
            display: 'flex',
            flexDirection: 'column',
            gap: '18px'
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
              <div style={{
                width: '40px',
                height: '40px',
                borderRadius: '10px',
                backgroundColor: 'rgba(239, 68, 68, 0.15)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                color: '#ef4444',
                fontSize: '22px',
                fontWeight: 'bold'
              }}>
                ⚠
              </div>
              <div>
                <h2 style={{ margin: 0, fontSize: '18px', fontWeight: '700' }}>Something went wrong</h2>
                <p style={{ margin: 0, fontSize: '13px', color: '#94a3b8' }}>An unexpected error occurred while rendering the question paper.</p>
              </div>
            </div>

            {this.state.error && (
              <div style={{
                padding: '12px 14px',
                backgroundColor: 'rgba(0, 0, 0, 0.3)',
                borderRadius: '8px',
                border: '1px solid rgba(255, 255, 255, 0.08)',
                fontSize: '12px',
                fontFamily: 'monospace',
                color: '#fca5a5',
                overflowX: 'auto',
                maxHeight: '150px'
              }}>
                {this.state.error.message || String(this.state.error)}
              </div>
            )}

            <div style={{ display: 'flex', gap: '10px', marginTop: '6px' }}>
              <button
                onClick={this.handleReload}
                style={{
                  flex: 1,
                  padding: '10px 16px',
                  borderRadius: '8px',
                  backgroundColor: '#334155',
                  color: '#fff',
                  border: 'none',
                  fontSize: '13px',
                  fontWeight: '600',
                  cursor: 'pointer'
                }}
              >
                Reload Page
              </button>
              <button
                onClick={this.handleResetDraft}
                style={{
                  flex: 1,
                  padding: '10px 16px',
                  borderRadius: '8px',
                  backgroundColor: '#4f46e5',
                  color: '#fff',
                  border: 'none',
                  fontSize: '13px',
                  fontWeight: '600',
                  cursor: 'pointer'
                }}
              >
                Clear Draft & Reset
              </button>
            </div>
          </div>
        </div>
      );
    }

    return this.props.children;
  }
}
