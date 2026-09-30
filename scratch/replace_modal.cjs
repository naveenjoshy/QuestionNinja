const fs = require('fs');
let content = fs.readFileSync('src/App.jsx', 'utf8');

// Replace handleFloatingFormulaClick
content = content.replace(/const handleFloatingFormulaClick = \(\) => \{[\s\S]*?\}\);\n  \};/, `const handleFloatingFormulaClick = () => {
    if (!activeInputInfo) return;
    setFormulaModal({ isOpen: true });
  };

  const handleInsertSymbol = (symbolToInsert) => {
    if (!activeInputInfo) return;
    const elementId = activeInputInfo.id;
    const el = document.getElementById(elementId);
    if (!el) return;

    const currentValue = el.value;
    const start = el.selectionStart;
    const end = el.selectionEnd;

    const newValue = currentValue.substring(0, start) + symbolToInsert + currentValue.substring(end);
    updateValueForId(elementId, newValue);

    // Restore focus and update cursor position
    setTimeout(() => {
      const inputEl = document.getElementById(elementId);
      if (inputEl) {
        inputEl.focus();
        const newCursorPos = start + symbolToInsert.length;
        inputEl.setSelectionRange(newCursorPos, newCursorPos);
        setActiveInputInfo({ id: elementId });
      }
    }, 50);
  };`);

// Replace Formula Editor Modal JSX
const oldModalMatch = /\{\/\* Formula Editor Modal \*\/\}\s*\{formulaModal\.isOpen && \([\s\S]*?\}\s*<\/button>\s*<\/div>\s*<\/div>\s*<\/div>\s*\)\}/;
const newModal = `{/* Formula Editor Modal */}
      {formulaModal.isOpen && (
        <div className="modal-overlay" onClick={() => setFormulaModal({ isOpen: false })} style={{ zIndex: 10000 }}>
          <div className="modal-content formula-modal-content" onClick={(e) => e.stopPropagation()} style={{ maxWidth: '450px' }}>
            <div className="modal-header">
              <h3 style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <span style={{ fontStyle: 'italic', fontWeight: 'bold', fontSize: '20px' }}>𝑓(x)</span> Insert Math Symbol
              </h3>
            </div>

            <div className="modal-body" style={{ display: 'flex', flexDirection: 'column', gap: '16px', maxHeight: '70vh', overflowY: 'auto', padding: '16px' }}>
              <div className="formula-toolbar">
                {FORMULA_BUTTONS.map((group) => (
                  <div key={group.label} className="formula-toolbar-group">
                    <span className="formula-toolbar-label">{group.label}</span>
                    <div className="formula-toolbar-buttons">
                      {group.buttons.map((btn, bIdx) => (
                        <button
                          key={bIdx}
                          className="formula-btn"
                          title={btn.hint}
                          onClick={() => handleInsertSymbol(btn.insert)}
                        >
                          {btn.symbol}
                        </button>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </div>

            <div className="modal-footer" style={{ gap: '8px' }}>
              <button className="btn btn-primary" onClick={() => setFormulaModal({ isOpen: false })}>
                Done
              </button>
            </div>
          </div>
        </div>
      )}`;

content = content.replace(oldModalMatch, newModal);
fs.writeFileSync('src/App.jsx', content);
