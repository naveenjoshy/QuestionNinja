const fs = require('fs');

let content = fs.readFileSync('src/App.jsx', 'utf8');

// 1. Update FORMULA_BUTTONS
const newFormulaButtons = `const FORMULA_BUTTONS = [
  { label: 'Basic', buttons: [
    { symbol: 'x²', insert: '²', hint: 'Superscript 2' },
    { symbol: 'x³', insert: '³', hint: 'Superscript 3' },
    { symbol: 'xⁿ', insert: '^', hint: 'Power' },
    { symbol: 'x₁', insert: '₁', hint: 'Subscript 1' },
    { symbol: 'x₂', insert: '₂', hint: 'Subscript 2' },
    { symbol: 'xₙ', insert: '_', hint: 'Subscript' },
    { symbol: '½', insert: '½', hint: 'Half' },
    { symbol: '¼', insert: '¼', hint: 'Quarter' },
    { symbol: '¾', insert: '¾', hint: 'Three quarters' },
    { symbol: '√', insert: '√', hint: 'Square root' },
    { symbol: '∛', insert: '∛', hint: 'Cube root' },
  ]},
  { label: 'Greek', buttons: [
    { symbol: 'α', insert: 'α', hint: 'Alpha' },
    { symbol: 'β', insert: 'β', hint: 'Beta' },
    { symbol: 'γ', insert: 'γ', hint: 'Gamma' },
    { symbol: 'δ', insert: 'δ', hint: 'Delta' },
    { symbol: 'θ', insert: 'θ', hint: 'Theta' },
    { symbol: 'λ', insert: 'λ', hint: 'Lambda' },
    { symbol: 'μ', insert: 'μ', hint: 'Mu' },
    { symbol: 'π', insert: 'π', hint: 'Pi' },
    { symbol: 'σ', insert: 'σ', hint: 'Sigma' },
    { symbol: 'φ', insert: 'φ', hint: 'Phi' },
    { symbol: 'ω', insert: 'ω', hint: 'Omega' },
    { symbol: 'Δ', insert: 'Δ', hint: 'Delta (uppercase)' },
    { symbol: 'Σ', insert: 'Σ', hint: 'Sigma (uppercase)' },
    { symbol: 'Ω', insert: 'Ω', hint: 'Omega (uppercase)' },
  ]},
  { label: 'Operators', buttons: [
    { symbol: '±', insert: '±', hint: 'Plus-minus' },
    { symbol: '×', insert: '×', hint: 'Multiply' },
    { symbol: '÷', insert: '÷', hint: 'Divide' },
    { symbol: '·', insert: '·', hint: 'Dot product' },
    { symbol: '≠', insert: '≠', hint: 'Not equal' },
    { symbol: '≤', insert: '≤', hint: 'Less than or equal' },
    { symbol: '≥', insert: '≥', hint: 'Greater than or equal' },
    { symbol: '≈', insert: '≈', hint: 'Approximately' },
    { symbol: '∞', insert: '∞', hint: 'Infinity' },
    { symbol: '∝', insert: '∝', hint: 'Proportional' },
    { symbol: '°', insert: '°', hint: 'Degree' },
  ]},
  { label: 'Calculus', buttons: [
    { symbol: '∫', insert: '∫', hint: 'Integral' },
    { symbol: '∂', insert: '∂', hint: 'Partial derivative' },
    { symbol: '∑', insert: '∑', hint: 'Summation' },
    { symbol: '∏', insert: '∏', hint: 'Product' },
  ]},
  { label: 'Trig & Log', buttons: [
    { symbol: 'sin', insert: 'sin ', hint: 'Sine' },
    { symbol: 'cos', insert: 'cos ', hint: 'Cosine' },
    { symbol: 'tan', insert: 'tan ', hint: 'Tangent' },
    { symbol: 'log', insert: 'log ', hint: 'Logarithm' },
    { symbol: 'ln', insert: 'ln ', hint: 'Natural log' },
  ]}
];`;

content = content.replace(/const FORMULA_BUTTONS = \[\s*\{ label: 'Basic'[\s\S]*?\]\s*}\s*\];/g, newFormulaButtons);

fs.writeFileSync('src/App.jsx', content);
