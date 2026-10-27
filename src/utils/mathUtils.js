import katex from 'katex';
import DOMPurify from 'dompurify';

// Helper: Fail-safe HTML entity escaping
export const escapeHTML = (str) =>
  String(str).replace(/[&<>'"]/g, (tag) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    "'": '&#39;',
    '"': '&quot;'
  }[tag] || tag));

// Helper: render square root with proper radical symbol and top overbar line
export const renderSquareRootHTML = (inner) => {
  if (!inner) return '<span class="math-sqrt-symbol">√</span>';
  const trimmed = inner.trim();
  if (!trimmed) {
    return '<span class="math-sqrt-symbol">√</span>';
  }
  return '<span class="math-sqrt-symbol">√</span>' + trimmed;
};

// Helper: render LaTeX formula to sanitized HTML string for live preview
export const renderLatex = (latex) => {
  if (!latex || !String(latex).trim()) return '';
  try {
    const rawHtml = katex.renderToString(String(latex).trim(), {
      throwOnError: false,
      displayMode: true,
      output: 'html'
    });
    const sanitizeFn = typeof DOMPurify?.sanitize === 'function' ? DOMPurify.sanitize : null;
    return sanitizeFn ? sanitizeFn(rawHtml, {
      ADD_TAGS: ['annotation', 'semantics', 'mrow', 'mi', 'mo', 'mn', 'msup', 'msub', 'mfrac', 'msqrt', 'span', 'sup', 'sub'],
      ADD_ATTR: ['aria-hidden', 'aria-expanded', 'style', 'class']
    }) : escapeHTML(rawHtml);
  } catch {
    return '<span style="color:red;">Invalid formula</span>';
  }
};

// High-performance LRU cache for sanitized math HTML
const mathRenderCache = new Map();
const MAX_MATH_CACHE_SIZE = 1000;

// Helper: render text that contains $...$ math blocks or raw math symbols (e.g. √, \sqrt, minus) (sanitized)
export const renderTextWithMath = (text) => {
  if (text === null || text === undefined) return '';
  const str = String(text);
  if (!str) return '';
  const cached = mathRenderCache.get(str);
  if (cached !== undefined) return cached;

  let result = str;

  // 1. Process explicit $...$ math blocks FIRST to isolate and protect KaTeX rendered math
  const mathBlocks = [];
  const mathToken = `___QN_MATH_${Date.now()}_${Math.random().toString(36).slice(2)}___`;
  result = result.replace(/\$([^$\n]+?)\$/g, (match, mathContent) => {
    try {
      const katexHTML = katex.renderToString(mathContent, {
        throwOnError: false,
        displayMode: false,
        output: 'html'
      });
      const placeholder = `${mathToken}${mathBlocks.length}___`;
      mathBlocks.push(katexHTML);
      return placeholder;
    } catch {
      return `<span style="color:red;">${escapeHTML(match)}</span>`;
    }
  });

  // 2. Format math subtraction and negative numbers with Unicode minus sign (U+2212) BEFORE creating HTML tags
  result = result.replace(/(\b[0-9a-zA-Z]+\b|\))\s*-\s*(\b[0-9a-zA-Z]+\b|\()/g, '$1 − $2');
  result = result.replace(/(^|\s)-([0-9a-zA-Z])/g, '$1−$2');

  // 3. Convert raw \sqrt{...} and Unicode √ OUTSIDE of $...$ into styled radical + overbar content
  result = result.replace(/\\sqrt\[([^\]]*)\]\{([^}]*)\}/g, (_, degree, inner) => `${degree}√${inner}`);
  result = result.replace(/\\sqrt\{([^}]*)\}/g, (_, inner) => renderSquareRootHTML(inner));
  result = result.replace(/√/g, '<span class="math-sqrt-symbol">√</span>');

  // 4. Convert Unicode cube root ∛ to styled span
  result = result.replace(/∛/g, '<span class="math-sqrt-symbol">∛</span>');

  // 5. Restore rendered $...$ math blocks safely using functional replacement to prevent $ substitution
  mathBlocks.forEach((html, i) => {
    result = result.replace(`${mathToken}${i}___`, () => html);
  });

  // Convert newlines to block lines for precise page-break slicing in PDF
  const lines = result.split('\n');
  result = lines.map(line => `<span class="paper-text-line" style="display: block; min-height: 1.2rem;">${line || '&nbsp;'}</span>`).join('');

  const sanitizeFn = typeof DOMPurify?.sanitize === 'function' ? DOMPurify.sanitize : null;
  const sanitized = sanitizeFn ? sanitizeFn(result, {
    ADD_TAGS: ['annotation', 'semantics', 'mrow', 'mi', 'mo', 'mn', 'msup', 'msub', 'mfrac', 'msqrt', 'span', 'sup', 'sub', 'br'],
    ADD_ATTR: ['aria-hidden', 'aria-expanded', 'style', 'class']
  }) : escapeHTML(result);

  if (mathRenderCache.size >= MAX_MATH_CACHE_SIZE) {
    const firstKey = mathRenderCache.keys().next().value;
    mathRenderCache.delete(firstKey);
  }
  mathRenderCache.set(text, sanitized);

  return sanitized;
};

// High-performance cache for LaTeX to plain text conversion
const latexCache = new Map();
const MAX_LATEX_CACHE_SIZE = 500;

// Pre-sorted symbol lookups for fast repeated string replacements
const GREEKS = {
  '\\alpha': 'α', '\\beta': 'β', '\\gamma': 'γ', '\\delta': 'δ', '\\epsilon': 'ε', '\\varepsilon': 'ϵ',
  '\\zeta': 'ζ', '\\eta': 'η', '\\theta': 'θ', '\\vartheta': 'ϑ', '\\iota': 'ι', '\\kappa': 'κ',
  '\\lambda': 'λ', '\\mu': 'μ', '\\nu': 'ν', '\\xi': 'ξ', '\\pi': 'π', '\\varpi': 'ϖ',
  '\\rho': 'ρ', '\\varrho': 'ϱ', '\\sigma': 'σ', '\\varsigma': 'ς', '\\tau': 'τ', '\\upsilon': 'υ',
  '\\phi': 'φ', '\\varphi': 'ϕ', '\\chi': 'χ', '\\psi': 'ψ', '\\omega': 'ω',
  '\\Gamma': 'Γ', '\\Delta': 'Δ', '\\Theta': 'Θ', '\\Lambda': 'Λ', '\\Xi': 'Ξ', '\\Pi': 'Π',
  '\\Sigma': 'Σ', '\\Upsilon': 'Υ', '\\Phi': 'Φ', '\\Psi': 'Ψ', '\\Omega': 'Ω'
};
const SORTED_GREEKS = Object.keys(GREEKS).sort((a, b) => b.length - a.length);

const OPS = {
  '\\pm': '±', '\\mp': '∓', '\\times': '×', '\\div': '÷', '\\cdot': '·', '\\star': '⋆', '\\circ': '∘',
  '\\bullet': '•', '\\cap': '∩', '\\cup': '∪', '\\uplus': '⊎', '\\sqcap': '⊓', '\\sqcup': '⊔',
  '\\vee': '∨', '\\wedge': '∧', '\\setminus': '∖', '\\diamond': '⋄',
  '\\leq': '≤', '\\geq': '≥', '\\neq': '≠', '\\equiv': '≡', '\\approx': '≈', '\\sim': '∼',
  '\\simeq': '≃', '\\asymp': '≍', '\\propto': '∝', '\\perp': '⊥', '\\parallel': '∥',
  '\\in': '∈', '\\notin': '∉', '\\subset': '⊂', '\\supset': '⊃', '\\subseteq': '⊆', '\\supseteq': '⊇',
  '\\forall': '∀', '\\exists': '∃', '\\nexists': '∄', '\\emptyset': '∅', '\\nabla': '∇',
  '\\partial': '∂', '\\int': '∫', '\\iint': '∬', '\\iiint': '∭', '\\oint': '∮', '\\sum': '∑',
  '\\prod': '∏', '\\coprod': '∐', '\\infty': '∞', '\\therefore': '∴', '\\because': '∵',
  '\\implies': '⇒', '\\iff': '⇔', '\\rightarrow': '→', '\\leftarrow': '←', '\\leftrightarrow': '↔',
  '\\Rightarrow': '⇒', '\\Leftarrow': '⇐', '\\Leftrightarrow': '⇔', '\\angle': '∠', '\\measuredangle': '∡',
  '\\degree': '°', '\\ldots': '…', '\\cdots': '⋯', '\\vdots': '⋮', '\\ddots': '⋱'
};
const SORTED_OPS = Object.keys(OPS).sort((a, b) => b.length - a.length);

const SUP_MAP = { '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹', 'n': 'ⁿ', '+': '⁺', '-': '⁻', '=': '⁼', '(': '⁽', ')': '⁾', 'a': 'ᵃ', 'b': 'ᵇ', 'c': 'ᶜ', 'x': 'ˣ', 'y': 'ʸ' };
const SUB_MAP = { '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄', '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉', '+': '₊', '-': '₋', '=': '₌', '(': '₍', ')': '₎', 'a': 'ₐ', 'e': 'ₑ', 'o': 'ₒ', 'x': 'ₓ' };

// Helper: convert LaTeX to readable plain text for DOCX with full math symbol support
export const latexToPlainText = (latex) => {
  if (!latex) return '';
  const cached = latexCache.get(latex);
  if (cached !== undefined) return cached;

  let text = latex;
  // Binomial coefficients / combinations
  text = text.replace(/\\binom\{([^}]*)\}\{([^}]*)\}/g, (_, n, k) => `C(${n}, ${k})`);
  // Fractions (keep clean separation for numerator / denominator)
  text = text.replace(/\\frac\{([^}]*)\}\{([^}]*)\}/g, (_, num, den) => `(${num}) / (${den})`);
  // nth root and square root
  text = text.replace(/\\sqrt\[([^\]]*)\]\{([^}]*)\}/g, (_, n, arg) => `ⁿ√(${arg})`);
  text = text.replace(/\\sqrt\{([^}]*)\}/g, (_, arg) => {
    const trimmed = arg ? arg.trim() : '';
    if (/^[0-9a-zA-Z]+$/.test(trimmed)) {
      return `√${trimmed}`;
    }
    return `√(${trimmed})`;
  });
  // Vector notation
  text = text.replace(/\\vec\{([^}]*)\}/g, (_, v) => `${v}⃗`);
  text = text.replace(/\\hat\{([^}]*)\}/g, (_, v) => `${v}̂`);
  text = text.replace(/\\bar\{([^}]*)\}/g, (_, v) => `${v}̄`);
  // Superscript
  text = text.replace(/\^\{([^}]*)\}/g, (_, p) => {
    return p.split('').map(c => SUP_MAP[c] || `^${c}`).join('');
  });
  text = text.replace(/\^([0-9n])/g, (_, c) => {
    return SUP_MAP[c] || `^${c}`;
  });
  // Subscript
  text = text.replace(/_\{([^}]*)\}/g, (_, p) => {
    return p.split('').map(c => SUB_MAP[c] || `_${c}`).join('');
  });
  text = text.replace(/_([0-9a-zA-Z])/g, (_, c) => {
    return SUB_MAP[c] || `_${c}`;
  });
  // Full Greek alphabet (lowercase & uppercase)
  for (const k of SORTED_GREEKS) {
    text = text.replaceAll(k, GREEKS[k]);
  }
  // Comprehensive Math, Arithmetic, Calculus & Logic operators
  for (const k of SORTED_OPS) {
    text = text.replaceAll(k, OPS[k]);
  }
  // Clean remaining commands
  text = text.replace(/\\(lim|log|sin|cos|tan|cot|sec|csc|ln|max|min|deg)/g, '$1');
  text = text.replace(/\\[a-zA-Z]+/g, '');
  text = text.replace(/[{}]/g, '');
  text = text.replace(/\s+/g, ' ').trim();

  if (latexCache.size >= MAX_LATEX_CACHE_SIZE) {
    const firstKey = latexCache.keys().next().value;
    latexCache.delete(firstKey);
  }
  latexCache.set(latex, text);

  return text;
};
