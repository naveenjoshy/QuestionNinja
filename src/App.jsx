import React, { useState, useEffect, useRef } from 'react';
import pageLogo from './assets/logo.png';
import schoolLogo from './assets/school_logo.png';
import {
  Plus,
  Trash2,
  ArrowUp,
  ArrowDown,
  Download,
  FileText,
  Settings,
  Layers,
  Image as ImageIcon,
  CheckCircle,
  AlertTriangle,
  Move,
  Maximize2,
  ChevronDown,
  ChevronUp,
  Sun,
  Moon,
  ArrowLeft,
  ArrowRight,
  Printer,
  Loader2,
  ExternalLink,
  Cloud,
  Shield,
  Upload,
  BookOpen,
  GraduationCap,
  School,
  Check
} from 'lucide-react';
import * as docx from 'docx';
import { jsPDF } from 'jspdf';
import katex from 'katex';
import 'katex/dist/katex.min.css';

// Helper: render LaTeX to HTML string
const renderLatex = (latex) => {
  try {
    return katex.renderToString(latex, {
      throwOnError: false,
      displayMode: true,
      output: 'html'
    });
  } catch {
    return `<span style="color:red;">Invalid formula</span>`;
  }
};

// Helper: check if text contains formulas or LaTeX math
const hasFormula = (text) => {
  if (!text) return false;
  return text.includes('$') || text.includes('\\') || /[\u0370-\u03FF\u2200-\u22FF]/.test(text);
};

// Helper: render text that contains $...$ math blocks using KaTeX
const renderTextWithMath = (text) => {
  if (!text) return '';
  let result = text.replace(/\$([^$\n]+?)\$/g, (match, mathContent) => {
    try {
      return katex.renderToString(mathContent, {
        throwOnError: false,
        displayMode: false,
        output: 'html'
      });
    } catch {
      return `<span style="color:red;">${match}</span>`;
    }
  });
  result = result.replace(/\n/g, '<br>');
  return result;
};

// Helper: convert text with $...$ math into an array of docx.TextRun objects
const docxTextRunsWithMath = (text, defaultOptions = {}) => {
  if (!text) return [];
  const parts = text.split('$');
  const runs = [];
  parts.forEach((part, index) => {
    if (index % 2 === 1) {
      // This is math
      const plainMath = latexToPlainText(part);
      runs.push(new docx.TextRun({
        text: plainMath,
        italics: true,
        font: 'Cambria Math',
        size: defaultOptions.size || 22,
        ...defaultOptions
      }));
    } else {
      // This is plain text, which may contain newlines '\n'
      const lines = part.split('\n');
      lines.forEach((line, lIdx) => {
        const runProps = {
          text: line,
          size: defaultOptions.size || 22,
          ...defaultOptions
        };
        if (lIdx > 0) {
          runProps.break = 1;
        }
        runs.push(new docx.TextRun(runProps));
      });
    }
  });
  return runs;
};

// Helper: convert LaTeX to readable plain text for DOCX
const latexToPlainText = (latex) => {
  if (!latex) return '';
  let text = latex;
  // Fractions
  text = text.replace(/\\frac\{([^}]*)\}\{([^}]*)\}/g, '($1)/($2)');
  // Square root
  text = text.replace(/\\sqrt\{([^}]*)\}/g, '√($1)');
  // Superscript
  text = text.replace(/\^\{([^}]*)\}/g, (_, p) => {
    const supMap = { '0':'⁰','1':'¹','2':'²','3':'³','4':'⁴','5':'⁵','6':'⁶','7':'⁷','8':'⁸','9':'⁹','n':'ⁿ','+':'⁺','-':'⁻' };
    return p.split('').map(c => supMap[c] || `^${c}`).join('');
  });
  text = text.replace(/\^([0-9n])/g, (_, c) => {
    const supMap = { '0':'⁰','1':'¹','2':'²','3':'³','4':'⁴','5':'⁵','6':'⁶','7':'⁷','8':'⁸','9':'⁹','n':'ⁿ' };
    return supMap[c] || `^${c}`;
  });
  // Subscript
  text = text.replace(/_\{([^}]*)\}/g, (_, p) => {
    const subMap = { '0':'₀','1':'₁','2':'₂','3':'₃','4':'₄','5':'₅','6':'₆','7':'₇','8':'₈','9':'₉' };
    return p.split('').map(c => subMap[c] || `_${c}`).join('');
  });
  text = text.replace(/_([0-9])/g, (_, c) => {
    const subMap = { '0':'₀','1':'₁','2':'₂','3':'₃','4':'₄','5':'₅','6':'₆','7':'₇','8':'₈','9':'₉' };
    return subMap[c] || `_${c}`;
  });
  // Greek letters
  const greeks = { '\\alpha':'α','\\beta':'β','\\gamma':'γ','\\delta':'δ','\\theta':'θ','\\lambda':'λ','\\mu':'μ','\\pi':'π','\\sigma':'σ','\\phi':'φ','\\omega':'ω','\\Delta':'Δ','\\Sigma':'Σ','\\Pi':'Π','\\Omega':'Ω' };
  for (const [k, v] of Object.entries(greeks)) {
    text = text.replaceAll(k, v);
  }
  // Operators
  const ops = { '\\pm':'±','\\times':'×','\\div':'÷','\\neq':'≠','\\leq':'≤','\\geq':'≥','\\infty':'∞','\\propto':'∝','\\approx':'≈','\\rightarrow':'→','\\leftarrow':'←','\\leftrightarrow':'↔','\\int':'∫','\\partial':'∂','\\sum':'Σ','\\prod':'∏','\\cdot':'·','\\ldots':'…' };
  for (const [k, v] of Object.entries(ops)) {
    text = text.replaceAll(k, v);
  }
  // Clean remaining commands
  text = text.replace(/\\(lim|log|sin|cos|tan|ln)/g, '$1');
  text = text.replace(/\\[a-zA-Z]+/g, '');
  text = text.replace(/[{}]/g, '');
  text = text.replace(/\s+/g, ' ').trim();
  return text;
};

// Formula toolbar button definitions
const FORMULA_BUTTONS = [
  { label: 'Basic', buttons: [
    { symbol: 'x²', latex: '^{2}', hint: 'Superscript' },
    { symbol: 'xₙ', latex: '_{n}', hint: 'Subscript' },
    { symbol: '⁄', latex: '\\frac{a}{b}', hint: 'Fraction', replace: true },
    { symbol: '√', latex: '\\sqrt{x}', hint: 'Square root', replace: true },
    { symbol: '∛', latex: '\\sqrt[3]{x}', hint: 'Cube root', replace: true },
    { symbol: 'xⁿ', latex: '^{n}', hint: 'nth power' },
  ]},
  { label: 'Greek', buttons: [
    { symbol: 'α', latex: '\\alpha', hint: 'Alpha' },
    { symbol: 'β', latex: '\\beta', hint: 'Beta' },
    { symbol: 'γ', latex: '\\gamma', hint: 'Gamma' },
    { symbol: 'δ', latex: '\\delta', hint: 'Delta' },
    { symbol: 'θ', latex: '\\theta', hint: 'Theta' },
    { symbol: 'λ', latex: '\\lambda', hint: 'Lambda' },
    { symbol: 'μ', latex: '\\mu', hint: 'Mu' },
    { symbol: 'π', latex: '\\pi', hint: 'Pi' },
    { symbol: 'σ', latex: '\\sigma', hint: 'Sigma (lowercase)' },
    { symbol: 'φ', latex: '\\phi', hint: 'Phi' },
    { symbol: 'ω', latex: '\\omega', hint: 'Omega' },
    { symbol: 'Δ', latex: '\\Delta', hint: 'Delta (uppercase)' },
    { symbol: 'Σ', latex: '\\Sigma', hint: 'Sigma (uppercase)' },
    { symbol: 'Π', latex: '\\Pi', hint: 'Pi (uppercase)' },
    { symbol: 'Ω', latex: '\\Omega', hint: 'Omega (uppercase)' },
  ]},
  { label: 'Operators', buttons: [
    { symbol: '±', latex: '\\pm', hint: 'Plus-minus' },
    { symbol: '×', latex: '\\times', hint: 'Multiply' },
    { symbol: '÷', latex: '\\div', hint: 'Divide' },
    { symbol: '·', latex: '\\cdot', hint: 'Dot product' },
    { symbol: '≠', latex: '\\neq', hint: 'Not equal' },
    { symbol: '≤', latex: '\\leq', hint: 'Less than or equal' },
    { symbol: '≥', latex: '\\geq', hint: 'Greater than or equal' },
    { symbol: '≈', latex: '\\approx', hint: 'Approximately' },
    { symbol: '∞', latex: '\\infty', hint: 'Infinity' },
    { symbol: '∝', latex: '\\propto', hint: 'Proportional' },
  ]},
  { label: 'Calculus', buttons: [
    { symbol: '∫', latex: '\\int_{a}^{b}', hint: 'Integral', replace: true },
    { symbol: '∂', latex: '\\partial', hint: 'Partial derivative' },
    { symbol: 'lim', latex: '\\lim_{x \\to \\infty}', hint: 'Limit', replace: true },
    { symbol: 'Σ', latex: '\\sum_{i=1}^{n}', hint: 'Summation', replace: true },
    { symbol: '∏', latex: '\\prod_{i=1}^{n}', hint: 'Product', replace: true },
    { symbol: 'log', latex: '\\log', hint: 'Logarithm' },
    { symbol: 'ln', latex: '\\ln', hint: 'Natural log' },
  ]},
  { label: 'Trig', buttons: [
    { symbol: 'sin', latex: '\\sin', hint: 'Sine' },
    { symbol: 'cos', latex: '\\cos', hint: 'Cosine' },
    { symbol: 'tan', latex: '\\tan', hint: 'Tangent' },
  ]},
  { label: 'Arrows', buttons: [
    { symbol: '→', latex: '\\rightarrow', hint: 'Right arrow' },
    { symbol: '←', latex: '\\leftarrow', hint: 'Left arrow' },
    { symbol: '↔', latex: '\\leftrightarrow', hint: 'Double arrow' },
  ]},
  { label: 'Templates', buttons: [
    { symbol: 'a/b', latex: '\\frac{a}{b}', hint: 'Fraction', replace: true },
    { symbol: '√x', latex: '\\sqrt{x}', hint: 'Square root', replace: true },
    { symbol: 'x²+y²', latex: 'x^{2} + y^{2}', hint: 'Sum of squares', replace: true },
    { symbol: 'Quadratic', latex: 'x = \\frac{-b \\pm \\sqrt{b^{2}-4ac}}{2a}', hint: 'Quadratic formula', replace: true },
    { symbol: 'E=mc²', latex: 'E = mc^{2}', hint: "Einstein's equation", replace: true },
  ]},
];

const DEFAULT_BRANDING = {
  logo: schoolLogo,
  logoWidth: 65,
  logoHeight: 65,
  logoPos: { x: 0, y: 0 },
  schoolName: 'Girijyothi CMI Public School',
  schoolAddress: 'Vazhathope, Idukki',
  fontFamily: 'Cinzel',
  headerLogoOnly: false,
  hideSchoolLogo: false
};

const normalizeLogo = (val) => {
  if (val === null || val === undefined) return null;
  if (typeof val === 'string') {
    const trimmed = val.trim();
    if (trimmed === '' || trimmed === 'null' || trimmed === 'none' || trimmed === 'undefined' || trimmed === 'false') {
      return null;
    }
    const lower = trimmed.toLowerCase();
    if (
      lower === 'default' ||
      lower === 'schoollogo' ||
      lower.includes('school_logo') ||
      lower.includes('school-logo')
    ) {
      return schoolLogo;
    }
    return trimmed;
  }
  return val;
};

const DEFAULT_METADATA = {
  title: 'FIRST TERM SUMMATIVE ASSESSMENT',
  subject: 'Computer Science & Programming',
  classDiv: 'Class X - Division A & B',
  maxMarks: 50,
  duration: '90 Minutes',
  separateAnswerSheet: true,
  language: 'english'
};

const DEFAULT_SECTIONS = [
  {
    id: 'sec-1',
    title: 'SECTION A: OBJECTIVE TYPE QUESTIONS (MCQ)',
    marks: 5,
    instructions: 'Answer all the questions. Each question carries 1 mark. Select the most appropriate option.',
    type: 'mcq',
    questions: [
      {
        id: 'q-1',
        text: 'Which of the following is NOT a high-level programming language?',
        marks: 1,
        options: ['Python', 'Assembly Language', 'Java', 'C++']
      },
      {
        id: 'q-2',
        text: 'Which data structure operates on a Last-In, First-Out (LIFO) basis?',
        marks: 1,
        options: ['Queue', 'Stack', 'Linked List', 'Array']
      }
    ]
  },
  {
    id: 'sec-1-tf',
    title: 'SECTION B: TRUE OR FALSE',
    marks: 5,
    instructions: 'State whether the following statements are True or False.',
    type: 'true_false',
    questions: [
      {
        id: 'q-2-tf-1',
        text: 'In Python, variables are dynamically typed and do not need to be declared.',
        marks: 1
      },
      {
        id: 'q-2-tf-2',
        text: 'HTML is a programming language used for logic execution.',
        marks: 1
      }
    ]
  },
  {
    id: 'sec-2',
    title: 'SECTION C: MATCH THE COMPONENTS',
    marks: 3,
    instructions: 'Match the computer hardware components with their appropriate primary functions.',
    type: 'match_following',
    questions: [
      {
        id: 'q-4',
        text: 'Associate components to functions:',
        marks: 3,
        matchPairs: [
          { premise: 'CPU', response: 'Instruction execution and logic processing' },
          { premise: 'RAM', response: 'Temporary high-speed volatile data storage' },
          { premise: 'SSD', response: 'Persistent high-speed non-volatile storage' }
        ],
        shuffleB: true
      }
    ]
  },
  {
    id: 'sec-3',
    title: 'SECTION D: SHORT ANSWER & ESSAYS',
    marks: 15,
    instructions: 'Answer all questions. Assign marks based on the depth and logic of your explanation.',
    type: 'essay',
    questions: [
      {
        id: 'q-5',
        text: 'Discuss cloud computing architectures and security:',
        marks: 10,
        blankLines: 12,
        subQuestions: [
          {
            id: 'sq-1',
            label: '(a)',
            text: 'Explain the key differences between Public, Private, and Hybrid Cloud architectures.',
            marks: 5,
            blankLines: 6
          },
          {
            id: 'sq-2',
            label: '(b)',
            text: 'List three critical security concerns when hosting sensitive data in public cloud servers.',
            marks: 5,
            blankLines: 6
          }
        ]
      },
      {
        id: 'q-6',
        text: 'Explain the difference between compiler and interpreter.',
        marks: 5,
        blankLines: 6
      }
    ]
  }
];

export default function App() {
  // App states
  const [branding, setBranding] = useState(DEFAULT_BRANDING);
  const [metadata, setMetadata] = useState(DEFAULT_METADATA);
  const [sections, setSections] = useState([]);
  const [activeTab, setActiveTab] = useState('branding'); // branding, metadata, sections
  const [collapsedSections, setCollapsedSections] = useState({});
  const [isPreviewOpen, setIsPreviewOpen] = useState(false);
  const [isDownloadOpen, setIsDownloadOpen] = useState(false);
  const [isDocsModalOpen, setIsDocsModalOpen] = useState(false);
  const [isDocsUploading, setIsDocsUploading] = useState(false);
  const [docsError, setDocsError] = useState('');
  const [formulaModal, setFormulaModal] = useState({ isOpen: false, latex: '', onSave: null });
  const [isAboutModalOpen, setIsAboutModalOpen] = useState(false);
  const [activeInputInfo, setActiveInputInfo] = useState(null);
  const [csvImportModal, setCsvImportModal] = useState({ isOpen: false, branding: null, metadata: null, sections: [], importSchool: false, importExam: true, importQuestions: true });
  const [isPdfExporting, setIsPdfExporting] = useState(false);
  const paperSheetRef = useRef(null);
  const formulaInputRef = useRef(null);
  const dropdownRef = useRef(null);

  // Hook to track focus on formula-enabled text inputs
  useEffect(() => {
    const handleFocus = (e) => {
      const target = e.target;
      if (target && (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA')) {
        if (target.id && target.id.startsWith('q__')) {
          setActiveInputInfo({ id: target.id });
        } else {
          setActiveInputInfo(null);
        }
      }
    };

    document.addEventListener('focusin', handleFocus);
    return () => {
      document.removeEventListener('focusin', handleFocus);
    };
  }, []);

  const updateValueForId = (id, newValue) => {
    const parts = id.split('__');
    const type = parts[1];
    const secId = parts[2];
    const qId = parts[3];

    if (type === 'text') {
      updateQuestion(secId, qId, { text: newValue });
    } else if (type === 'opt') {
      const oIdx = parseInt(parts[4], 10);
      setSections(prev => prev.map(s => {
        if (s.id !== secId) return s;
        return {
          ...s,
          questions: s.questions.map(q => {
            if (q.id !== qId) return q;
            const newOpts = [...q.options];
            newOpts[oIdx] = newValue;
            return { ...q, options: newOpts };
          })
        };
      }));
    } else if (type === 'matcha') {
      const pIdx = parseInt(parts[4], 10);
      setSections(prev => prev.map(s => {
        if (s.id !== secId) return s;
        return {
          ...s,
          questions: s.questions.map(q => {
            if (q.id !== qId) return q;
            const newPairs = [...q.matchPairs];
            newPairs[pIdx] = { ...newPairs[pIdx], premise: newValue };
            return { ...q, matchPairs: newPairs };
          })
        };
      }));
    } else if (type === 'matchb') {
      const pIdx = parseInt(parts[4], 10);
      setSections(prev => prev.map(s => {
        if (s.id !== secId) return s;
        return {
          ...s,
          questions: s.questions.map(q => {
            if (q.id !== qId) return q;
            const newPairs = [...q.matchPairs];
            newPairs[pIdx] = { ...newPairs[pIdx], response: newValue };
            return { ...q, matchPairs: newPairs };
          })
        };
      }));
    } else if (type === 'tblh') {
      const hIdx = parseInt(parts[4], 10);
      setSections(prev => prev.map(s => {
        if (s.id !== secId) return s;
        return {
          ...s,
          questions: s.questions.map(q => {
            if (q.id !== qId) return q;
            const newData = { ...q.tableData };
            const newHeaders = [...newData.headers];
            newHeaders[hIdx] = newValue;
            newData.headers = newHeaders;
            return { ...q, tableData: newData };
          })
        };
      }));
    } else if (type === 'tblc') {
      const rIdx = parseInt(parts[4], 10);
      const cIdx = parseInt(parts[5], 10);
      setSections(prev => prev.map(s => {
        if (s.id !== secId) return s;
        return {
          ...s,
          questions: s.questions.map(q => {
            if (q.id !== qId) return q;
            const newData = { ...q.tableData };
            const newRows = newData.rows.map(r => [...r]);
            newRows[rIdx][cIdx] = newValue;
            newData.rows = newRows;
            return { ...q, tableData: newData };
          })
        };
      }));
    } else if (type === 'sqtext') {
      const sqId = parts[4];
      updateSubQuestion(secId, qId, sqId, { text: newValue });
    }
  };

  const handleFloatingFormulaClick = () => {
    if (!activeInputInfo) return;
    const elementId = activeInputInfo.id;
    const el = document.getElementById(elementId);
    if (!el) return;

    const currentValue = el.value;
    const start = el.selectionStart;
    const end = el.selectionEnd;

    setFormulaModal({
      isOpen: true,
      latex: '',
      onSave: (latex) => {
        const formulaString = `$${latex}$`;
        const newValue = currentValue.substring(0, start) + formulaString + currentValue.substring(end);
        updateValueForId(elementId, newValue);

        // Restore focus and cursor
        setTimeout(() => {
          const inputEl = document.getElementById(elementId);
          if (inputEl) {
            inputEl.focus();
            const newCursorPos = start + formulaString.length;
            inputEl.setSelectionRange(newCursorPos, newCursorPos);
            setActiveInputInfo({ id: elementId });
          }
        }, 50);
      }
    });
  };

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
        setIsDownloadOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, []);

  const [theme, setTheme] = useState(() => {
    return localStorage.getItem('question_ninja_theme') || 'dark';
  });

  useEffect(() => {
    localStorage.setItem('question_ninja_theme', theme);
  }, [theme]);

  const toggleTheme = () => {
    setTheme(prev => prev === 'dark' ? 'light' : 'dark');
  };

  const toggleSectionCollapse = (secId) => {
    setCollapsedSections(prev => ({
      ...prev,
      [secId]: !prev[secId]
    }));
  };

  // Logo interaction states
  const [isDragging, setIsDragging] = useState(false);
  const [isResizing, setIsResizing] = useState(false);
  const dragStart = useRef({ x: 0, y: 0 });
  const resizeStart = useRef({ x: 0, width: 0 });
  const logoRef = useRef(null);

  // Load from local storage on mount
  useEffect(() => {
    const savedDraft = localStorage.getItem('question_ninja_draft');
    if (savedDraft) {
      try {
        const parsed = JSON.parse(savedDraft);
        if (parsed.branding) {
          setBranding(prev => ({
            ...DEFAULT_BRANDING,
            ...parsed.branding,
            logo: parsed.branding.logo === null ? null : (normalizeLogo(parsed.branding.logo) || schoolLogo)
          }));
        }
        if (parsed.metadata) setMetadata(prev => ({ ...DEFAULT_METADATA, ...parsed.metadata }));
        if (parsed.sections) setSections(parsed.sections);
      } catch (e) {
        console.error('Error loading saved draft from localStorage', e);
      }
    }
  }, []);

  // Save to local storage on change
  useEffect(() => {
    const stateToSave = { branding, metadata, sections };
    localStorage.setItem('question_ninja_draft', JSON.stringify(stateToSave));
  }, [branding, metadata, sections]);

  // Handle Logo Upload with validation
  const handleLogoUpload = (e) => {
    const file = e.target.files[0];
    if (file) {
      if (!file.type.startsWith('image/')) {
        alert('Please select a valid image file (PNG, JPEG, etc.).');
        return;
      }
      if (file.size > 3 * 1024 * 1024) { // 3MB limit
        alert('Selected image exceeds the 3MB size limit. Please choose a smaller image.');
        return;
      }
      const reader = new FileReader();
      reader.onload = (uploadEvent) => {
        setBranding(prev => ({
          ...prev,
          logo: uploadEvent.target.result
        }));
      };
      reader.readAsDataURL(file);
    }
  };

  const removeLogo = () => {
    setBranding(prev => ({ ...prev, logo: null }));
  };

  // Logo pointer events (Custom Drag / Resize logic)
  const handleLogoPointerDown = (e) => {
    if (e.target.classList.contains('resize-handle')) {
      setIsResizing(true);
      resizeStart.current = {
        x: e.clientX,
        y: e.clientY,
        width: branding.logoWidth || 100,
        height: branding.logoHeight || 100
      };
    } else {
      setIsDragging(true);
      dragStart.current = {
        x: e.clientX - branding.logoPos.x,
        y: e.clientY - branding.logoPos.y
      };
    }
    e.target.setPointerCapture(e.pointerId);
  };

  const handleLogoPointerMove = (e) => {
    if (isDragging) {
      const newX = e.clientX - dragStart.current.x;
      const newY = e.clientY - dragStart.current.y;
      setBranding(prev => ({
        ...prev,
        logoPos: { x: newX, y: newY }
      }));
    } else if (isResizing) {
      const deltaX = e.clientX - resizeStart.current.x;
      const deltaY = e.clientY - resizeStart.current.y;
      const newWidth = Math.max(40, Math.min(400, resizeStart.current.width + deltaX));
      const newHeight = Math.max(40, Math.min(400, resizeStart.current.height + deltaY));
      setBranding(prev => ({
        ...prev,
        logoWidth: newWidth,
        logoHeight: newHeight
      }));
    }
  };

  const handleLogoPointerUp = (e) => {
    setIsDragging(false);
    setIsResizing(false);
    e.target.releasePointerCapture(e.pointerId);
  };

  // Helper to format marks to show at most 2 decimal places
  const formatMarks = (val) => {
    const num = Number(val) || 0;
    if (Number.isInteger(num)) return String(num);
    return num.toFixed(2);
  };

  // Helper calculation functions
  const getQuestionMarks = (q) => {
    if (q && q.subQuestions && q.subQuestions.length > 0) {
      return q.subQuestions.reduce((sum, sq) => sum + (Number(sq.marks) || 0), 0);
    }
    return Number(q?.marks) || 0;
  };

  const getSectionTotalMarks = (section) => {
    const rawTotal = section.questions.reduce((total, q) => total + getQuestionMarks(q), 0);
    return Math.round(rawTotal * 100) / 100;
  };

  const getExamCurrentTotalMarks = () => {
    const rawTotal = sections.reduce((total, s) => total + getSectionTotalMarks(s), 0);
    return Math.round(rawTotal * 100) / 100;
  };

  const hasBlankQuestions = () => {
    return sections.some(s => s.questions.some(q => {
      if (q.subQuestions && q.subQuestions.length > 0) {
        return q.subQuestions.some(sq => !sq.text || !sq.text.trim());
      }
      return !q.text || !q.text.trim();
    }));
  };

  const hasZeroMarkQuestions = () => {
    return sections.some(s => s.questions.some(q => {
      if (q.subQuestions && q.subQuestions.length > 0) {
        return q.subQuestions.some(sq => Number(sq.marks) === 0);
      }
      return Number(q.marks) === 0;
    }));
  };

  const hasQuestions = () => {
    return sections.some(s => s.questions && s.questions.length > 0);
  };

  const canFitSingleLine = (options) => {
    if (!options || options.length === 0) return false;
    if (options.some(opt => typeof opt === 'string' && opt.includes('\n'))) return false;
    const totalChars = options.reduce((acc, opt) => acc + (opt ? String(opt).length : 0), 0);
    const maxOptLen = Math.max(...options.map(opt => (opt ? String(opt).length : 0)));
    return options.length <= 4 && totalChars <= 36 && maxOptLen <= 12;
  };


  // State update helpers for Sections & Questions
  const addSection = () => {
    const newSection = {
      id: `sec-${Date.now()}`,
      title: `SECTION ${String.fromCharCode(65 + sections.length)}: NEW SECTION`,
      marks: 10,
      instructions: 'Answer all questions. Each question carries equal marks.',
      type: 'essay',
      questions: []
    };
    setSections([...sections, newSection]);
  };

  const deleteSection = (secId) => {
    setSections(sections.filter(s => s.id !== secId));
  };

  const updateSectionMeta = (secId, field, value) => {
    setSections(sections.map(s => {
      if (s.id === secId) {
        return { ...s, [field]: value };
      }
      return s;
    }));
  };

  const updateSectionType = (secId, newType) => {
    setSections(sections.map(sec => {
      if (sec.id === secId) {
        const updatedQuestions = sec.questions.map(q => {
          const updatedQ = { ...q };
          if (newType === 'mcq' && !updatedQ.options) {
            updatedQ.options = ['', '', '', ''];
          }
          if (newType === 'match_following' && !updatedQ.matchPairs) {
            updatedQ.matchPairs = [
              { premise: 'Item A', response: 'Match A' },
              { premise: 'Item B', response: 'Match B' }
            ];
            updatedQ.shuffleB = true;
          }
          if (newType === 'essay' && updatedQ.blankLines === undefined) {
            updatedQ.blankLines = 5;
          }
          if (newType === 'table') {
            if (!updatedQ.tableData) {
              updatedQ.tableRows = 3;
              updatedQ.tableCols = 3;
              updatedQ.tableData = {
                headers: ['Column 1', 'Column 2', 'Column 3'],
                rows: [
                  ['', '', ''],
                  ['', '', '']
                ]
              };
            }
          }
          return updatedQ;
        });
        return {
          ...sec,
          type: newType,
          questions: updatedQuestions
        };
      }
      return sec;
    }));
  };

  const addQuestion = (secId) => {
    setSections(sections.map(s => {
      if (s.id === secId) {
        const type = s.type || 'essay';
        const defaultQuestion = {
          id: `q-${Date.now()}`,
          text: '',
          marks: 1
        };

        if (type === 'mcq') {
          defaultQuestion.options = ['', '', '', ''];
        } else if (type === 'essay') {
          defaultQuestion.blankLines = 5;
        } else if (type === 'match_following') {
          defaultQuestion.matchPairs = [
            { premise: 'Item A', response: 'Match A' },
            { premise: 'Item B', response: 'Match B' }
          ];
          defaultQuestion.shuffleB = true;
        } else if (type === 'table') {
          defaultQuestion.tableRows = 3;
          defaultQuestion.tableCols = 3;
          defaultQuestion.tableData = {
            headers: ['Column 1', 'Column 2', 'Column 3'],
            rows: [
              ['', '', ''],
              ['', '', '']
            ]
          };
        }

        return {
          ...s,
          questions: [...s.questions, defaultQuestion]
        };
      }
      return s;
    }));
  };

  const deleteQuestion = (secId, qId) => {
    setSections(sections.map(s => {
      if (s.id === secId) {
        return {
          ...s,
          questions: s.questions.filter(q => q.id !== qId)
        };
      }
      return s;
    }));
  };

  const updateQuestion = (secId, qId, updatedFields) => {
    setSections(sections.map(s => {
      if (s.id === secId) {
        return {
          ...s,
          questions: s.questions.map(q => {
            if (q.id === qId) {
              return { ...q, ...updatedFields };
            }
            return q;
          })
        };
      }
      return s;
    }));
  };

  const addSubQuestion = (secId, qId) => {
    setSections(prev => prev.map(s => {
      if (s.id !== secId) return s;
      return {
        ...s,
        questions: s.questions.map(q => {
          if (q.id !== qId) return q;
          const subQs = q.subQuestions || [];
          const labelIndex = subQs.length;
          const defaultLabel = `(${String.fromCharCode(97 + (labelIndex % 26))})`;
          const newSubQ = {
            id: `sq-${Date.now()}`,
            label: defaultLabel,
            text: '',
            marks: 1,
            blankLines: 4,
            image: '',
            imageWidth: 300,
            imageHeight: 200
          };
          const updatedSubQs = [...subQs, newSubQ];
          const newTotalMarks = updatedSubQs.reduce((sum, sq) => sum + (Number(sq.marks) || 0), 0);
          return {
            ...q,
            subQuestions: updatedSubQs,
            marks: newTotalMarks
          };
        })
      };
    }));
  };

  const updateSubQuestion = (secId, qId, sqId, updatedFields) => {
    setSections(prev => prev.map(s => {
      if (s.id !== secId) return s;
      return {
        ...s,
        questions: s.questions.map(q => {
          if (q.id !== qId) return q;
          if (!q.subQuestions) return q;
          const updatedSubQs = q.subQuestions.map(sq => {
            if (sq.id !== sqId) return sq;
            return { ...sq, ...updatedFields };
          });
          const newTotalMarks = updatedSubQs.reduce((sum, sq) => sum + (Number(sq.marks) || 0), 0);
          return {
            ...q,
            subQuestions: updatedSubQs,
            marks: newTotalMarks
          };
        })
      };
    }));
  };

  const deleteSubQuestion = (secId, qId, sqId) => {
    setSections(prev => prev.map(s => {
      if (s.id !== secId) return s;
      return {
        ...s,
        questions: s.questions.map(q => {
          if (q.id !== qId) return q;
          if (!q.subQuestions) return q;
          const updatedSubQs = q.subQuestions.filter(sq => sq.id !== sqId);
          const newTotalMarks = updatedSubQs.reduce((sum, sq) => sum + (Number(sq.marks) || 0), 0);
          return {
            ...q,
            subQuestions: updatedSubQs,
            marks: updatedSubQs.length > 0 ? newTotalMarks : q.marks
          };
        })
      };
    }));
  };

  const moveSubQuestion = (secId, qId, sqIndex, direction) => {
    setSections(prev => prev.map(s => {
      if (s.id !== secId) return s;
      return {
        ...s,
        questions: s.questions.map(q => {
          if (q.id !== qId) return q;
          if (!q.subQuestions) return q;
          if (direction === 'up' && sqIndex === 0) return q;
          if (direction === 'down' && sqIndex === q.subQuestions.length - 1) return q;

          const nextIndex = direction === 'up' ? sqIndex - 1 : sqIndex + 1;
          const newSubQs = [...q.subQuestions];
          const temp = newSubQs[sqIndex];
          newSubQs[sqIndex] = newSubQs[nextIndex];
          newSubQs[nextIndex] = temp;
          return { ...q, subQuestions: newSubQs };
        })
      };
    }));
  };

  const handlePasteImage = (e, secId, qId, sqId = null) => {
    const section = sections.find(s => s.id === secId);
    if (!section) return;

    const items = e.clipboardData?.items;
    if (!items) return;

    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      if (item.type.indexOf('image') !== -1) {
        const file = item.getAsFile();
        if (file) {
          e.preventDefault();
          const reader = new FileReader();
          reader.onload = (uploadEvent) => {
            if (sqId) {
              updateSubQuestion(secId, qId, sqId, { image: uploadEvent.target.result });
            } else {
              updateQuestion(secId, qId, { image: uploadEvent.target.result });
            }
          };
          reader.readAsDataURL(file);
          break;
        }
      }
    }
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    e.currentTarget.classList.add('drag-over');
  };

  const handleDragLeave = (e) => {
    e.currentTarget.classList.remove('drag-over');
  };

  const handleDropImage = (e, secId, qId, sqId = null) => {
    e.preventDefault();
    e.currentTarget.classList.remove('drag-over');
    const files = e.dataTransfer?.files;
    if (files && files.length > 0) {
      const file = files[0];
      if (file.type.startsWith('image/')) {
        const reader = new FileReader();
        reader.onload = (uploadEvent) => {
          if (sqId) {
            updateSubQuestion(secId, qId, sqId, { image: uploadEvent.target.result });
          } else {
            updateQuestion(secId, qId, { image: uploadEvent.target.result });
          }
        };
        reader.readAsDataURL(file);
      }
    }
  };

  const moveSection = (index, direction) => {
    if (direction === 'up' && index === 0) return;
    if (direction === 'down' && index === sections.length - 1) return;

    const nextIndex = direction === 'up' ? index - 1 : index + 1;
    const newSections = [...sections];
    const temp = newSections[index];
    newSections[index] = newSections[nextIndex];
    newSections[nextIndex] = temp;
    setSections(newSections);
  };

  const moveQuestion = (secId, qIndex, direction) => {
    const section = sections.find(s => s.id === secId);
    if (!section) return;

    if (direction === 'up' && qIndex === 0) return;
    if (direction === 'down' && qIndex === section.questions.length - 1) return;

    const nextIndex = direction === 'up' ? qIndex - 1 : qIndex + 1;
    const newQuestions = [...section.questions];
    const temp = newQuestions[qIndex];
    newQuestions[qIndex] = newQuestions[nextIndex];
    newQuestions[nextIndex] = temp;

    setSections(sections.map(s => {
      if (s.id === secId) {
        return { ...s, questions: newQuestions };
      }
      return s;
    }));
  };

  // Reset demo
  const loadDemo = () => {
    if (window.confirm('This will overwrite your current progress with demo content. Proceed?')) {
      setBranding(DEFAULT_BRANDING);
      setMetadata(DEFAULT_METADATA);
      setSections(DEFAULT_SECTIONS);
    }
  };

  const resetAll = () => {
    if (window.confirm('Are you sure you want to clear all questions in the draft?')) {
      setSections([]);
    }
  };

  // Helper to format export filename as 'Class Name - Subject Name - Export Date - Time'
  const generateExportFilename = (ext) => {
    const className = (metadata.classDiv || 'Class').trim();
    const subjectName = (metadata.subject || 'Subject').trim();

    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const dateStr = `${year}-${month}-${day}`;

    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const seconds = String(now.getSeconds()).padStart(2, '0');
    const timeStr = `${hours}-${minutes}-${seconds}`;

    const sanitize = (str) => str.replace(/[\/\\:*?"<>|]/g, '').replace(/\s+/g, ' ').trim();

    const cleanClass = sanitize(className) || 'Class';
    const cleanSubject = sanitize(subjectName) || 'Subject';

    const baseName = `${cleanClass} - ${cleanSubject} - ${dateStr} - ${timeStr}`;
    return ext ? `${baseName}.${ext.replace(/^\./, '')}` : baseName;
  };

  // CSV Export & Import Features
  const exportToCSV = () => {
    if (!hasQuestions()) {
      alert("No questions added to export.");
      return;
    }
    const headers = [
      'School Logo',
      'School Logo Width',
      'School Logo Height',
      'Hide School Logo',
      'Header Logo Only',
      'School Name',
      'School Address',
      'Font Family',
      'Exam Title',
      'Subject',
      'Class Div',
      'Max Marks',
      'Duration',
      'Separate Answer Sheet',
      'Language',
      'Section Title',
      'Section Marks',
      'Section Instructions',
      'Question Type',
      'Question Text',
      'Question Marks',
      'Options',
      'Blank Lines',
      'Match Pairs',
      'Image Data',
      'Image Width',
      'Image Height',
      'Sub Questions'
    ];

    const getBrandingMetaCols = () => [
      branding.logo === schoolLogo ? 'school_logo.png' : (branding.logo || ''),
      branding.logoWidth !== undefined && branding.logoWidth !== null ? branding.logoWidth : '',
      branding.logoHeight !== undefined && branding.logoHeight !== null ? branding.logoHeight : '',
      branding.hideSchoolLogo ? 'true' : 'false',
      branding.headerLogoOnly ? 'true' : 'false',
      branding.schoolName || '',
      branding.schoolAddress || '',
      branding.fontFamily || '',
      metadata.title || '',
      metadata.subject || '',
      metadata.classDiv || '',
      metadata.maxMarks !== undefined && metadata.maxMarks !== null ? metadata.maxMarks : '',
      metadata.duration || '',
      metadata.separateAnswerSheet ? 'true' : 'false',
      metadata.language || ''
    ];

    const rows = [];
    if (sections.length === 0) {
      rows.push([
        ...getBrandingMetaCols(),
        '', '', '', '', '', '', '', '', '', '', '', '', ''
      ]);
    } else {
      sections.forEach((sec) => {
        if (sec.questions.length === 0) {
          rows.push([
            ...getBrandingMetaCols(),
            sec.title || '',
            sec.marks !== undefined && sec.marks !== null ? sec.marks : '',
            sec.instructions || '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            ''
          ]);
        } else {
          sec.questions.forEach((q) => {
            let optionsStr = '';
            if (sec.type === 'mcq' && q.options) {
              optionsStr = q.options.join(';');
            }

            let matchPairsStr = '';
            if (sec.type === 'match_following' && q.matchPairs) {
              const hasImages = q.matchPairs.some(p => typeof p === 'object' && (p.premiseImage || p.responseImage));
              if (hasImages) {
                matchPairsStr = JSON.stringify(q.matchPairs);
              } else {
                matchPairsStr = q.matchPairs.map(p => `${typeof p === 'string' ? p : (p.premise || '')}=${typeof p === 'string' ? p : (p.response || '')}`).join(';');
              }
            }

            let subQsStr = '';
            if (q.subQuestions && q.subQuestions.length > 0) {
              subQsStr = JSON.stringify(q.subQuestions);
            }

            rows.push([
              ...getBrandingMetaCols(),
              sec.title || '',
              sec.marks !== undefined && sec.marks !== null ? sec.marks : '',
              sec.instructions || '',
              sec.type || 'essay',
              q.text || '',
              getQuestionMarks(q),
              optionsStr,
              q.blankLines !== undefined && q.blankLines !== null ? q.blankLines : '',
              matchPairsStr,
              q.image || '',
              q.imageWidth || '',
              q.imageHeight || '',
              subQsStr
            ]);
          });
        }
      });
    }

    const csvContent = [
      headers.map(h => `"${h.replace(/"/g, '""')}"`).join(','),
      ...rows.map(row => row.map(val => {
        const valStr = val === undefined || val === null ? '' : String(val);
        return `"${valStr.replace(/"/g, '""')}"`;
      }).join(','))
    ].join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', generateExportFilename('csv'));
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const parseCSV = (text) => {
    const lines = [];
    let row = [""];
    let insideQuote = false;

    for (let i = 0; i < text.length; i++) {
      const char = text[i];
      const nextChar = text[i + 1];

      if (char === '"') {
        if (insideQuote && nextChar === '"') {
          row[row.length - 1] += '"';
          i++;
        } else {
          insideQuote = !insideQuote;
        }
      } else if (char === ',' && !insideQuote) {
        row.push("");
      } else if ((char === '\r' || char === '\n') && !insideQuote) {
        if (char === '\r' && nextChar === '\n') {
          i++;
        }
        lines.push(row);
        row = [""];
      } else {
        row[row.length - 1] += char;
      }
    }
    if (row.length > 1 || row[0] !== "") {
      lines.push(row);
    }
    return lines;
  };

  const importFromCSV = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const text = event.target.result;
        const parsedRows = parseCSV(text);
        if (parsedRows.length < 2) {
          alert('Invalid CSV file.');
          return;
        }

        const headerRow = parsedRows[0].map(h => (h || '').trim().toLowerCase());
        const getColIdx = (name) => headerRow.indexOf(name.toLowerCase());

        // Header column mappings
        const schoolLogoIdx = getColIdx('School Logo');
        const schoolLogoWidthIdx = getColIdx('School Logo Width');
        const schoolLogoHeightIdx = getColIdx('School Logo Height');
        const hideSchoolLogoIdx = getColIdx('Hide School Logo');
        const headerLogoOnlyIdx = getColIdx('Header Logo Only');
        const schoolNameIdx = getColIdx('School Name');
        const schoolAddressIdx = getColIdx('School Address');
        const fontFamilyIdx = getColIdx('Font Family');

        const examTitleIdx = getColIdx('Exam Title');
        const subjectIdx = getColIdx('Subject');
        const classDivIdx = getColIdx('Class Div');
        const maxMarksIdx = getColIdx('Max Marks');
        const durationIdx = getColIdx('Duration');
        const separateAnswerSheetIdx = getColIdx('Separate Answer Sheet');
        const languageIdx = getColIdx('Language');

        // Section & Question column mappings with legacy index fallbacks
        const secTitleIdx = getColIdx('Section Title') !== -1 ? getColIdx('Section Title') : 0;
        const secMarksIdx = getColIdx('Section Marks') !== -1 ? getColIdx('Section Marks') : 1;
        const secInstructionsIdx = getColIdx('Section Instructions') !== -1 ? getColIdx('Section Instructions') : 2;
        const qTypeIdx = getColIdx('Question Type') !== -1 ? getColIdx('Question Type') : 3;
        const qTextIdx = getColIdx('Question Text') !== -1 ? getColIdx('Question Text') : 4;
        const qMarksIdx = getColIdx('Question Marks') !== -1 ? getColIdx('Question Marks') : 5;
        const optionsIdx = getColIdx('Options') !== -1 ? getColIdx('Options') : 6;
        const blankLinesIdx = getColIdx('Blank Lines') !== -1 ? getColIdx('Blank Lines') : 7;
        const matchPairsIdx = getColIdx('Match Pairs') !== -1 ? getColIdx('Match Pairs') : 8;
        const imageDataIdx = getColIdx('Image Data') !== -1 ? getColIdx('Image Data') : 9;
        const imageWidthIdx = getColIdx('Image Width') !== -1 ? getColIdx('Image Width') : 10;
        const imageHeightIdx = getColIdx('Image Height') !== -1 ? getColIdx('Image Height') : 11;
        const subQsIdx = getColIdx('Sub Questions') !== -1 ? getColIdx('Sub Questions') : 12;

        const importedSections = [];
        let importedBranding = null;
        let importedMetadata = null;
        let currentSection = null;

        for (let i = 1; i < parsedRows.length; i++) {
          const row = parsedRows[i];
          if (!row || row.length === 0) continue;

          // Extract Branding if present and not yet extracted
          if (!importedBranding && (schoolNameIdx !== -1 || schoolAddressIdx !== -1 || schoolLogoIdx !== -1)) {
            const logoVal = schoolLogoIdx !== -1 ? row[schoolLogoIdx] : undefined;
            const logoWidthVal = schoolLogoWidthIdx !== -1 ? row[schoolLogoWidthIdx] : undefined;
            const logoHeightVal = schoolLogoHeightIdx !== -1 ? row[schoolLogoHeightIdx] : undefined;
            const hideLogoVal = hideSchoolLogoIdx !== -1 ? row[hideSchoolLogoIdx] : undefined;
            const headerLogoOnlyVal = headerLogoOnlyIdx !== -1 ? row[headerLogoOnlyIdx] : undefined;
            const schoolNameVal = schoolNameIdx !== -1 ? row[schoolNameIdx] : undefined;
            const schoolAddressVal = schoolAddressIdx !== -1 ? row[schoolAddressIdx] : undefined;
            const fontFamilyVal = fontFamilyIdx !== -1 ? row[fontFamilyIdx] : undefined;

            if (
              logoVal !== undefined || schoolNameVal !== undefined || schoolAddressVal !== undefined ||
              logoWidthVal !== undefined || fontFamilyVal !== undefined
            ) {
              importedBranding = {};
              if (logoVal !== undefined && logoVal !== '') importedBranding.logo = normalizeLogo(logoVal);
              if (logoWidthVal !== undefined && logoWidthVal !== '' && !isNaN(logoWidthVal)) importedBranding.logoWidth = Number(logoWidthVal);
              if (logoHeightVal !== undefined && logoHeightVal !== '' && !isNaN(logoHeightVal)) importedBranding.logoHeight = Number(logoHeightVal);
              if (hideLogoVal !== undefined && hideLogoVal !== '') importedBranding.hideSchoolLogo = hideLogoVal === 'true';
              if (headerLogoOnlyVal !== undefined && headerLogoOnlyVal !== '') importedBranding.headerLogoOnly = headerLogoOnlyVal === 'true';
              if (schoolNameVal !== undefined && schoolNameVal !== '') importedBranding.schoolName = schoolNameVal;
              if (schoolAddressVal !== undefined && schoolAddressVal !== '') importedBranding.schoolAddress = schoolAddressVal;
              if (fontFamilyVal !== undefined && fontFamilyVal !== '') importedBranding.fontFamily = fontFamilyVal;
            }
          }

          // Extract Metadata if present and not yet extracted
          if (!importedMetadata && (examTitleIdx !== -1 || subjectIdx !== -1 || classDivIdx !== -1 || maxMarksIdx !== -1)) {
            const titleVal = examTitleIdx !== -1 ? row[examTitleIdx] : undefined;
            const subjectVal = subjectIdx !== -1 ? row[subjectIdx] : undefined;
            const classDivVal = classDivIdx !== -1 ? row[classDivIdx] : undefined;
            const maxMarksVal = maxMarksIdx !== -1 ? row[maxMarksIdx] : undefined;
            const durationVal = durationIdx !== -1 ? row[durationIdx] : undefined;
            const sepAnsVal = separateAnswerSheetIdx !== -1 ? row[separateAnswerSheetIdx] : undefined;
            const langVal = languageIdx !== -1 ? row[languageIdx] : undefined;

            if (
              titleVal !== undefined || subjectVal !== undefined || classDivVal !== undefined ||
              maxMarksVal !== undefined || durationVal !== undefined
            ) {
              importedMetadata = {};
              if (titleVal !== undefined && titleVal !== '') importedMetadata.title = titleVal;
              if (subjectVal !== undefined && subjectVal !== '') importedMetadata.subject = subjectVal;
              if (classDivVal !== undefined && classDivVal !== '') importedMetadata.classDiv = classDivVal;
              if (maxMarksVal !== undefined && maxMarksVal !== '' && !isNaN(maxMarksVal)) importedMetadata.maxMarks = Number(maxMarksVal);
              if (durationVal !== undefined && durationVal !== '') importedMetadata.duration = durationVal;
              if (sepAnsVal !== undefined && sepAnsVal !== '') importedMetadata.separateAnswerSheet = sepAnsVal === 'true';
              if (langVal !== undefined && langVal !== '') importedMetadata.language = langVal;
            }
          }

          const secTitle = row[secTitleIdx] || '';
          const secMarks = Math.max(0, Math.round((Number(row[secMarksIdx]) || 0) * 100) / 100);
          const secInstructions = row[secInstructionsIdx] || '';
          const qType = row[qTypeIdx] || '';
          const qText = row[qTextIdx] || '';
          const qMarks = Math.max(0, Math.round((Number(row[qMarksIdx]) || 0) * 100) / 100);
          const optionsStr = row[optionsIdx] || '';
          const blankLinesVal = row[blankLinesIdx] || '';
          const matchPairsStr = row[matchPairsIdx] || '';
          const imageData = row[imageDataIdx] || '';
          const imageWidthVal = row[imageWidthIdx] || '';
          const imageHeightVal = row[imageHeightIdx] || '';
          const subQsStr = row[subQsIdx] || '';

          if (!secTitle && !qText) continue;

          if (secTitle && (!currentSection || currentSection.title !== secTitle)) {
            currentSection = {
              id: `sec-${Date.now()}-${i}`,
              title: secTitle,
              marks: secMarks,
              instructions: secInstructions,
              type: qType || 'essay',
              questions: []
            };
            importedSections.push(currentSection);
          }

          if (!currentSection) {
            currentSection = {
              id: `sec-${Date.now()}-${i}`,
              title: 'Imported Section',
              marks: 0,
              instructions: '',
              type: qType || 'essay',
              questions: []
            };
            importedSections.push(currentSection);
          }

          if (qType && qText) {
            const q = {
              id: `q-${Date.now()}-${i}`,
              type: qType,
              text: qText,
              marks: qMarks
            };

            if (subQsStr) {
              try {
                q.subQuestions = JSON.parse(subQsStr);
                q.marks = q.subQuestions.reduce((sum, sq) => sum + (Number(sq.marks) || 0), 0);
              } catch (e) {
                console.warn('Failed to parse subQuestions from CSV', e);
              }
            }

            if (qType === 'mcq') {
              q.options = optionsStr ? optionsStr.split(';') : ['', '', '', ''];
            } else if (qType === 'essay') {
              q.blankLines = (blankLinesVal !== '' && !isNaN(blankLinesVal)) ? Math.max(0, parseInt(blankLinesVal, 10)) : 5;
            } else if (qType === 'match_following') {
              if (matchPairsStr) {
                const trimmed = matchPairsStr.trim();
                if (trimmed.startsWith('[') || trimmed.startsWith('{')) {
                  try {
                    const parsed = JSON.parse(trimmed);
                    q.matchPairs = Array.isArray(parsed) ? parsed.map(p => ({
                      premise: typeof p === 'string' ? p : (p.premise || ''),
                      premiseImage: typeof p === 'object' ? (p.premiseImage || '') : '',
                      response: typeof p === 'string' ? p : (p.response || ''),
                      responseImage: typeof p === 'object' ? (p.responseImage || '') : ''
                    })) : [];
                  } catch (e) {
                    q.matchPairs = matchPairsStr.split(';').map(pair => {
                      const parts = pair.split('=');
                      return { premise: parts[0] || '', premiseImage: '', response: parts[1] || '', responseImage: '' };
                    });
                  }
                } else {
                  q.matchPairs = matchPairsStr.split(';').map(pair => {
                    const parts = pair.split('=');
                    return { premise: parts[0] || '', premiseImage: '', response: parts[1] || '', responseImage: '' };
                  });
                }
              } else {
                q.matchPairs = [];
              }
            }

            if (imageData) {
              q.image = imageData;
              q.imageWidth = Number(imageWidthVal) || 300;
              q.imageHeight = Number(imageHeightVal) || 200;
            }

            if (q.type === 'image') {
              q.type = 'essay';
            }

            currentSection.questions.push(q);
          }
        }

        const hasSections = importedSections.length > 0;
        const hasBranding = importedBranding && Object.keys(importedBranding).length > 0;
        const hasMeta = importedMetadata && Object.keys(importedMetadata).length > 0;

        if (hasSections || hasBranding || hasMeta) {
          setCsvImportModal({
            isOpen: true,
            branding: hasBranding ? importedBranding : null,
            metadata: hasMeta ? importedMetadata : null,
            sections: hasSections ? importedSections : [],
            importSchool: false,
            importExam: hasMeta,
            importQuestions: hasSections
          });
        } else {
          alert('No valid sections, questions, or branding/exam details found in the CSV.');
        }
      } catch (err) {
        console.error(err);
        alert('Error parsing CSV file. Please make sure the format is correct.');
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  const confirmCsvImport = () => {
    const { importSchool, importExam, importQuestions, branding: iBranding, metadata: iMetadata, sections: iSections } = csvImportModal;
    if (!importSchool && !importExam && !importQuestions) {
      alert('Please select at least one category to import.');
      return;
    }
    if (importQuestions && iSections.length > 0) {
      setSections(iSections);
    }
    if (importSchool && iBranding && Object.keys(iBranding).length > 0) {
      setBranding(prev => ({
        ...prev,
        ...iBranding,
        logo: iBranding.logo !== undefined ? normalizeLogo(iBranding.logo) : prev.logo
      }));
    }
    if (importExam && iMetadata && Object.keys(iMetadata).length > 0) {
      setMetadata(prev => ({ ...prev, ...iMetadata }));
    }
    setCsvImportModal({ isOpen: false, branding: null, metadata: null, sections: [], importSchool: false, importExam: true, importQuestions: true });
  };

  // Print PDF Trigger
  const triggerPrint = () => {
    if (hasBlankQuestions()) {
      if (!window.confirm("Some questions have empty text. Are you sure you want to print?")) {
        return;
      }
    }
    window.print();
  };

  const triggerPdfExport = async () => {
    if (!hasQuestions()) {
      alert("No questions added to export.");
      return;
    }
    if (hasBlankQuestions()) {
      if (!window.confirm("Some questions have empty text. Are you sure you want to export?")) {
        return;
      }
    }
    const el = paperSheetRef.current;
    if (!el) {
      alert('Please open the preview first.');
      return;
    }
    setIsPdfExporting(true);
    let footer = null;
    try {
      // Hide the paper-footer element during capture so we can stamp custom page numbers
      footer = el.querySelector('.paper-footer');
      if (footer) footer.style.display = 'none';

      const filename = generateExportFilename('pdf');

      const pdf = new jsPDF({
        unit: 'mm',
        format: 'a4',
        orientation: 'portrait',
        compress: true
      });

      const elementWidth = el.offsetWidth || el.clientWidth || 794;

      await pdf.html(el, {
        x: 0,
        y: 0,
        width: 210, // A4 width in mm
        windowWidth: elementWidth,
        margin: [10, 0, 18, 0], // top, right, bottom, left in mm
        autoPaging: 'text',
        html2canvas: {
          useCORS: true,
          allowTaint: true,
          backgroundColor: '#ffffff',
          logging: false
        }
      });

      const totalPages = pdf.internal.getNumberOfPages();
      for (let i = 1; i <= totalPages; i++) {
        pdf.setPage(i);
        pdf.setFontSize(13);
        pdf.setTextColor(120, 120, 120);
        pdf.text(`Page ${i} of ${totalPages}`, 198, 287, { align: 'right' });
      }

      pdf.save(filename);
    } catch (err) {
      console.error('PDF export failed:', err);
      alert('PDF export failed. Please try again or use Print instead.');
    } finally {
      if (footer) footer.style.display = '';
      setIsPdfExporting(false);
    }
  };

  // Helper to convert Data URL to Uint8Array for docx ImageRun
  const dataURLToUint8Array = (dataURL) => {
    if (!dataURL) return null;
    const parts = dataURL.split(';base64,');
    if (parts.length < 2) return null;
    const base64 = parts[1];
    try {
      const raw = window.atob(base64);
      const rawLength = raw.length;
      const array = new Uint8Array(new ArrayBuffer(rawLength));
      for (let i = 0; i < rawLength; i++) {
        array[i] = raw.charCodeAt(i);
      }
      return array;
    } catch (e) {
      console.error("Failed to decode base64 image", e);
      return null;
    }
  };

  const getFontFamily = () => {
    if (metadata.language === 'malayalam') return 'Manjari';
    if (metadata.language === 'hindi') return 'Noto Sans Devanagari';
    return branding.fontFamily === 'Inter' ? 'Calibri' : 'Times New Roman';
  };

  const imageToUint8Array = async (src) => {
    if (!src) return null;
    const resolvedSrc = normalizeLogo(src);
    if (!resolvedSrc) return null;
    if (resolvedSrc.startsWith('data:')) {
      const uint8 = dataURLToUint8Array(resolvedSrc);
      if (uint8) return uint8;
    }
    try {
      const response = await fetch(resolvedSrc);
      const blob = await response.blob();
      const arrayBuffer = await blob.arrayBuffer();
      return new Uint8Array(arrayBuffer);
    } catch (e) {
      console.error("Failed to fetch image from URL for docx export", e);
      if (resolvedSrc !== schoolLogo) {
        return imageToUint8Array(schoolLogo);
      }
      return null;
    }
  };

  // DOCX Export Implementation
  const generateDocxBlob = async () => {
    // Create the School branding header
    const headerChildren = [];

    // Fetch and load logo data if active and not hidden
    let logoData = null;
    if (branding.logo && !branding.hideSchoolLogo) {
      logoData = await imageToUint8Array(branding.logo);
    }

    const schoolDetailsParagraphs = [];
    if (!branding.headerLogoOnly && branding.schoolName) {
      schoolDetailsParagraphs.push(
        new docx.Paragraph({
          alignment: docx.AlignmentType.CENTER,
          children: [
            new docx.TextRun({
              text: (branding.schoolName || '').toUpperCase(),
              bold: true,
              size: 28,
              font: getFontFamily()
            })
          ],
          spacing: { after: 120 }
        })
      );
    }

    if (!branding.headerLogoOnly && branding.schoolAddress) {
      const addressLines = branding.schoolAddress.split('\n');
      addressLines.forEach(line => {
        schoolDetailsParagraphs.push(
          new docx.Paragraph({
            alignment: docx.AlignmentType.CENTER,
            children: [
              new docx.TextRun({
                text: line || '',
                size: 20,
                font: getFontFamily()
              })
            ],
            spacing: { after: 80 }
          })
        );
      });
    }

    if (logoData) {
      const logoRun = new docx.ImageRun({
        data: logoData,
        transformation: {
          width: branding.logoWidth || 100,
          height: branding.logoHeight || 100
        }
      });

      if (branding.headerLogoOnly) {
        headerChildren.push(
          new docx.Paragraph({
            alignment: docx.AlignmentType.CENTER,
            children: [logoRun],
            spacing: { before: -200, after: 200 }
          })
        );
      } else {
        // Construct side-by-side layout using borderless table with explicit DXA widths (prevents Mobile Google Docs 1-letter column collapse)
        const headerTable = new docx.Table({
          width: { size: 9000, type: docx.WidthType.DXA },
          columnWidths: [1800, 7200],
          borders: {
            top: { style: docx.BorderStyle.NONE, size: 0 },
            bottom: { style: docx.BorderStyle.NONE, size: 0 },
            left: { style: docx.BorderStyle.NONE, size: 0 },
            right: { style: docx.BorderStyle.NONE, size: 0 },
            insideHorizontal: { style: docx.BorderStyle.NONE, size: 0 },
            insideVertical: { style: docx.BorderStyle.NONE, size: 0 }
          },
          rows: [
            new docx.TableRow({
              children: [
                new docx.TableCell({
                  width: { size: 1800, type: docx.WidthType.DXA },
                  borders: {
                    top: { style: docx.BorderStyle.NONE, size: 0 },
                    bottom: { style: docx.BorderStyle.NONE, size: 0 },
                    left: { style: docx.BorderStyle.NONE, size: 0 },
                    right: { style: docx.BorderStyle.NONE, size: 0 }
                  },
                  children: [
                    new docx.Paragraph({
                      alignment: docx.AlignmentType.CENTER,
                      children: [logoRun]
                    })
                  ]
                }),
                new docx.TableCell({
                  width: { size: 7200, type: docx.WidthType.DXA },
                  borders: {
                    top: { style: docx.BorderStyle.NONE, size: 0 },
                    bottom: { style: docx.BorderStyle.NONE, size: 0 },
                    left: { style: docx.BorderStyle.NONE, size: 0 },
                    right: { style: docx.BorderStyle.NONE, size: 0 }
                  },
                  children: schoolDetailsParagraphs
                })
              ]
            })
          ]
        });
        headerChildren.push(headerTable);
      }
    } else {
      headerChildren.push(...schoolDetailsParagraphs);
    }

    // Divider line (Double bottom border matching PDF)
    headerChildren.push(
      new docx.Paragraph({
        border: {
          bottom: {
            style: docx.BorderStyle.DOUBLE,
            size: 18,
            color: '000000',
            space: 6
          }
        },
        spacing: { after: 200 }
      })
    );

    // Helper to generate a metadata row with labels on the left and input data aligned to the right margin
    const createMetaParagraph = (label, value) => {
      return new docx.Paragraph({
        tabStops: [
          {
            type: docx.TabStopType.RIGHT,
            position: docx.TabStopPosition.MAX
          }
        ],
        children: [
          new docx.TextRun({ text: label, bold: true, size: 22 }),
          new docx.TextRun({ text: `\t${value || ''}`, size: 22 })
        ],
        spacing: { after: 120 }
      });
    };

    headerChildren.push(createMetaParagraph('Examination: ', metadata.title));
    headerChildren.push(createMetaParagraph('Subject: ', metadata.subject));
    headerChildren.push(createMetaParagraph('Class: ', metadata.classDiv));
    headerChildren.push(createMetaParagraph('Max Marks: ', formatMarks(metadata.maxMarks)));
    headerChildren.push(createMetaParagraph('Duration: ', metadata.duration));

    // Bottom border for metadata
    headerChildren.push(
      new docx.Paragraph({
        border: {
          bottom: {
            style: docx.BorderStyle.SINGLE,
            size: 8,
            color: '000000',
            space: 6
          }
        },
        spacing: { after: 240 }
      })
    );

    // Now populate sections and questions
    let absoluteQuestionCount = 1;

    for (const sec of sections) {
      const hasInstructions = !!sec.instructions;
      // Section header with bottom border line matching PDF (below instructions if present)
      headerChildren.push(
        new docx.Paragraph({
          alignment: docx.AlignmentType.JUSTIFY,
          border: hasInstructions ? undefined : {
            bottom: {
              style: docx.BorderStyle.SINGLE,
              size: 8,
              color: '000000',
              space: 4
            }
          },
          tabStops: [
            {
              type: docx.TabStopType.RIGHT,
              position: docx.TabStopPosition.MAX
            }
          ],
          spacing: { before: 240, after: hasInstructions ? 40 : 80 },
          children: [
            new docx.TextRun({
              text: (sec.title || '').toUpperCase(),
              bold: true,
              size: 24,
              font: getFontFamily()
            }),
            new docx.TextRun({
              text: `\t[${formatMarks(sec.marks)} Marks]`,
              bold: true,
              size: 22,
              font: getFontFamily()
            })
          ]
        })
      );

      // Section instructions
      if (sec.instructions) {
        headerChildren.push(
          new docx.Paragraph({
            alignment: docx.AlignmentType.JUSTIFY,
            border: {
              bottom: {
                style: docx.BorderStyle.SINGLE,
                size: 8,
                color: '000000',
                space: 4
              }
            },
            spacing: { before: 40, after: 80 },
            children: [
              new docx.TextRun({
                text: sec.instructions || '',
                italic: true,
                size: 20,
                font: getFontFamily()
              })
            ]
          })
        );
      }

      // Add each question
      for (const q of sec.questions) {
        const qNum = `Q${absoluteQuestionCount}.`;
        absoluteQuestionCount++;

        // Add question text matching PDF (Q1. and (1 M))
        const qLines = (q.text || '').split('\n');
        headerChildren.push(
          new docx.Paragraph({
            alignment: docx.AlignmentType.JUSTIFY,
            spacing: { before: 120, after: 40 },
            tabStops: [
              {
                type: docx.TabStopType.RIGHT,
                position: docx.TabStopPosition.MAX
              }
            ],
            children: [
              new docx.TextRun({
                text: `${qNum}  `,
                bold: true,
                size: 22
              }),
              ...docxTextRunsWithMath(qLines[0] || ''),
              new docx.TextRun({
                text: `\t(${formatMarks(getQuestionMarks(q))} M)`,
                italic: true,
                size: 20
              })
            ]
          })
        );

        for (let lIdx = 1; lIdx < qLines.length; lIdx++) {
          headerChildren.push(
            new docx.Paragraph({
              indent: { left: 450 },
              spacing: { after: 40 },
              children: docxTextRunsWithMath(qLines[lIdx])
            })
          );
        }

        // Formatting specific question types
        if (sec.type === 'mcq' && q.options) {
          const createOptionParagraphs = (letter, text) => {
            if (!text) return [new docx.Paragraph({ children: [] })];
            const lines = text.split('\n');
            const paragraphs = [];
            paragraphs.push(
              new docx.Paragraph({
                spacing: { after: 40 },
                children: [
                  new docx.TextRun({ text: `(${letter})  `, bold: true, size: 22 }),
                  ...docxTextRunsWithMath(lines[0] || '')
                ]
              })
            );
            for (let i = 1; i < lines.length; i++) {
              paragraphs.push(
                new docx.Paragraph({
                  indent: { left: 450 },
                  spacing: { after: 40 },
                  children: docxTextRunsWithMath(lines[i])
                })
              );
            }
            return paragraphs;
          };

          const isSingle = canFitSingleLine(q.options);

          if (isSingle) {
            const numOpts = q.options.length;
            const colWidth = Math.floor(9000 / numOpts);
            const cells = q.options.map((opt, oIdx) => {
              const letter = String.fromCharCode(65 + oIdx);
              return new docx.TableCell({
                width: { size: colWidth, type: docx.WidthType.DXA },
                borders: {
                  top: { style: docx.BorderStyle.NONE, size: 0 },
                  bottom: { style: docx.BorderStyle.NONE, size: 0 },
                  left: { style: docx.BorderStyle.NONE, size: 0 },
                  right: { style: docx.BorderStyle.NONE, size: 0 }
                },
                children: createOptionParagraphs(letter, opt)
              });
            });

            headerChildren.push(
              new docx.Table({
                width: { size: 9000, type: docx.WidthType.DXA },
                columnWidths: Array(numOpts).fill(colWidth),
                borders: {
                  top: { style: docx.BorderStyle.NONE, size: 0 },
                  bottom: { style: docx.BorderStyle.NONE, size: 0 },
                  left: { style: docx.BorderStyle.NONE, size: 0 },
                  right: { style: docx.BorderStyle.NONE, size: 0 },
                  insideHorizontal: { style: docx.BorderStyle.NONE, size: 0 },
                  insideVertical: { style: docx.BorderStyle.NONE, size: 0 }
                },
                rows: [new docx.TableRow({ children: cells })]
              })
            );
          } else {
            const optRows = [];
            for (let i = 0; i < q.options.length; i += 2) {
              const leftLetter = String.fromCharCode(65 + i);
              const leftText = q.options[i] || '';
              const rightLetter = i + 1 < q.options.length ? String.fromCharCode(65 + i + 1) : '';
              const rightText = i + 1 < q.options.length ? (q.options[i + 1] || '') : '';

              const leftChildren = createOptionParagraphs(leftLetter, leftText);
              const rightChildren = rightLetter ? createOptionParagraphs(rightLetter, rightText) : [new docx.Paragraph({ children: [] })];

              optRows.push(
                new docx.TableRow({
                  children: [
                    new docx.TableCell({
                      width: { size: 4500, type: docx.WidthType.DXA },
                      borders: {
                        top: { style: docx.BorderStyle.NONE, size: 0 },
                        bottom: { style: docx.BorderStyle.NONE, size: 0 },
                        left: { style: docx.BorderStyle.NONE, size: 0 },
                        right: { style: docx.BorderStyle.NONE, size: 0 }
                      },
                      children: leftChildren
                    }),
                    new docx.TableCell({
                      width: { size: 4500, type: docx.WidthType.DXA },
                      borders: {
                        top: { style: docx.BorderStyle.NONE, size: 0 },
                        bottom: { style: docx.BorderStyle.NONE, size: 0 },
                        left: { style: docx.BorderStyle.NONE, size: 0 },
                        right: { style: docx.BorderStyle.NONE, size: 0 }
                      },
                      children: rightChildren
                    })
                  ]
                })
              );
            }

            headerChildren.push(
              new docx.Table({
                width: { size: 9000, type: docx.WidthType.DXA },
                columnWidths: [4500, 4500],
                borders: {
                  top: { style: docx.BorderStyle.NONE, size: 0 },
                  bottom: { style: docx.BorderStyle.NONE, size: 0 },
                  left: { style: docx.BorderStyle.NONE, size: 0 },
                  right: { style: docx.BorderStyle.NONE, size: 0 },
                  insideHorizontal: { style: docx.BorderStyle.NONE, size: 0 },
                  insideVertical: { style: docx.BorderStyle.NONE, size: 0 }
                },
                rows: optRows
              })
            );
          }
        }

        else if (sec.type === 'essay') {
          if (q.subQuestions && q.subQuestions.length > 0) {
            for (let sqIdx = 0; sqIdx < q.subQuestions.length; sqIdx++) {
              const sq = q.subQuestions[sqIdx];
              const sqLabel = sq.label || `(${String.fromCharCode(97 + (sqIdx % 26))})`;
              const sqLines = (sq.text || '').split('\n');
              headerChildren.push(
                new docx.Paragraph({
                  alignment: docx.AlignmentType.JUSTIFY,
                  indent: { left: 360 },
                  spacing: { before: 80, after: 40 },
                  tabStops: [
                    {
                      type: docx.TabStopType.RIGHT,
                      position: docx.TabStopPosition.MAX
                    }
                  ],
                  children: [
                    new docx.TextRun({
                      text: `${sqLabel}  `,
                      bold: true,
                      size: 22
                    }),
                    ...docxTextRunsWithMath(sqLines[0] || ''),
                    new docx.TextRun({
                      text: `\t(${formatMarks(sq.marks)} M)`,
                      italic: true,
                      size: 20
                    })
                  ]
                })
              );
              for (let lIdx = 1; lIdx < sqLines.length; lIdx++) {
                headerChildren.push(
                  new docx.Paragraph({
                    indent: { left: 720 },
                    spacing: { after: 40 },
                    children: docxTextRunsWithMath(sqLines[lIdx])
                  })
                );
              }

              if (sq.image) {
                const sqImageBytes = await imageToUint8Array(sq.image);
                if (sqImageBytes) {
                  headerChildren.push(
                    new docx.Paragraph({
                      indent: { left: 720 },
                      spacing: { before: 80, after: 80 },
                      children: [
                        new docx.ImageRun({
                          data: sqImageBytes,
                          transformation: {
                            width: sq.imageWidth || 300,
                            height: sq.imageHeight || 200
                          }
                        })
                      ]
                    })
                  );
                }
              }

              if (!metadata.separateAnswerSheet) {
                const sqLines = (sq.blankLines !== undefined && sq.blankLines !== '') ? sq.blankLines : 4;
                for (let i = 0; i < sqLines; i++) {
                  headerChildren.push(
                    new docx.Paragraph({
                      indent: { left: 360 },
                      border: {
                        bottom: {
                          style: docx.BorderStyle.SINGLE,
                          size: 8,
                          color: '333333',
                          space: 2
                        }
                      },
                      spacing: { before: 40, after: 80 },
                      children: [new docx.TextRun({ text: '' })]
                    })
                  );
                }
              }
            }
          } else if (!metadata.separateAnswerSheet) {
            // Renders specified blank lines
            const linesCount = (q.blankLines !== undefined && q.blankLines !== '') ? q.blankLines : 5;
            for (let i = 0; i < linesCount; i++) {
              headerChildren.push(
                new docx.Paragraph({
                  border: {
                    bottom: {
                      style: docx.BorderStyle.SINGLE,
                      size: 8,
                      color: '333333',
                      space: 2
                    }
                  },
                  spacing: { before: 40, after: 80 },
                  children: [new docx.TextRun({ text: '' })]
                })
              );
            }
          }
        }

        else if (sec.type === 'true_false' && !metadata.separateAnswerSheet) {
          headerChildren.push(
            new docx.Paragraph({
              indent: { left: 720 },
              spacing: { after: 60 },
              children: [
                new docx.TextRun({
                  text: '[    ] True        [    ] False',
                  bold: true,
                  size: 20
                })
              ]
            })
          );
        }

        else if (sec.type === 'match_following' && q.matchPairs) {
          const columnA = q.matchPairs.map(p => ({ text: p.premise, image: p.premiseImage }));
          const shuffledList = getShuffledList(q);
          const columnB = q.matchPairs.map((pair, pIdx) => {
            const itemB = shuffledList[pIdx] || (typeof pair === 'string' ? { response: pair } : pair);
            return {
              text: typeof itemB === 'string' ? itemB : (itemB.response || ''),
              image: typeof itemB === 'object' ? (itemB.responseImage || '') : ''
            };
          });

          const tableRows = [];

          // Header row
          tableRows.push(
            new docx.TableRow({
              children: [
                new docx.TableCell({
                  width: { size: 4500, type: docx.WidthType.DXA },
                  borders: {
                    top: { style: docx.BorderStyle.NONE, size: 0 },
                    bottom: { style: docx.BorderStyle.SINGLE, size: 4, color: '000000' },
                    left: { style: docx.BorderStyle.NONE, size: 0 },
                    right: { style: docx.BorderStyle.NONE, size: 0 }
                  },
                  children: [
                    new docx.Paragraph({
                      spacing: { before: 60, after: 60 },
                      children: [new docx.TextRun({ text: 'Column A', bold: true, size: 22 })]
                    })
                  ]
                }),
                new docx.TableCell({
                  width: { size: 4500, type: docx.WidthType.DXA },
                  borders: {
                    top: { style: docx.BorderStyle.NONE, size: 0 },
                    bottom: { style: docx.BorderStyle.SINGLE, size: 4, color: '000000' },
                    left: { style: docx.BorderStyle.NONE, size: 0 },
                    right: { style: docx.BorderStyle.NONE, size: 0 }
                  },
                  children: [
                    new docx.Paragraph({
                      spacing: { before: 60, after: 60 },
                      children: [new docx.TextRun({ text: 'Column B', bold: true, size: 22 })]
                    })
                  ]
                })
              ]
            })
          );

          const romanNum = (idx) => {
            const r = ['i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x'];
            return r[idx] || (idx + 1).toString();
          };

          for (let index = 0; index < columnA.length; index++) {
            const itemA = columnA[index];
            const itemB = columnB[index];

            const cellAChildren = [
              new docx.Paragraph({
                spacing: { before: 40, after: 40 },
                children: [
                  new docx.TextRun({ text: `${index + 1}. `, bold: true, size: 22 }),
                  ...docxTextRunsWithMath(itemA.text || '')
                ]
              })
            ];
            if (itemA.image) {
              const imgBytes = await imageToUint8Array(itemA.image);
              if (imgBytes) {
                cellAChildren.push(
                  new docx.Paragraph({
                    spacing: { before: 20, after: 40 },
                    children: [
                      new docx.ImageRun({
                        data: imgBytes,
                        transformation: { width: 120, height: 80 }
                      })
                    ]
                  })
                );
              }
            }

            const cellBChildren = [
              new docx.Paragraph({
                spacing: { before: 40, after: 40 },
                children: [
                  new docx.TextRun({ text: `${romanNum(index)}. `, bold: true, size: 22 }),
                  ...docxTextRunsWithMath(itemB.text || '')
                ]
              })
            ];
            if (itemB.image) {
              const imgBytes = await imageToUint8Array(itemB.image);
              if (imgBytes) {
                cellBChildren.push(
                  new docx.Paragraph({
                    spacing: { before: 20, after: 40 },
                    children: [
                      new docx.ImageRun({
                        data: imgBytes,
                        transformation: { width: 120, height: 80 }
                      })
                    ]
                  })
                );
              }
            }

            tableRows.push(
              new docx.TableRow({
                children: [
                  new docx.TableCell({
                    width: { size: 4500, type: docx.WidthType.DXA },
                    borders: {
                      top: { style: docx.BorderStyle.NONE, size: 0 },
                      bottom: { style: docx.BorderStyle.NONE, size: 0 },
                      left: { style: docx.BorderStyle.NONE, size: 0 },
                      right: { style: docx.BorderStyle.NONE, size: 0 }
                    },
                    children: cellAChildren
                  }),
                  new docx.TableCell({
                    width: { size: 4500, type: docx.WidthType.DXA },
                    borders: {
                      top: { style: docx.BorderStyle.NONE, size: 0 },
                      bottom: { style: docx.BorderStyle.NONE, size: 0 },
                      left: { style: docx.BorderStyle.NONE, size: 0 },
                      right: { style: docx.BorderStyle.NONE, size: 0 }
                    },
                    children: cellBChildren
                  })
                ]
              })
            );
          }

          headerChildren.push(
            new docx.Table({
              width: { size: 9000, type: docx.WidthType.DXA },
              columnWidths: [4500, 4500],
              borders: {
                top: { style: docx.BorderStyle.NONE, size: 0 },
                bottom: { style: docx.BorderStyle.NONE, size: 0 },
                left: { style: docx.BorderStyle.NONE, size: 0 },
                right: { style: docx.BorderStyle.NONE, size: 0 },
                insideHorizontal: { style: docx.BorderStyle.NONE, size: 0 },
                insideVertical: { style: docx.BorderStyle.NONE, size: 0 }
              },
              rows: tableRows
            })
          );
        }

        else if (q.image && sec.type !== 'match_following') {
          const imageBytes = await imageToUint8Array(q.image);
          if (imageBytes) {
            headerChildren.push(
              new docx.Paragraph({
                indent: { left: 720 },
                spacing: { before: 120, after: 120 },
                children: [
                  new docx.ImageRun({
                    data: imageBytes,
                    transformation: {
                      width: q.imageWidth || 300,
                      height: q.imageHeight || 200
                    }
                  })
                ]
              })
            );
          }
        }

        else if (sec.type === 'table' && q.tableData) {
          // Build a bordered table in DOCX with bold headers and regular body
          const tblRows = [];

          // Header row
          const headerCells = q.tableData.headers.map(h =>
            new docx.TableCell({
              shading: { fill: 'E8E8E8' },
              children: [
                new docx.Paragraph({
                  spacing: { before: 40, after: 40 },
                  children: docxTextRunsWithMath(h || '', { bold: true })
                })
              ]
            })
          );
          tblRows.push(new docx.TableRow({ children: headerCells }));

          // Body rows
          q.tableData.rows.forEach(row => {
            const bodyCells = row.map(cell =>
              new docx.TableCell({
                children: [
                  new docx.Paragraph({
                    spacing: { before: 40, after: 40 },
                    children: docxTextRunsWithMath(cell || '')
                  })
                ]
              })
            );
            tblRows.push(new docx.TableRow({ children: bodyCells }));
          });

          headerChildren.push(
            new docx.Table({
              width: { size: 9000, type: docx.WidthType.DXA },
              indent: { size: 360, type: docx.WidthType.DXA },
              rows: tblRows
            })
          );
        }
      }
    }

    const doc = new docx.Document({
      styles: {
        default: {
          document: {
            run: {
              font: getFontFamily()
            }
          }
        }
      },
      sections: [{
        properties: {},
        footers: {
          default: new docx.Footer({
            children: [
              new docx.Paragraph({
                alignment: docx.AlignmentType.RIGHT,
                children: [
                  new docx.TextRun({
                    text: 'Page ',
                    size: 20
                  }),
                  new docx.SimpleField("PAGE"),
                  new docx.TextRun({
                    text: ' of ',
                    size: 20
                  }),
                  new docx.SimpleField("NUMPAGES")
                ]
              })
            ]
          })
        },
        children: headerChildren
      }]
    });

    return await docx.Packer.toBlob(doc);
  };

  const triggerDocxExport = async () => {
    if (!hasQuestions()) {
      alert("No questions added to export.");
      return;
    }
    if (hasBlankQuestions()) {
      if (!window.confirm("Some questions have empty text. Are you sure you want to export?")) {
        return;
      }
    }
    try {
      const blob = await generateDocxBlob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = generateExportFilename('docx');
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Error exporting DOCX:', err);
      alert('Failed to generate DOCX file.');
    }
  };
  const handleGoogleDocsWebPreview = async () => {
    if (hasBlankQuestions()) {
      if (!window.confirm("Some questions have empty text. Are you sure you want to preview?")) {
        return;
      }
    }
    setIsDocsUploading(true);
    setDocsError('');
    try {
      const blob = await generateDocxBlob();
      const formData = new FormData();
      formData.append('file', blob, generateExportFilename('docx'));

      const response = await fetch('https://tmpfiles.org/api/v1/upload', {
        method: 'POST',
        body: formData
      });

      if (!response.ok) {
        throw new Error(`Upload failed with status: ${response.status}`);
      }

      const resData = await response.json();
      if (resData.status !== 'success' || !resData.data || !resData.data.url) {
        throw new Error(resData.message || 'Invalid response from file upload server');
      }

      const uploadUrl = resData.data.url;
      const directUrl = uploadUrl.replace('tmpfiles.org/', 'tmpfiles.org/dl/');
      const googleDocsUrl = `https://docs.google.com/viewer?url=${encodeURIComponent(directUrl)}`;
      window.open(googleDocsUrl, '_blank');
      setIsDocsModalOpen(false);
    } catch (err) {
      console.error('Error opening in Google Docs:', err);
      setDocsError(err.message || 'Failed to upload document for Google Docs preview.');
    } finally {
      setIsDocsUploading(false);
    }
  };

  // Helper functions for shuffling Columns
  const getShuffledList = (q) => {
    if (!q.matchPairs) return [];
    if (!q._shuffledB || q._shuffledB.length !== q.matchPairs.length) {
      const list = q.matchPairs.map(p => (
        typeof p === 'string'
          ? { response: p, responseImage: '' }
          : { response: p.response || '', responseImage: p.responseImage || '' }
      ));
      if (q.shuffleB) {
        list.sort(() => 0.5 - Math.random());
      }
      q._shuffledB = list;
    }
    return q._shuffledB;
  };

  return (
    <div className={`app-container theme-${theme}`}>
      {/* Editor Panel (Left sidebar) */}
      <div className="editor-panel">
        <div className="editor-header">
          <div>
            <h1>
              <img src={pageLogo} style={{ width: '26px', height: '26px', objectFit: 'contain', marginRight: '2px' }} alt="Logo" />
              <span>QuestionNinja</span>
            </h1>
            <p>Where Great Questions Begin</p>
          </div>
          <div style={{ display: 'flex', gap: '8px' }}>
            <button className="btn btn-secondary btn-sm" onClick={toggleTheme} title="Toggle Light/Dark Theme">
              {theme === 'dark' ? <Sun size={14} /> : <Moon size={14} />}
            </button>
            <button className="btn btn-primary btn-sm" onClick={() => setIsPreviewOpen(true)}>
              <Maximize2 size={14} /> Live Preview
            </button>
          </div>
        </div>

        {/* Tab Selection */}
        <div style={{ display: 'flex', borderBottom: '1px solid var(--border-color)', backgroundColor: 'rgba(0,0,0,0.1)' }}>
          <button
            className={`btn btn-sm ${activeTab === 'branding' ? 'btn-primary' : 'btn-secondary'}`}
            style={{ flex: 1, borderRadius: 0, border: 'none', borderBottom: activeTab === 'branding' ? '2px solid var(--accent)' : 'none' }}
            onClick={() => setActiveTab('branding')}
          >
            <ImageIcon size={14} /> School Details
          </button>
          <button
            className={`btn btn-sm ${activeTab === 'metadata' ? 'btn-primary' : 'btn-secondary'}`}
            style={{ flex: 1, borderRadius: 0, border: 'none', borderBottom: activeTab === 'metadata' ? '2px solid var(--accent)' : 'none' }}
            onClick={() => setActiveTab('metadata')}
          >
            <Settings size={14} /> Exam Details
          </button>
          <button
            className={`btn btn-sm ${activeTab === 'sections' ? 'btn-primary' : 'btn-secondary'}`}
            style={{ flex: 1, borderRadius: 0, border: 'none', borderBottom: activeTab === 'sections' ? '2px solid var(--accent)' : 'none' }}
            onClick={() => setActiveTab('sections')}
          >
            <Layers size={14} /> Questions
          </button>
        </div>

        <div className="editor-content">
          {/* TAB 1: SCHOOL DETAILS */}
          {activeTab === 'branding' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
              <div className="editor-card">
                <div className="editor-card-header">
                  <div className="editor-card-title">School Details</div>
                </div>

                <div className="form-group">
                  <label>Upload Logo</label>
                  <input
                    type="file"
                    accept="image/*"
                    onChange={handleLogoUpload}
                    style={{ fontSize: '13px', color: 'var(--text-secondary)' }}
                  />
                  {branding.logo && (
                    <div style={{ marginTop: '12px', display: 'flex', alignItems: 'center', gap: '14px', background: 'var(--bg-card, #f8fafc)', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--border-color, #e2e8f0)' }}>
                      <div style={{
                        width: '64px',
                        height: '64px',
                        borderRadius: '6px',
                        border: '1px solid #cbd5e1',
                        padding: '4px',
                        background: '#ffffff',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        flexShrink: 0,
                        boxShadow: '0 1px 3px rgba(0,0,0,0.1)'
                      }}>
                        <img
                          src={branding.logo}
                          alt="Logo Preview"
                          onError={(e) => {
                            e.currentTarget.onerror = null;
                            if (branding.logo !== schoolLogo) {
                              setBranding(prev => ({ ...prev, logo: schoolLogo }));
                            }
                          }}
                          style={{ maxWidth: '100%', maxHeight: '100%', objectFit: 'contain' }}
                        />
                      </div>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                        <span style={{ fontSize: '12px', fontWeight: 600, color: 'var(--text-primary)' }}>Logo Active</span>
                        <button className="btn btn-sm btn-danger" onClick={removeLogo} style={{ alignSelf: 'flex-start' }}>
                          Remove Logo
                        </button>
                      </div>
                    </div>
                  )}
                </div>

                <div className="form-group">
                  <label>Institution Name</label>
                  <input
                    type="text"
                    value={branding.schoolName}
                    onChange={(e) => setBranding({ ...branding, schoolName: e.target.value })}
                    placeholder="e.g. Greenwood High School"
                  />
                </div>

                <div className="form-group">
                  <label>Address / Sub-header</label>
                  <textarea
                    value={branding.schoolAddress}
                    onChange={(e) => setBranding({ ...branding, schoolAddress: e.target.value })}
                    placeholder="Enter institution address and contacts..."
                  />
                </div>




              </div>

              <div className="warning-badge" style={{ padding: '12px' }}>
                <Move size={16} />
                <span>Tip: You can click & drag the school logo directly on the A4 page preview to position it, or drag the bottom-right handle to scale it!</span>
              </div>
            </div>
          )}

          {/* TAB 2: EXAM DETAILS */}
          {activeTab === 'metadata' && (
            <div className="editor-card">
              <div className="editor-card-header">
                <div className="editor-card-title">Exam Details</div>
              </div>

              <div className="form-group">
                <label>Examination Title</label>
                <input
                  type="text"
                  value={metadata.title}
                  onChange={(e) => setMetadata({ ...metadata, title: e.target.value })}
                  placeholder="e.g. Term 1 Summative Assessment"
                />
              </div>

              <div className="form-group">
                <label>Subject Name</label>
                <input
                  type="text"
                  value={metadata.subject}
                  onChange={(e) => setMetadata({ ...metadata, subject: e.target.value })}
                  placeholder="e.g. Computer Science"
                />
              </div>

              <div className="form-group">
                <label>Class & Division</label>
                <input
                  type="text"
                  value={metadata.classDiv}
                  onChange={(e) => setMetadata({ ...metadata, classDiv: e.target.value })}
                  placeholder="e.g. Class X - Div A"
                />
              </div>

              <div className="form-row">
                <div className="form-group">
                  <label>Max Marks</label>
                  <input
                    type="number"
                    step="any"
                    min="0"
                    value={metadata.maxMarks}
                    onChange={(e) => setMetadata({ ...metadata, maxMarks: Math.max(0, Number(e.target.value)) })}
                    onBlur={(e) => setMetadata({ ...metadata, maxMarks: Math.max(0, Math.round(Number(e.target.value) * 100) / 100) })}
                  />
                </div>
                <div className="form-group">
                  <label>Time Duration</label>
                  <input
                    type="text"
                    value={metadata.duration}
                    onChange={(e) => setMetadata({ ...metadata, duration: e.target.value })}
                    placeholder="e.g. 2 Hours"
                  />
                </div>
              </div>

              <div className="form-group">
                <label>Question Paper Language</label>
                <select
                  value={metadata.language || 'english'}
                  onChange={(e) => setMetadata({ ...metadata, language: e.target.value })}
                >
                  <option value="english">English (Default)</option>
                  <option value="malayalam">Malayalam</option>
                  <option value="hindi">Hindi</option>
                </select>
              </div>



              {/* Validation Badges */}
              <div style={{ marginTop: '12px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
                {getExamCurrentTotalMarks() !== metadata.maxMarks ? (
                  <div className="warning-badge">
                    <AlertTriangle size={14} />
                    <span>Marks Mismatch: Current Questions = {formatMarks(getExamCurrentTotalMarks())} marks (Target = {formatMarks(metadata.maxMarks)} marks).</span>
                  </div>
                ) : (
                  <div className="warning-badge" style={{ backgroundColor: 'rgba(16,185,129,0.15)', color: 'var(--success)', borderColor: 'rgba(16,185,129,0.3)' }}>
                    <CheckCircle size={14} />
                    <span>Perfect: Sum of all questions matches targets!</span>
                  </div>
                )}
                {hasZeroMarkQuestions() && (
                  <div className="warning-badge" style={{ backgroundColor: 'rgba(239, 68, 68, 0.15)', color: 'var(--danger)', borderColor: 'rgba(239, 68, 68, 0.3)' }}>
                    <AlertTriangle size={14} />
                    <span>Warning: One or more questions have 0 marks.</span>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* TAB 3: SECTIONS & QUESTIONS MANAGEMENT */}
          {activeTab === 'sections' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
              {sections.map((sec, sIdx) => {
                const secTotal = getSectionTotalMarks(sec);
                const isOverProvisioned = secTotal > sec.marks;
                const isUnderProvisioned = secTotal < sec.marks;

                return (
                  <div key={sec.id} className="editor-card" style={{ borderLeft: `4px solid var(--accent)` }}>
                    <div className="editor-card-header">
                      <div className="editor-card-title" onClick={() => toggleSectionCollapse(sec.id)} style={{ cursor: 'pointer', flex: 1 }}>
                        <Layers size={14} />
                        <span>Section {String.fromCharCode(65 + sIdx)}</span>
                        {collapsedSections[sec.id] && (
                          <span style={{ fontSize: '11px', color: 'var(--text-muted)', fontWeight: 'normal', marginLeft: '6px' }}>
                            ({sec.questions.length} questions, {secTotal}M)
                          </span>
                        )}
                      </div>
                      <div style={{ display: 'flex', gap: '6px' }}>
                        <button className="btn-icon-only" onClick={() => toggleSectionCollapse(sec.id)} title={collapsedSections[sec.id] ? "Expand Section" : "Collapse Section"}>
                          {collapsedSections[sec.id] ? <ChevronDown size={14} /> : <ChevronUp size={14} />}
                        </button>
                        <button className="btn-icon-only" onClick={() => moveSection(sIdx, 'up')} disabled={sIdx === 0}>
                          <ArrowUp size={14} />
                        </button>
                        <button className="btn-icon-only" onClick={() => moveSection(sIdx, 'down')} disabled={sIdx === sections.length - 1}>
                          <ArrowDown size={14} />
                        </button>
                        <button className="btn-icon-only" style={{ color: 'var(--danger)' }} onClick={() => deleteSection(sec.id)}>
                          <Trash2 size={14} />
                        </button>
                      </div>
                    </div>

                    {!collapsedSections[sec.id] && (
                      <>
                        <div className="form-group">
                          <label>Section Title</label>
                          <input
                            type="text"
                            value={sec.title}
                            onChange={(e) => updateSectionMeta(sec.id, 'title', e.target.value)}
                          />
                        </div>

                        <div className="form-row">
                          <div className="form-group">
                            <label>Declared Section Marks</label>
                            <input
                              type="number"
                              step="any"
                              min="0"
                              value={sec.marks}
                              onChange={(e) => updateSectionMeta(sec.id, 'marks', Math.max(0, Number(e.target.value)))}
                              onBlur={(e) => updateSectionMeta(sec.id, 'marks', Math.max(0, Math.round(Number(e.target.value) * 100) / 100))}
                            />
                          </div>
                          <div className="form-group">
                            <label>Current Questions Total</label>
                            <div style={{
                              padding: '10px 14px',
                              backgroundColor: 'rgba(0,0,0,0.2)',
                              borderRadius: 'var(--radius-sm)',
                              fontSize: '14px',
                              fontWeight: 'bold',
                              color: isOverProvisioned ? 'var(--warning)' : isUnderProvisioned ? 'var(--danger)' : 'var(--success)'
                            }}>
                              {formatMarks(secTotal)} / {formatMarks(sec.marks)} Marks
                            </div>
                          </div>
                        </div>

                        <div className="form-group">
                          <label>Notes / Section Instructions</label>
                          <textarea
                            value={sec.instructions}
                            onChange={(e) => updateSectionMeta(sec.id, 'instructions', e.target.value)}
                            placeholder="e.g. Answer any 5 of the following 7 questions"
                          />
                        </div>

                        <div className="form-group">
                          <label>Section Question Type</label>
                          <select
                            value={sec.type || 'essay'}
                            onChange={(e) => updateSectionType(sec.id, e.target.value)}
                          >
                            <option value="essay">Short Answer / Essay / Fill Blank</option>
                            <option value="mcq">Multiple Choice (MCQ)</option>
                            <option value="true_false">True / False</option>
                            <option value="match_following">Match the Following</option>
                            <option value="table">Table Question</option>
                          </select>
                        </div>

                        {isOverProvisioned && (
                          <div className="warning-badge" style={{ fontSize: '11px' }}>
                            <AlertTriangle size={12} />
                            <span>Choice Provisioning: Question total ({formatMarks(secTotal)}) exceeds Declared section marks ({formatMarks(sec.marks)}). Permitted for optional choices.</span>
                          </div>
                        )}

                        {/* Questions Area */}
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', marginTop: '16px', borderTop: '1px solid var(--border-color)', paddingTop: '16px' }}>
                          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                            <h4 style={{ fontSize: '13px', fontWeight: 'bold', textTransform: 'uppercase' }}>Questions ({sec.questions.length})</h4>
                          </div>

                          {sec.questions.map((q, qIdx) => (
                            <div key={q.id} onPaste={(e) => handlePasteImage(e, sec.id, q.id)} style={{ padding: '12px', backgroundColor: 'var(--bg-editor)', borderRadius: 'var(--radius-sm)', display: 'flex', flexDirection: 'column', gap: '10px', border: '1px solid var(--border-color)' }}>
                              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                <span style={{ fontSize: '12px', fontWeight: 'bold', color: 'var(--accent)' }}>
                                  Q{qIdx + 1} ({(sec.type || 'essay').toUpperCase()})
                                </span>
                                <div style={{ display: 'flex', gap: '4px' }}>
                                  <button className="btn-icon-only" style={{ padding: '4px' }} onClick={() => moveQuestion(sec.id, qIdx, 'up')} disabled={qIdx === 0}>
                                    <ArrowUp size={12} />
                                  </button>
                                  <button className="btn-icon-only" style={{ padding: '4px' }} onClick={() => moveQuestion(sec.id, qIdx, 'down')} disabled={qIdx === sec.questions.length - 1}>
                                    <ArrowDown size={12} />
                                  </button>
                                  <button className="btn-icon-only" style={{ padding: '4px', color: 'var(--danger)' }} onClick={() => deleteQuestion(sec.id, q.id)}>
                                    <Trash2 size={12} />
                                  </button>
                                </div>
                              </div>

                              <div className="form-group">
                                <label style={{ fontSize: '10px' }}>Question Text</label>
                                <textarea
                                  id={`q__text__${sec.id}__${q.id}`}
                                  value={q.text}
                                  placeholder="Enter question details here..."
                                  style={{
                                    minHeight: '60px',
                                    fontSize: '13px',
                                    borderColor: (!q.text || !q.text.trim()) ? 'var(--danger)' : 'var(--border-color)'
                                  }}
                                  onChange={(e) => updateQuestion(sec.id, q.id, { text: e.target.value })}
                                  onFocus={(e) => {
                                    if (e.target.value === 'New Question details here...') {
                                      updateQuestion(sec.id, q.id, { text: '' });
                                    }
                                  }}
                                />
                                {(!q.text || !q.text.trim()) && (
                                  <span style={{ color: 'var(--danger)', fontSize: '11px', marginTop: '4px', display: 'flex', alignItems: 'center', gap: '4px' }}>
                                    <AlertTriangle size={12} /> Please enter the question text.
                                  </span>
                                )}

                                {/* Question Image Fields - positioned directly below Question Text */}
                                {sec.type !== 'match_following' && (
                                  <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', marginTop: '8px' }}>
                                    <label style={{ fontSize: '10px', fontWeight: 'bold' }}>Question Image (Optional)</label>
                                    {q.image ? (
                                      <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                        <div 
                                          className="editor-image-preview-container"
                                          onDragOver={handleDragOver}
                                          onDragLeave={handleDragLeave}
                                          onDrop={(e) => handleDropImage(e, sec.id, q.id)}
                                        >
                                          <img
                                            src={q.image}
                                            alt="Question"
                                            style={{
                                              maxWidth: '100%',
                                              maxHeight: '150px',
                                              objectFit: 'contain',
                                              borderRadius: 'var(--radius-sm)',
                                              border: '1px solid var(--border-color)',
                                              padding: '2px'
                                            }}
                                          />
                                        </div>
                                        <div style={{ display: 'flex', gap: '6px' }}>
                                          <label className="btn btn-secondary btn-sm" style={{ cursor: 'pointer', margin: 0, padding: '2px 6px', fontSize: '10px' }}>
                                            Change Image
                                            <input
                                              type="file"
                                              accept="image/*"
                                              style={{ display: 'none' }}
                                              onChange={(e) => {
                                                const file = e.target.files[0];
                                                if (file) {
                                                  const reader = new FileReader();
                                                  reader.onload = (uploadEvent) => {
                                                    updateQuestion(sec.id, q.id, { image: uploadEvent.target.result });
                                                  };
                                                  reader.readAsDataURL(file);
                                                }
                                              }}
                                            />
                                          </label>
                                          <button
                                            className="btn btn-danger btn-sm"
                                            style={{ padding: '2px 6px', fontSize: '10px' }}
                                            onClick={() => updateQuestion(sec.id, q.id, { image: '' })}
                                          >
                                            Remove Image
                                          </button>
                                        </div>
                                        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px' }}>
                                          <div className="form-group" style={{ margin: 0 }}>
                                            <label style={{ fontSize: '9px' }}>Width (px)</label>
                                            <input
                                              type="number"
                                              value={q.imageWidth || 300}
                                              style={{ padding: '2px 6px', fontSize: '11px' }}
                                              onChange={(e) => updateQuestion(sec.id, q.id, { imageWidth: Number(e.target.value) || 0 })}
                                            />
                                          </div>
                                          <div className="form-group" style={{ margin: 0 }}>
                                            <label style={{ fontSize: '9px' }}>Height (px)</label>
                                            <input
                                              type="number"
                                              value={q.imageHeight || 200}
                                              style={{ padding: '2px 6px', fontSize: '11px' }}
                                              onChange={(e) => updateQuestion(sec.id, q.id, { imageHeight: Number(e.target.value) || 0 })}
                                            />
                                          </div>
                                        </div>
                                      </div>
                                    ) : (
                                      <label 
                                        className="image-upload-dashed-zone"
                                        style={{ padding: '8px 12px', gap: '4px' }}
                                        onDragOver={handleDragOver}
                                        onDragLeave={handleDragLeave}
                                        onDrop={(e) => handleDropImage(e, sec.id, q.id)}
                                      >
                                        <ImageIcon size={14} className="text-secondary" />
                                        <span style={{ fontSize: '10px' }}>Upload Question Image (click, drag, or Ctrl+V)</span>
                                        <input
                                          type="file"
                                          accept="image/*"
                                          style={{ display: 'none' }}
                                          onChange={(e) => {
                                            const file = e.target.files[0];
                                            if (file) {
                                              const reader = new FileReader();
                                              reader.onload = (uploadEvent) => {
                                                updateQuestion(sec.id, q.id, { image: uploadEvent.target.result });
                                              };
                                              reader.readAsDataURL(file);
                                            }
                                          }}
                                        />
                                      </label>
                                    )}
                                  </div>
                                )}

                                {/* Match the Following Specific Fields */}
                                {sec.type === 'match_following' && q.matchPairs && (
                                  <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', marginTop: '10px' }}>
                                    {/* ... rest of logic */}
                                  </div>
                                )}
                                {(sec.type === 'essay') && (
                                  <button
                                    className="btn btn-secondary btn-sm"
                                    style={{ alignSelf: 'flex-start', marginTop: '4px', fontSize: '11px', padding: '4px 8px' }}
                                    onClick={() => {
                                      const text = q.text + ' _______';
                                      updateQuestion(sec.id, q.id, { text });
                                    }}
                                  >
                                    Insert Blank
                                  </button>
                                )}
                              </div>

                              <div className="form-group">
                                <label style={{ fontSize: '10px' }}>Question Marks</label>
                                {q.subQuestions && q.subQuestions.length > 0 ? (
                                  <div style={{
                                    padding: '6px 10px',
                                    fontSize: '13px',
                                    fontWeight: 'bold',
                                    backgroundColor: 'rgba(99, 102, 241, 0.15)',
                                    borderRadius: 'var(--radius-sm)',
                                    border: '1px solid var(--accent)',
                                    color: 'var(--accent)'
                                  }}>
                                    {formatMarks(getQuestionMarks(q))} Marks (Calculated from sub-questions)
                                  </div>
                                ) : (
                                  <input
                                    type="number"
                                    step="any"
                                    min="0"
                                    value={q.marks}
                                    style={{ padding: '6px 10px', fontSize: '13px' }}
                                    onChange={(e) => updateQuestion(sec.id, q.id, { marks: Math.max(0, Number(e.target.value)) })}
                                    onBlur={(e) => updateQuestion(sec.id, q.id, { marks: Math.max(0, Math.round(Number(e.target.value) * 100) / 100) })}
                                  />
                                )}
                              </div>

                              {/* MCQ Specific Fields */}
                              {sec.type === 'mcq' && q.options && (
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                    <label style={{ fontSize: '10px', fontWeight: 'bold' }}>Options</label>
                                    <button
                                      type="button"
                                      className="btn-secondary"
                                      style={{ padding: '2px 8px', fontSize: '11px', borderRadius: '4px' }}
                                      onClick={() => {
                                        const newOpts = [...q.options, ''];
                                        updateQuestion(sec.id, q.id, { options: newOpts });
                                      }}
                                    >
                                      + Add Option
                                    </button>
                                  </div>
                                  {q.options.map((opt, oIdx) => (
                                    <div key={oIdx} style={{ display: 'flex', gap: '6px', width: '100%', alignItems: 'flex-start' }}>
                                      <span style={{ fontSize: '13px', fontWeight: 'bold', paddingTop: '6px', minWidth: '18px' }}>
                                        {String.fromCharCode(65 + oIdx)}.
                                      </span>
                                      <textarea
                                        id={`q__opt__${sec.id}__${q.id}__${oIdx}`}
                                        value={opt}
                                        rows={opt && opt.includes('\n') ? Math.max(2, opt.split('\n').length) : 1}
                                        placeholder={`Option ${String.fromCharCode(65 + oIdx)}`}
                                        className="mcq-option-textarea"
                                        style={{
                                          padding: '6px 8px',
                                          fontSize: '12px',
                                          flex: 1,
                                          minHeight: '34px',
                                          resize: 'vertical',
                                          fontFamily: 'inherit'
                                        }}
                                        onChange={(e) => {
                                          const newOpts = [...q.options];
                                          newOpts[oIdx] = e.target.value;
                                          updateQuestion(sec.id, q.id, { options: newOpts });
                                        }}
                                      />
                                      {q.options.length > 2 && (
                                        <button
                                          type="button"
                                          title="Remove Option"
                                          style={{
                                            background: 'none',
                                            border: 'none',
                                            color: 'var(--danger, #ef4444)',
                                            cursor: 'pointer',
                                            fontSize: '14px',
                                            padding: '6px 4px',
                                            lineHeight: 1
                                          }}
                                          onClick={() => {
                                            const newOpts = q.options.filter((_, idx) => idx !== oIdx);
                                            updateQuestion(sec.id, q.id, { options: newOpts });
                                          }}
                                        >
                                          ✕
                                        </button>
                                      )}
                                    </div>
                                  ))}
                                </div>
                              )}

                              {/* Essay Specific Fields & Sub-Questions */}
                              {sec.type === 'essay' && (
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                                  {(!q.subQuestions || q.subQuestions.length === 0) && (
                                    <div className="form-group">
                                      <label style={{ fontSize: '10px' }}>Blank lines for printing</label>
                                      <input
                                        type="number"
                                        min="0"
                                        value={q.blankLines}
                                        style={{ padding: '6px 10px', fontSize: '13px' }}
                                        onChange={(e) => updateQuestion(sec.id, q.id, { blankLines: e.target.value === '' ? '' : Math.max(0, parseInt(e.target.value, 10) || 0) })}
                                      />
                                    </div>
                                  )}

                                  {/* Sub-Questions Management Area */}
                                  <div style={{
                                    border: '1px solid var(--border-color)',
                                    borderRadius: 'var(--radius-sm)',
                                    padding: '10px',
                                    backgroundColor: 'rgba(0,0,0,0.15)',
                                    display: 'flex',
                                    flexDirection: 'column',
                                    gap: '10px'
                                  }}>
                                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                      <span style={{ fontSize: '11px', fontWeight: 'bold', textTransform: 'uppercase', color: 'var(--text-primary)' }}>
                                        Sub-Questions {q.subQuestions && q.subQuestions.length > 0 ? `(${q.subQuestions.length})` : ''}
                                      </span>
                                      {(!q.subQuestions || q.subQuestions.length === 0) && (
                                        <button
                                          className="btn btn-secondary btn-sm"
                                          style={{ padding: '3px 8px', fontSize: '11px' }}
                                          onClick={() => addSubQuestion(sec.id, q.id)}
                                        >
                                          + Add Sub-Question
                                        </button>
                                      )}
                                    </div>

                                    {q.subQuestions && q.subQuestions.map((sq, sqIdx) => (
                                      <div
                                        key={sq.id}
                                        style={{
                                          padding: '10px',
                                          backgroundColor: 'var(--bg-card)',
                                          borderRadius: 'var(--radius-sm)',
                                          border: '1px solid var(--border-color)',
                                          display: 'flex',
                                          flexDirection: 'column',
                                          gap: '8px'
                                        }}
                                      >
                                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                          <span style={{ fontSize: '11px', fontWeight: 'bold', color: 'var(--accent)' }}>
                                            Sub-Question {sq.label || `(${String.fromCharCode(97 + (sqIdx % 26))})`}
                                          </span>
                                          <div style={{ display: 'flex', gap: '4px' }}>
                                            <button
                                              className="btn-icon-only"
                                              style={{ padding: '2px 4px' }}
                                              onClick={() => moveSubQuestion(sec.id, q.id, sqIdx, 'up')}
                                              disabled={sqIdx === 0}
                                              title="Move Up"
                                            >
                                              <ArrowUp size={12} />
                                            </button>
                                            <button
                                              className="btn-icon-only"
                                              style={{ padding: '2px 4px' }}
                                              onClick={() => moveSubQuestion(sec.id, q.id, sqIdx, 'down')}
                                              disabled={sqIdx === q.subQuestions.length - 1}
                                              title="Move Down"
                                            >
                                              <ArrowDown size={12} />
                                            </button>
                                            <button
                                              className="btn-icon-only"
                                              style={{ padding: '2px 4px', color: 'var(--danger)' }}
                                              onClick={() => deleteSubQuestion(sec.id, q.id, sq.id)}
                                              title="Delete Sub-Question"
                                            >
                                              <Trash2 size={12} />
                                            </button>
                                          </div>
                                        </div>

                                        <div style={{ display: 'grid', gridTemplateColumns: '100px 1fr', gap: '8px' }}>
                                          <div className="form-group" style={{ margin: 0 }}>
                                            <label style={{ fontSize: '9px' }}>Label</label>
                                            <input
                                              type="text"
                                              value={sq.label || `(${String.fromCharCode(97 + (sqIdx % 26))})`}
                                              style={{ padding: '4px 6px', fontSize: '12px' }}
                                              onChange={(e) => updateSubQuestion(sec.id, q.id, sq.id, { label: e.target.value })}
                                            />
                                          </div>

                                          <div className="form-group" style={{ margin: 0 }}>
                                            <label style={{ fontSize: '9px' }}>Sub-Question Text</label>
                                            <textarea
                                              id={`q__sqtext__${sec.id}__${q.id}__${sq.id}`}
                                              value={sq.text}
                                              placeholder="Enter sub-question text..."
                                              style={{
                                                minHeight: '45px',
                                                fontSize: '12px',
                                                borderColor: (!sq.text || !sq.text.trim()) ? 'var(--danger)' : 'var(--border-color)'
                                              }}
                                              onChange={(e) => updateSubQuestion(sec.id, q.id, sq.id, { text: e.target.value })}
                                            />
                                            {(!sq.text || !sq.text.trim()) && (
                                              <span style={{ color: 'var(--danger)', fontSize: '10px', marginTop: '2px', display: 'flex', alignItems: 'center', gap: '3px' }}>
                                                <AlertTriangle size={10} /> Please enter sub-question text.
                                              </span>
                                            )}
                                            <button
                                              className="btn btn-secondary btn-sm"
                                              style={{ alignSelf: 'flex-start', marginTop: '4px', fontSize: '10px', padding: '2px 6px' }}
                                              onClick={() => {
                                                const text = sq.text + ' _______';
                                                updateSubQuestion(sec.id, q.id, sq.id, { text });
                                              }}
                                            >
                                              Insert Blank
                                            </button>
                                          </div>
                                        </div>

                                        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
                                          <div className="form-group" style={{ margin: 0 }}>
                                            <label style={{ fontSize: '9px' }}>Sub-Question Marks</label>
                                            <input
                                              type="number"
                                              step="any"
                                              min="0"
                                              value={sq.marks}
                                              style={{ padding: '4px 6px', fontSize: '12px' }}
                                              onChange={(e) => updateSubQuestion(sec.id, q.id, sq.id, { marks: Math.max(0, Number(e.target.value)) })}
                                              onBlur={(e) => updateSubQuestion(sec.id, q.id, sq.id, { marks: Math.max(0, Math.round(Number(e.target.value) * 100) / 100) })}
                                            />
                                          </div>
                                          <div className="form-group" style={{ margin: 0 }}>
                                            <label style={{ fontSize: '9px' }}>Blank Lines for printing</label>
                                            <input
                                              type="number"
                                              min="0"
                                              value={sq.blankLines !== undefined ? sq.blankLines : 4}
                                              style={{ padding: '4px 6px', fontSize: '12px' }}
                                              onChange={(e) => updateSubQuestion(sec.id, q.id, sq.id, { blankLines: e.target.value === '' ? '' : Math.max(0, parseInt(e.target.value, 10) || 0) })}
                                            />
                                          </div>
                                        </div>

                                        {/* Sub-Question Image Controls */}
                                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', marginTop: '6px' }}>
                                          <label style={{ fontSize: '9px', fontWeight: 'bold' }}>Sub-Question Image (Optional)</label>
                                          {sq.image ? (
                                            <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                                              <div 
                                                className="editor-image-preview-container"
                                                onDragOver={handleDragOver}
                                                onDragLeave={handleDragLeave}
                                                onDrop={(e) => handleDropImage(e, sec.id, q.id, sq.id)}
                                              >
                                                <img
                                                  src={sq.image}
                                                  alt="Sub-question diagram"
                                                  style={{
                                                    maxWidth: '100%',
                                                    maxHeight: '120px',
                                                    objectFit: 'contain',
                                                    borderRadius: 'var(--radius-sm)',
                                                    border: '1px solid var(--border-color)',
                                                    padding: '2px'
                                                  }}
                                                />
                                              </div>
                                              <div style={{ display: 'flex', gap: '6px' }}>
                                                <label className="btn btn-secondary btn-sm" style={{ cursor: 'pointer', margin: 0, padding: '2px 6px', fontSize: '10px' }}>
                                                  Change Image
                                                  <input
                                                    type="file"
                                                    accept="image/*"
                                                    style={{ display: 'none' }}
                                                    onChange={(e) => {
                                                      const file = e.target.files[0];
                                                      if (file) {
                                                        const reader = new FileReader();
                                                        reader.onload = (uploadEvent) => {
                                                          updateSubQuestion(sec.id, q.id, sq.id, { image: uploadEvent.target.result });
                                                        };
                                                        reader.readAsDataURL(file);
                                                      }
                                                    }}
                                                  />
                                                </label>
                                                <button
                                                  className="btn btn-danger btn-sm"
                                                  style={{ padding: '2px 6px', fontSize: '10px' }}
                                                  onClick={() => updateSubQuestion(sec.id, q.id, sq.id, { image: '' })}
                                                >
                                                  Remove Image
                                                </button>
                                              </div>
                                              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px' }}>
                                                <div className="form-group" style={{ margin: 0 }}>
                                                  <label style={{ fontSize: '9px' }}>Width (px)</label>
                                                  <input
                                                    type="number"
                                                    value={sq.imageWidth || 300}
                                                    style={{ padding: '2px 6px', fontSize: '11px' }}
                                                    onChange={(e) => updateSubQuestion(sec.id, q.id, sq.id, { imageWidth: Number(e.target.value) || 0 })}
                                                  />
                                                </div>
                                                <div className="form-group" style={{ margin: 0 }}>
                                                  <label style={{ fontSize: '9px' }}>Height (px)</label>
                                                  <input
                                                    type="number"
                                                    value={sq.imageHeight || 200}
                                                    style={{ padding: '2px 6px', fontSize: '11px' }}
                                                    onChange={(e) => updateSubQuestion(sec.id, q.id, sq.id, { imageHeight: Number(e.target.value) || 0 })}
                                                  />
                                                </div>
                                              </div>
                                            </div>
                                          ) : (
                                            <label 
                                              className="image-upload-dashed-zone"
                                              style={{ padding: '8px 12px', gap: '4px' }}
                                              onDragOver={handleDragOver}
                                              onDragLeave={handleDragLeave}
                                              onDrop={(e) => handleDropImage(e, sec.id, q.id, sq.id)}
                                            >
                                              <ImageIcon size={14} className="text-secondary" />
                                              <span style={{ fontSize: '10px' }}>Upload Sub-Question Image (click, drag, or Ctrl+V)</span>
                                              <input
                                                type="file"
                                                accept="image/*"
                                                style={{ display: 'none' }}
                                                onChange={(e) => {
                                                  const file = e.target.files[0];
                                                  if (file) {
                                                    const reader = new FileReader();
                                                    reader.onload = (uploadEvent) => {
                                                      updateSubQuestion(sec.id, q.id, sq.id, { image: uploadEvent.target.result });
                                                    };
                                                    reader.readAsDataURL(file);
                                                  }
                                                }}
                                              />
                                            </label>
                                          )}
                                        </div>
                                      </div>
                                    ))}
                                    {q.subQuestions && q.subQuestions.length > 0 && (
                                      <button
                                        className="btn btn-secondary btn-sm"
                                        style={{ alignSelf: 'flex-end', padding: '4px 10px', fontSize: '11px' }}
                                        onClick={() => addSubQuestion(sec.id, q.id)}
                                      >
                                        + Add Sub-Question
                                      </button>
                                    )}
                                  </div>
                                </div>
                              )}

                              {/* Question Image Fields */}
                              {sec.type !== 'match_following' && (
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                                  <label style={{ fontSize: '10px', fontWeight: 'bold' }}>Question Image (Optional)</label>
                                  {q.image ? (
                                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                      <div 
                                        className="editor-image-preview-container"
                                        onDragOver={handleDragOver}
                                        onDragLeave={handleDragLeave}
                                        onDrop={(e) => handleDropImage(e, sec.id, q.id)}
                                      >
                                        <img
                                          src={q.image}
                                          alt="Question"
                                          style={{
                                            maxWidth: '100%',
                                            maxHeight: '150px',
                                            objectFit: 'contain',
                                            borderRadius: 'var(--radius-sm)',
                                            border: '1px solid var(--border-color)',
                                            padding: '2px'
                                          }}
                                        />
                                      </div>
                                      <div style={{ display: 'flex', gap: '6px' }}>
                                        <label className="btn btn-secondary btn-sm" style={{ cursor: 'pointer', margin: 0, padding: '2px 6px', fontSize: '10px' }}>
                                          Change Image
                                          <input
                                            type="file"
                                            accept="image/*"
                                            style={{ display: 'none' }}
                                            onChange={(e) => {
                                              const file = e.target.files[0];
                                              if (file) {
                                                const reader = new FileReader();
                                                reader.onload = (uploadEvent) => {
                                                  updateQuestion(sec.id, q.id, { image: uploadEvent.target.result });
                                                };
                                                reader.readAsDataURL(file);
                                              }
                                            }}
                                          />
                                        </label>
                                        <button
                                          className="btn btn-danger btn-sm"
                                          style={{ padding: '2px 6px', fontSize: '10px' }}
                                          onClick={() => updateQuestion(sec.id, q.id, { image: '' })}
                                        >
                                          Remove Image
                                        </button>
                                      </div>
                                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px' }}>
                                        <div className="form-group" style={{ margin: 0 }}>
                                          <label style={{ fontSize: '9px' }}>Width (px)</label>
                                          <input
                                            type="number"
                                            value={q.imageWidth || 300}
                                            style={{ padding: '2px 6px', fontSize: '11px' }}
                                            onChange={(e) => updateQuestion(sec.id, q.id, { imageWidth: Number(e.target.value) || 0 })}
                                          />
                                        </div>
                                        <div className="form-group" style={{ margin: 0 }}>
                                          <label style={{ fontSize: '9px' }}>Height (px)</label>
                                          <input
                                            type="number"
                                            value={q.imageHeight || 200}
                                            style={{ padding: '2px 6px', fontSize: '11px' }}
                                            onChange={(e) => updateQuestion(sec.id, q.id, { imageHeight: Number(e.target.value) || 0 })}
                                          />
                                        </div>
                                      </div>
                                    </div>
                                  ) : (
                                    <label 
                                      className="image-upload-dashed-zone"
                                      style={{ padding: '8px 12px', gap: '4px' }}
                                      onDragOver={handleDragOver}
                                      onDragLeave={handleDragLeave}
                                      onDrop={(e) => handleDropImage(e, sec.id, q.id)}
                                    >
                                      <ImageIcon size={14} className="text-secondary" />
                                      <span style={{ fontSize: '10px' }}>Upload Question Image (click, drag, or Ctrl+V)</span>
                                      <input
                                        type="file"
                                        accept="image/*"
                                        style={{ display: 'none' }}
                                        onChange={(e) => {
                                          const file = e.target.files[0];
                                          if (file) {
                                            const reader = new FileReader();
                                            reader.onload = (uploadEvent) => {
                                              updateQuestion(sec.id, q.id, { image: uploadEvent.target.result });
                                            };
                                            reader.readAsDataURL(file);
                                          }
                                        }}
                                      />
                                    </label>
                                  )}
                                </div>
                              )}

                              {/* Match the Following Specific Fields */}
                              {sec.type === 'match_following' && q.matchPairs && (
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                    <label style={{ fontSize: '10px', fontWeight: 'bold' }}>Match Pairs (Column A & Column B)</label>
                                    <button
                                      type="button"
                                      className="btn btn-secondary btn-sm"
                                      style={{ padding: '2px 6px', fontSize: '10px' }}
                                      onClick={() => {
                                        updateQuestion(sec.id, q.id, {
                                          matchPairs: [...q.matchPairs, { premise: '', premiseImage: '', response: '', responseImage: '' }],
                                          _shuffledB: null
                                        });
                                      }}
                                    >
                                      + Pair
                                    </button>
                                  </div>

                                  {/* Column Headings Banner */}
                                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px', padding: '4px 8px', backgroundColor: 'rgba(99, 102, 241, 0.1)', borderRadius: '4px', fontSize: '11px', fontWeight: 'bold' }}>
                                    <span>Column A</span>
                                    <span>Column B</span>
                                  </div>

                                  {q.matchPairs.map((pair, pIdx) => (
                                    <div key={pIdx} style={{ display: 'flex', flexDirection: 'column', gap: '6px', border: '1px solid var(--border-color)', padding: '8px', borderRadius: '6px', backgroundColor: 'rgba(0,0,0,0.05)' }}>
                                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                        <span style={{ fontSize: '11px', fontWeight: 'bold', color: 'var(--accent)' }}>Pair {pIdx + 1}</span>
                                        {q.matchPairs.length > 1 && (
                                          <button
                                            type="button"
                                            className="btn btn-danger btn-sm"
                                            style={{ padding: '2px 6px', fontSize: '10px' }}
                                            onClick={() => {
                                              updateQuestion(sec.id, q.id, {
                                                matchPairs: q.matchPairs.filter((_, idx) => idx !== pIdx),
                                                _shuffledB: null
                                              });
                                            }}
                                          >
                                            <Trash2 size={12} /> Remove
                                          </button>
                                        )}
                                      </div>

                                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
                                        {/* Column A Item */}
                                        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                                          <span style={{ fontSize: '10px', color: 'var(--text-secondary)' }}>Col A ({pIdx + 1})</span>
                                          <input
                                            id={`q__matcha__${sec.id}__${q.id}__${pIdx}`}
                                            type="text"
                                            placeholder="Premise text"
                                            value={pair.premise || ''}
                                            style={{ padding: '4px 8px', fontSize: '12px' }}
                                            onChange={(e) => {
                                              const newPairs = [...q.matchPairs];
                                              newPairs[pIdx] = { ...newPairs[pIdx], premise: e.target.value };
                                              updateQuestion(sec.id, q.id, { matchPairs: newPairs, _shuffledB: null });
                                            }}
                                          />
                                          {pair.premiseImage ? (
                                            <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginTop: '2px' }}>
                                              <img
                                                src={pair.premiseImage}
                                                alt="Col A Preview"
                                                style={{ width: '40px', height: '40px', objectFit: 'contain', border: '1px solid var(--border-color)', borderRadius: '4px', backgroundColor: '#fff' }}
                                              />
                                              <button
                                                type="button"
                                                style={{ background: 'none', border: 'none', color: 'var(--danger)', cursor: 'pointer', fontSize: '11px' }}
                                                onClick={() => {
                                                  const newPairs = [...q.matchPairs];
                                                  newPairs[pIdx] = { ...newPairs[pIdx], premiseImage: '' };
                                                  updateQuestion(sec.id, q.id, { matchPairs: newPairs, _shuffledB: null });
                                                }}
                                              >
                                                ✕ Remove Image
                                              </button>
                                            </div>
                                          ) : (
                                            <label style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', cursor: 'pointer', fontSize: '11px', color: 'var(--accent)', marginTop: '2px' }}>
                                              <ImageIcon size={13} />
                                              <span>Add Image</span>
                                              <input
                                                type="file"
                                                accept="image/*"
                                                style={{ display: 'none' }}
                                                onChange={(e) => {
                                                  const file = e.target.files[0];
                                                  if (file) {
                                                    const reader = new FileReader();
                                                    reader.onload = (uploadEvent) => {
                                                      const newPairs = [...q.matchPairs];
                                                      newPairs[pIdx] = { ...newPairs[pIdx], premiseImage: uploadEvent.target.result };
                                                      updateQuestion(sec.id, q.id, { matchPairs: newPairs, _shuffledB: null });
                                                    };
                                                    reader.readAsDataURL(file);
                                                  }
                                                }}
                                              />
                                            </label>
                                          )}
                                        </div>

                                        {/* Column B Item */}
                                        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                                          <span style={{ fontSize: '10px', color: 'var(--text-secondary)' }}>Col B ({String.fromCharCode(97 + (pIdx % 26))})</span>
                                          <input
                                            id={`q__matchb__${sec.id}__${q.id}__${pIdx}`}
                                            type="text"
                                            placeholder="Response text"
                                            value={pair.response || ''}
                                            style={{ padding: '4px 8px', fontSize: '12px' }}
                                            onChange={(e) => {
                                              const newPairs = [...q.matchPairs];
                                              newPairs[pIdx] = { ...newPairs[pIdx], response: e.target.value };
                                              updateQuestion(sec.id, q.id, { matchPairs: newPairs, _shuffledB: null });
                                            }}
                                          />
                                          {pair.responseImage ? (
                                            <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginTop: '2px' }}>
                                              <img
                                                src={pair.responseImage}
                                                alt="Col B Preview"
                                                style={{ width: '40px', height: '40px', objectFit: 'contain', border: '1px solid var(--border-color)', borderRadius: '4px', backgroundColor: '#fff' }}
                                              />
                                              <button
                                                type="button"
                                                style={{ background: 'none', border: 'none', color: 'var(--danger)', cursor: 'pointer', fontSize: '11px' }}
                                                onClick={() => {
                                                  const newPairs = [...q.matchPairs];
                                                  newPairs[pIdx] = { ...newPairs[pIdx], responseImage: '' };
                                                  updateQuestion(sec.id, q.id, { matchPairs: newPairs, _shuffledB: null });
                                                }}
                                              >
                                                ✕ Remove Image
                                              </button>
                                            </div>
                                          ) : (
                                            <label style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', cursor: 'pointer', fontSize: '11px', color: 'var(--accent)', marginTop: '2px' }}>
                                              <ImageIcon size={13} />
                                              <span>Add Image</span>
                                              <input
                                                type="file"
                                                accept="image/*"
                                                style={{ display: 'none' }}
                                                onChange={(e) => {
                                                  const file = e.target.files[0];
                                                  if (file) {
                                                    const reader = new FileReader();
                                                    reader.onload = (uploadEvent) => {
                                                      const newPairs = [...q.matchPairs];
                                                      newPairs[pIdx] = { ...newPairs[pIdx], responseImage: uploadEvent.target.result };
                                                      updateQuestion(sec.id, q.id, { matchPairs: newPairs, _shuffledB: null });
                                                    };
                                                    reader.readAsDataURL(file);
                                                  }
                                                }}
                                              />
                                            </label>
                                          )}
                                        </div>
                                      </div>
                                    </div>
                                  ))}

                                  <div style={{ display: 'flex', gap: '8px', alignItems: 'center', marginTop: '4px' }}>
                                    <input
                                      type="checkbox"
                                      id={`shuffle-${q.id}`}
                                      checked={q.shuffleB}
                                      onChange={(e) => updateQuestion(sec.id, q.id, { shuffleB: e.target.checked, _shuffledB: null })}
                                    />
                                    <label htmlFor={`shuffle-${q.id}`} style={{ fontSize: '11px', textTransform: 'none' }}>
                                      Shuffle Column B in preview/exports
                                    </label>
                                  </div>
                                </div>
                              )}

                              {/* Question Image Fields */}
                              {(sec.type === 'image' || (q.image && sec.type !== 'match_following')) && (
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                                  <label style={{ fontSize: '10px', fontWeight: 'bold' }}>Question Image</label>
                                  {q.image ? (
                                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                      <div 
                                        className="editor-image-preview-container"
                                        onDragOver={handleDragOver}
                                        onDragLeave={handleDragLeave}
                                        onDrop={(e) => handleDropImage(e, sec.id, q.id)}
                                      >
                                        <img
                                          src={q.image}
                                          alt="Question"
                                          style={{
                                            maxWidth: '100%',
                                            maxHeight: '150px',
                                            objectFit: 'contain',
                                            borderRadius: 'var(--radius-sm)',
                                            border: '1px solid var(--border-color)',
                                            padding: '4px',
                                            backgroundColor: 'rgba(255,255,255,0.05)'
                                          }}
                                        />
                                      </div>
                                      <div style={{ display: 'flex', gap: '8px' }}>
                                        <label className="btn btn-secondary btn-sm" style={{ cursor: 'pointer', margin: 0, padding: '4px 8px', fontSize: '11px' }}>
                                          Change Image
                                          <input
                                            type="file"
                                            accept="image/*"
                                            style={{ display: 'none' }}
                                            onChange={(e) => {
                                              const file = e.target.files[0];
                                              if (file) {
                                                const reader = new FileReader();
                                                reader.onload = (uploadEvent) => {
                                                  updateQuestion(sec.id, q.id, { image: uploadEvent.target.result });
                                                };
                                                reader.readAsDataURL(file);
                                              }
                                            }}
                                          />
                                        </label>
                                        <button
                                          className="btn btn-danger btn-sm"
                                          style={{ padding: '4px 8px', fontSize: '11px' }}
                                          onClick={() => updateQuestion(sec.id, q.id, { image: '' })}
                                        >
                                          Remove Image
                                        </button>
                                      </div>
                                    </div>
                                  ) : (
                                    <label 
                                      className="image-upload-dashed-zone"
                                      onDragOver={handleDragOver}
                                      onDragLeave={handleDragLeave}
                                      onDrop={(e) => handleDropImage(e, sec.id, q.id)}
                                    >
                                      <ImageIcon size={20} className="text-secondary" />
                                      <span>Upload Question Image, drag & drop, or paste here</span>
                                      <span style={{ fontSize: '11px', color: 'var(--text-secondary)', opacity: 0.7 }}>
                                        (Click to browse, drag file, or Ctrl+V / Cmd+V to paste)
                                      </span>
                                      <input
                                        type="file"
                                        accept="image/*"
                                        style={{ display: 'none' }}
                                        onChange={(e) => {
                                          const file = e.target.files[0];
                                          if (file) {
                                            const reader = new FileReader();
                                            reader.onload = (uploadEvent) => {
                                              updateQuestion(sec.id, q.id, { image: uploadEvent.target.result });
                                            };
                                            reader.readAsDataURL(file);
                                          }
                                        }}
                                      />
                                    </label>
                                  )}

                                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
                                    <div className="form-group" style={{ margin: 0 }}>
                                      <label style={{ fontSize: '10px' }}>Image Width (px)</label>
                                      <input
                                        type="number"
                                        value={q.imageWidth || 300}
                                        style={{ padding: '4px 8px', fontSize: '12px' }}
                                        onChange={(e) => updateQuestion(sec.id, q.id, { imageWidth: Number(e.target.value) || 0 })}
                                      />
                                    </div>
                                    <div className="form-group" style={{ margin: 0 }}>
                                      <label style={{ fontSize: '10px' }}>Image Height (px)</label>
                                      <input
                                        type="number"
                                        value={q.imageHeight || 200}
                                        style={{ padding: '4px 8px', fontSize: '12px' }}
                                        onChange={(e) => updateQuestion(sec.id, q.id, { imageHeight: Number(e.target.value) || 0 })}
                                      />
                                    </div>
                                  </div>
                                </div>
                              )}

                              {/* Table Question Specific Fields */}
                              {sec.type === 'table' && q.tableData && (
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                                  <label style={{ fontSize: '10px', fontWeight: 'bold' }}>Table Configuration</label>
                                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
                                    <div className="form-group" style={{ margin: 0 }}>
                                      <label style={{ fontSize: '10px' }}>Rows (excl. header)</label>
                                      <input
                                        type="number"
                                        min="1"
                                        max="20"
                                        value={q.tableRows - 1}
                                        style={{ padding: '4px 8px', fontSize: '12px' }}
                                        onChange={(e) => {
                                          const bodyRows = Math.max(1, Math.min(20, Number(e.target.value) || 1));
                                          const totalRows = bodyRows + 1;
                                          const cols = q.tableCols || 3;
                                          const newData = { ...q.tableData };
                                          const currentBodyRows = newData.rows || [];
                                          const newBodyRows = [];
                                          for (let r = 0; r < bodyRows; r++) {
                                            if (r < currentBodyRows.length) {
                                              const existingRow = [...currentBodyRows[r]];
                                              while (existingRow.length < cols) existingRow.push('');
                                              newBodyRows.push(existingRow.slice(0, cols));
                                            } else {
                                              newBodyRows.push(Array(cols).fill(''));
                                            }
                                          }
                                          newData.rows = newBodyRows;
                                          updateQuestion(sec.id, q.id, { tableRows: totalRows, tableData: newData });
                                        }}
                                      />
                                    </div>
                                    <div className="form-group" style={{ margin: 0 }}>
                                      <label style={{ fontSize: '10px' }}>Columns</label>
                                      <input
                                        type="number"
                                        min="1"
                                        max="10"
                                        value={q.tableCols}
                                        style={{ padding: '4px 8px', fontSize: '12px' }}
                                        onChange={(e) => {
                                          const cols = Math.max(1, Math.min(10, Number(e.target.value) || 1));
                                          const newData = { ...q.tableData };
                                          const oldHeaders = newData.headers || [];
                                          const newHeaders = [];
                                          for (let c = 0; c < cols; c++) {
                                            newHeaders.push(c < oldHeaders.length ? oldHeaders[c] : `Column ${c + 1}`);
                                          }
                                          newData.headers = newHeaders;
                                          newData.rows = (newData.rows || []).map(row => {
                                            const newRow = [];
                                            for (let c = 0; c < cols; c++) {
                                              newRow.push(c < row.length ? row[c] : '');
                                            }
                                            return newRow;
                                          });
                                          updateQuestion(sec.id, q.id, { tableCols: cols, tableData: newData });
                                        }}
                                      />
                                    </div>
                                  </div>

                                  {/* Table header cells */}
                                  <label style={{ fontSize: '10px', fontWeight: 'bold' }}>Header Row (Bold)</label>
                                  <div style={{ display: 'grid', gridTemplateColumns: `repeat(${q.tableCols}, 1fr)`, gap: '6px' }}>
                                    {q.tableData.headers.map((h, hIdx) => (
                                      <input
                                        key={hIdx}
                                        id={`q__tblh__${sec.id}__${q.id}__${hIdx}`}
                                        type="text"
                                        value={h}
                                        placeholder={`Header ${hIdx + 1}`}
                                        style={{ padding: '4px 6px', fontSize: '11px', fontWeight: 'bold' }}
                                        onChange={(e) => {
                                          const newData = { ...q.tableData };
                                          const newHeaders = [...newData.headers];
                                          newHeaders[hIdx] = e.target.value;
                                          newData.headers = newHeaders;
                                          updateQuestion(sec.id, q.id, { tableData: newData });
                                        }}
                                      />
                                    ))}
                                  </div>

                                  {/* Table body cells */}
                                  <label style={{ fontSize: '10px', fontWeight: 'bold' }}>Body Rows (Regular)</label>
                                  {q.tableData.rows.map((row, rIdx) => (
                                    <div key={rIdx} style={{ display: 'grid', gridTemplateColumns: `repeat(${q.tableCols}, 1fr)`, gap: '6px' }}>
                                      {row.map((cell, cIdx) => (
                                        <input
                                          key={cIdx}
                                          id={`q__tblc__${sec.id}__${q.id}__${rIdx}__${cIdx}`}
                                          type="text"
                                          value={cell}
                                          placeholder={`R${rIdx + 1}C${cIdx + 1}`}
                                          style={{ padding: '4px 6px', fontSize: '11px' }}
                                          onChange={(e) => {
                                            const newData = { ...q.tableData };
                                            const newRows = newData.rows.map(r => [...r]);
                                            newRows[rIdx][cIdx] = e.target.value;
                                            newData.rows = newRows;
                                            updateQuestion(sec.id, q.id, { tableData: newData });
                                          }}
                                        />
                                      ))}
                                    </div>
                                  ))}
                                </div>
                              )}
                            </div>
                          ))}

                          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '12px', flexWrap: 'wrap', gap: '8px' }}>
                            <div style={{
                              display: 'flex',
                              alignItems: 'center',
                              gap: '8px',
                              fontSize: '12px',
                              fontWeight: 'bold'
                            }}>
                              <span style={{ fontSize: '11px', color: 'var(--text-muted)', textTransform: 'uppercase' }}>
                                Current Section Total:
                              </span>
                              <span style={{
                                padding: '4px 10px',
                                backgroundColor: 'rgba(0,0,0,0.2)',
                                borderRadius: 'var(--radius-sm)',
                                fontSize: '13px',
                                color: isOverProvisioned ? 'var(--warning)' : isUnderProvisioned ? 'var(--danger)' : 'var(--success)'
                              }}>
                                {formatMarks(secTotal)} / {formatMarks(sec.marks)} Marks
                              </span>
                            </div>
                            <button
                              className="btn btn-secondary btn-sm"
                              onClick={() => addQuestion(sec.id)}
                            >
                              + Add Question
                            </button>
                          </div>
                        </div>
                      </>
                    )}
                  </div>
                );
              })}

              <button className="btn btn-secondary" onClick={addSection}>
                <Plus size={16} /> Add New Section
              </button>

              {/* Validation Badges */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', marginTop: '12px' }}>
                {getExamCurrentTotalMarks() !== metadata.maxMarks ? (
                  <div className="warning-badge" style={{ display: 'flex', alignSelf: 'flex-start' }}>
                    <AlertTriangle size={14} />
                    <span>Marks Mismatch: Current Questions = {formatMarks(getExamCurrentTotalMarks())} marks (Target = {formatMarks(metadata.maxMarks)} marks).</span>
                  </div>
                ) : (
                  <div className="warning-badge" style={{ display: 'flex', alignSelf: 'flex-start', backgroundColor: 'rgba(16,185,129,0.15)', color: 'var(--success)', borderColor: 'rgba(16,185,129,0.3)' }}>
                    <CheckCircle size={14} />
                    <span>Perfect: Sum of all questions matches targets!</span>
                  </div>
                )}
                {hasZeroMarkQuestions() && (
                  <div className="warning-badge" style={{ display: 'flex', alignSelf: 'flex-start', backgroundColor: 'rgba(239, 68, 68, 0.15)', color: 'var(--danger)', borderColor: 'rgba(239, 68, 68, 0.3)' }}>
                    <AlertTriangle size={14} />
                    <span>Warning: One or more questions have 0 marks.</span>
                  </div>
                )}
              </div>
            </div>
          )}
        </div>

        {/* Global Action Bar */}
        <div className="action-bar" style={{ flexDirection: 'column', gap: '8px' }}>
          {/* Wizard Navigation */}
          <div style={{ display: 'flex', gap: '8px', width: '100%' }}>
            {activeTab !== 'branding' && (
              <button className="btn btn-secondary" onClick={() => {
                if (activeTab === 'metadata') setActiveTab('branding');
                else if (activeTab === 'sections') setActiveTab('metadata');
              }} style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '6px' }}>
                <ArrowLeft size={14} /> Back
              </button>
            )}
            {activeTab !== 'sections' && (
              <button className="btn btn-primary" onClick={() => {
                if (activeTab === 'branding') setActiveTab('metadata');
                else if (activeTab === 'metadata') setActiveTab('sections');
              }} style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '6px' }}>
                Next <ArrowRight size={14} />
              </button>
            )}
          </div>

          <div style={{ display: 'flex', gap: '8px', width: '100%' }}>
            <button className="btn btn-secondary btn-sm" onClick={loadDemo} style={{ flex: 1 }}>
              Demo Data
            </button>
            <button className="btn btn-danger btn-sm" onClick={resetAll} style={{ flex: 1 }}>
              Clear Draft
            </button>
          </div>
          <div style={{ display: 'flex', gap: '8px', width: '100%' }}>
            <button className="btn btn-primary btn-sm" onClick={exportToCSV} style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '6px' }}>
              <Download size={12} /> Export CSV
            </button>
            <label className="btn btn-secondary btn-sm" style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '6px', cursor: 'pointer', margin: 0 }}>
              <Plus size={12} /> Import CSV
              <input type="file" accept=".csv" onChange={importFromCSV} style={{ display: 'none' }} />
            </label>
          </div>
        </div>

      </div>

      {/* Preview Modal */}
      {isPreviewOpen && (
        <div className="modal-overlay preview-overlay" onClick={() => setIsPreviewOpen(false)}>
          <div className="modal-content preview-modal-content" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h3>Question Paper Live Preview</h3>
            </div>

            {/* Live Preview Controls Bar */}
            <div className="preview-options-bar" style={{
              display: 'flex',
              gap: '24px',
              padding: '12px 24px',
              borderBottom: '1px solid var(--border-color)',
              backgroundColor: 'var(--bg-editor)',
              alignItems: 'center',
              flexWrap: 'wrap'
            }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <input
                  type="checkbox"
                  id="preview-hideSchoolLogo"
                  checked={branding.hideSchoolLogo || false}
                  onChange={(e) => setBranding({ ...branding, hideSchoolLogo: e.target.checked })}
                  style={{ width: 'auto', cursor: 'pointer', margin: 0 }}
                />
                <label htmlFor="preview-hideSchoolLogo" style={{ cursor: 'pointer', marginBottom: 0, userSelect: 'none', fontSize: '13px', fontWeight: '500', color: 'var(--text-primary)' }}>
                  Hide School Logo
                </label>
              </div>

              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <input
                  type="checkbox"
                  id="preview-headerLogoOnly"
                  checked={branding.headerLogoOnly || false}
                  onChange={(e) => setBranding({ ...branding, headerLogoOnly: e.target.checked })}
                  style={{ width: 'auto', cursor: 'pointer', margin: 0 }}
                />
                <label htmlFor="preview-headerLogoOnly" style={{ cursor: 'pointer', marginBottom: 0, userSelect: 'none', fontSize: '13px', fontWeight: '500', color: 'var(--text-primary)' }}>
                  Hide School Name & Address
                </label>
              </div>

              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <input
                  type="checkbox"
                  id="preview-separateAnswerSheet"
                  checked={metadata.separateAnswerSheet ?? true}
                  onChange={(e) => setMetadata({ ...metadata, separateAnswerSheet: e.target.checked })}
                  style={{ width: 'auto', cursor: 'pointer', margin: 0 }}
                />
                <label htmlFor="preview-separateAnswerSheet" style={{ cursor: 'pointer', marginBottom: 0, userSelect: 'none', fontSize: '13px', fontWeight: '500', color: 'var(--text-primary)' }}>
                  Write answers on a separate sheet (do not print blanks)
                </label>
              </div>

              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <label htmlFor="preview-fontFamily" style={{ marginBottom: 0, userSelect: 'none', fontSize: '13px', fontWeight: '500', color: 'var(--text-primary)' }}>
                  Branding Font:
                </label>
                <select
                  id="preview-fontFamily"
                  value={branding.fontFamily}
                  onChange={(e) => setBranding({ ...branding, fontFamily: e.target.value })}
                  style={{ width: 'auto', padding: '4px 8px', fontSize: '12px', height: 'auto', border: '1px solid var(--border-color)', borderRadius: '4px', backgroundColor: 'var(--bg-card)', color: 'var(--text-primary)' }}
                >
                  <option value="Inter">Inter (Clean Modern)</option>
                  <option value="Playfair">Playfair Display (Elegant Serif)</option>
                  <option value="Montserrat">Montserrat (Geometric Sans)</option>
                  <option value="Courier">Courier Prime (Monospace / Classic)</option>
                  <option value="Merriweather">Merriweather (Soft Serif)</option>
                  <option value="Cinzel">Cinzel (Regal / Classical)</option>
                </select>
              </div>
            </div>

            {getExamCurrentTotalMarks() !== metadata.maxMarks && (
              <div style={{ padding: '12px 24px', borderBottom: '1px solid var(--border-color)', backgroundColor: 'var(--bg-editor)', display: 'flex' }} className="warning-badge-container print-hide">
                <div className="warning-badge" style={{ display: 'flex', width: '100%', boxSizing: 'border-box' }}>
                  <AlertTriangle size={14} />
                  <span>Marks Mismatch: Current Questions = {formatMarks(getExamCurrentTotalMarks())} marks (Target = {formatMarks(metadata.maxMarks)} marks).</span>
                </div>
              </div>
            )}

            {hasZeroMarkQuestions() && (
              <div style={{ padding: '12px 24px', borderBottom: '1px solid var(--border-color)', backgroundColor: 'var(--bg-editor)', display: 'flex' }} className="warning-badge-container print-hide">
                <div className="warning-badge" style={{ display: 'flex', width: '100%', boxSizing: 'border-box', backgroundColor: 'rgba(239, 68, 68, 0.1)', color: 'var(--danger)', borderColor: 'rgba(239, 68, 68, 0.2)' }}>
                  <AlertTriangle size={14} />
                  <span>Warning: One or more questions have 0 marks assigned.</span>
                </div>
              </div>
            )}

            <div className="modal-body">
              {/* Dynamic A4 Preview Sheet */}
              <div ref={paperSheetRef} className={`paper-sheet lang-${metadata.language || 'english'}`}>

                {/* Header Layout */}
                <div className={`paper-header font-${branding.fontFamily}`}>
                  {branding.logo && !branding.hideSchoolLogo && (
                    <div
                      ref={logoRef}
                      className={`brand-logo-container ${isDragging ? 'dragging' : ''} ${branding.headerLogoOnly ? 'centered' : ''}`}
                      style={branding.headerLogoOnly ? {
                        width: `${branding.logoWidth || 100}px`,
                        height: `${branding.logoHeight || 100}px`
                      } : {
                        width: `${branding.logoWidth || 100}px`,
                        height: `${branding.logoHeight || 100}px`,
                        left: `${branding.logoPos.x}px`,
                        top: `${branding.logoPos.y}px`
                      }}
                      onPointerDown={branding.headerLogoOnly ? undefined : handleLogoPointerDown}
                      onPointerMove={branding.headerLogoOnly ? undefined : handleLogoPointerMove}
                      onPointerUp={branding.headerLogoOnly ? undefined : handleLogoPointerUp}
                    >
                      {!branding.headerLogoOnly && <div className="drag-indicator">Drag to move</div>}
                      <img
                        src={branding.logo}
                        className="brand-logo-img"
                        alt="School Logo"
                        onError={(e) => {
                          e.currentTarget.onerror = null;
                          if (branding.logo !== schoolLogo) {
                            setBranding(prev => ({ ...prev, logo: schoolLogo }));
                          }
                        }}
                      />
                    </div>
                  )}

                  {!branding.headerLogoOnly && (
                    <div className="school-details">
                      {branding.schoolName && <h1 className="school-name-render">{branding.schoolName}</h1>}
                      {branding.schoolAddress && <p className="school-address-render">{branding.schoolAddress}</p>}
                    </div>
                  )}
                </div>

                {/* Exam Specs Row */}
                <div className="exam-meta-grid">
                  <div className="exam-meta-item full-width">
                    <span className="exam-meta-label">Examination:</span>
                    <span>{metadata.title || '_______________________'}</span>
                  </div>
                  <div className="exam-meta-item full-width">
                    <span className="exam-meta-label">Subject:</span>
                    <span>{metadata.subject || '_______________________'}</span>
                  </div>
                  <div className="exam-meta-item full-width">
                    <span className="exam-meta-label">Class:</span>
                    <span>{metadata.classDiv || '_______________________'}</span>
                  </div>
                  <div className="exam-meta-item full-width">
                    <span className="exam-meta-label">Max Marks:</span>
                    <span>{formatMarks(metadata.maxMarks)}</span>
                  </div>
                  <div className="exam-meta-item full-width">
                    <span className="exam-meta-label">Duration:</span>
                    <span>{metadata.duration || '_______________________'}</span>
                  </div>
                </div>

                {/* Render Sections & Questions */}
                <div className="paper-sections-container">
                  {sections.length === 0 ? (
                    <div style={{ textAlign: 'center', color: '#999', padding: '40px 0', fontSize: '14px', fontStyle: 'italic' }}>
                      No sections added yet. Use the "Questions" tab in the editor to create sections and populate questions.
                    </div>
                  ) : (
                    sections.map((sec, sIdx) => (
                      <div key={sec.id} className="paper-section">
                        <div className="paper-section-header">
                          <div className="paper-section-title-row">
                            <h2 className="paper-section-title">{sec.title}</h2>
                            <span className="paper-section-marks">[{formatMarks(sec.marks)} Marks]</span>
                          </div>
                          {sec.instructions && (
                            <p className="paper-section-instructions">{sec.instructions}</p>
                          )}
                        </div>

                        <div className="paper-questions-list">
                          {sec.questions.map((q, qIdx) => {
                            // We compute the global question count.
                            // Let's count all questions from previous sections
                            let previousQuestionsCount = 0;
                            for (let i = 0; i < sIdx; i++) {
                              previousQuestionsCount += sections[i].questions.length;
                            }
                            const globalNum = previousQuestionsCount + qIdx + 1;

                            return (
                              <div key={q.id} className="paper-question-item">
                                <span className="paper-question-number">Q{globalNum}.</span>
                                <div className="paper-question-body">
                                  <p style={{ fontWeight: '500', textAlign: 'justify', textAlignLast: 'left' }} dangerouslySetInnerHTML={{ __html: renderTextWithMath(q.text) }} />

                                  {/* MCQ Options */}
                                  {sec.type === 'mcq' && q.options && (
                                    <div className={`paper-mcq-options ${canFitSingleLine(q.options) ? 'single-line' : ''}`}>
                                      {q.options.map((opt, oIdx) => (
                                        <div key={oIdx} className="paper-mcq-option">
                                          <span style={{ fontWeight: '600', flexShrink: 0 }}>({String.fromCharCode(65 + oIdx)})</span>
                                          <span className="paper-mcq-option-text" dangerouslySetInnerHTML={{ __html: renderTextWithMath(opt) }} />
                                        </div>
                                      ))}
                                    </div>
                                  )}

                                  {/* Essay spaces and Sub-Questions */}
                                  {sec.type === 'essay' && (
                                    <>
                                      {q.subQuestions && q.subQuestions.length > 0 ? (
                                        <div className="paper-subquestions-list" style={{ marginTop: '8px', paddingLeft: '16px', display: 'flex', flexDirection: 'column', gap: '10px' }}>
                                          {q.subQuestions.map((sq, sqIdx) => (
                                            <div key={sq.id} className="paper-subquestion-item" style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                                              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: '12px' }}>
                                                <div style={{ display: 'flex', gap: '6px', flex: 1 }}>
                                                  <span style={{ fontWeight: '600', flexShrink: 0 }}>{sq.label || `(${String.fromCharCode(97 + (sqIdx % 26))})`}</span>
                                                  <span style={{ flex: 1, textAlign: 'justify', textAlignLast: 'left' }} dangerouslySetInnerHTML={{ __html: renderTextWithMath(sq.text) }} />
                                                </div>
                                                <span className="paper-question-marks" style={{ fontStyle: 'italic', fontSize: '12px' }}>({formatMarks(sq.marks)} M)</span>
                                              </div>
                                              {!metadata.separateAnswerSheet && (
                                                <div style={{ display: 'flex', flexDirection: 'column', gap: '12px', marginTop: '4px', paddingLeft: '20px' }}>
                                                  {Array.from({ length: (sq.blankLines !== undefined && sq.blankLines !== '') ? sq.blankLines : 4 }).map((_, lineIdx) => (
                                                    <div key={lineIdx} className="paper-answer-line"></div>
                                                  ))}
                                                </div>
                                              )}
                                            </div>
                                          ))}
                                        </div>
                                      ) : (
                                        !metadata.separateAnswerSheet && (
                                          <div style={{ display: 'flex', flexDirection: 'column', gap: '14px', marginTop: '8px' }}>
                                            {Array.from({ length: (q.blankLines !== undefined && q.blankLines !== '') ? q.blankLines : 5 }).map((_, lineIdx) => (
                                              <div key={lineIdx} className="paper-answer-line"></div>
                                            ))}
                                          </div>
                                        )
                                      )}
                                    </>
                                  )}

                                  {/* True/False selection */}
                                  {sec.type === 'true_false' && !metadata.separateAnswerSheet && (
                                    <div className="paper-tf-options">
                                      <span>[   ] True</span>
                                      <span>[   ] False</span>
                                    </div>
                                  )}

                                  {/* Match the Following columns */}
                                  {sec.type === 'match_following' && q.matchPairs && (
                                    <table className="paper-match-table">
                                      <thead>
                                        <tr>
                                          <th>Column A</th>
                                          <th style={{ paddingLeft: '20px' }}>Column B</th>
                                        </tr>
                                      </thead>
                                      <tbody>
                                        {q.matchPairs.map((pair, pIdx) => {
                                          const shuffledList = getShuffledList(q);
                                          const roman = (idx) => {
                                            const r = ['i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x'];
                                            return r[idx] || (idx + 1).toString();
                                          };
                                          const itemB = shuffledList[pIdx] || (typeof pair === 'string' ? { response: pair } : pair);
                                          const respText = typeof itemB === 'string' ? itemB : (itemB.response || '');
                                          const respImg = typeof itemB === 'object' ? itemB.responseImage : '';

                                          return (
                                            <tr key={pIdx}>
                                              <td style={{ padding: '6px 0', verticalAlign: 'top' }}>
                                                <div style={{ display: 'flex', gap: '6px', alignItems: 'flex-start' }}>
                                                  <span style={{ fontWeight: '600', flexShrink: 0 }}>{pIdx + 1}.</span>
                                                  <div>
                                                    {pair.premise && <span dangerouslySetInnerHTML={{ __html: renderTextWithMath(pair.premise) }} />}
                                                    {pair.premiseImage && (
                                                      <img
                                                        src={pair.premiseImage}
                                                        alt={`Col A ${pIdx + 1}`}
                                                        style={{ maxHeight: '90px', maxWidth: '140px', objectFit: 'contain', display: 'block', marginTop: '4px', borderRadius: '4px' }}
                                                      />
                                                    )}
                                                  </div>
                                                </div>
                                              </td>
                                              <td style={{ padding: '6px 0', paddingLeft: '20px', verticalAlign: 'top' }}>
                                                <div style={{ display: 'flex', gap: '6px', alignItems: 'flex-start' }}>
                                                  <span style={{ fontWeight: '600', flexShrink: 0 }}>{roman(pIdx)}.</span>
                                                  <div>
                                                    {respText && <span dangerouslySetInnerHTML={{ __html: renderTextWithMath(respText) }} />}
                                                    {respImg && (
                                                      <img
                                                        src={respImg}
                                                        alt={`Col B ${pIdx + 1}`}
                                                        style={{ maxHeight: '90px', maxWidth: '140px', objectFit: 'contain', display: 'block', marginTop: '4px', borderRadius: '4px' }}
                                                      />
                                                    )}
                                                  </div>
                                                </div>
                                              </td>
                                            </tr>
                                          );
                                        })}
                                      </tbody>
                                    </table>
                                  )}

                                  {/* Question Image render */}
                                  {q.image && sec.type !== 'match_following' && (
                                    <div className="paper-image-container" style={{ marginTop: '8px', display: 'flex', justifyContent: 'flex-start' }}>
                                      <img
                                        src={q.image}
                                        alt={`Question ${globalNum}`}
                                        style={{
                                          width: `${q.imageWidth || 300}px`,
                                          height: `${q.imageHeight || 200}px`,
                                          objectFit: 'contain',
                                          maxWidth: '100%'
                                        }}
                                      />
                                    </div>
                                  )}

                                  {/* Table question render */}
                                  {sec.type === 'table' && q.tableData && (
                                    <table className="paper-table-question">
                                      <thead>
                                        <tr>
                                          {q.tableData.headers.map((h, hIdx) => (
                                            <th key={hIdx} dangerouslySetInnerHTML={{ __html: renderTextWithMath(h) }} />
                                          ))}
                                        </tr>
                                      </thead>
                                      <tbody>
                                        {q.tableData.rows.map((row, rIdx) => (
                                          <tr key={rIdx}>
                                            {row.map((cell, cIdx) => (
                                              <td key={cIdx} dangerouslySetInnerHTML={{ __html: renderTextWithMath(cell) }} />
                                            ))}
                                          </tr>
                                        ))}
                                      </tbody>
                                    </table>
                                  )}
                                </div>
                                <span className="paper-question-marks">({formatMarks(getQuestionMarks(q))} M)</span>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    ))
                  )}
                </div>

                {/* Page number footer */}
                <div className="paper-footer"></div>
              </div>
            </div>

            <div className="modal-footer" style={{ gap: '10px' }}>
              <button
                className="btn btn-primary"
                disabled={isPdfExporting || !hasQuestions()}
                title={hasQuestions() ? "Download PDF" : "Add at least one question to download"}
                onClick={triggerPdfExport}
                style={{ display: 'flex', alignItems: 'center', gap: '6px' }}
              >
                {isPdfExporting ? <Loader2 size={16} className="animate-spin" /> : <FileText size={16} />}
                {isPdfExporting ? 'Generating PDF...' : 'Download PDF'}
              </button>

              <button
                className="btn btn-docx"
                disabled={!hasQuestions()}
                title={hasQuestions() ? "Download Word (DOCX)" : "Add at least one question to download"}
                onClick={triggerDocxExport}
                style={{ display: 'flex', alignItems: 'center', gap: '6px' }}
              >
                <FileText size={16} style={{ color: '#ffffff' }} />
                Download Word (DOCX)
              </button>

              <button className="btn btn-danger" onClick={() => setIsPreviewOpen(false)}>
                Close
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Google Docs Integration Modal */}
      {isDocsModalOpen && (
        <div className="modal-overlay" style={{ zIndex: 1100 }} onClick={() => setIsDocsModalOpen(false)}>
          <div className="modal-content docs-modal-content" onClick={(e) => e.stopPropagation()} style={{ maxWidth: '640px' }}>
            <div className="modal-header">
              <h3 style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <svg viewBox="0 0 24 24" width="20" height="20" style={{ fill: '#4285f4' }}>
                  <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
                </svg>
                Open in Google Docs
              </h3>
            </div>
            <div className="modal-body docs-modal-body" style={{ flexDirection: 'column', gap: '20px', backgroundColor: 'var(--bg-sidebar)', padding: '24px' }}>
              <p style={{ fontSize: '14px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                Google Docs requires your question paper (.docx) to be uploaded to your Google account. Select your preferred method below:
              </p>

              <div className="docs-options-grid" style={{ display: 'flex', flexDirection: 'column', gap: '16px', width: '100%' }}>
                {/* Option 1: Web Preview (Secure Upload) */}
                <div className="docs-option-card" style={{
                  border: '1px solid var(--border-color)',
                  borderRadius: 'var(--radius-md)',
                  padding: '16px',
                  backgroundColor: 'var(--bg-card)',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '12px'
                }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <Cloud size={18} style={{ color: '#4285f4' }} />
                    <strong style={{ fontSize: '15px', color: 'var(--text-primary)' }}>Option A: Instant Web Preview</strong>
                  </div>
                  <p style={{ fontSize: '13px', color: 'var(--text-secondary)', lineHeight: '1.4' }}>
                    Temporarily uploads a secure copy of your document (valid for 1 hour) so Google's viewer can load it. Once open, click <strong>"Open with Google Docs"</strong> at the top to edit.
                  </p>
                  
                  <div className="docs-warning-alert" style={{
                    backgroundColor: 'rgba(245, 158, 11, 0.1)',
                    borderLeft: '3px solid var(--warning)',
                    padding: '8px 12px',
                    borderRadius: '4px',
                    fontSize: '12px',
                    color: 'var(--warning)',
                    lineHeight: '1.4'
                  }}>
                    <strong>⚠️ Privacy Note:</strong> Do not use this for actual confidential school exams, as it uploads the document to a temporary public URL.
                  </div>

                  <button
                    className="btn btn-primary"
                    disabled={isDocsUploading}
                    onClick={handleGoogleDocsWebPreview}
                    style={{ alignSelf: 'flex-start', marginTop: '4px', gap: '8px', minWidth: '160px' }}
                  >
                    {isDocsUploading ? (
                      <>
                        <Loader2 size={16} className="animate-spin" />
                        <span>Uploading...</span>
                      </>
                    ) : (
                      <>
                        <ExternalLink size={16} />
                        <span>Proceed to Google Docs</span>
                      </>
                    )}
                  </button>

                  {docsError && (
                    <div style={{ color: 'var(--danger)', fontSize: '12px', marginTop: '4px' }}>
                      {docsError}
                    </div>
                  )}
                </div>

                {/* Option 2: Offline Import (100% Private) */}
                <div className="docs-option-card" style={{
                  border: '1px solid var(--border-color)',
                  borderRadius: 'var(--radius-md)',
                  padding: '16px',
                  backgroundColor: 'var(--bg-card)',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '12px'
                }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <Shield size={18} style={{ color: 'var(--success)' }} />
                    <strong style={{ fontSize: '15px', color: 'var(--text-primary)' }}>Option B: Offline Import (100% Private)</strong>
                  </div>
                  <p style={{ fontSize: '13px', color: 'var(--text-secondary)', lineHeight: '1.4' }}>
                    Download the Word (.docx) file locally to your machine, then manually upload or drag it directly into Google Drive or Docs.
                  </p>
                  
                  <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
                    <button
                      className="btn btn-secondary"
                      onClick={() => {
                        triggerDocxExport();
                        window.open('https://drive.google.com/', '_blank');
                      }}
                      style={{ fontSize: '13px', display: 'inline-flex', alignItems: 'center', gap: '6px' }}
                    >
                      <Download size={14} /> Download & Open Drive
                    </button>
                    <button
                      className="btn btn-secondary"
                      onClick={() => {
                        triggerDocxExport();
                        window.open('https://docs.google.com/document/', '_blank');
                      }}
                      style={{ fontSize: '13px', display: 'inline-flex', alignItems: 'center', gap: '6px' }}
                    >
                      <Download size={14} /> Download & Open Docs
                    </button>
                  </div>
                </div>
              </div>
            </div>
            <div className="modal-footer" style={{ borderTop: '1px solid var(--border-color)', backgroundColor: 'var(--bg-card)' }}>
              <button className="btn btn-secondary" onClick={() => setIsDocsModalOpen(false)}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}
      {/* Floating Formula FAB */}
      {!isPreviewOpen && activeInputInfo && (
        <button
          type="button"
          className="floating-formula-btn print-hide"
          onClick={handleFloatingFormulaClick}
          title="Insert formula into selected inputbox"
        >
          <span style={{ fontStyle: 'italic', fontWeight: 'bold' }}>𝑓</span> Formula
        </button>
      )}

      {/* Formula Editor Modal */}
      {formulaModal.isOpen && (
        <div className="modal-overlay" onClick={() => setFormulaModal({ ...formulaModal, isOpen: false })} style={{ zIndex: 10000 }}>
          <div className="modal-content formula-modal-content" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h3 style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <span style={{ fontStyle: 'italic', fontWeight: 'bold', fontSize: '20px' }}>𝑓(x)</span> Formula Editor
              </h3>
            </div>

            <div className="modal-body" style={{ display: 'flex', flexDirection: 'column', gap: '16px', maxHeight: '70vh', overflowY: 'auto' }}>
              {/* Toolbar */}
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
                          onClick={() => {
                            if (btn.replace) {
                              setFormulaModal(prev => ({ ...prev, latex: btn.latex }));
                            } else {
                              setFormulaModal(prev => ({ ...prev, latex: prev.latex + ' ' + btn.latex }));
                            }
                            formulaInputRef.current?.focus();
                          }}
                        >
                          {btn.symbol}
                        </button>
                      ))}
                    </div>
                  </div>
                ))}
              </div>

              {/* LaTeX Input */}
              <div className="form-group" style={{ margin: 0 }}>
                <label style={{ fontSize: '11px', fontWeight: 'bold' }}>LaTeX Formula</label>
                <textarea
                  ref={formulaInputRef}
                  value={formulaModal.latex}
                  onChange={(e) => setFormulaModal(prev => ({ ...prev, latex: e.target.value }))}
                  placeholder="Type LaTeX here, e.g. x = \frac{-b \pm \sqrt{b^2-4ac}}{2a}"
                  style={{ minHeight: '80px', fontSize: '14px', fontFamily: 'monospace' }}
                />
              </div>

              {/* Live Preview */}
              <div className="formula-live-preview">
                <label style={{ fontSize: '11px', fontWeight: 'bold', marginBottom: '6px', display: 'block' }}>Live Preview</label>
                <div className="formula-preview-box">
                  {formulaModal.latex ? (
                    <div dangerouslySetInnerHTML={{ __html: renderLatex(formulaModal.latex) }} />
                  ) : (
                    <span style={{ color: 'var(--text-muted)', fontStyle: 'italic' }}>Formula preview will appear here...</span>
                  )}
                </div>
              </div>
            </div>

            <div className="modal-footer" style={{ gap: '8px' }}>
              <button className="btn btn-secondary" onClick={() => setFormulaModal({ isOpen: false, latex: '', onSave: null })}>
                Cancel
              </button>
              <button
                className="btn btn-primary"
                onClick={() => {
                  if (formulaModal.onSave) {
                    formulaModal.onSave(formulaModal.latex);
                  }
                  setFormulaModal({ isOpen: false, latex: '', onSave: null });
                }}
              >
                Insert Formula
              </button>
            </div>
          </div>
        </div>
      )}

      {/* About Developer Modal */}
      {/* CSV Import Options Modal */}
      {csvImportModal.isOpen && (
        <div className="modal-overlay" style={{ zIndex: 1200 }} onClick={() => setCsvImportModal({ ...csvImportModal, isOpen: false })}>
          <div className="modal-content" style={{ maxWidth: '420px' }} onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h3 style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Upload size={18} style={{ color: 'var(--accent)' }} />
                Import Options
              </h3>
            </div>
            <div className="modal-body" style={{ display: 'flex', flexDirection: 'column', gap: '12px', padding: '20px' }}>
              <p style={{ fontSize: '13px', color: 'var(--text-muted)', margin: 0, lineHeight: '1.5' }}>
                Select which sections to import from the CSV file:
              </p>

              {/* School Details */}
              <label
                style={{
                  display: 'flex', alignItems: 'center', gap: '12px', padding: '12px 14px',
                  backgroundColor: csvImportModal.importSchool ? 'var(--accent-glow)' : 'var(--bg-main)',
                  borderRadius: '10px', border: `1.5px solid ${csvImportModal.importSchool ? 'var(--accent)' : 'var(--border-color)'}`,
                  cursor: csvImportModal.branding ? 'pointer' : 'not-allowed', opacity: csvImportModal.branding ? 1 : 0.45,
                  transition: 'all 0.2s ease'
                }}
              >
                <input
                  type="checkbox"
                  checked={csvImportModal.importSchool}
                  disabled={!csvImportModal.branding}
                  onChange={(e) => setCsvImportModal({ ...csvImportModal, importSchool: e.target.checked })}
                  style={{ display: 'none' }}
                />
                <div style={{
                  width: '22px', height: '22px', borderRadius: '6px', flexShrink: 0,
                  border: `2px solid ${csvImportModal.importSchool ? 'var(--accent)' : 'var(--border-color)'}`,
                  backgroundColor: csvImportModal.importSchool ? 'var(--accent)' : 'transparent',
                  display: 'flex', alignItems: 'center', justifyContent: 'center', transition: 'all 0.2s ease'
                }}>
                  {csvImportModal.importSchool && <Check size={14} style={{ color: '#fff' }} />}
                </div>
                <School size={18} style={{ color: csvImportModal.importSchool ? 'var(--accent)' : 'var(--text-muted)', flexShrink: 0 }} />
                <div style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                  <span style={{ fontSize: '14px', fontWeight: '600', color: 'var(--text-primary)' }}>School Details</span>
                  <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                    {csvImportModal.branding ? 'School name, address, logo & font' : 'Not found in CSV'}
                  </span>
                </div>
              </label>

              {/* Exam Details */}
              <label
                style={{
                  display: 'flex', alignItems: 'center', gap: '12px', padding: '12px 14px',
                  backgroundColor: csvImportModal.importExam ? 'var(--accent-glow)' : 'var(--bg-main)',
                  borderRadius: '10px', border: `1.5px solid ${csvImportModal.importExam ? 'var(--accent)' : 'var(--border-color)'}`,
                  cursor: csvImportModal.metadata ? 'pointer' : 'not-allowed', opacity: csvImportModal.metadata ? 1 : 0.45,
                  transition: 'all 0.2s ease'
                }}
              >
                <input
                  type="checkbox"
                  checked={csvImportModal.importExam}
                  disabled={!csvImportModal.metadata}
                  onChange={(e) => setCsvImportModal({ ...csvImportModal, importExam: e.target.checked })}
                  style={{ display: 'none' }}
                />
                <div style={{
                  width: '22px', height: '22px', borderRadius: '6px', flexShrink: 0,
                  border: `2px solid ${csvImportModal.importExam ? 'var(--accent)' : 'var(--border-color)'}`,
                  backgroundColor: csvImportModal.importExam ? 'var(--accent)' : 'transparent',
                  display: 'flex', alignItems: 'center', justifyContent: 'center', transition: 'all 0.2s ease'
                }}>
                  {csvImportModal.importExam && <Check size={14} style={{ color: '#fff' }} />}
                </div>
                <GraduationCap size={18} style={{ color: csvImportModal.importExam ? 'var(--accent)' : 'var(--text-muted)', flexShrink: 0 }} />
                <div style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                  <span style={{ fontSize: '14px', fontWeight: '600', color: 'var(--text-primary)' }}>Exam Details</span>
                  <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                    {csvImportModal.metadata ? 'Title, subject, class, marks & duration' : 'Not found in CSV'}
                  </span>
                </div>
              </label>

              {/* Questions */}
              <label
                style={{
                  display: 'flex', alignItems: 'center', gap: '12px', padding: '12px 14px',
                  backgroundColor: csvImportModal.importQuestions ? 'var(--accent-glow)' : 'var(--bg-main)',
                  borderRadius: '10px', border: `1.5px solid ${csvImportModal.importQuestions ? 'var(--accent)' : 'var(--border-color)'}`,
                  cursor: csvImportModal.sections.length > 0 ? 'pointer' : 'not-allowed', opacity: csvImportModal.sections.length > 0 ? 1 : 0.45,
                  transition: 'all 0.2s ease'
                }}
              >
                <input
                  type="checkbox"
                  checked={csvImportModal.importQuestions}
                  disabled={csvImportModal.sections.length === 0}
                  onChange={(e) => setCsvImportModal({ ...csvImportModal, importQuestions: e.target.checked })}
                  style={{ display: 'none' }}
                />
                <div style={{
                  width: '22px', height: '22px', borderRadius: '6px', flexShrink: 0,
                  border: `2px solid ${csvImportModal.importQuestions ? 'var(--accent)' : 'var(--border-color)'}`,
                  backgroundColor: csvImportModal.importQuestions ? 'var(--accent)' : 'transparent',
                  display: 'flex', alignItems: 'center', justifyContent: 'center', transition: 'all 0.2s ease'
                }}>
                  {csvImportModal.importQuestions && <Check size={14} style={{ color: '#fff' }} />}
                </div>
                <BookOpen size={18} style={{ color: csvImportModal.importQuestions ? 'var(--accent)' : 'var(--text-muted)', flexShrink: 0 }} />
                <div style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                  <span style={{ fontSize: '14px', fontWeight: '600', color: 'var(--text-primary)' }}>Questions</span>
                  <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                    {csvImportModal.sections.length > 0
                      ? `${csvImportModal.sections.length} section(s), ${csvImportModal.sections.reduce((sum, s) => sum + s.questions.length, 0)} question(s)`
                      : 'Not found in CSV'}
                  </span>
                </div>
              </label>
            </div>
            <div className="modal-footer" style={{ display: 'flex', gap: '10px', padding: '16px 20px' }}>
              <button
                className="btn btn-secondary"
                style={{ flex: 1 }}
                onClick={() => setCsvImportModal({ ...csvImportModal, isOpen: false })}
              >
                Cancel
              </button>
              <button
                className="btn btn-primary"
                style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '6px' }}
                onClick={confirmCsvImport}
              >
                <Upload size={14} /> Import Selected
              </button>
            </div>
          </div>
        </div>
      )}

      {isAboutModalOpen && (
        <div className="modal-backdrop" onClick={() => setIsAboutModalOpen(false)}>
          <div className="modal-content" style={{ maxWidth: '440px' }} onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h3 className="modal-title" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="var(--accent)" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" />
                  <circle cx="12" cy="7" r="4" />
                </svg>
                About Developer
              </h3>
              <button className="modal-close" onClick={() => setIsAboutModalOpen(false)}>×</button>
            </div>
            <div className="modal-body" style={{ display: 'flex', flexDirection: 'column', gap: '16px', padding: '20px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '14px', borderBottom: '1px solid var(--border-color)', paddingBottom: '16px' }}>
                <div style={{ width: '52px', height: '52px', borderRadius: '50%', backgroundColor: 'var(--accent-glow)', display: 'flex', alignItems: 'center', justifyContent: 'center', border: '2px solid var(--accent)', color: 'var(--accent)', fontWeight: 'bold', fontSize: '20px' }}>
                  NJ
                </div>
                <div>
                  <h4 style={{ margin: 0, fontSize: '18px', fontWeight: '700', color: 'var(--text-main)' }}>Naveen Joshy</h4>
                  <span style={{ fontSize: '13px', color: 'var(--text-muted)' }}>QuestionNinja Creator & Developer</span>
                </div>
              </div>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '12px', fontSize: '14px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '12px', padding: '10px 14px', backgroundColor: 'var(--bg-main)', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="var(--accent)" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <rect width="20" height="16" x="2" y="4" rx="2" />
                    <path d="m22 7-8.97 5.7a1.94 1.94 0 0 1-2.06 0L2 7" />
                  </svg>
                  <div>
                    <div style={{ fontSize: '11px', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Mail ID</div>
                    <a href="mailto:naveenjoshy64@gmail.com" style={{ color: 'var(--accent)', textDecoration: 'none', fontWeight: '500' }}>naveenjoshy64@gmail.com</a>
                  </div>
                </div>

                <div style={{ display: 'flex', alignItems: 'center', gap: '12px', padding: '10px 14px', backgroundColor: 'var(--bg-main)', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="var(--accent)" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z" />
                  </svg>
                  <div>
                    <div style={{ fontSize: '11px', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Phone</div>
                    <a href="tel:+919400489149" style={{ color: 'var(--text-main)', textDecoration: 'none', fontWeight: '500' }}>+91 9400489149</a>
                  </div>
                </div>

                <div style={{ display: 'flex', alignItems: 'center', gap: '12px', padding: '10px 14px', backgroundColor: 'var(--bg-main)', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="var(--accent)" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M15 22v-4a4.8 4.8 0 0 0-1-3.5c3 0 6-2 6-5.5.08-1.25-.27-2.48-1-3.5.28-1.15.28-2.35 0-3.5 0 0-1 0-3 1.5-2.64-.5-5.36-.5-8 0C6 2 5 2 5 2c-.3 1.15-.3 2.35 0 3.5A5.403 5.403 0 0 0 4 9c0 3.5 3 5.5 6 5.5-.39.49-.68 1.05-.85 1.65-.17.6-.22 1.23-.15 1.85v4" />
                    <path d="M9 18c-4.51 2-5-2-7-2" />
                  </svg>
                  <div>
                    <div style={{ fontSize: '11px', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>GitHub</div>
                    <a href="https://github.com/naveenjoshy" target="_blank" rel="noopener noreferrer" style={{ color: 'var(--accent)', textDecoration: 'none', fontWeight: '500' }}>github.com/naveenjoshy</a>
                  </div>
                </div>
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-primary" style={{ width: '100%' }} onClick={() => setIsAboutModalOpen(false)}>
                Close
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Fixed App Footer */}
      <footer className="app-footer-fixed">
        <span>© {new Date().getFullYear()} Mr. Naveen Joshy. All Rights Reserved.</span>
        <span>Contact: <a href="#about" onClick={(e) => { e.preventDefault(); setIsAboutModalOpen(true); }} style={{ color: 'var(--accent)', textDecoration: 'none', fontWeight: '500', cursor: 'pointer' }}>naveenjoshy64@gmail.com</a></span>
      </footer>
    </div>
  );
}



//hello this is new update from naveen, please check it.