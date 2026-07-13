import React, { useState, useEffect, useRef } from 'react';
import {
  Plus,
  Trash2,
  ArrowUp,
  ArrowDown,
  Download,
  FileText,
  Settings,
  BookOpen,
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
  Shield
} from 'lucide-react';
import * as docx from 'docx';
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

// Helper: render text that contains $...$ math blocks using KaTeX
const renderTextWithMath = (text) => {
  if (!text) return '';
  return text.replace(/\$([^$\n]+?)\$/g, (match, mathContent) => {
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
};

// Helper: convert text with $...$ math into an array of docx.TextRun objects
const docxTextRunsWithMath = (text, defaultOptions = {}) => {
  if (!text) return [];
  const parts = text.split('$');
  return parts.map((part, index) => {
    if (index % 2 === 1) {
      // This is math
      const plainMath = latexToPlainText(part);
      return new docx.TextRun({
        text: plainMath,
        italics: true,
        font: 'Cambria Math',
        size: defaultOptions.size || 22,
        ...defaultOptions
      });
    } else {
      // This is plain text
      return new docx.TextRun({
        text: part,
        size: defaultOptions.size || 22,
        ...defaultOptions
      });
    }
  });
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
  logo: '',
  logoWidth: 100,
  logoHeight: 100,
  logoPos: { x: 0, y: 0 },
  schoolName: 'Girijyothi CMI Public School',
  schoolAddress: 'Vazhathope, Idukki',
  fontFamily: 'Cinzel',
  headerLogoOnly: false,
  hideSchoolLogo: false
};

const DEFAULT_METADATA = {
  title: 'FIRST TERM SUMMATIVE ASSESSMENT',
  subject: 'Computer Science & Programming',
  classDiv: 'Class X - Division A & B',
  maxMarks: 50,
  duration: '90 Minutes',
  separateAnswerSheet: false,
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
        text: 'Discuss the security implications of cloud computing and explain the key differences between Public, Private, and Hybrid Cloud architectures.',
        marks: 10,
        blankLines: 12
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
  const [sections, setSections] = useState(DEFAULT_SECTIONS);
  const [activeTab, setActiveTab] = useState('branding'); // branding, metadata, sections
  const [collapsedSections, setCollapsedSections] = useState({});
  const [isPreviewOpen, setIsPreviewOpen] = useState(false);
  const [isDownloadOpen, setIsDownloadOpen] = useState(false);
  const [isDocsModalOpen, setIsDocsModalOpen] = useState(false);
  const [isDocsUploading, setIsDocsUploading] = useState(false);
  const [docsError, setDocsError] = useState('');
  const [formulaModal, setFormulaModal] = useState({ isOpen: false, latex: '', onSave: null });
  const [activeInputInfo, setActiveInputInfo] = useState(null);
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
        if (parsed.branding) setBranding(parsed.branding);
        if (parsed.metadata) setMetadata(parsed.metadata);
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

  // Handle Logo Upload
  const handleLogoUpload = (e) => {
    const file = e.target.files[0];
    if (file) {
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
    setBranding(prev => ({ ...prev, logo: '' }));
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

  // Helper calculation functions
  const getSectionTotalMarks = (section) => {
    return section.questions.reduce((total, q) => total + (Number(q.marks) || 0), 0);
  };

  const getExamCurrentTotalMarks = () => {
    return sections.reduce((total, s) => total + getSectionTotalMarks(s), 0);
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
            updatedQ.options = ['Option A', 'Option B', 'Option C', 'Option D'];
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
          if (newType === 'image') {
            if (updatedQ.image === undefined) updatedQ.image = '';
            if (updatedQ.imageWidth === undefined) updatedQ.imageWidth = 300;
            if (updatedQ.imageHeight === undefined) updatedQ.imageHeight = 200;
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
          text: 'New Question details here...',
          marks: 1
        };

        if (type === 'mcq') {
          defaultQuestion.options = ['Option A', 'Option B', 'Option C', 'Option D'];
        } else if (type === 'essay') {
          defaultQuestion.blankLines = 5;
        } else if (type === 'match_following') {
          defaultQuestion.matchPairs = [
            { premise: 'Item A', response: 'Match A' },
            { premise: 'Item B', response: 'Match B' }
          ];
          defaultQuestion.shuffleB = true;
        } else if (type === 'image') {
          defaultQuestion.image = '';
          defaultQuestion.imageWidth = 300;
          defaultQuestion.imageHeight = 200;
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
    if (window.confirm('Are you sure you want to clear the entire draft?')) {
      setBranding({
        logo: '',
        logoWidth: 100,
        logoHeight: 100,
        logoPos: { x: 0, y: 0 },
        schoolName: '',
        schoolAddress: '',
        fontFamily: 'Inter',
        headerLogoOnly: false,
        hideSchoolLogo: false
      });
      setMetadata({
        title: '',
        subject: '',
        classDiv: '',
        maxMarks: 100,
        duration: '',
        separateAnswerSheet: false,
        language: 'english'
      });
      setSections([]);
      localStorage.removeItem('question_ninja_draft');
    }
  };

  // CSV Export & Import Features
  const exportToCSV = () => {
    const headers = [
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
      'Image Height'
    ];

    const rows = [];
    sections.forEach((sec) => {
      if (sec.questions.length === 0) {
        rows.push([
          sec.title,
          sec.marks,
          sec.instructions,
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
            matchPairsStr = q.matchPairs.map(p => `${p.premise}=${p.response}`).join(';');
          }

          rows.push([
            sec.title,
            sec.marks,
            sec.instructions,
            sec.type || 'essay',
            q.text,
            q.marks,
            optionsStr,
            q.blankLines || '',
            matchPairsStr,
            q.image || '',
            q.imageWidth || '',
            q.imageHeight || ''
          ]);
        });
      }
    });

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
    link.setAttribute('download', `${metadata.title || 'question_paper'}.csv`);
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

        const importedSections = [];
        let currentSection = null;

        for (let i = 1; i < parsedRows.length; i++) {
          const row = parsedRows[i];
          if (row.length < 3) continue;

          const secTitle = row[0];
          const secMarks = Number(row[1]) || 0;
          const secInstructions = row[2];
          const qType = row[3];
          const qText = row[4];
          const qMarks = Number(row[5]) || 0;
          const optionsStr = row[6];
          const blankLinesVal = row[7];
          const matchPairsStr = row[8];
          const imageData = row[9];
          const imageWidthVal = row[10];
          const imageHeightVal = row[11];

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

            if (qType === 'mcq') {
              q.options = optionsStr ? optionsStr.split(';') : ['', '', '', ''];
            } else if (qType === 'essay') {
              q.blankLines = Number(blankLinesVal) || 5;
            } else if (qType === 'match_following') {
              q.matchPairs = matchPairsStr
                ? matchPairsStr.split(';').map(pair => {
                  const parts = pair.split('=');
                  return { premise: parts[0] || '', response: parts[1] || '' };
                })
                : [];
            } else if (qType === 'image') {
              q.image = imageData || '';
              q.imageWidth = Number(imageWidthVal) || 300;
              q.imageHeight = Number(imageHeightVal) || 200;
            }

            currentSection.questions.push(q);
          }
        }

        if (importedSections.length > 0) {
          if (window.confirm(`Successfully parsed ${importedSections.length} sections. Import and replace current layout?`)) {
            setSections(importedSections);
          }
        } else {
          alert('No valid sections or questions found in the CSV.');
        }
      } catch (err) {
        console.error(err);
        alert('Error parsing CSV file. Please make sure the format is correct.');
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  // Print PDF Trigger
  const triggerPrint = () => {
    window.print();
  };

  const triggerPdfExport = () => {
    alert("To save as a PDF file, please select 'Save as PDF' under the 'Destination' selection in the browser print window.");
    window.print();
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
    if (src.startsWith('data:')) {
      return dataURLToUint8Array(src);
    }
    try {
      const response = await fetch(src);
      const blob = await response.blob();
      const arrayBuffer = await blob.arrayBuffer();
      return new Uint8Array(arrayBuffer);
    } catch (e) {
      console.error("Failed to fetch image from URL for docx export", e);
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
            spacing: { after: 200 }
          })
        );
      } else {
        // Construct side-by-side layout using borderless table
        const headerTable = new docx.Table({
          width: { size: 100, type: docx.WidthType.PERCENTAGE },
          borders: {
            top: { style: docx.BorderStyle.NONE },
            bottom: { style: docx.BorderStyle.NONE },
            left: { style: docx.BorderStyle.NONE },
            right: { style: docx.BorderStyle.NONE },
            insideHorizontal: { style: docx.BorderStyle.NONE },
            insideVertical: { style: docx.BorderStyle.NONE }
          },
          rows: [
            new docx.TableRow({
              children: [
                new docx.TableCell({
                  width: { size: 20, type: docx.WidthType.PERCENTAGE },
                  borders: {
                    top: { style: docx.BorderStyle.NONE },
                    bottom: { style: docx.BorderStyle.NONE },
                    left: { style: docx.BorderStyle.NONE },
                    right: { style: docx.BorderStyle.NONE }
                  },
                  children: [
                    new docx.Paragraph({
                      alignment: docx.AlignmentType.CENTER,
                      children: [logoRun]
                    })
                  ]
                }),
                new docx.TableCell({
                  width: { size: 80, type: docx.WidthType.PERCENTAGE },
                  borders: {
                    top: { style: docx.BorderStyle.NONE },
                    bottom: { style: docx.BorderStyle.NONE },
                    left: { style: docx.BorderStyle.NONE },
                    right: { style: docx.BorderStyle.NONE }
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

    // Divider line
    headerChildren.push(
      new docx.Paragraph({
        alignment: docx.AlignmentType.CENTER,
        children: [
          new docx.TextRun({
            text: '_________________________________________________________________________________',
            bold: true,
            color: '000000'
          })
        ],
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
    headerChildren.push(createMetaParagraph('Class & Div: ', metadata.classDiv));
    headerChildren.push(createMetaParagraph('Max Marks: ', String(metadata.maxMarks || 0)));
    headerChildren.push(createMetaParagraph('Time Allowed: ', metadata.duration));

    // Bottom border for metadata
    headerChildren.push(
      new docx.Paragraph({
        alignment: docx.AlignmentType.CENTER,
        children: [
          new docx.TextRun({
            text: '_________________________________________________________________________________',
            bold: true,
            color: '000000'
          })
        ],
        spacing: { after: 300 }
      })
    );

    // Now populate sections and questions
    let absoluteQuestionCount = 1;

    sections.forEach((sec) => {
      // Section header
      headerChildren.push(
        new docx.Paragraph({
          spacing: { before: 240, after: 80 },
          children: [
            new docx.TextRun({
              text: (sec.title || '').toUpperCase(),
              bold: true,
              size: 24,
              font: getFontFamily()
            }),
            new docx.TextRun({
              text: `\t(Total: ${sec.marks || 0} Marks)`,
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
            spacing: { after: 180 },
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
      sec.questions.forEach((q) => {
        const qNum = `${absoluteQuestionCount}.`;
        absoluteQuestionCount++;

        // Add question text
        headerChildren.push(
          new docx.Paragraph({
            spacing: { before: 120, after: 80 },
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
              ...docxTextRunsWithMath(q.text || ''),
              new docx.TextRun({
                text: `\t[${q.marks || 0} Marks]`,
                bold: true,
                size: 20
              })
            ]
          })
        );

        // Formatting specific question types
        if (sec.type === 'mcq' && q.options) {
          // Render MCQ choices in a 2-column table (two options per row)
          const optRows = [];
          for (let i = 0; i < q.options.length; i += 2) {
            const leftLetter = String.fromCharCode(65 + i);
            const leftText = q.options[i] || '';
            const rightLetter = i + 1 < q.options.length ? String.fromCharCode(65 + i + 1) : '';
            const rightText = i + 1 < q.options.length ? (q.options[i + 1] || '') : '';

            const cells = [
              new docx.TableCell({
                width: { size: 50, type: docx.WidthType.PERCENTAGE },
                borders: {
                  top: { style: docx.BorderStyle.NONE, size: 0 },
                  bottom: { style: docx.BorderStyle.NONE, size: 0 },
                  left: { style: docx.BorderStyle.NONE, size: 0 },
                  right: { style: docx.BorderStyle.NONE, size: 0 }
                },
                children: [
                  new docx.Paragraph({
                    indent: { left: 360 },
                    spacing: { after: 40 },
                    children: [
                      new docx.TextRun({ text: `(${leftLetter})  `, size: 22 }),
                      ...docxTextRunsWithMath(leftText)
                    ]
                  })
                ]
              })
            ];

            if (rightText || rightLetter) {
              cells.push(
                new docx.TableCell({
                  width: { size: 50, type: docx.WidthType.PERCENTAGE },
                  borders: {
                    top: { style: docx.BorderStyle.NONE, size: 0 },
                    bottom: { style: docx.BorderStyle.NONE, size: 0 },
                    left: { style: docx.BorderStyle.NONE, size: 0 },
                    right: { style: docx.BorderStyle.NONE, size: 0 }
                  },
                  children: [
                    new docx.Paragraph({
                      spacing: { after: 40 },
                      children: [
                        new docx.TextRun({ text: `(${rightLetter})  `, size: 22 }),
                        ...docxTextRunsWithMath(rightText)
                      ]
                    })
                  ]
                })
              );
            } else {
              cells.push(
                new docx.TableCell({
                  width: { size: 50, type: docx.WidthType.PERCENTAGE },
                  borders: {
                    top: { style: docx.BorderStyle.NONE, size: 0 },
                    bottom: { style: docx.BorderStyle.NONE, size: 0 },
                    left: { style: docx.BorderStyle.NONE, size: 0 },
                    right: { style: docx.BorderStyle.NONE, size: 0 }
                  },
                  children: [new docx.Paragraph({ children: [] })]
                })
              );
            }

            optRows.push(new docx.TableRow({ children: cells }));
          }

          headerChildren.push(
            new docx.Table({
              width: { size: 100, type: docx.WidthType.PERCENTAGE },
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

        else if (sec.type === 'essay') {
          if (!metadata.separateAnswerSheet) {
            // Renders specified blank lines
            const linesCount = q.blankLines || 5;
            for (let i = 0; i < linesCount; i++) {
              headerChildren.push(
                new docx.Paragraph({
                  spacing: { after: 120 },
                  children: [
                    new docx.TextRun({
                      text: '____________________________________________________________________________',
                      color: 'E0E0E0'
                    })
                  ]
                })
              );
            }
          }
        }

        else if (sec.type === 'true_false') {
          headerChildren.push(
            new docx.Paragraph({
              indent: { left: 720 },
              spacing: { after: 60 },
              children: [
                new docx.TextRun({
                  text: metadata.separateAnswerSheet ? '(True / False)' : '[    ] True        [    ] False',
                  bold: true,
                  size: 20
                })
              ]
            })
          );
        }

        else if (sec.type === 'match_following' && q.matchPairs) {
          // Build match list
          const columnA = q.matchPairs.map(p => p.premise);
          let columnB = q.matchPairs.map(p => p.response);
          if (q.shuffleB) {
            columnB = [...columnB].sort(() => Math.random() - 0.5);
          }

          // Renders a simple side-by-side matches listing
          for (let index = 0; index < columnA.length; index++) {
            const romanNum = (idx) => {
              const r = ['i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x'];
              return r[idx] || (idx + 1).toString();
            };

            headerChildren.push(
              new docx.Paragraph({
                indent: { left: 720 },
                spacing: { after: 60 },
                children: [
                  new docx.TextRun({
                    text: `${index + 1}. `,
                    size: 22
                  }),
                  ...docxTextRunsWithMath(columnA[index] || ''),
                  new docx.TextRun({
                    text: `\t\t\t\t\t\t${romanNum(index)}. `,
                    size: 22
                  }),
                  ...docxTextRunsWithMath(columnB[index] || '')
                ]
              })
            );
          }
        }

        else if (sec.type === 'image' && q.image) {
          const imageBytes = dataURLToUint8Array(q.image);
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
              width: { size: 100, type: docx.WidthType.PERCENTAGE },
              indent: { size: 360, type: docx.WidthType.DXA },
              rows: tblRows
            })
          );
        }
      });
    });

    const doc = new docx.Document({
      features: {
        updateFields: true
      },
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
                  new docx.TextRun({
                    children: [docx.PageNumber.CURRENT],
                    size: 20
                  }),
                  new docx.TextRun({
                    text: ' of ',
                    size: 20
                  }),
                  new docx.TextRun({
                    children: [docx.PageNumber.TOTAL_PAGES],
                    size: 20
                  })
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
    try {
      const blob = await generateDocxBlob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const safeName = (metadata.title || 'QuestionPaper').replace(/[^a-z0-9]/gi, '_').toLowerCase();
      a.download = `${safeName}.docx`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Error exporting DOCX:', err);
      alert('Failed to generate DOCX file.');
    }
  };
  const handleGoogleDocsWebPreview = async () => {
    setIsDocsUploading(true);
    setDocsError('');
    try {
      const blob = await generateDocxBlob();
      const safeName = (metadata.title || 'QuestionPaper').replace(/[^a-z0-9]/gi, '_').toLowerCase();
      
      const formData = new FormData();
      formData.append('file', blob, `${safeName}.docx`);

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
      const list = q.matchPairs.map(p => p.response);
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
              <BookOpen size={24} className="text-accent" />
              <span>QuestionNinja</span>
            </h1>
            <p>Question Paper Designer for Schools</p>
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
                    <button className="btn btn-sm btn-danger" onClick={removeLogo} style={{ marginTop: '8px' }}>
                      Remove Logo
                    </button>
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
                    value={metadata.maxMarks}
                    onChange={(e) => setMetadata({ ...metadata, maxMarks: Number(e.target.value) })}
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
                    <span>Marks Mismatch: Current Questions = {getExamCurrentTotalMarks()} marks (Target = {metadata.maxMarks} marks).</span>
                  </div>
                ) : (
                  <div className="warning-badge" style={{ backgroundColor: 'rgba(16,185,129,0.15)', color: 'var(--success)', borderColor: 'rgba(16,185,129,0.3)' }}>
                    <CheckCircle size={14} />
                    <span>Perfect: Sum of all questions matches targets!</span>
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
                              value={sec.marks}
                              onChange={(e) => updateSectionMeta(sec.id, 'marks', Number(e.target.value))}
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
                              {secTotal} / {sec.marks} Marks
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
                            <option value="image">Image Question</option>
                            <option value="table">Table Question</option>
                          </select>
                        </div>

                        {isOverProvisioned && (
                          <div className="warning-badge" style={{ fontSize: '11px' }}>
                            <AlertTriangle size={12} />
                            <span>Choice Provisioning: Question total ({secTotal}) exceeds Declared section marks ({sec.marks}). Permitted for optional choices.</span>
                          </div>
                        )}

                        {/* Questions Area */}
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', marginTop: '16px', borderTop: '1px solid var(--border-color)', paddingTop: '16px' }}>
                          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                            <h4 style={{ fontSize: '13px', fontWeight: 'bold', textTransform: 'uppercase' }}>Questions ({sec.questions.length})</h4>
                          </div>

                          {sec.questions.map((q, qIdx) => (
                            <div key={q.id} style={{ padding: '12px', backgroundColor: 'var(--bg-editor)', borderRadius: 'var(--radius-sm)', display: 'flex', flexDirection: 'column', gap: '10px', border: '1px solid var(--border-color)' }}>
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
                                  style={{ minHeight: '60px', fontSize: '13px' }}
                                  onChange={(e) => updateQuestion(sec.id, q.id, { text: e.target.value })}
                                />
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
                                <input
                                  type="number"
                                  value={q.marks}
                                  style={{ padding: '6px 10px', fontSize: '13px' }}
                                  onChange={(e) => updateQuestion(sec.id, q.id, { marks: Number(e.target.value) })}
                                />
                              </div>

                              {/* MCQ Specific Fields */}
                              {sec.type === 'mcq' && q.options && (
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                  <label style={{ fontSize: '10px', fontWeight: 'bold' }}>Options</label>
                                  {q.options.map((opt, oIdx) => (
                                    <div key={oIdx} style={{ display: 'flex', gap: '6px', width: '100%' }}>
                                      <span style={{ fontSize: '13px', alignSelf: 'center' }}>{String.fromCharCode(65 + oIdx)}.</span>
                                      <input
                                        id={`q__opt__${sec.id}__${q.id}__${oIdx}`}
                                        type="text"
                                        value={opt}
                                        style={{ padding: '4px 8px', fontSize: '12px', flex: 1 }}
                                        onChange={(e) => {
                                          const newOpts = [...q.options];
                                          newOpts[oIdx] = e.target.value;
                                          updateQuestion(sec.id, q.id, { options: newOpts });
                                        }}
                                      />
                                    </div>
                                  ))}
                                </div>
                              )}

                              {/* Essay Specific Fields */}
                              {sec.type === 'essay' && (
                                <div className="form-group">
                                  <label style={{ fontSize: '10px' }}>Blank lines for printing</label>
                                  <input
                                    type="number"
                                    value={q.blankLines}
                                    style={{ padding: '6px 10px', fontSize: '13px' }}
                                    onChange={(e) => updateQuestion(sec.id, q.id, { blankLines: Number(e.target.value) })}
                                  />
                                </div>
                              )}

                              {/* Match the Following Specific Fields */}
                              {sec.type === 'match_following' && q.matchPairs && (
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                    <label style={{ fontSize: '10px', fontWeight: 'bold' }}>Match Pairs</label>
                                    <button
                                      className="btn btn-secondary btn-sm"
                                      style={{ padding: '2px 6px', fontSize: '10px' }}
                                      onClick={() => {
                                        updateQuestion(sec.id, q.id, {
                                          matchPairs: [...q.matchPairs, { premise: 'New Item', response: 'New Match' }]
                                        });
                                      }}
                                    >
                                      + Pair
                                    </button>
                                  </div>
                                  {q.matchPairs.map((pair, pIdx) => (
                                    <div key={pIdx} style={{ display: 'grid', gridTemplateColumns: '1fr 1fr auto', gap: '6px', alignItems: 'center' }}>
                                      <input
                                        id={`q__matcha__${sec.id}__${q.id}__${pIdx}`}
                                        type="text"
                                        placeholder="Premise (Col A)"
                                        value={pair.premise}
                                        style={{ padding: '4px 8px', fontSize: '12px' }}
                                        onChange={(e) => {
                                          const newPairs = [...q.matchPairs];
                                          newPairs[pIdx].premise = e.target.value;
                                          updateQuestion(sec.id, q.id, { matchPairs: newPairs });
                                        }}
                                      />
                                      <input
                                        id={`q__matchb__${sec.id}__${q.id}__${pIdx}`}
                                        type="text"
                                        placeholder="Response (Col B)"
                                        value={pair.response}
                                        style={{ padding: '4px 8px', fontSize: '12px' }}
                                        onChange={(e) => {
                                          const newPairs = [...q.matchPairs];
                                          newPairs[pIdx].response = e.target.value;
                                          updateQuestion(sec.id, q.id, { matchPairs: newPairs });
                                        }}
                                      />
                                      <button
                                        className="btn btn-danger btn-sm"
                                        style={{ padding: '4px' }}
                                        onClick={() => {
                                          updateQuestion(sec.id, q.id, {
                                            matchPairs: q.matchPairs.filter((_, idx) => idx !== pIdx)
                                          });
                                        }}
                                      >
                                        <Trash2 size={12} />
                                      </button>
                                    </div>
                                  ))}

                                  <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                                    <input
                                      type="checkbox"
                                      id={`shuffle-${q.id}`}
                                      checked={q.shuffleB}
                                      onChange={(e) => updateQuestion(sec.id, q.id, { shuffleB: e.target.checked })}
                                    />
                                    <label htmlFor={`shuffle-${q.id}`} style={{ fontSize: '11px', textTransform: 'none' }}>
                                      Shuffle Column B in preview/exports
                                    </label>
                                  </div>
                                </div>
                              )}

                              {/* Image Question Specific Fields */}
                              {sec.type === 'image' && (
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                                  <label style={{ fontSize: '10px', fontWeight: 'bold' }}>Question Image</label>
                                  {q.image ? (
                                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                      <div className="editor-image-preview-container">
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
                                    <label className="image-upload-dashed-zone">
                                      <ImageIcon size={20} className="text-secondary" />
                                      <span>Upload Question Image</span>
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

                          <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: '8px' }}>
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
                    <span>Marks Mismatch: Current Questions = {getExamCurrentTotalMarks()} marks (Target = {metadata.maxMarks} marks).</span>
                  </div>
                ) : (
                  <div className="warning-badge" style={{ display: 'flex', alignSelf: 'flex-start', backgroundColor: 'rgba(16,185,129,0.15)', color: 'var(--success)', borderColor: 'rgba(16,185,129,0.3)' }}>
                    <CheckCircle size={14} />
                    <span>Perfect: Sum of all questions matches targets!</span>
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
                  checked={metadata.separateAnswerSheet || false}
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
                  <span>Marks Mismatch: Current Questions = {getExamCurrentTotalMarks()} marks (Target = {metadata.maxMarks} marks).</span>
                </div>
              </div>
            )}

            <div className="modal-body">
              {/* Dynamic A4 Preview Sheet */}
              <div className={`paper-sheet lang-${metadata.language || 'english'}`}>

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
                      <img src={branding.logo} className="brand-logo-img" alt="School Logo" />
                      {!branding.headerLogoOnly && <div className="resize-handle"></div>}
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
                    <span>{metadata.maxMarks}</span>
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
                          <h2 className="paper-section-title">{sec.title}</h2>
                          <span className="paper-section-marks">[{sec.marks} Marks]</span>
                        </div>
                        {sec.instructions && (
                          <p className="paper-section-instructions">{sec.instructions}</p>
                        )}

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
                                  <p style={{ fontWeight: '500' }} dangerouslySetInnerHTML={{ __html: renderTextWithMath(q.text) }} />

                                  {/* MCQ Options */}
                                  {sec.type === 'mcq' && q.options && (
                                    <div className="paper-mcq-options">
                                      {q.options.map((opt, oIdx) => (
                                        <div key={oIdx} className="paper-mcq-option">
                                          <span style={{ fontWeight: '600' }}>({String.fromCharCode(65 + oIdx)})</span>
                                          <span dangerouslySetInnerHTML={{ __html: renderTextWithMath(opt) }} />
                                        </div>
                                      ))}
                                    </div>
                                  )}

                                  {/* Essay spaces */}
                                  {sec.type === 'essay' && !metadata.separateAnswerSheet && (
                                    <div style={{ display: 'flex', flexDirection: 'column', gap: '14px', marginTop: '8px' }}>
                                      {Array.from({ length: q.blankLines || 5 }).map((_, lineIdx) => (
                                        <div key={lineIdx} style={{ borderBottom: '1px dotted #ccc', height: '14px' }}></div>
                                      ))}
                                    </div>
                                  )}

                                  {/* True/False selection */}
                                  {sec.type === 'true_false' && (
                                    <div className="paper-tf-options">
                                      {metadata.separateAnswerSheet ? (
                                        <span>(True / False)</span>
                                      ) : (
                                        <>
                                          <span>[   ] True</span>
                                          <span>[   ] False</span>
                                        </>
                                      )}
                                    </div>
                                  )}

                                  {/* Match the Following columns */}
                                  {sec.type === 'match_following' && q.matchPairs && (
                                    <table className="paper-match-table">
                                      <tbody>
                                        {q.matchPairs.map((pair, pIdx) => {
                                          const shuffledList = getShuffledList(q);
                                          const roman = (idx) => {
                                            const r = ['i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x'];
                                            return r[idx] || (idx + 1).toString();
                                          };
                                          return (
                                            <tr key={pIdx}>
                                              <td style={{ padding: '4px 0' }}>
                                                {pIdx + 1}. <span dangerouslySetInnerHTML={{ __html: renderTextWithMath(pair.premise) }} />
                                              </td>
                                              <td style={{ padding: '4px 0', paddingLeft: '20px' }}>
                                                {roman(pIdx)}. <span dangerouslySetInnerHTML={{ __html: renderTextWithMath(shuffledList[pIdx] || pair.response) }} />
                                              </td>
                                            </tr>
                                          );
                                        })}
                                      </tbody>
                                    </table>
                                  )}

                                  {/* Image question render */}
                                  {sec.type === 'image' && q.image && (
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
                                <span className="paper-question-marks">({q.marks} M)</span>
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
              <button className="btn btn-primary" onClick={triggerPrint} style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <Printer size={16} /> Print
              </button>

              <div className="dropdown-container" ref={dropdownRef}>
                <button
                  className="btn btn-secondary"
                  onClick={() => setIsDownloadOpen(!isDownloadOpen)}
                  style={{ display: 'flex', alignItems: 'center', gap: '6px' }}
                >
                  <Download size={16} /> Download <ChevronDown size={14} />
                </button>
                {isDownloadOpen && (
                  <div className="dropdown-menu">
                    <button className="dropdown-item" onClick={() => { setIsDownloadOpen(false); triggerPdfExport(); }} style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                      <FileText size={16} className="text-accent" /> Download PDF
                    </button>
                    <button className="dropdown-item" onClick={() => { setIsDownloadOpen(false); triggerDocxExport(); }} style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                      <FileText size={16} style={{ color: '#2b579a' }} /> Download Word (DOCX)
                    </button>
                  </div>
                )}
              </div>

              <button
                className="btn btn-secondary"
                onClick={() => setIsDocsModalOpen(true)}
                style={{ display: 'flex', alignItems: 'center', gap: '6px', borderColor: 'rgba(66, 133, 244, 0.4)' }}
              >
                <svg viewBox="0 0 24 24" width="16" height="16" style={{ fill: '#4285f4' }}>
                  <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
                </svg>
                <span>Open in Google Docs</span>
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
      {activeInputInfo && (
        <button
          type="button"
          className="floating-formula-btn"
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
    </div>
  );
}



//hello this is new update from naveen, please check it.