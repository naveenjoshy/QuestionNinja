import React, { useState, useEffect, useRef, useCallback, useMemo } from 'react';
import pageLogo from './assets/logo.png';
import schoolLogo from './assets/school_logo.png';
import developerPhoto from './assets/developer_photo.jpeg';
import qrCode from './assets/qrcode.jpeg';
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
  Eye,
  ArrowLeft,
  ArrowRight,
  Loader2,
  Shield,
  Upload,
  BookOpen,
  GraduationCap,
  School,
  Check,
  X,
  Printer
} from 'lucide-react';
import {
  FORMULA_BUTTONS,
  DEFAULT_BRANDING,
  DEFAULT_METADATA,
  DEFAULT_SECTIONS,
  _PRESET_TEMPLATES
} from './config/constants';
import {
  renderTextWithMath
} from './utils/mathUtils';
import {
  normalizeLogo,
  formatMarks,
  getQuestionMarks,
  getSectionTotalMarks,
  getExamCurrentTotalMarks,
  getExamTotalQuestionsCount,
  hasBlankQuestions,
  hasZeroMarkQuestions,
  getOptionText,
  getOptionImage,
  getOptionImageWidth,
  getOptionImageHeight,
  canFitSingleLine
} from './utils/formatUtils';
import { exportToCSV, generateExportFilename, importCsvData } from './utils/csvUtils';

export default function App() {
  // App states
  const [branding, setBranding] = useState(DEFAULT_BRANDING);
  const [metadata, setMetadata] = useState(DEFAULT_METADATA);
  const [sections, setSections] = useState([]);
  const [activeTab, setActiveTab] = useState('branding'); // branding, metadata, sections
  const [collapsedSections, setCollapsedSections] = useState({});
  const [isPreviewOpen, setIsPreviewOpen] = useState(false);
  const [isDocsModalOpen, setIsDocsModalOpen] = useState(false);
  const [formulaModal, setFormulaModal] = useState({ isOpen: false, latex: '', onSave: null });
  const [isAboutModalOpen, setIsAboutModalOpen] = useState(false);
  const [activeInputInfo, setActiveInputInfo] = useState(null);
  const [csvImportModal, setCsvImportModal] = useState({ isOpen: false, branding: null, metadata: null, sections: [], importSchool: false, importExam: true, importQuestions: true });
  const [isPdfExporting, setIsPdfExporting] = useState(false);
  const [previewZoom, setPreviewZoom] = useState(0.85);
  const [showLivePreview, setShowLivePreview] = useState(true);
  const paperSheetRef = useRef(null);
  

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
  };

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
    try {
      const savedDraft = localStorage.getItem('question_ninja_draft');
      if (savedDraft) {
        const parsed = JSON.parse(savedDraft);
        if (parsed.branding) {
          setBranding(() => ({
            ...DEFAULT_BRANDING,
            ...parsed.branding,
            logo: parsed.branding.logo === null ? null : (normalizeLogo(parsed.branding.logo) || schoolLogo)
          }));
        }
        if (parsed.metadata) setMetadata(() => ({ ...DEFAULT_METADATA, ...parsed.metadata }));
        if (parsed.sections) setSections(parsed.sections);
      }
    } catch (e) {
      console.error('Error loading saved draft from localStorage', e);
    }
  }, []);

  // Save to local storage on change
  useEffect(() => {
    try {
      const stateToSave = { branding, metadata, sections };
      localStorage.setItem('question_ninja_draft', JSON.stringify(stateToSave));
    } catch (err) {
      console.warn('Unable to auto-save draft to localStorage (quota exceeded or storage restricted):', err);
    }
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








  const hasQuestions = useCallback(() => {
    return sections.some(s => s.questions && s.questions.length > 0);
  }, [sections]);

  const examCurrentTotalMarks = useMemo(() => getExamCurrentTotalMarks(sections), [sections]);
  const examTotalQuestionsCount = useMemo(() => getExamTotalQuestionsCount(sections), [sections]);
  const hasBlank = useMemo(() => hasBlankQuestions(sections), [sections]);
  const hasZeroMark = useMemo(() => hasZeroMarkQuestions(sections), [sections]);





  const updateOptionText = (secId, qId, oIdx, text) => {
    setSections(prev => prev.map(sec => {
      if (sec.id !== secId) return sec;
      return {
        ...sec,
        questions: sec.questions.map(q => {
          if (q.id !== qId) return q;
          const newOpts = [...(q.options || [])];
          const cur = newOpts[oIdx];
          if (typeof cur === 'object' && cur !== null) {
            newOpts[oIdx] = { ...cur, text };
          } else {
            newOpts[oIdx] = text;
          }
          return { ...q, options: newOpts };
        })
      };
    }));
  };

  const updateOptionImage = (secId, qId, oIdx, imageProps) => {
    setSections(prev => prev.map(sec => {
      if (sec.id !== secId) return sec;
      return {
        ...sec,
        questions: sec.questions.map(q => {
          if (q.id !== qId) return q;
          const newOpts = [...(q.options || [])];
          const cur = newOpts[oIdx];
          const text = typeof cur === 'object' && cur !== null ? cur.text : (cur || '');
          const existingObj = typeof cur === 'object' && cur !== null ? cur : { text };
          newOpts[oIdx] = { ...existingObj, ...imageProps };
          return { ...q, options: newOpts };
        })
      };
    }));
  };



  // State update helpers for Sections & Questions
  const addSection = () => {
    const newSection = {
      id: `sec-${Date.now()}`,
      title: `SECTION ${String.fromCharCode(65 + sections.length)}: NEW SECTION`,
      marks: 10,
      instructions: 'Answer all questions. Each question carries equal marks.',
      type: 'essay',
      pageBreakBefore: false,
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
        const finalValue = field === 'title' ? String(value || '').toUpperCase() : value;
        return { ...s, [field]: finalValue };
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
          marks: 1,
          pageBreakBefore: false
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

  const handlePasteImage = (e, secId, qId, sqId = null, oIdx = null) => {
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
          if (file.size > 3 * 1024 * 1024) {
            alert('Pasted image exceeds the 3MB size limit. Please choose a smaller image.');
            return;
          }
          const reader = new FileReader();
          reader.onload = (uploadEvent) => {
            if (oIdx !== null && oIdx !== undefined) {
              updateOptionImage(secId, qId, oIdx, { image: uploadEvent.target.result });
            } else if (sqId) {
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

  const handleDropImage = (e, secId, qId, sqId = null, oIdx = null) => {
    e.preventDefault();
    e.currentTarget.classList.remove('drag-over');
    const files = e.dataTransfer?.files;
    if (files && files.length > 0) {
      const file = files[0];
      if (file.type.startsWith('image/')) {
        if (file.size > 3 * 1024 * 1024) {
          alert('Dropped image exceeds the 3MB size limit. Please choose a smaller image.');
          return;
        }
        const reader = new FileReader();
        reader.onload = (uploadEvent) => {
          if (oIdx !== null && oIdx !== undefined) {
            updateOptionImage(secId, qId, oIdx, { image: uploadEvent.target.result });
          } else if (sqId) {
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

  // Direct printing handler with validation and automatic title branding
  const handleDirectPrint = useCallback(() => {
    if (!hasQuestions()) {
      alert("No questions added to print. Please add questions first.");
      return;
    }
    if (hasBlank) {
      if (!window.confirm("Some questions have empty text. Are you sure you want to print?")) {
        return;
      }
    }

    const doPrint = () => {
      const originalTitle = document.title;
      try {
        const examName = generateExportFilename('', metadata).replace(/\.$/, '');
        if (examName) {
          document.title = examName;
        }
      } catch (err) {
        console.warn("Could not set document title for printing:", err);
      }

      window.print();

      setTimeout(() => {
        document.title = originalTitle;
      }, 1000);
    };

    if (!isPreviewOpen) {
      setIsPreviewOpen(true);
      setTimeout(doPrint, 250);
    } else {
      doPrint();
    }
  }, [hasQuestions, hasBlank, isPreviewOpen, metadata]);

  // Keyboard shortcuts: Ctrl+P / Cmd+P for direct printing, Escape to close preview modal
  useEffect(() => {
    const handleKeyDown = (e) => {
      if ((e.metaKey || e.ctrlKey) && (e.key === 'p' || e.key === 'P')) {
        e.preventDefault();
        handleDirectPrint();
      } else if (e.key === 'Escape' && isPreviewOpen) {
        setIsPreviewOpen(false);
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [handleDirectPrint, isPreviewOpen]);

  // CSV Export & Import Features

  const importFromCSV = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const text = event.target.result;
        const { branding, metadata, sections } = importCsvData(text);

        if (!branding && !metadata && (!sections || sections.length === 0)) {
          alert('No valid sections, questions, or branding/exam details found in the CSV.');
          return;
        }

        setCsvImportModal({
          isOpen: true,
          branding: branding || null,
          metadata: metadata || null,
          sections: sections || [],
          importSchool: false,
          importExam: !!metadata,
          importQuestions: !!(sections && sections.length > 0)
        });
      } catch (err) {
        console.error(err);
        alert(err.message || 'Error parsing CSV file. Please make sure the format is correct.');
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
    if (importQuestions && iSections && iSections.length > 0) {
      setSections(iSections);
      setActiveTab('sections');
      setCollapsedSections({});
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

  // Dynamic DOCX Export Handler
  const handleDocxExport = async () => {
    try {
      const { triggerDocxExport } = await import("./utils/exportUtils");
      await triggerDocxExport(branding, metadata, sections, hasQuestions, () => hasBlankQuestions(sections));
    } catch (err) {
      console.error("Error importing or exporting DOCX:", err);
      alert("Failed to initialize Word export.");
    }
  };

  // Dynamic PDF Export Handler
  const handlePdfExport = async () => {
    try {
      const targetElement = paperSheetRef.current || document.querySelector(".exam-paper-container");
      const { triggerPdfExport } = await import("./utils/exportUtils");
      await triggerPdfExport(targetElement, metadata, hasQuestions, () => hasBlankQuestions(sections), setIsPdfExporting);
    } catch (err) {
      console.error("Error importing or exporting PDF:", err);
      alert("Failed to initialize PDF export.");
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
    <div className={`app-container studio-container theme-${theme}`}>
      {/* Studio Top Header Bar */}
      <div className="studio-header print-hide">
        <div style={{ display: 'flex', alignItems: 'center', gap: '14px' }}>
          <div style={{
            width: '40px',
            height: '40px',
            borderRadius: '12px',
            background: 'linear-gradient(135deg, rgba(99, 102, 241, 0.3) 0%, rgba(79, 70, 229, 0.2) 100%)',
            border: '1px solid rgba(99, 102, 241, 0.35)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            boxShadow: '0 4px 14px rgba(99, 102, 241, 0.25)'
          }}>
            <img src={pageLogo} style={{ width: '26px', height: '26px', objectFit: 'contain' }} alt="Logo" />
          </div>
          <div>
            <h1 style={{
              fontSize: '18px',
              fontWeight: 800,
              margin: 0,
              letterSpacing: '-0.3px',
              background: theme === 'dark'
                ? 'linear-gradient(135deg, #ffffff 0%, #a5b4fc 100%)'
                : 'linear-gradient(135deg, #1e1b4b 0%, #4338ca 100%)',
              WebkitBackgroundClip: 'text',
              WebkitTextFillColor: 'transparent'
            }}>
              QuestionNinja Studio
            </h1>
            <p style={{ fontSize: '12px', color: 'var(--text-secondary)', margin: 0, fontWeight: 500 }}>
              Professional Question Paper Generator
            </p>
          </div>
        </div>

        {/* Action Controls & Responsive Toggle */}
        <div style={{ display: 'flex', gap: '10px', alignItems: 'center', marginLeft: 'auto' }}>
          {/* Side-by-side Live Preview Toggle (Desktop Only) */}
          <button
            className={`btn btn-sm side-preview-toggle-btn ${showLivePreview ? 'btn-primary' : 'btn-secondary'}`}
            onClick={() => setShowLivePreview(!showLivePreview)}
            title={showLivePreview ? "Hide side preview to expand editor" : "Show side-by-side live paper preview"}
            style={{ alignItems: 'center', gap: '6px' }}
          >
            <Eye size={14} />
            <span>{showLivePreview ? 'Side Preview: ON' : 'Side Preview: OFF'}</span>
          </button>

          <button
            className="btn btn-secondary btn-sm"
            onClick={toggleTheme}
            title={theme === 'dark' ? "Switch to Light Mode" : "Switch to Dark Mode"}
            aria-label="Toggle Light/Dark Theme"
            style={{ padding: '6px 10px', display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}
          >
            {theme === 'dark' ? <Sun size={16} /> : <Moon size={16} />}
          </button>

          <button className="btn btn-primary btn-sm" onClick={() => setIsPreviewOpen(true)} title="Full Paper Preview Modal">
            <Maximize2 size={15} />
            <span>Preview</span>
          </button>
        </div>
      </div>

      {/* Main Studio Workspace */}
      <div className="studio-workspace">
        {/* Left Column: Question & Exam Builder */}
        <div className={`studio-editor-column ${!showLivePreview ? 'full-width' : ''}`}>
          {/* Marks & Status Validation Meter */}
          <div className="marks-meter-bar">
            <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
              <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-secondary)' }}>
                Marks: <strong style={{ color: 'var(--text-primary)' }}>{metadata.maxMarks || 0} M</strong>
              </span>
              <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-secondary)' }}>
                Sections: <strong style={{ color: 'var(--text-primary)' }}>{sections ? sections.length : 0}</strong>
              </span>
              <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-secondary)' }}>
                Questions: <strong style={{ color: 'var(--text-primary)' }}>{examTotalQuestionsCount}</strong>
              </span>
            </div>
            <div>
              {examCurrentTotalMarks === metadata.maxMarks ? (
                <span className="marks-meter-badge marks-meter-success">
                  <CheckCircle size={14} /> Allocated: {formatMarks(examCurrentTotalMarks)} M (Match!)
                </span>
              ) : (
                <span className="marks-meter-badge marks-meter-warning">
                  <AlertTriangle size={14} /> Allocated: {formatMarks(examCurrentTotalMarks)} / {metadata.maxMarks} M
                </span>
              )}
            </div>
          </div>



          {/* Tab Selection */}
          <div className="nav-tab-container">
            <button
              className={`nav-tab-btn ${activeTab === 'branding' ? 'active' : ''}`}
              onClick={() => setActiveTab('branding')}
            >
              <ImageIcon size={15} />
              <span>School Details</span>
            </button>
            <button
              className={`nav-tab-btn ${activeTab === 'metadata' ? 'active' : ''}`}
              onClick={() => setActiveTab('metadata')}
            >
              <Settings size={15} />
              <span>Exam Details</span>
            </button>
            <button
              className={`nav-tab-btn ${activeTab === 'sections' ? 'active' : ''}`}
              onClick={() => setActiveTab('sections')}
            >
              <Layers size={15} />
              <span>Questions ({examTotalQuestionsCount})</span>
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
                          <span style={{ fontSize: '12px', fontWeight: 600, color: 'var(--text-primary)' }}>Active Logo</span>
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
                  <label>General Instructions</label>
                  <textarea
                    value={metadata.instructions || ''}
                    onChange={(e) => setMetadata({ ...metadata, instructions: e.target.value })}
                    placeholder="e.g. Answer all questions. Each question carries 1 mark."
                  />
                </div>              {/* Validation Badges */}
                <div style={{ marginTop: '12px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
                  {examCurrentTotalMarks !== metadata.maxMarks ? (
                    <div className="warning-badge">
                      <AlertTriangle size={14} />
                      <span>Marks Mismatch: Current Questions = {formatMarks(examCurrentTotalMarks)} marks (Target = {formatMarks(metadata.maxMarks)} marks).</span>
                    </div>
                  ) : (
                    <div className="warning-badge" style={{ backgroundColor: 'rgba(16,185,129,0.15)', color: 'var(--success)', borderColor: 'rgba(16,185,129,0.3)' }}>
                      <CheckCircle size={14} />
                      <span>Perfect: Sum of all questions matches targets!</span>
                    </div>
                  )}
                  {hasZeroMark && (
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
                      <div className="editor-card-header" style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%' }}>
                          <div
                            onClick={() => toggleSectionCollapse(sec.id)}
                            style={{ cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '8px' }}
                          >
                            <Layers size={16} style={{ color: 'var(--accent)' }} />
                            <span style={{ fontSize: '15.5px', fontWeight: '700', color: 'var(--text-primary)' }}>
                              Section {String.fromCharCode(65 + sIdx)}
                            </span>
                            {sec.pageBreakBefore && (
                              <span style={{ fontSize: '10px', fontWeight: '600', backgroundColor: 'rgba(99, 102, 241, 0.15)', color: 'var(--accent)', padding: '2px 6px', borderRadius: '4px', border: '1px solid rgba(99, 102, 241, 0.3)' }}>
                                ✂ Page Break
                              </span>
                            )}
                            {collapsedSections[sec.id] && (
                              <span style={{ fontSize: '11px', color: 'var(--text-muted)', fontWeight: 'normal', marginLeft: '6px' }}>
                                ({sec.questions.length} questions, {secTotal}M)
                              </span>
                            )}
                          </div>

                          <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }}>
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

                        <div
                          onClick={() => toggleSectionCollapse(sec.id)}
                          style={{ cursor: 'pointer', display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: '8px', fontSize: '12px', color: 'var(--text-secondary)' }}
                        >
                          {sec.title && (
                            <span style={{ fontWeight: '600', color: 'var(--text-primary)', letterSpacing: '0.2px' }}>{sec.title}</span>
                          )}
                          <span style={{ color: 'var(--accent)', fontSize: '11px', fontWeight: '500', backgroundColor: 'rgba(99, 102, 241, 0.08)', padding: '2px 8px', borderRadius: '12px', border: '1px solid rgba(99, 102, 241, 0.15)' }}>
                            {sec.type === 'mcq' ? 'Multiple Choice (MCQ)' :
                              sec.type === 'match_following' ? 'Match the Following' :
                                sec.type === 'true_false' ? 'True / False' :
                                  sec.type === 'table' ? 'Table Question' :
                                    'Short Answer / Essay'}
                          </span>
                        </div>
                      </div>

                      {!collapsedSections[sec.id] && (
                        <>
                          <div className="form-group">
                            <label>Section Title</label>
                            <input
                              type="text"
                              value={(sec.title || '').toUpperCase()}
                              onChange={(e) => updateSectionMeta(sec.id, 'title', e.target.value.toUpperCase())}
                              style={{ textTransform: 'uppercase' }}
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
                              placeholder="e.g. Answer any 5 of the following 7 questions&#10;Each question carries equal marks"
                              rows={3}
                              style={{ minHeight: '65px', resize: 'vertical' }}
                            />
                          </div>

                          <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginTop: '4px', marginBottom: '8px' }}>
                            <input
                              type="checkbox"
                              id={`sec-pagebreak-${sec.id}`}
                              checked={!!sec.pageBreakBefore}
                              onChange={(e) => updateSectionMeta(sec.id, 'pageBreakBefore', e.target.checked)}
                              style={{ cursor: 'pointer', margin: 0, width: 'auto' }}
                            />
                            <label
                              htmlFor={`sec-pagebreak-${sec.id}`}
                              style={{ fontSize: '12px', fontWeight: '500', cursor: 'pointer', margin: 0, userSelect: 'none', color: 'var(--text-primary)' }}
                            >
                              Insert Page Break before this section
                            </label>
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
                                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                    <span style={{ fontSize: '12px', fontWeight: 'bold', color: 'var(--accent)' }}>
                                      Q{qIdx + 1} ({(sec.type || 'essay').toUpperCase()})
                                    </span>
                                    {q.pageBreakBefore && (
                                      <span style={{ fontSize: '10px', fontWeight: '600', backgroundColor: 'rgba(99, 102, 241, 0.15)', color: 'var(--accent)', padding: '2px 6px', borderRadius: '4px', border: '1px solid rgba(99, 102, 241, 0.3)' }}>
                                        ✂ Page Break
                                      </span>
                                    )}
                                  </div>
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

                                  <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginTop: '6px' }}>
                                    <input
                                      type="checkbox"
                                      id={`pagebreak-${sec.id}-${q.id}`}
                                      checked={!!q.pageBreakBefore}
                                      onChange={(e) => updateQuestion(sec.id, q.id, { pageBreakBefore: e.target.checked })}
                                      style={{ cursor: 'pointer', margin: 0, width: 'auto' }}
                                    />
                                    <label
                                      htmlFor={`pagebreak-${sec.id}-${q.id}`}
                                      style={{ fontSize: '11px', fontWeight: '500', cursor: 'pointer', margin: 0, userSelect: 'none', color: 'var(--text-primary)' }}
                                    >
                                      Insert Page Break before this question
                                    </label>
                                  </div>

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
                                      <div key={oIdx} style={{ display: 'flex', flexDirection: 'column', gap: '6px', padding: '8px', border: '1px solid var(--border-color)', borderRadius: 'var(--radius-sm)', backgroundColor: 'rgba(0,0,0,0.1)' }}>
                                        <div style={{ display: 'flex', gap: '6px', width: '100%', alignItems: 'flex-start' }}>
                                          <span style={{ fontSize: '13px', fontWeight: 'bold', paddingTop: '6px', minWidth: '18px' }}>
                                            {String.fromCharCode(65 + oIdx)}.
                                          </span>
                                          <textarea
                                            id={`q__opt__${sec.id}__${q.id}__${oIdx}`}
                                            value={getOptionText(opt)}
                                            rows={getOptionText(opt) && getOptionText(opt).includes('\n') ? Math.max(2, getOptionText(opt).split('\n').length) : 1}
                                            placeholder={`Option ${String.fromCharCode(65 + oIdx)} text...`}
                                            className="mcq-option-textarea"
                                            style={{
                                              padding: '6px 8px',
                                              fontSize: '12px',
                                              flex: 1,
                                              minHeight: '34px',
                                              resize: 'vertical',
                                              fontFamily: 'inherit'
                                            }}
                                            onChange={(e) => updateOptionText(sec.id, q.id, oIdx, e.target.value)}
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

                                        {/* Option Image Controls */}
                                        {getOptionImage(opt) ? (
                                          <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', marginLeft: '24px', marginTop: '2px' }}>
                                            <div
                                              className="editor-image-preview-container"
                                              onDragOver={handleDragOver}
                                              onDragLeave={handleDragLeave}
                                              onDrop={(e) => handleDropImage(e, sec.id, q.id, null, oIdx)}
                                            >
                                              <img
                                                src={getOptionImage(opt)}
                                                alt={`Option ${String.fromCharCode(65 + oIdx)}`}
                                                style={{
                                                  maxWidth: '100%',
                                                  maxHeight: '100px',
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
                                                        updateOptionImage(sec.id, q.id, oIdx, { image: uploadEvent.target.result });
                                                      };
                                                      reader.readAsDataURL(file);
                                                    }
                                                  }}
                                                />
                                              </label>
                                              <button
                                                type="button"
                                                className="btn btn-danger btn-sm"
                                                style={{ padding: '2px 6px', fontSize: '10px' }}
                                                onClick={() => updateOptionImage(sec.id, q.id, oIdx, { image: '' })}
                                              >
                                                Remove Image
                                              </button>
                                            </div>
                                            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px' }}>
                                              <div className="form-group" style={{ margin: 0 }}>
                                                <label style={{ fontSize: '9px' }}>Width (px)</label>
                                                <input
                                                  type="number"
                                                  value={getOptionImageWidth(opt)}
                                                  style={{ padding: '2px 6px', fontSize: '11px' }}
                                                  onChange={(e) => updateOptionImage(sec.id, q.id, oIdx, { imageWidth: Number(e.target.value) || 0 })}
                                                />
                                              </div>
                                              <div className="form-group" style={{ margin: 0 }}>
                                                <label style={{ fontSize: '9px' }}>Height (px)</label>
                                                <input
                                                  type="number"
                                                  value={getOptionImageHeight(opt)}
                                                  style={{ padding: '2px 6px', fontSize: '11px' }}
                                                  onChange={(e) => updateOptionImage(sec.id, q.id, oIdx, { imageHeight: Number(e.target.value) || 0 })}
                                                />
                                              </div>
                                            </div>
                                          </div>
                                        ) : (
                                          <div style={{ marginLeft: '24px' }}>
                                            <label
                                              className="image-upload-dashed-zone"
                                              style={{ padding: '4px 8px', gap: '4px', fontSize: '10px' }}
                                              onDragOver={handleDragOver}
                                              onDragLeave={handleDragLeave}
                                              onDrop={(e) => handleDropImage(e, sec.id, q.id, null, oIdx)}
                                            >
                                              <ImageIcon size={12} className="text-secondary" />
                                              <span>Upload Option Image (click, drag, or Ctrl+V)</span>
                                              <input
                                                type="file"
                                                accept="image/*"
                                                style={{ display: 'none' }}
                                                onChange={(e) => {
                                                  const file = e.target.files[0];
                                                  if (file) {
                                                    const reader = new FileReader();
                                                    reader.onload = (uploadEvent) => {
                                                      updateOptionImage(sec.id, q.id, oIdx, { image: uploadEvent.target.result });
                                                    };
                                                    reader.readAsDataURL(file);
                                                  }
                                                }}
                                              />
                                            </label>
                                          </div>
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

                                              {/* Sub-Question Image Controls - positioned directly below text */}
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
                                        </div>
                                      ))}
                                      <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                                        <input
                                          type="checkbox"
                                          checked={!!q.pageBreakBefore}
                                          onChange={(e) => updateQuestion(sec.id, q.id, { pageBreakBefore: e.target.checked })}
                                        />
                                        <label style={{ fontSize: '11px' }}>Page break before question</label>
                                      </div>
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

                                {/* Match the Following Specific Fields */}
                                {sec.type === 'match_following' && q.matchPairs && (
                                  <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                      <label style={{ fontSize: '10px', fontWeight: 'bold' }}>Match Pairs (Column A & Column B)</label>
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

                                    <button
                                      type="button"
                                      className="btn btn-secondary btn-sm"
                                      style={{ width: '100%', padding: '6px 12px', fontSize: '11px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '4px', marginTop: '2px' }}
                                      onClick={() => {
                                        updateQuestion(sec.id, q.id, {
                                          matchPairs: [...q.matchPairs, { premise: '', premiseImage: '', response: '', responseImage: '' }],
                                          _shuffledB: null
                                        });
                                      }}
                                    >
                                      <Plus size={13} /> Add Pair
                                    </button>

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
                  {examCurrentTotalMarks !== metadata.maxMarks ? (
                    <div className="warning-badge" style={{ display: 'flex', alignSelf: 'flex-start' }}>
                      <AlertTriangle size={14} />
                      <span>Marks Mismatch: Current Questions = {formatMarks(examCurrentTotalMarks)} marks (Target = {formatMarks(metadata.maxMarks)} marks).</span>
                    </div>
                  ) : (
                    <div className="warning-badge" style={{ display: 'flex', alignSelf: 'flex-start', backgroundColor: 'rgba(16,185,129,0.15)', color: 'var(--success)', borderColor: 'rgba(16,185,129,0.3)' }}>
                      <CheckCircle size={14} />
                      <span>Perfect: Sum of all questions matches targets!</span>
                    </div>
                  )}
                  {hasZeroMark && (
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
          <div className="action-bar studio-action-bar">
            {/* Wizard Navigation */}
            <div className="action-bar-group action-bar-wizard">
              {activeTab !== 'branding' && (
                <button
                  className="btn btn-secondary btn-sm action-bar-btn"
                  onClick={() => {
                    if (activeTab === 'metadata') setActiveTab('branding');
                    else if (activeTab === 'sections') setActiveTab('metadata');
                  }}
                >
                  <ArrowLeft size={14} /> Back
                </button>
              )}
              {activeTab !== 'sections' && (
                <button
                  className="btn btn-primary btn-sm action-bar-btn"
                  onClick={() => {
                    if (activeTab === 'branding') setActiveTab('metadata');
                    else if (activeTab === 'metadata') setActiveTab('sections');
                  }}
                >
                  Next <ArrowRight size={14} />
                </button>
              )}
            </div>

            <div className="action-bar-group action-bar-demo">
              <button className="btn btn-secondary btn-sm action-bar-btn" onClick={loadDemo}>
                Demo Data
              </button>
              <button className="btn btn-danger btn-sm action-bar-btn" onClick={resetAll}>
                Clear Draft
              </button>
            </div>
            <div className="action-bar-group action-bar-data">
              <button className="btn btn-primary btn-sm action-bar-btn" onClick={() => exportToCSV(branding, metadata, sections, getOptionText)}>
                <Download size={12} /> Export CSV
              </button>
              <label className="btn btn-secondary btn-sm action-bar-btn" style={{ cursor: 'pointer', margin: 0 }}>
                <Plus size={12} /> Import CSV
                <input type="file" accept=".csv" onChange={importFromCSV} style={{ display: 'none' }} />
              </label>
            </div>
          </div>
        </div>

        {/* Right Column: Real-Time Side-by-Side Live A4 Paper Sheet View */}
        <div className={`studio-preview-column ${!showLivePreview ? 'hidden' : ''}`}>
          {/* Live Preview Controls Header */}
          <div style={{
            padding: '12px 20px',
            borderBottom: '1px solid var(--border-color)',
            background: 'var(--bg-sidebar)',
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            gap: '12px',
            flexWrap: 'wrap'
          }}>
            {/* Zoom Controls */}
            <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
              <button
                className="btn btn-secondary btn-sm"
                style={{ padding: '4px 8px' }}
                onClick={() => setPreviewZoom(prev => Math.max(0.4, Number((prev - 0.05).toFixed(2))))}
                title="Zoom Out (5%)"
              >
                -
              </button>
              <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-primary)', minWidth: '48px', textAlign: 'center' }}>
                {Math.round(previewZoom * 100)}%
              </span>
              <button
                className="btn btn-secondary btn-sm"
                style={{ padding: '4px 8px' }}
                onClick={() => setPreviewZoom(prev => Math.min(1.5, Number((prev + 0.05).toFixed(2))))}
                title="Zoom In (5%)"
              >
                +
              </button>
              <button
                className="btn btn-secondary btn-sm"
                style={{ padding: '4px 8px', fontSize: '11px' }}
                onClick={() => setPreviewZoom(0.85)}
                title="Reset Zoom to 85%"
              >
                Reset
              </button>
            </div>

            {/* Quick Actions */}
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <button
                className="btn btn-print btn-sm"
                onClick={handleDirectPrint}
                disabled={!hasQuestions()}
                title="Direct Print Paper (Ctrl+P / ⌘P)"
                style={{ display: 'flex', alignItems: 'center', gap: '6px' }}
              >
                <Printer size={14} />
                <span>Print</span>
              </button>
              <button
                className="btn btn-primary btn-sm"
                onClick={() => setIsPreviewOpen(true)}
                title="Open Full Screen Preview Modal"
                style={{ display: 'flex', alignItems: 'center', gap: '6px' }}
              >
                <Maximize2 size={14} />
                <span>Full Preview</span>
              </button>
            </div>
          </div>

          {/* Paper Zoom Container */}
          <div style={{ padding: '24px 12px', flex: 1, display: 'flex', justifyContent: 'center', overflowY: 'auto' }}>
            <div style={{ transform: `scale(${previewZoom})`, transformOrigin: 'top center', transition: 'transform 0.2s ease' }}>
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
                  {metadata.instructions && (
                    <div className="exam-meta-item full-width" style={{ marginTop: '8px', display: 'block', fontSize: '15px', fontWeight: 'normal', wordBreak: 'break-word', overflowWrap: 'break-word' }}>
                      <div style={{ fontWeight: 'bold', marginBottom: '4px' }}>General Instructions:</div>
                      <div className="general-instructions-render" style={{ whiteSpace: 'pre-wrap' }}>
                        {(metadata.instructions || '').split('\n').map((line, idx) => (
                          <span key={idx} className="paper-text-line" style={{ display: 'block', minHeight: '1.2rem' }}>{line || ' '}</span>
                        ))}
                      </div>
                    </div>
                  )}
                </div>

                {/* Render Sections & Questions */}
                <div className="paper-sections-container">
                  {sections.length === 0 ? (
                    <div style={{ textAlign: 'center', color: '#999', padding: '40px 0', fontSize: '14px', fontStyle: 'italic' }}>
                      No sections added yet. Use the "Questions" tab in the editor to create sections and populate questions.
                    </div>
                  ) : (
                    sections.map((sec, sIdx) => {
                      const firstQHasBreak = sec.questions && sec.questions[0] && sec.questions[0].pageBreakBefore;
                      const hasSecBreak = sec.pageBreakBefore || firstQHasBreak;
                      return (
                        <React.Fragment key={sec.id}>
                          {hasSecBreak && (
                            <div className="preview-page-break-indicator print-hide" style={{
                              display: 'flex',
                              alignItems: 'center',
                              justifyContent: 'center',
                              gap: '8px',
                              margin: '20px 0 14px 0',
                              padding: '6px 12px',
                              borderTop: '2px dashed var(--accent)',
                              borderBottom: '2px dashed var(--accent)',
                              backgroundColor: 'rgba(99, 102, 241, 0.12)',
                              color: 'var(--accent)',
                              fontSize: '11px',
                              fontWeight: 'bold',
                              borderRadius: '4px',
                              letterSpacing: '0.5px'
                            }}>
                              <span>✂ PAGE BREAK BEFORE SECTION {String.fromCharCode(65 + sIdx)} ✂</span>
                            </div>
                          )}
                          <div className={`paper-section ${hasSecBreak ? 'page-break-before' : ''}`}>
                            <div className="paper-section-header">
                              <div className="paper-section-title-row">
                                <h2 className="paper-section-title">{(sec.title || '').toUpperCase()}</h2>
                                <span className="paper-section-marks">[{formatMarks(sec.marks)} Marks]</span>
                              </div>
                              {sec.instructions && (
                                <div className="paper-section-instructions" style={{ whiteSpace: 'pre-wrap' }}>
                                  {(sec.instructions || '').split('\n').map((line, idx) => (
                                    <span key={idx} className="paper-text-line" style={{ display: 'block', minHeight: '1.2rem' }}>{line || ' '}</span>
                                  ))}
                                </div>
                              )}
                            </div>

                            <div className="paper-questions-list">
                              {sec.questions.map((q, qIdx) => {
                                let previousQuestionsCount = 0;
                                for (let i = 0; i < sIdx; i++) {
                                  previousQuestionsCount += sections[i].questions.length;
                                }
                                const globalNum = previousQuestionsCount + qIdx + 1;
                                const showQBreak = q.pageBreakBefore && qIdx > 0;

                                return (
                                  <React.Fragment key={q.id}>
                                    {showQBreak && (
                                      <div className="preview-page-break-indicator print-hide" style={{
                                        display: 'flex',
                                        alignItems: 'center',
                                        justifyContent: 'center',
                                        gap: '8px',
                                        margin: '16px 0 12px 0',
                                        padding: '6px 12px',
                                        borderTop: '2px dashed var(--accent)',
                                        borderBottom: '2px dashed var(--accent)',
                                        backgroundColor: 'rgba(99, 102, 241, 0.08)',
                                        color: 'var(--accent)',
                                        fontSize: '11px',
                                        fontWeight: 'bold',
                                        borderRadius: '4px',
                                        letterSpacing: '0.5px'
                                      }}>
                                        <span>✂ PAGE BREAK BEFORE QUESTION Q{globalNum} ✂</span>
                                      </div>
                                    )}
                                    <div className={`paper-question-item ${showQBreak ? 'page-break-before' : ''}`}>
                                      <span className="paper-question-number">Q{globalNum}.</span>
                                      <div className="paper-question-body">
                                        <p style={{ fontWeight: '500', textAlign: 'justify', textAlignLast: 'left' }} dangerouslySetInnerHTML={{ __html: renderTextWithMath(q.text) }} />

                                        {/* Question Image render */}
                                        {q.image && sec.type !== 'match_following' && (
                                          <div className="paper-image-container" style={{ marginTop: '8px', marginBottom: '8px', display: 'flex', justifyContent: 'flex-start' }}>
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

                                        {/* MCQ Options */}
                                        {sec.type === 'mcq' && q.options && (
                                          <div className={`paper-mcq-options ${canFitSingleLine(q.options) ? 'single-line' : ''}`}>
                                            {q.options.map((opt, oIdx) => (
                                              <div key={oIdx} className="paper-mcq-option" style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                                                <div style={{ display: 'flex', gap: '6px', alignItems: 'flex-start' }}>
                                                  <span style={{ fontWeight: '600', flexShrink: 0 }}>({String.fromCharCode(65 + oIdx)})</span>
                                                  {getOptionText(opt) && (
                                                    <span className="paper-mcq-option-text" dangerouslySetInnerHTML={{ __html: renderTextWithMath(getOptionText(opt)) }} />
                                                  )}
                                                </div>
                                                {getOptionImage(opt) && (
                                                  <div style={{ marginTop: '4px', marginLeft: '22px' }}>
                                                    <img
                                                      src={getOptionImage(opt)}
                                                      alt={`Option ${String.fromCharCode(65 + oIdx)}`}
                                                      style={{
                                                        width: getOptionImageWidth(opt) ? `${getOptionImageWidth(opt)}px` : 'auto',
                                                        height: getOptionImageHeight(opt) ? `${getOptionImageHeight(opt)}px` : 'auto',
                                                        maxWidth: '100%',
                                                        maxHeight: '180px',
                                                        objectFit: 'contain'
                                                      }}
                                                    />
                                                  </div>
                                                )}
                                              </div>
                                            ))}
                                          </div>
                                        )}

                                        {/* Essay spaces and Sub-Questions */}
                                        {sec.type === 'essay' && (
                                          <>
                                            {q.subQuestions && q.subQuestions.length > 0 ? (
                                              <div className="paper-subquestions-list" style={{ marginTop: '4px', paddingLeft: '16px', display: 'flex', flexDirection: 'column', gap: '6px' }}>
                                                {q.subQuestions.map((sq, sqIdx) => (
                                                  <div key={sq.id} className="paper-subquestion-item" style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                                                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: '12px' }}>
                                                      <div style={{ display: 'flex', gap: '6px', flex: 1 }}>
                                                        <span style={{ fontWeight: '600', flexShrink: 0 }}>{sq.label || `(${String.fromCharCode(97 + (sqIdx % 26))})`}</span>
                                                        <span style={{ flex: 1, textAlign: 'justify', textAlignLast: 'left' }} dangerouslySetInnerHTML={{ __html: renderTextWithMath(sq.text) }} />
                                                      </div>
                                                      <span className="paper-question-marks" style={{ fontStyle: 'italic', fontSize: '11px' }}>({formatMarks(sq.marks)} M)</span>
                                                    </div>
                                                    {sq.image && (
                                                      <div className="paper-subquestion-image-container" style={{ marginTop: '4px', paddingLeft: '18px' }}>
                                                        <img
                                                          src={sq.image}
                                                          alt={`Sub-question ${sq.label || sqIdx + 1}`}
                                                          style={{
                                                            width: `${sq.imageWidth || 300}px`,
                                                            height: `${sq.imageHeight || 200}px`,
                                                            objectFit: 'contain',
                                                            maxWidth: '100%'
                                                          }}
                                                        />
                                                      </div>
                                                    )}
                                                    {!metadata.separateAnswerSheet && (
                                                      <div style={{ display: 'flex', flexDirection: 'column', gap: '5px', marginTop: '2px', paddingLeft: '20px' }}>
                                                        {Array.from({ length: (sq.blankLines !== undefined && sq.blankLines !== '') ? sq.blankLines : 2 }).map((_, lineIdx) => (
                                                          <div key={lineIdx} className="paper-answer-line"></div>
                                                        ))}
                                                      </div>
                                                    )}
                                                  </div>
                                                ))}
                                              </div>
                                            ) : (
                                              !metadata.separateAnswerSheet && (
                                                <div style={{ display: 'flex', flexDirection: 'column', gap: '5px', marginTop: '3px' }}>
                                                  {Array.from({ length: (q.blankLines !== undefined && q.blankLines !== '') ? q.blankLines : 2 }).map((_, lineIdx) => (
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
                                                const premiseText = typeof pair === 'string' ? pair : ((pair && pair.premise) || '');
                                                const premiseImg = typeof pair === 'object' && pair !== null ? (pair.premiseImage || '') : '';
                                                const itemB = shuffledList[pIdx] || (typeof pair === 'string' ? { response: pair } : pair);
                                                const respText = typeof itemB === 'string' ? itemB : ((itemB && itemB.response) || '');
                                                const respImg = typeof itemB === 'object' && itemB !== null ? (itemB.responseImage || '') : '';

                                                return (
                                                  <tr key={pIdx}>
                                                    <td style={{ padding: '6px 0', verticalAlign: 'top' }}>
                                                      <div style={{ display: 'flex', gap: '6px', alignItems: 'flex-start' }}>
                                                        <span style={{ fontWeight: '600', flexShrink: 0 }}>{pIdx + 1}.</span>
                                                        <div>
                                                          {premiseText && <span dangerouslySetInnerHTML={{ __html: renderTextWithMath(premiseText) }} />}
                                                          {premiseImg && (
                                                            <img
                                                              src={premiseImg}
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
                                  </React.Fragment>
                                );
                              })}
                            </div>
                          </div>
                        </React.Fragment>
                      );
                    })
                  )}
                </div>

                {/* Page number footer */}
                <div className="paper-footer print-hide"></div>
              </div>
            </div>
          </div>
        </div>
      </div>
      {isPreviewOpen && (
        <div className="modal-overlay preview-overlay" onClick={() => setIsPreviewOpen(false)}>
          <div className="modal-content preview-modal-content" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <Eye size={20} style={{ color: 'var(--accent)' }} />
                <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 600 }}>Question Paper Live Preview</h3>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <button
                  className="btn btn-secondary btn-sm"
                  onClick={() => setIsPreviewOpen(false)}
                  title="Close Preview (Esc)"
                  style={{ display: 'flex', alignItems: 'center', gap: '4px', padding: '6px 10px' }}
                >
                  <X size={16} />
                </button>
              </div>
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

            {examCurrentTotalMarks !== metadata.maxMarks && (
              <div style={{ padding: '12px 24px', borderBottom: '1px solid var(--border-color)', backgroundColor: 'var(--bg-editor)', display: 'flex' }} className="warning-badge-container print-hide">
                <div className="warning-badge" style={{ display: 'flex', width: '100%', boxSizing: 'border-box' }}>
                  <AlertTriangle size={14} />
                  <span>Marks Mismatch: Current Questions = {formatMarks(examCurrentTotalMarks)} marks (Target = {formatMarks(metadata.maxMarks)} marks).</span>
                </div>
              </div>
            )}

            {hasZeroMark && (
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
                  {metadata.instructions && (
                    <div className="exam-meta-item full-width" style={{ marginTop: '8px', display: 'block', fontSize: '15px', fontWeight: 'normal', wordBreak: 'break-word', overflowWrap: 'break-word' }}>
                      <div style={{ fontWeight: 'bold', marginBottom: '4px' }}>General Instructions:</div>
                      <div style={{ whiteSpace: 'pre-wrap' }}>{metadata.instructions}</div>
                    </div>
                  )}
                </div>

                {/* Render Sections & Questions */}
                <div className="paper-sections-container">
                  {sections.length === 0 ? (
                    <div style={{ textAlign: 'center', color: '#999', padding: '40px 0', fontSize: '14px', fontStyle: 'italic' }}>
                      No sections added yet. Use the "Questions" tab in the editor to create sections and populate questions.
                    </div>
                  ) : (
                    sections.map((sec, sIdx) => {
                      const firstQHasBreak = sec.questions && sec.questions[0] && sec.questions[0].pageBreakBefore;
                      const hasSecBreak = sec.pageBreakBefore || firstQHasBreak;
                      return (
                        <React.Fragment key={sec.id}>
                          {hasSecBreak && (
                            <div className="preview-page-break-indicator print-hide" style={{
                              display: 'flex',
                              alignItems: 'center',
                              justifyContent: 'center',
                              gap: '8px',
                              margin: '20px 0 14px 0',
                              padding: '6px 12px',
                              borderTop: '2px dashed var(--accent)',
                              borderBottom: '2px dashed var(--accent)',
                              backgroundColor: 'rgba(99, 102, 241, 0.12)',
                              color: 'var(--accent)',
                              fontSize: '11px',
                              fontWeight: 'bold',
                              borderRadius: '4px',
                              letterSpacing: '0.5px'
                            }}>
                              <span>✂ PAGE BREAK BEFORE SECTION {String.fromCharCode(65 + sIdx)} ✂</span>
                            </div>
                          )}
                          <div className={`paper-section ${hasSecBreak ? 'page-break-before' : ''}`}>
                            <div className="paper-section-header">
                              <div className="paper-section-title-row">
                                <h2 className="paper-section-title">{(sec.title || '').toUpperCase()}</h2>
                                <span className="paper-section-marks">[{formatMarks(sec.marks)} Marks]</span>
                              </div>
                              {sec.instructions && (
                                <div className="paper-section-instructions" style={{ whiteSpace: 'pre-wrap' }}>
                                  {(sec.instructions || '').split('\n').map((line, idx) => (
                                    <span key={idx} className="paper-text-line" style={{ display: 'block', minHeight: '1.2rem' }}>{line || ' '}</span>
                                  ))}
                                </div>
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
                                const showQBreak = q.pageBreakBefore && qIdx > 0;

                                return (
                                  <React.Fragment key={q.id}>
                                    {showQBreak && (
                                      <div className="preview-page-break-indicator print-hide" style={{
                                        display: 'flex',
                                        alignItems: 'center',
                                        justifyContent: 'center',
                                        gap: '8px',
                                        margin: '16px 0 12px 0',
                                        padding: '6px 12px',
                                        borderTop: '2px dashed var(--accent)',
                                        borderBottom: '2px dashed var(--accent)',
                                        backgroundColor: 'rgba(99, 102, 241, 0.08)',
                                        color: 'var(--accent)',
                                        fontSize: '11px',
                                        fontWeight: 'bold',
                                        borderRadius: '4px',
                                        letterSpacing: '0.5px'
                                      }}>
                                        <span>✂ PAGE BREAK BEFORE QUESTION Q{globalNum} ✂</span>
                                      </div>
                                    )}
                                    <div className={`paper-question-item ${showQBreak ? 'page-break-before' : ''}`}>
                                      <span className="paper-question-number">Q{globalNum}.</span>
                                      <div className="paper-question-body">
                                        <p style={{ fontWeight: '500', textAlign: 'justify', textAlignLast: 'left' }} dangerouslySetInnerHTML={{ __html: renderTextWithMath(q.text) }} />

                                        {/* Question Image render - positioned directly below Question Text */}
                                        {q.image && sec.type !== 'match_following' && (
                                          <div className="paper-image-container" style={{ marginTop: '8px', marginBottom: '8px', display: 'flex', justifyContent: 'flex-start' }}>
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

                                        {/* MCQ Options */}
                                        {sec.type === 'mcq' && q.options && (
                                          <div className={`paper-mcq-options ${canFitSingleLine(q.options) ? 'single-line' : ''}`}>
                                            {q.options.map((opt, oIdx) => (
                                              <div key={oIdx} className="paper-mcq-option" style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                                                <div style={{ display: 'flex', gap: '6px', alignItems: 'flex-start' }}>
                                                  <span style={{ fontWeight: '600', flexShrink: 0 }}>({String.fromCharCode(65 + oIdx)})</span>
                                                  {getOptionText(opt) && (
                                                    <span className="paper-mcq-option-text" dangerouslySetInnerHTML={{ __html: renderTextWithMath(getOptionText(opt)) }} />
                                                  )}
                                                </div>
                                                {getOptionImage(opt) && (
                                                  <div style={{ marginTop: '4px', marginLeft: '22px' }}>
                                                    <img
                                                      src={getOptionImage(opt)}
                                                      alt={`Option ${String.fromCharCode(65 + oIdx)}`}
                                                      style={{
                                                        width: getOptionImageWidth(opt) ? `${getOptionImageWidth(opt)}px` : 'auto',
                                                        height: getOptionImageHeight(opt) ? `${getOptionImageHeight(opt)}px` : 'auto',
                                                        maxWidth: '100%',
                                                        maxHeight: '180px',
                                                        objectFit: 'contain'
                                                      }}
                                                    />
                                                  </div>
                                                )}
                                              </div>
                                            ))}
                                          </div>
                                        )}

                                        {/* Essay spaces and Sub-Questions */}
                                        {sec.type === 'essay' && (
                                          <>
                                            {q.subQuestions && q.subQuestions.length > 0 ? (
                                              <div className="paper-subquestions-list" style={{ marginTop: '4px', paddingLeft: '16px', display: 'flex', flexDirection: 'column', gap: '6px' }}>
                                                {q.subQuestions.map((sq, sqIdx) => (
                                                  <div key={sq.id} className="paper-subquestion-item" style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                                                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: '12px' }}>
                                                      <div style={{ display: 'flex', gap: '6px', flex: 1 }}>
                                                        <span style={{ fontWeight: '600', flexShrink: 0 }}>{sq.label || `(${String.fromCharCode(97 + (sqIdx % 26))})`}</span>
                                                        <span style={{ flex: 1, textAlign: 'justify', textAlignLast: 'left' }} dangerouslySetInnerHTML={{ __html: renderTextWithMath(sq.text) }} />
                                                      </div>
                                                      <span className="paper-question-marks" style={{ fontStyle: 'italic', fontSize: '11px' }}>({formatMarks(sq.marks)} M)</span>
                                                    </div>
                                                    {sq.image && (
                                                      <div className="paper-subquestion-image-container" style={{ marginTop: '4px', paddingLeft: '18px' }}>
                                                        <img
                                                          src={sq.image}
                                                          alt={`Sub-question ${sq.label || sqIdx + 1}`}
                                                          style={{
                                                            width: `${sq.imageWidth || 300}px`,
                                                            height: `${sq.imageHeight || 200}px`,
                                                            objectFit: 'contain',
                                                            maxWidth: '100%'
                                                          }}
                                                        />
                                                      </div>
                                                    )}
                                                    {!metadata.separateAnswerSheet && (
                                                      <div style={{ display: 'flex', flexDirection: 'column', gap: '5px', marginTop: '2px', paddingLeft: '20px' }}>
                                                        {Array.from({ length: (sq.blankLines !== undefined && sq.blankLines !== '') ? sq.blankLines : 2 }).map((_, lineIdx) => (
                                                          <div key={lineIdx} className="paper-answer-line"></div>
                                                        ))}
                                                      </div>
                                                    )}
                                                  </div>
                                                ))}
                                              </div>
                                            ) : (
                                              !metadata.separateAnswerSheet && (
                                                <div style={{ display: 'flex', flexDirection: 'column', gap: '5px', marginTop: '3px' }}>
                                                  {Array.from({ length: (q.blankLines !== undefined && q.blankLines !== '') ? q.blankLines : 2 }).map((_, lineIdx) => (
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
                                                const premiseText = typeof pair === 'string' ? pair : ((pair && pair.premise) || '');
                                                const premiseImg = typeof pair === 'object' && pair !== null ? (pair.premiseImage || '') : '';
                                                const itemB = shuffledList[pIdx] || (typeof pair === 'string' ? { response: pair } : pair);
                                                const respText = typeof itemB === 'string' ? itemB : ((itemB && itemB.response) || '');
                                                const respImg = typeof itemB === 'object' && itemB !== null ? (itemB.responseImage || '') : '';

                                                return (
                                                  <tr key={pIdx}>
                                                    <td style={{ padding: '6px 0', verticalAlign: 'top' }}>
                                                      <div style={{ display: 'flex', gap: '6px', alignItems: 'flex-start' }}>
                                                        <span style={{ fontWeight: '600', flexShrink: 0 }}>{pIdx + 1}.</span>
                                                        <div>
                                                          {premiseText && <span dangerouslySetInnerHTML={{ __html: renderTextWithMath(premiseText) }} />}
                                                          {premiseImg && (
                                                            <img
                                                              src={premiseImg}
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
                                  </React.Fragment>
                                );
                              })}
                            </div>
                          </div>
                        </React.Fragment>
                      );
                    })
                  )}
                </div>

                {/* Page number footer */}
                <div className="paper-footer print-hide"></div>
              </div>
            </div>

            <div className="modal-footer" style={{ gap: '10px', flexWrap: 'wrap' }}>
              <button
                className="btn btn-print"
                disabled={!hasQuestions()}
                title={hasQuestions() ? "Print directly to connected printer or Save as PDF (Ctrl+P / ⌘P)" : "Add at least one question to print"}
                onClick={handleDirectPrint}
                style={{ display: 'flex', alignItems: 'center', gap: '6px' }}
              >
                <Printer size={16} />
                <span>Direct Print</span>
              </button>

              <button
                className="btn btn-primary"
                disabled={isPdfExporting || !hasQuestions()}
                title={hasQuestions() ? "Download PDF" : "Add at least one question to export"}
                onClick={handlePdfExport}
                style={{ display: 'flex', alignItems: 'center', gap: '6px' }}
              >
                {isPdfExporting ? <Loader2 size={16} className="animate-spin" /> : <FileText size={16} />}
                {isPdfExporting ? 'Generating PDF...' : 'Download PDF'}
              </button>

              <button
                className="btn btn-docx"
                disabled={!hasQuestions()}
                title={hasQuestions() ? "Download Word (DOCX)" : "Add at least one question to download"}
                onClick={handleDocxExport}
                style={{ display: 'flex', alignItems: 'center', gap: '6px' }}
              >
                <FileText size={16} style={{ color: '#ffffff' }} />
                <span>Download Word (DOCX)</span>
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
                  <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z" />
                </svg>
                Open in Google Docs
              </h3>
            </div>
            <div className="modal-body docs-modal-body" style={{ flexDirection: 'column', gap: '20px', backgroundColor: 'var(--bg-sidebar)', padding: '24px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <Shield size={22} style={{ color: 'var(--success)' }} />
                <div>
                  <strong style={{ fontSize: '15px', color: 'var(--text-primary)' }}>100% Private & Secure Workflow</strong>
                  <p style={{ fontSize: '13px', color: 'var(--text-secondary)', margin: '2px 0 0 0' }}>
                    Your exam questions are strictly confidential. QuestionNinja generates your Word (.docx) document 100% on your device without sending any data to public third-party servers.
                  </p>
                </div>
              </div>

              <div className="docs-options-grid" style={{ display: 'flex', flexDirection: 'column', gap: '14px', width: '100%' }}>
                <div className="docs-option-card" style={{
                  border: '1px solid var(--border-color)',
                  borderRadius: 'var(--radius-md)',
                  padding: '18px',
                  backgroundColor: 'var(--bg-card)',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '14px'
                }}>
                  <div style={{ fontSize: '13px', color: 'var(--text-primary)', lineHeight: '1.6' }}>
                    <strong>Simple 2-Step Process:</strong>
                    <ol style={{ margin: '8px 0 0 16px', padding: 0 }}>
                      <li>Click a button below to <strong>download your question paper</strong> and launch Google Docs or Drive.</li>
                      <li>Drag and drop the downloaded <strong>.docx file</strong> into Google Docs to start editing immediately.</li>
                    </ol>
                  </div>

                  <div style={{ display: 'flex', gap: '12px', flexWrap: 'wrap', marginTop: '6px' }}>
                    <button
                      className="btn btn-primary"
                      onClick={() => {
                        handleDocxExport();
                        window.open('https://docs.google.com/document/', '_blank', 'noopener,noreferrer');
                        setIsDocsModalOpen(false);
                      }}
                      style={{ fontSize: '13px', display: 'inline-flex', alignItems: 'center', gap: '8px' }}
                    >
                      <Download size={15} /> Download & Open Google Docs
                    </button>
                    <button
                      className="btn btn-secondary"
                      onClick={() => {
                        handleDocxExport();
                        window.open('https://drive.google.com/', '_blank', 'noopener,noreferrer');
                        setIsDocsModalOpen(false);
                      }}
                      style={{ fontSize: '13px', display: 'inline-flex', alignItems: 'center', gap: '8px' }}
                    >
                      <Download size={15} /> Download & Open Google Drive
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
        <div className="about-modal-overlay" onClick={() => setIsAboutModalOpen(false)}>
          <div className="about-modal-content" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '16px 20px' }}>
              <h3 className="modal-title" style={{ display: 'flex', alignItems: 'center', gap: '8px', margin: 0 }}>
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="var(--accent)" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" />
                  <circle cx="12" cy="7" r="4" />
                </svg>
                About Developer
              </h3>
              <button
                type="button"
                className="about-modal-close-btn"
                onClick={() => setIsAboutModalOpen(false)}
                title="Close"
                aria-label="Close About Developer"
              >
                <X size={18} />
              </button>
            </div>
            <div className="about-modal-body">
              <div style={{ display: 'flex', alignItems: 'center', gap: '14px', borderBottom: '1px solid var(--border-color)', paddingBottom: '16px' }}>
                <img src={developerPhoto} alt="Naveen Joshy" style={{ width: '64px', height: '64px', borderRadius: '50%', objectFit: 'cover', objectPosition: 'top', border: '2px solid var(--accent)', flexShrink: 0 }} />
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
                    <a href="tel:+919400489149" style={{ color: 'var(--accent)', textDecoration: 'none', fontWeight: '500' }}>+91 9400489149</a>
                  </div>
                </div>

                <div style={{ display: 'flex', alignItems: 'center', gap: '12px', padding: '10px 14px', backgroundColor: 'var(--bg-main)', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="var(--accent)" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8v.5z" />
                  </svg>
                  <div>
                    <div style={{ fontSize: '11px', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>WhatsApp</div>
                    <a href="https://wa.me/919400489149" target="_blank" rel="noopener noreferrer" style={{ color: 'var(--accent)', textDecoration: 'none', fontWeight: '500' }}>+91 9400489149</a>
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

                <div style={{ display: 'flex', alignItems: 'center', gap: '12px', padding: '10px 14px', backgroundColor: 'var(--bg-main)', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="var(--accent)" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <rect width="20" height="20" x="2" y="2" rx="5" ry="5" />
                    <path d="M16 11.37A4 4 0 1 1 12.63 8 4 4 0 0 1 16 11.37z" />
                    <line x1="17.5" x2="17.51" y1="6.5" y2="6.5" />
                  </svg>
                  <div>
                    <div style={{ fontSize: '11px', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Instagram</div>
                    <a href="https://www.instagram.com/naveenjoshy/" target="_blank" rel="noopener noreferrer" style={{ color: 'var(--accent)', textDecoration: 'none', fontWeight: '500' }}>instagram.com/naveenjoshy</a>
                  </div>
                </div>

                <div style={{ display: 'flex', alignItems: 'center', gap: '12px', padding: '10px 14px', backgroundColor: 'var(--bg-main)', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="var(--accent)" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M18 2h-3a5 5 0 0 0-5 5v3H7v4h3v8h4v-8h3l1-4h-4V7a1 1 0 0 1 1-1h3z" />
                  </svg>
                  <div>
                    <div style={{ fontSize: '11px', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Facebook</div>
                    <a href="https://www.facebook.com/naveenjoshy" target="_blank" rel="noopener noreferrer" style={{ color: 'var(--accent)', textDecoration: 'none', fontWeight: '500' }}>facebook.com/naveenjoshy</a>
                  </div>
                </div>

                <div style={{ display: 'flex', alignItems: 'center', gap: '12px', padding: '10px 14px', backgroundColor: 'var(--bg-main)', borderRadius: '8px', border: '1px solid var(--border-color)' }}>
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="var(--accent)" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M4 4l16 16" />
                    <path d="M4 20 20 4" />
                  </svg>
                  <div>
                    <div style={{ fontSize: '11px', color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>X (Twitter)</div>
                    <a href="https://x.com/naveenjoshy64" target="_blank" rel="noopener noreferrer" style={{ color: 'var(--accent)', textDecoration: 'none', fontWeight: '500' }}>x.com/naveenjoshy64</a>
                  </div>
                </div>
              </div>

              {/* Buy Me a Coffee */}
              <div style={{ borderTop: '1px solid var(--border-color)', paddingTop: '16px', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '10px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <span style={{ fontSize: '18px' }}>☕</span>
                  <span style={{ fontSize: '14px', fontWeight: '700', color: 'var(--text-main)', letterSpacing: '0.3px' }}>Buy Me a Coffee</span>
                </div>
                <p style={{ margin: 0, fontSize: '12px', color: 'var(--text-muted)', textAlign: 'center' }}>Scan to support this project!</p>
                <img src={qrCode} alt="Buy Me a Coffee QR Code" style={{ display: 'block', margin: '0 auto', width: '160px', height: '160px', objectFit: 'contain', borderRadius: '12px', border: '1px solid var(--border-color)', padding: '6px', backgroundColor: '#fff' }} />
              </div>
            </div>
            <div className="modal-footer" style={{ padding: '14px 20px', flexShrink: 0 }}>
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