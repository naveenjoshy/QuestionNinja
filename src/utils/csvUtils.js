import { normalizeLogo } from './formatUtils.js';

export const sanitizeImageURI = (val) => {
  if (!val || typeof val !== 'string') return '';
  const trimmed = val.trim();
  if (!trimmed) return '';
  // Strictly allow base64 image data URIs or blob URLs
  if (trimmed.startsWith('data:image/') || trimmed.startsWith('blob:')) {
    return trimmed;
  }
  const lower = trimmed.toLowerCase();
  // Block any executable or untrusted script/data schemes
  if (
    lower.startsWith('javascript:') ||
    lower.startsWith('vbscript:') ||
    lower.startsWith('data:') ||
    lower.startsWith('file:')
  ) {
    return '';
  }
  // Allow safe relative/local image asset paths
  if (/^[a-zA-Z0-9_\-./]+\.(png|jpe?g|webp|svg|gif|ico)$/i.test(trimmed)) {
    return trimmed;
  }
  // Allow standard web URLs
  if (/^https?:\/\/[^\s$.?#].[^\s]*$/i.test(trimmed)) {
    return trimmed;
  }
  return '';
};

export const generateExportFilename = (ext, metadata = {}) => {
  const className = (metadata?.classDiv || 'Class').trim();
  const subjectName = (metadata?.subject || 'Subject').trim();

  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const dateStr = `${year}-${month}-${day}`;

  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  const seconds = String(now.getSeconds()).padStart(2, '0');
  const timeStr = `${hours}-${minutes}-${seconds}`;

  const sanitize = (str) => str.replace(/[/\\:*?"<>|]/g, '').replace(/\s+/g, ' ').trim();

  const cleanClass = sanitize(className) || 'Class';
  const cleanSubject = sanitize(subjectName) || 'Subject';

  const baseName = `${cleanClass} - ${cleanSubject} - ${dateStr} - ${timeStr}`;
  return ext ? `${baseName}.${ext.replace(/^\./, '')}` : baseName;
};

export const exportToCSV = (branding = {}, metadata = {}, sections = [], getOptionText) => {
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
    'General Instructions',
    'Section Title',
    'Section Marks',
    'Section Instructions',
    'Section Page Break Before',
    'Question Type',
    'Question Text',
    'Question Marks',
    'Options',
    'Blank Lines',
    'Match Pairs',
    'Image Data',
    'Image Width',
    'Image Height',
    'Sub Questions',
    'Shuffle Column B',
    'Table Data',
    'Page Break Before'
  ];

  const safeBranding = branding || {};
  const safeMetadata = metadata || {};
  const safeSections = Array.isArray(sections) ? sections : [];

  const safeGetOptionText = typeof getOptionText === 'function'
    ? getOptionText
    : (opt) => (typeof opt === 'object' && opt !== null ? (opt.text ?? '') : String(opt ?? ''));

  const getBrandingMetaCols = () => [
    safeBranding.logo === 'school_logo.png' ? 'school_logo.png' : (safeBranding.logo || ''),
    safeBranding.logoWidth !== undefined && safeBranding.logoWidth !== null ? safeBranding.logoWidth : '',
    safeBranding.logoHeight !== undefined && safeBranding.logoHeight !== null ? safeBranding.logoHeight : '',
    safeBranding.hideSchoolLogo ? 'true' : 'false',
    safeBranding.headerLogoOnly ? 'true' : 'false',
    safeBranding.schoolName || '',
    safeBranding.schoolAddress || '',
    safeBranding.fontFamily || '',
    safeMetadata.title || '',
    safeMetadata.subject || '',
    safeMetadata.classDiv || '',
    safeMetadata.maxMarks !== undefined && safeMetadata.maxMarks !== null ? safeMetadata.maxMarks : '',
    safeMetadata.duration || '',
    safeMetadata.separateAnswerSheet ? 'true' : 'false',
    safeMetadata.language || '',
    safeMetadata.instructions || ''
  ];

  const rows = [];
  if (safeSections.length === 0) {
    rows.push([
      ...getBrandingMetaCols(),
      ...Array(17).fill('')
    ]);
  } else {
    safeSections.forEach((sec, sIdx) => {
      if (!sec) return;
      const secTitle = (sec.title || '').trim() ? (sec.title || '').trim().toUpperCase() : `SECTION ${sIdx + 1}`;
      const secMarksVal = sec.marks !== undefined && sec.marks !== null ? sec.marks : '';
      const secInstructions = sec.instructions || '';
      const secPageBreak = sec.pageBreakBefore ? 'true' : 'false';

      const questions = Array.isArray(sec.questions) ? sec.questions.filter(Boolean) : [];

      if (questions.length === 0) {
        rows.push([
          ...getBrandingMetaCols(),
          secTitle,
          secMarksVal,
          secInstructions,
          secPageBreak,
          ...Array(13).fill('')
        ]);
      } else {
        questions.forEach((q) => {
          if (!q) return;

          let optionsStr = '';
          const isMcq = sec.type === 'mcq' || q.type === 'mcq' || (Array.isArray(q.options) && q.options.length > 0);
          if (isMcq && Array.isArray(q.options) && q.options.length > 0) {
            const hasImagesOrObjects = q.options.some(opt => typeof opt === 'object' && opt !== null && (opt.image || opt.imageWidth || opt.imageHeight));
            if (hasImagesOrObjects) {
              optionsStr = JSON.stringify(q.options);
            } else {
              optionsStr = q.options.map(opt => safeGetOptionText(opt)).join(';');
            }
          }

          let matchPairsStr = '';
          const isMatch = sec.type === 'match_following' || q.type === 'match_following' || (Array.isArray(q.matchPairs) && q.matchPairs.length > 0);
          if (isMatch && Array.isArray(q.matchPairs) && q.matchPairs.length > 0) {
            const hasImages = q.matchPairs.some(p => typeof p === 'object' && p !== null && (p.premiseImage || p.responseImage || p.leftImage || p.rightImage));
            if (hasImages) {
              matchPairsStr = JSON.stringify(q.matchPairs);
            } else {
              matchPairsStr = q.matchPairs.map(p => {
                if (typeof p === 'string') return p;
                const left = p?.premise !== undefined ? p.premise : (p?.premiseText !== undefined ? p.premiseText : (p?.left || ''));
                const right = p?.response !== undefined ? p.response : (p?.responseText !== undefined ? p.responseText : (p?.right || ''));
                return `${left}=${right}`;
              }).join(';');
            }
          }

          let subQsStr = '';
          if (Array.isArray(q.subQuestions) && q.subQuestions.length > 0) {
            subQsStr = JSON.stringify(q.subQuestions);
          }

          let tableDataStr = '';
          const isTable = sec.type === 'table' || q.type === 'table';
          if (isTable && q.tableData) {
            tableDataStr = typeof q.tableData === 'object' ? JSON.stringify(q.tableData) : String(q.tableData);
          }

          const qEffectiveMarks = (Array.isArray(q.subQuestions) && q.subQuestions.length > 0)
            ? q.subQuestions.reduce((sum, sq) => sum + (Number(sq?.marks) || 0), 0)
            : (q.marks !== undefined && q.marks !== null && !isNaN(q.marks) ? Number(q.marks) : 0);

          rows.push([
            ...getBrandingMetaCols(),
            secTitle,
            secMarksVal,
            secInstructions,
            secPageBreak,
            q.type || sec.type || 'essay',
            q.text || '',
            qEffectiveMarks,
            optionsStr,
            q.blankLines !== undefined && q.blankLines !== null ? q.blankLines : '',
            matchPairsStr,
            q.image || '',
            q.imageWidth || '',
            q.imageHeight || '',
            subQsStr,
            (q.shuffleB || q.shuffleColumnB) ? 'true' : 'false',
            tableDataStr,
            q.pageBreakBefore ? 'true' : 'false'
          ]);
        });
      }
    });
  }

  const sanitizeCSVCell = (val) => {
    if (val === undefined || val === null || val === '') return '""';
    let str = String(val);
    if (/^[=+\-@\t\r]/.test(str) && isNaN(Number(str))) {
      str = "'" + str;
    }
    return `"${str.replace(/"/g, '""')}"`;
  };

  const csvContent = [
    headers.map(h => `"${h.replace(/"/g, '""')}"`).join(','),
    ...rows.map(row => row.map(sanitizeCSVCell).join(','))
  ].join('\r\n');

  if (typeof window !== 'undefined' && typeof document !== 'undefined') {
    const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', generateExportFilename('csv', safeMetadata));
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 100);
  }

  return csvContent;
};

export const parseCSV = (text) => {
  if (!text) return [];
  const clean = text.charCodeAt(0) === 0xFEFF ? text.slice(1) : text;
  const lines = [];
  let row = [''];
  let inQuotes = false;
  let fieldStarted = false;
  let isQuotedField = false;
  let leadingWs = '';

  for (let i = 0; i < clean.length; i++) {
    const char = clean[i];
    const nextChar = clean[i + 1];

    if (!fieldStarted) {
      if (char === ' ' || char === '\t') {
        leadingWs += char;
        continue;
      }
      fieldStarted = true;
      if (char === '"') {
        isQuotedField = true;
        inQuotes = true;
        leadingWs = '';
        continue;
      } else {
        isQuotedField = false;
        inQuotes = false;
        if (leadingWs) {
          row[row.length - 1] += leadingWs;
          leadingWs = '';
        }
      }
    }

    if (isQuotedField) {
      if (char === '"') {
        if (inQuotes && nextChar === '"') {
          row[row.length - 1] += '"';
          i++;
        } else if (inQuotes) {
          inQuotes = false;
        } else {
          row[row.length - 1] += '"';
        }
      } else if (!inQuotes && (char === ' ' || char === '\t')) {
        // Discard trailing whitespace outside quotes before delimiter
        continue;
      } else if ((char === '\r' || char === '\n') && inQuotes) {
        if (char === '\r' && nextChar === '\n') {
          i++;
        }
        row[row.length - 1] += '\n';
      } else if (char === ',' && !inQuotes) {
        row.push('');
        fieldStarted = false;
        isQuotedField = false;
        inQuotes = false;
        leadingWs = '';
      } else if ((char === '\r' || char === '\n') && !inQuotes) {
        if (char === '\r' && nextChar === '\n') {
          i++;
        }
        lines.push(row);
        row = [''];
        fieldStarted = false;
        isQuotedField = false;
        inQuotes = false;
        leadingWs = '';
      } else {
        row[row.length - 1] += char;
      }
    } else {
      if (char === ',') {
        row.push('');
        fieldStarted = false;
        isQuotedField = false;
        inQuotes = false;
        leadingWs = '';
      } else if (char === '\r' || char === '\n') {
        if (char === '\r' && nextChar === '\n') {
          i++;
        }
        lines.push(row);
        row = [''];
        fieldStarted = false;
        isQuotedField = false;
        inQuotes = false;
        leadingWs = '';
      } else {
        row[row.length - 1] += char;
      }
    }
  }

  if (row.length > 1 || row[0] !== '') {
    lines.push(row);
  }
  return lines;
};

export const cleanOptionText = (val) => {
  if (val === undefined || val === null) return '';
  const str = String(val).trim();
  const cleaned = str.replace(/^(?:\([A-Za-z0-9]\)|[A-Za-z0-9][.):])\s*/, '').trim();
  return cleaned !== '' ? cleaned : str;
};

export const parseOptionsString = (optionsStr) => {
  if (!optionsStr) return [];
  const trimmed = String(optionsStr).trim();
  if (!trimmed) return [];

  if (trimmed.startsWith('[') || trimmed.startsWith('{')) {
    try {
      const parsed = JSON.parse(trimmed);
      if (Array.isArray(parsed)) {
        return parsed.map(opt => {
          if (typeof opt === 'object' && opt !== null) {
            return {
              text: cleanOptionText(opt.text ?? opt.value ?? ''),
              image: sanitizeImageURI(opt.image),
              imageWidth: opt.imageWidth || 150,
              imageHeight: opt.imageHeight || 100
            };
          }
          return cleanOptionText(String(opt));
        });
      }
    } catch {
      // Fall through to delimiter split
    }
  }

  let parts = [];
  if (trimmed.includes(';')) {
    parts = trimmed.split(';');
  } else if (trimmed.includes('\n')) {
    parts = trimmed.split('\n');
  } else if (trimmed.includes('|')) {
    parts = trimmed.split('|');
  } else if (trimmed.includes(',')) {
    parts = trimmed.split(',');
  } else {
    const prefixRegex = /(?:^|\s+)(?:\([A-Da-d1-4]\)|[A-Da-d1-4][.):])\s+/g;
    const matches = [...trimmed.matchAll(prefixRegex)];
    if (matches.length >= 2) {
      for (let mIdx = 0; mIdx < matches.length; mIdx++) {
        const start = matches[mIdx].index + matches[mIdx][0].length;
        const end = mIdx + 1 < matches.length ? matches[mIdx + 1].index : trimmed.length;
        parts.push(trimmed.slice(start, end).trim());
      }
    } else {
      parts = [trimmed];
    }
  }

  const cleanedParts = parts.map(p => cleanOptionText(p)).filter(p => p !== '');
  return cleanedParts.length > 0 ? cleanedParts : [trimmed];
};

export const inferTypeFromText = (text) => {
  if (!text) return null;
  const lower = String(text).trim().toLowerCase().replace(/[-_]/g, ' ');
  if (
    lower.includes('mcq') ||
    lower.includes('multiple choice') ||
    lower.includes('objective') ||
    lower.includes('choose the correct') ||
    lower.includes('choose the best') ||
    lower.includes('choose correct') ||
    lower.includes('single choice') ||
    lower.includes('choice question') ||
    lower.includes('quiz')
  ) {
    return 'mcq';
  }
  if (lower.includes('true or false') || lower.includes('true/false') || lower.includes('true false') || lower.includes('state true')) {
    return 'true_false';
  }
  if (lower.includes('match the following') || lower.includes('match column') || lower.includes('matching') || lower.includes('match pairs')) {
    return 'match_following';
  }
  if (lower.includes('table') || lower.includes('tabular')) {
    return 'table';
  }
  return null;
};

export const normalizeQType = (rawType, defaultType = 'essay') => {
  if (!rawType) return defaultType;
  const lower = String(rawType).trim().toLowerCase().replace(/[-_]/g, ' ');
  if (
    lower.includes('mcq') ||
    lower.includes('multiple choice') ||
    lower.includes('objective') ||
    lower.includes('single choice') ||
    lower.includes('choice') ||
    lower.includes('select') ||
    lower.includes('quiz') ||
    lower.includes('option')
  ) {
    return 'mcq';
  }
  if (lower.includes('true') || lower.includes('false') || lower.includes('tf')) return 'true_false';
  if (lower.includes('match') || lower.includes('following') || lower.includes('pair')) return 'match_following';
  if (lower.includes('table')) return 'table';
  return defaultType;
};

export const importCsvData = (text) => {
  const parsedRows = parseCSV(text);
  if (parsedRows.length < 2) {
    throw new Error('Invalid CSV file.');
  }

  const headerRow = parsedRows[0].map(h => (h || '').trim().toLowerCase());
  const hasRecognizedHeaders = headerRow.some(h =>
    ['title', 'subject', 'section', 'question', 'marks', 'type', 'options', 'logo', 'blank', 'pairs'].some(keyword => h.includes(keyword))
  );

  const findCol = (exactList, includesList = [], excludesList = []) => {
    // 1. Exact match first
    for (const name of exactList) {
      const target = name.toLowerCase();
      const idx = headerRow.findIndex(h => h === target);
      if (idx !== -1) return idx;
    }
    // 2. Includes match with exclusions
    for (const prefix of includesList) {
      const target = prefix.toLowerCase();
      const idx = headerRow.findIndex(h => {
        if (!h.includes(target)) return false;
        return !excludesList.some(ex => h.includes(ex.toLowerCase()));
      });
      if (idx !== -1) return idx;
    }
    return -1;
  };

  const schoolLogoIdx = findCol(['school logo', 'logo'], ['logo'], ['width', 'height', 'hide']);
  const schoolLogoWidthIdx = findCol(['school logo width', 'logo width'], ['logo width']);
  const schoolLogoHeightIdx = findCol(['school logo height', 'logo height'], ['logo height']);
  const hideSchoolLogoIdx = findCol(['hide school logo', 'hide logo'], ['hide logo']);
  const headerLogoOnlyIdx = findCol(['header logo only', 'logo only'], ['logo only']);
  const schoolNameIdx = findCol(['school name', 'school'], ['school'], ['logo', 'address']);
  const schoolAddressIdx = findCol(['school address', 'address'], ['address']);
  const fontFamilyIdx = findCol(['font family', 'font'], ['font']);

  const examTitleIdx = findCol(['exam title', 'paper title', 'exam name'], ['exam title'], ['section']);
  const subjectIdx = findCol(['subject', 'course'], ['subject']);
  const classDivIdx = findCol(['class div', 'grade', 'class', 'division', 'standard'], ['class']);
  const maxMarksIdx = findCol(['max marks', 'maximum marks', 'total marks', 'exam marks'], ['max mark', 'maximum mark', 'total mark'], ['question', 'section']);
  const durationIdx = findCol(['duration', 'time', 'exam duration'], ['duration']);
  const separateAnswerSheetIdx = findCol(['separate answer sheet', 'answer sheet'], ['answer sheet']);
  const languageIdx = findCol(['language', 'lang'], ['language', 'lang']);
  const generalInstructionsIdx = findCol(['general instructions', 'general instruction', 'exam instructions'], ['general instruction']);

  const secTitleIdx = findCol(
    ['section title', 'section name', 'section', 'sec title', 'sec name', 'sec'],
    ['section title', 'sec title'],
    ['mark', 'instruction', 'break', 'page', 'logo', 'exam']
  ) !== -1 ? findCol(
    ['section title', 'section name', 'section', 'sec title', 'sec name', 'sec'],
    ['section title', 'sec title'],
    ['mark', 'instruction', 'break', 'page', 'logo', 'exam']
  ) : (hasRecognizedHeaders ? -1 : 0);

  const secMarksIdx = findCol(['section marks', 'sec marks', 'sec mark', 'section mark'], ['section mark', 'sec mark'], ['question']) !== -1
    ? findCol(['section marks', 'sec marks', 'sec mark', 'section mark'], ['section mark', 'sec mark'], ['question'])
    : (hasRecognizedHeaders ? -1 : 1);

  const secInstructionsIdx = findCol(['section instructions', 'section instruction', 'sec instructions'], ['section instruction', 'sec instruction']) !== -1
    ? findCol(['section instructions', 'section instruction', 'sec instructions'], ['section instruction', 'sec instruction'])
    : (hasRecognizedHeaders ? -1 : 2);

  const secPageBreakBeforeIdx = findCol(['section page break before', 'section page break', 'sec page break before', 'sec page break'], ['section page break', 'sec page break']);

  const qTypeIdx = findCol(['question type', 'q type', 'qtype', 'type'], ['question type', 'q type'], ['section']) !== -1
    ? findCol(['question type', 'q type', 'qtype', 'type'], ['question type', 'q type'], ['section'])
    : (hasRecognizedHeaders ? -1 : 3);

  const qTextIdx = findCol(
    ['question text', 'question', 'text', 'prompt', 'problem', 'question details', 'question description', 'q text', 'qtext', 'question body', 'body'],
    ['question text', 'q text'],
    ['type', 'mark', 'score', 'option', 'choice', 'sub', 'pair', 'image', 'table', 'break']
  ) !== -1 ? findCol(
    ['question text', 'question', 'text', 'prompt', 'problem', 'question details', 'question description', 'q text', 'qtext', 'question body', 'body'],
    ['question text', 'q text'],
    ['type', 'mark', 'score', 'option', 'choice', 'sub', 'pair', 'image', 'table', 'break']
  ) : (hasRecognizedHeaders ? -1 : 4);

  const qMarksIdx = findCol(
    ['question marks', 'q marks', 'q mark', 'question mark', 'marks', 'mark', 'points', 'point', 'score'],
    ['question mark', 'q mark'],
    ['section', 'total', 'max']
  ) !== -1 ? findCol(
    ['question marks', 'q marks', 'q mark', 'question mark', 'marks', 'mark', 'points', 'point', 'score'],
    ['question mark', 'q mark'],
    ['section', 'total', 'max']
  ) : (hasRecognizedHeaders ? -1 : 5);

  const isMultiOptionHeader = (h) => {
    const trimmed = (h || '').trim().toLowerCase();
    return /^(?:option|opt|choice)\s*[a-e1-5]$/i.test(trimmed) || /^[a-e]$/i.test(trimmed);
  };

  const getSpecificColIdx = (exactNames, prefixNames = []) => {
    for (const name of exactNames) {
      const target = name.toLowerCase();
      const idx = headerRow.findIndex(h => h === target);
      if (idx !== -1) return idx;
    }
    for (const prefix of prefixNames) {
      const target = prefix.toLowerCase();
      const idx = headerRow.findIndex(h => h.startsWith(target) || h.includes(target));
      if (idx !== -1) return idx;
    }
    return -1;
  };

  const optAIdx = getSpecificColIdx(['option a', 'option 1', 'opt a', 'choice a', 'optiona', 'opt1', 'choice 1', 'a'], ['option a', 'opt a', 'choice a']);
  const optBIdx = getSpecificColIdx(['option b', 'option 2', 'opt b', 'choice b', 'optionb', 'opt2', 'choice 2', 'b'], ['option b', 'opt b', 'choice b']);
  const optCIdx = getSpecificColIdx(['option c', 'option 3', 'opt c', 'choice c', 'optionc', 'opt3', 'choice 3', 'c'], ['option c', 'opt c', 'choice c']);
  const optDIdx = getSpecificColIdx(['option d', 'option 4', 'opt d', 'choice d', 'optiond', 'opt4', 'choice 4', 'd'], ['option d', 'opt d', 'choice d']);
  const optEIdx = getSpecificColIdx(['option e', 'option 5', 'opt e', 'choice e', 'optione', 'opt5', 'choice 5', 'e'], ['option e', 'opt e', 'choice e']);

  const multiOptionCols = [];
  if (optAIdx !== -1 && optBIdx !== -1 && optAIdx !== optBIdx) {
    multiOptionCols.push(optAIdx, optBIdx);
    if (optCIdx !== -1 && !multiOptionCols.includes(optCIdx)) multiOptionCols.push(optCIdx);
    if (optDIdx !== -1 && !multiOptionCols.includes(optDIdx)) multiOptionCols.push(optDIdx);
    if (optEIdx !== -1 && !multiOptionCols.includes(optEIdx)) multiOptionCols.push(optEIdx);
  }

  let rawOptionsIdx = -1;
  const exactOptIdx = headerRow.findIndex(h => h === 'options' || h === 'choices' || h === 'question options');
  if (exactOptIdx !== -1) {
    rawOptionsIdx = exactOptIdx;
  } else {
    rawOptionsIdx = headerRow.findIndex(h => (h.includes('options') || h.includes('choices')) && !isMultiOptionHeader(h));
  }
  const optionsIdx = rawOptionsIdx !== -1 ? rawOptionsIdx : (hasRecognizedHeaders ? -1 : 6);

  const blankLinesIdx = findCol(['blank lines', 'blank line', 'blanks', 'lines'], ['blank line']) !== -1
    ? findCol(['blank lines', 'blank line', 'blanks', 'lines'], ['blank line'])
    : (hasRecognizedHeaders ? -1 : 7);

  const matchPairsIdx = findCol(['match pairs', 'matching pairs', 'match column', 'match items', 'pairs'], ['match pair', 'matching', 'pairs']) !== -1
    ? findCol(['match pairs', 'matching pairs', 'match column', 'match items', 'pairs'], ['match pair', 'matching', 'pairs'])
    : (hasRecognizedHeaders ? -1 : 8);

  const imageDataIdx = findCol(['image data', 'image url', 'image src', 'image', 'diagram'], ['image data', 'image url', 'diagram'], ['width', 'height', 'logo']) !== -1
    ? findCol(['image data', 'image url', 'image src', 'image', 'diagram'], ['image data', 'image url', 'diagram'], ['width', 'height', 'logo'])
    : (hasRecognizedHeaders ? -1 : 9);

  const imageWidthIdx = findCol(['image width', 'q image width'], ['image width'], ['logo']) !== -1
    ? findCol(['image width', 'q image width'], ['image width'], ['logo'])
    : (hasRecognizedHeaders ? -1 : 10);

  const imageHeightIdx = findCol(['image height', 'q image height'], ['image height'], ['logo']) !== -1
    ? findCol(['image height', 'q image height'], ['image height'], ['logo'])
    : (hasRecognizedHeaders ? -1 : 11);

  const subQsIdx = findCol(['sub questions', 'subquestions', 'sub-questions', 'sub questions data', 'sub questions json'], ['sub question', 'subquestion']) !== -1
    ? findCol(['sub questions', 'subquestions', 'sub-questions', 'sub questions data', 'sub questions json'], ['sub question', 'subquestion'])
    : (hasRecognizedHeaders ? -1 : 12);

  const shuffleColBIdx = findCol(['shuffle column b', 'shuffle b', 'shuffle'], ['shuffle']);
  const tableDataIdx = findCol(['table data', 'table'], ['table data']);
  const pageBreakBeforeIdx = findCol(['page break before', 'page break', 'question page break'], ['page break'], ['section']);

  const cleanCell = (val) => {
    if (val === undefined || val === null) return '';
    const str = String(val).trim();
    if (/^'([=+\-@\t\r])/.test(str)) {
      return str.slice(1);
    }
    return str;
  };

  const importedSections = [];
  let importedBranding = null;
  let importedMetadata = null;
  let currentSection = null;

  for (let i = 1; i < parsedRows.length; i++) {
    const row = parsedRows[i];
    if (!row || row.length === 0) continue;

    if (!importedBranding && (schoolNameIdx !== -1 || schoolAddressIdx !== -1 || schoolLogoIdx !== -1)) {
      const logoVal = schoolLogoIdx !== -1 ? cleanCell(row[schoolLogoIdx]) : undefined;
      const logoWidthVal = schoolLogoWidthIdx !== -1 ? cleanCell(row[schoolLogoWidthIdx]) : undefined;
      const logoHeightVal = schoolLogoHeightIdx !== -1 ? cleanCell(row[schoolLogoHeightIdx]) : undefined;
      const hideLogoVal = hideSchoolLogoIdx !== -1 ? cleanCell(row[hideSchoolLogoIdx]) : undefined;
      const headerLogoOnlyVal = headerLogoOnlyIdx !== -1 ? cleanCell(row[headerLogoOnlyIdx]) : undefined;
      const schoolNameVal = schoolNameIdx !== -1 ? cleanCell(row[schoolNameIdx]) : undefined;
      const schoolAddressVal = schoolAddressIdx !== -1 ? cleanCell(row[schoolAddressIdx]) : undefined;
      const fontFamilyVal = fontFamilyIdx !== -1 ? cleanCell(row[fontFamilyIdx]) : undefined;

      if (
        (logoVal !== undefined && logoVal !== '') ||
        (schoolNameVal !== undefined && schoolNameVal !== '') ||
        (schoolAddressVal !== undefined && schoolAddressVal !== '') ||
        (logoWidthVal !== undefined && logoWidthVal !== '') ||
        (fontFamilyVal !== undefined && fontFamilyVal !== '')
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

    if (!importedMetadata && (examTitleIdx !== -1 || subjectIdx !== -1 || classDivIdx !== -1 || maxMarksIdx !== -1)) {
      const titleVal = examTitleIdx !== -1 ? cleanCell(row[examTitleIdx]) : undefined;
      const subjectVal = subjectIdx !== -1 ? cleanCell(row[subjectIdx]) : undefined;
      const classDivVal = classDivIdx !== -1 ? cleanCell(row[classDivIdx]) : undefined;
      const maxMarksVal = maxMarksIdx !== -1 ? cleanCell(row[maxMarksIdx]) : undefined;
      const durationVal = durationIdx !== -1 ? cleanCell(row[durationIdx]) : undefined;
      const sepAnsVal = separateAnswerSheetIdx !== -1 ? cleanCell(row[separateAnswerSheetIdx]) : undefined;
      const langVal = languageIdx !== -1 ? cleanCell(row[languageIdx]) : undefined;
      const generalInstrVal = generalInstructionsIdx !== -1 ? cleanCell(row[generalInstructionsIdx]) : undefined;

      if (
        (titleVal !== undefined && titleVal !== '') ||
        (subjectVal !== undefined && subjectVal !== '') ||
        (classDivVal !== undefined && classDivVal !== '') ||
        (maxMarksVal !== undefined && maxMarksVal !== '') ||
        (durationVal !== undefined && durationVal !== '')
      ) {
        importedMetadata = {};
        if (titleVal !== undefined && titleVal !== '') importedMetadata.title = titleVal;
        if (subjectVal !== undefined && subjectVal !== '') importedMetadata.subject = subjectVal;
        if (classDivVal !== undefined && classDivVal !== '') importedMetadata.classDiv = classDivVal;
        if (maxMarksVal !== undefined && maxMarksVal !== '' && !isNaN(maxMarksVal)) importedMetadata.maxMarks = Number(maxMarksVal);
        if (durationVal !== undefined && durationVal !== '') importedMetadata.duration = durationVal;
        if (sepAnsVal !== undefined && sepAnsVal !== '') importedMetadata.separateAnswerSheet = sepAnsVal === 'true';
        if (langVal !== undefined && langVal !== '') importedMetadata.language = langVal;
        if (generalInstrVal !== undefined && generalInstrVal !== '') importedMetadata.instructions = generalInstrVal;
      }
    }

    const secTitle = secTitleIdx !== -1 ? cleanCell(row[secTitleIdx]) : '';
    const rawSecMarks = secMarksIdx !== -1 ? cleanCell(row[secMarksIdx]) : '';
    const secMarks = rawSecMarks !== '' && !isNaN(rawSecMarks) ? Math.max(0, Math.round(Number(rawSecMarks) * 100) / 100) : 0;
    const secInstructions = secInstructionsIdx !== -1 ? cleanCell(row[secInstructionsIdx]) : '';
    const secPageBreakVal = secPageBreakBeforeIdx !== -1 ? cleanCell(row[secPageBreakBeforeIdx]) : '';

    const rawQType = qTypeIdx !== -1 ? cleanCell(row[qTypeIdx]) : '';
    const qText = qTextIdx !== -1 ? cleanCell(row[qTextIdx]) : '';
    const rawQMarks = qMarksIdx !== -1 ? cleanCell(row[qMarksIdx]) : '';
    const qMarks = (rawQMarks !== '' && !isNaN(rawQMarks)) ? Math.max(0, Math.round(Number(rawQMarks) * 100) / 100) : 1;

    let extractedOptions = [];
    if (multiOptionCols.length >= 2) {
      const fromCols = multiOptionCols
        .map(cIdx => cleanCell(row[cIdx]))
        .filter(val => val !== '')
        .map(val => cleanOptionText(val));
      if (fromCols.length >= 2) {
        extractedOptions = fromCols;
      }
    }

    const optionsStr = (optionsIdx !== -1 && !multiOptionCols.includes(optionsIdx)) ? cleanCell(row[optionsIdx]) : '';
    if (extractedOptions.length === 0 && optionsIdx === headerRow.length - 1 && row.length > headerRow.length) {
      const tailCells = row.slice(optionsIdx).map(c => cleanOptionText(cleanCell(c))).filter(Boolean);
      if (tailCells.length >= 2) {
        extractedOptions = tailCells;
      }
    }
    if (extractedOptions.length === 0 && optionsStr) {
      extractedOptions = parseOptionsString(optionsStr);
    }

    const blankLinesVal = blankLinesIdx !== -1 ? cleanCell(row[blankLinesIdx]) : '';
    const matchPairsStr = matchPairsIdx !== -1 ? cleanCell(row[matchPairsIdx]) : '';
    const imageData = imageDataIdx !== -1 ? cleanCell(row[imageDataIdx]) : '';
    const imageWidthVal = imageWidthIdx !== -1 ? cleanCell(row[imageWidthIdx]) : '';
    const imageHeightVal = imageHeightIdx !== -1 ? cleanCell(row[imageHeightIdx]) : '';
    const subQsStr = subQsIdx !== -1 ? cleanCell(row[subQsIdx]) : '';
    const shuffleColBVal = shuffleColBIdx !== -1 ? cleanCell(row[shuffleColBIdx]) : '';
    const tableDataStr = tableDataIdx !== -1 ? cleanCell(row[tableDataIdx]) : '';
    const pageBreakBeforeVal = pageBreakBeforeIdx !== -1 ? cleanCell(row[pageBreakBeforeIdx]) : '';

    const hasOptions = extractedOptions.length > 0;
    const hasQuestionData = Boolean(
      (qText && qText.trim() !== '') ||
      (rawQType && rawQType.trim() !== '') ||
      (rawQMarks !== '') ||
      (imageData && imageData.trim() !== '') ||
      (blankLinesVal !== '') ||
      (matchPairsStr && matchPairsStr.trim() !== '') ||
      (optionsStr && optionsStr.trim() !== '') ||
      hasOptions ||
      (tableDataStr && tableDataStr.trim() !== '') ||
      (subQsStr && subQsStr.trim() !== '')
    );

    if (!secTitle && !hasQuestionData) continue;

    const inferredSecType = inferTypeFromText(secTitle) || inferTypeFromText(secInstructions);
    const initialSecType = normalizeQType(rawQType, '') || (hasOptions ? 'mcq' : (inferredSecType || 'essay'));

    if (secTitle && (!currentSection || currentSection.title.toUpperCase() !== secTitle.toUpperCase())) {
      currentSection = {
        id: `sec-${Date.now()}-${i}`,
        title: secTitle.toUpperCase(),
        marks: secMarks,
        instructions: secInstructions,
        type: initialSecType,
        pageBreakBefore: secPageBreakVal === 'true',
        questions: []
      };
      importedSections.push(currentSection);
    }

    if (!currentSection) {
      currentSection = {
        id: `sec-${Date.now()}-${i}`,
        title: 'SECTION 1',
        marks: secMarks,
        instructions: secInstructions,
        type: initialSecType,
        questions: []
      };
      importedSections.push(currentSection);
    }

    let effectiveQType = normalizeQType(rawQType, '');
    if (!effectiveQType) {
      if (hasOptions) {
        effectiveQType = 'mcq';
      } else if (matchPairsStr) {
        effectiveQType = 'match_following';
      } else if (tableDataStr) {
        effectiveQType = 'table';
      } else if (currentSection && currentSection.type) {
        effectiveQType = currentSection.type;
      } else {
        effectiveQType = inferredSecType || 'essay';
      }
    } else if (hasOptions && effectiveQType !== 'mcq') {
      effectiveQType = 'mcq';
    }

    if (currentSection.questions.length === 0 && currentSection.type === 'essay' && effectiveQType !== 'essay') {
      currentSection.type = effectiveQType;
    }

    if (hasQuestionData) {
      const q = {
        id: `q-${Date.now()}-${i}`,
        type: effectiveQType,
        text: qText || (effectiveQType === 'match_following' ? 'Match the following items:' : ''),
        marks: qMarks
      };

      if (subQsStr) {
        let parsedSuccessfully = false;
        try {
          const parsedSub = typeof subQsStr === 'string' ? JSON.parse(subQsStr) : subQsStr;
          if (Array.isArray(parsedSub)) {
            q.subQuestions = parsedSub.map((sq, sqIdx) => ({
              id: sq?.id || `sq-${Date.now()}-${sqIdx}`,
              label: sq?.label || `(${String.fromCharCode(97 + (sqIdx % 26))})`,
              text: sq?.text || '',
              marks: Number(sq?.marks) || 1,
              blankLines: sq?.blankLines !== undefined && sq?.blankLines !== '' ? Math.max(0, Number(sq.blankLines)) : 2,
              image: sanitizeImageURI(sq?.image),
              imageWidth: Number(sq?.imageWidth) || 300,
              imageHeight: Number(sq?.imageHeight) || 200
            }));
            q.marks = q.subQuestions.reduce((sum, sq) => sum + (Number(sq.marks) || 0), 0);
            parsedSuccessfully = true;
          }
        } catch {
          parsedSuccessfully = false;
        }

        if (!parsedSuccessfully) {
          const trimmedSub = subQsStr.trim();
          let parts = [];
          if (trimmedSub.includes(';') || trimmedSub.includes('\n') || trimmedSub.includes('|')) {
            parts = trimmedSub.split(/[;\n|]+/).map(p => p.trim()).filter(Boolean);
          } else {
            const labelRegex = /(?:^|\s+)(?:\([a-z0-9]\)|[a-z0-9][.):])\s+/gi;
            const matches = [...trimmedSub.matchAll(labelRegex)];
            if (matches.length >= 2) {
              for (let mIdx = 0; mIdx < matches.length; mIdx++) {
                const start = matches[mIdx].index;
                const end = mIdx + 1 < matches.length ? matches[mIdx + 1].index : trimmedSub.length;
                parts.push(trimmedSub.slice(start, end).trim());
              }
            } else if (trimmedSub) {
              parts = [trimmedSub];
            }
          }

          if (parts.length > 0) {
            q.subQuestions = parts.map((part, sqIdx) => {
              let label = `(${String.fromCharCode(97 + (sqIdx % 26))})`;
              let subText = part;
              let subMarks = 1;

              const labelMatch = subText.match(/^(\([a-z0-9]+\)|[a-z0-9]+[.):])\s*/i);
              if (labelMatch) {
                const innerLetter = labelMatch[1].replace(/[():. ]/g, '').toLowerCase();
                label = `(${innerLetter})`;
                subText = subText.slice(labelMatch[0].length).trim();
              }

              const marksMatch = subText.match(/(?:\[|\()(\d+(?:\.\d+)?)\s*(?:marks?|m)?(?:\]|\))$/i);
              if (marksMatch) {
                subMarks = Number(marksMatch[1]) || 1;
                subText = subText.slice(0, marksMatch.index).trim();
              }

              return {
                id: `sq-${Date.now()}-${sqIdx}`,
                label,
                text: subText,
                marks: subMarks,
                blankLines: 2,
                image: '',
                imageWidth: 300,
                imageHeight: 200
              };
            });
            q.marks = q.subQuestions.reduce((sum, sq) => sum + (Number(sq.marks) || 0), 0);
          }
        }
      }

      if (effectiveQType === 'match_following') {
        q.shuffleB = true;
        q.shuffleColumnB = true;
      } else if (shuffleColBVal !== '') {
        const isTrue = shuffleColBVal === 'true' || shuffleColBVal === '1';
        q.shuffleB = isTrue;
        q.shuffleColumnB = isTrue;
      } else {
        q.shuffleB = false;
        q.shuffleColumnB = false;
      }

      if (pageBreakBeforeVal !== '') {
        q.pageBreakBefore = pageBreakBeforeVal === 'true';
      }

      if (tableDataStr) {
        try {
          const parsedTbl = typeof tableDataStr === 'string' ? JSON.parse(tableDataStr) : tableDataStr;
          if (parsedTbl) {
            q.tableData = parsedTbl;
            if (parsedTbl.rows) {
              q.tableRows = parsedTbl.rows.length + 1;
              q.tableCols = parsedTbl.headers ? parsedTbl.headers.length : 3;
            }
          }
        } catch {
          // Table data parsing fallback
        }
      }
      if (effectiveQType === 'table' && !q.tableData) {
        q.tableRows = 3;
        q.tableCols = 3;
        q.tableData = {
          headers: ['Column 1', 'Column 2', 'Column 3'],
          rows: [
            ['', '', ''],
            ['', '', '']
          ]
        };
      }

      if (effectiveQType === 'mcq') {
        if (extractedOptions.length > 0) {
          q.options = extractedOptions;
        } else {
          q.options = ['', '', '', ''];
        }
      } else if (effectiveQType === 'essay') {
        q.blankLines = (blankLinesVal !== '' && !isNaN(blankLinesVal)) ? Math.max(0, parseInt(blankLinesVal, 10)) : 5;
      } else if (effectiveQType === 'match_following') {
        if (matchPairsStr) {
          const trimmed = matchPairsStr.trim();
          let matchParsed = false;
          if (trimmed.startsWith('[') || trimmed.startsWith('{')) {
            try {
              const parsed = JSON.parse(trimmed);
              if (Array.isArray(parsed)) {
                q.matchPairs = parsed.map(p => {
                  if (typeof p === 'string') {
                    const parts = p.split('=');
                    return { premise: parts[0] || '', premiseImage: '', response: parts[1] || '', responseImage: '' };
                  }
                  return {
                    premise: p.premise !== undefined ? p.premise : (p.premiseText !== undefined ? p.premiseText : (p.left || '')),
                    premiseImage: sanitizeImageURI(p.premiseImage || p.leftImage),
                    response: p.response !== undefined ? p.response : (p.responseText !== undefined ? p.responseText : (p.right || '')),
                    responseImage: sanitizeImageURI(p.responseImage || p.rightImage)
                  };
                });
                matchParsed = true;
              }
            } catch {
              matchParsed = false;
            }
          }
          if (!matchParsed) {
            q.matchPairs = matchPairsStr.split(';').map(pair => {
              const parts = pair.split('=');
              return { premise: parts[0] || '', premiseImage: '', response: parts[1] || '', responseImage: '' };
            });
          }
        } else {
          q.matchPairs = [
            { premise: 'Column A Item 1', premiseImage: '', response: 'Column B Match 1', responseImage: '' }
          ];
        }
      }

      const safeImageData = sanitizeImageURI(imageData);
      if (safeImageData) {
        q.image = safeImageData;
        q.imageWidth = Number(imageWidthVal) || 300;
        q.imageHeight = Number(imageHeightVal) || 200;
      }

      currentSection.questions.push(q);
      if (currentSection.marks === 0) {
        currentSection.marks = currentSection.questions.reduce((sum, item) => sum + (Number(item.marks) || 0), 0);
      }
    }
  }

  return {
    branding: importedBranding,
    metadata: importedMetadata,
    sections: importedSections
  };
};
