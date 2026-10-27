import { normalizeLogo } from './formatUtils';

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

export const exportToCSV = (branding, metadata, sections, getOptionText) => {
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

  const getBrandingMetaCols = () => [
    branding.logo === 'school_logo.png' ? 'school_logo.png' : (branding.logo || ''),
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
    metadata.language || '',
    metadata.instructions || ''
  ];

  const rows = [];
  if (sections.length === 0) {
    rows.push([
      ...getBrandingMetaCols(),
      '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''
    ]);
  } else {
    sections.forEach((sec) => {
      if (sec.questions.length === 0) {
        rows.push([
          ...getBrandingMetaCols(),
          sec.title || '',
          sec.marks !== undefined && sec.marks !== null ? sec.marks : '',
          sec.instructions || '',
          sec.pageBreakBefore ? 'true' : 'false',
          '',
          '',
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
            const hasImagesOrObjects = q.options.some(opt => typeof opt === 'object' && opt !== null);
            if (hasImagesOrObjects) {
              optionsStr = JSON.stringify(q.options);
            } else {
              optionsStr = q.options.map(opt => getOptionText(opt)).join(';');
            }
          }

          let matchPairsStr = '';
          if (sec.type === 'match_following' && q.matchPairs) {
            const hasImagesOrObjects = q.matchPairs.some(p => typeof p === 'object' && p !== null);
            if (hasImagesOrObjects) {
              matchPairsStr = JSON.stringify(q.matchPairs);
            } else {
              matchPairsStr = q.matchPairs.map(p => `${typeof p === 'string' ? p : (p.premise || '')}=${typeof p === 'string' ? p : (p.response || '')}`).join(';');
            }
          }

          let subQsStr = '';
          if (q.subQuestions && q.subQuestions.length > 0) {
            subQsStr = JSON.stringify(q.subQuestions);
          }

          let tableDataStr = '';
          if (sec.type === 'table' && q.tableData) {
            tableDataStr = JSON.stringify(q.tableData);
          }

          rows.push([
            ...getBrandingMetaCols(),
            (sec.title || '').toUpperCase(),
            sec.marks !== undefined && sec.marks !== null ? sec.marks : '',
            sec.instructions || '',
            sec.pageBreakBefore ? 'true' : 'false',
            sec.type || 'essay',
            q.text || '',
            q.marks || 0,
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
    if (val === undefined || val === null) return '""';
    let str = String(val);
    if (/^[=+\-@\t\r]/.test(str)) {
      str = "'" + str;
    }
    return `"${str.replace(/"/g, '""')}"`;
  };

  const csvContent = [
    headers.map(h => `"${h.replace(/"/g, '""')}"`).join(','),
    ...rows.map(row => row.map(sanitizeCSVCell).join(','))
  ].join('\n');

  if (typeof window !== 'undefined' && typeof document !== 'undefined') {
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', generateExportFilename('csv', metadata));
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 100);
  }

  return csvContent;
};

export const parseCSV = (text) => {
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

export const normalizeQType = (rawType, defaultType = 'essay') => {
  if (!rawType) return defaultType;
  const lower = String(rawType).trim().toLowerCase().replace(/[-_]/g, ' ');
  if (lower.includes('mcq') || lower.includes('multiple choice')) return 'mcq';
  if (lower.includes('true') || lower.includes('false') || lower.includes('tf')) return 'true_false';
  if (lower.includes('match') || lower.includes('following') || lower.includes('pair')) return 'match_following';
  if (lower.includes('table')) return 'table';
  return 'essay';
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

  const getColIdx = (...names) => {
    // 1. Try exact matches first
    for (const name of names) {
      const target = name.toLowerCase();
      const idx = headerRow.findIndex(h => h === target);
      if (idx !== -1) return idx;
    }
    // 2. Try contains matches
    for (const name of names) {
      const target = name.toLowerCase();
      const idx = headerRow.findIndex(h => h.includes(target));
      if (idx !== -1) return idx;
    }
    return -1;
  };

  const schoolLogoIdx = getColIdx('School Logo', 'logo');
  const schoolLogoWidthIdx = getColIdx('School Logo Width', 'logo width');
  const schoolLogoHeightIdx = getColIdx('School Logo Height', 'logo height');
  const hideSchoolLogoIdx = getColIdx('Hide School Logo', 'hide logo');
  const headerLogoOnlyIdx = getColIdx('Header Logo Only');
  const schoolNameIdx = getColIdx('School Name', 'school');
  const schoolAddressIdx = getColIdx('School Address', 'address');
  const fontFamilyIdx = getColIdx('Font Family', 'font');

  const examTitleIdx = getColIdx('Exam Title', 'title');
  const subjectIdx = getColIdx('Subject');
  const classDivIdx = getColIdx('Class Div', 'grade', 'class');
  const maxMarksIdx = getColIdx('Max Marks', 'total marks');
  const durationIdx = getColIdx('Duration', 'time');
  const separateAnswerSheetIdx = getColIdx('Separate Answer Sheet', 'answer sheet');
  const languageIdx = getColIdx('Language', 'lang');
  const generalInstructionsIdx = getColIdx('General Instructions');

  const secTitleIdx = getColIdx('Section Title', 'section') !== -1 ? getColIdx('Section Title', 'section') : (hasRecognizedHeaders ? -1 : 0);
  const secMarksIdx = getColIdx('Section Marks') !== -1 ? getColIdx('Section Marks') : (hasRecognizedHeaders ? -1 : 1);
  const secInstructionsIdx = getColIdx('Section Instructions', 'instructions') !== -1 ? getColIdx('Section Instructions', 'instructions') : (hasRecognizedHeaders ? -1 : 2);
  const secPageBreakBeforeIdx = getColIdx('Section Page Break Before');
  const qTypeIdx = getColIdx('Question Type', 'type') !== -1 ? getColIdx('Question Type', 'type') : (hasRecognizedHeaders ? -1 : 3);
  const qTextIdx = getColIdx('Question Text', 'text', 'question') !== -1 ? getColIdx('Question Text', 'text', 'question') : (hasRecognizedHeaders ? -1 : 4);
  const qMarksIdx = getColIdx('Question Marks', 'marks') !== -1 ? getColIdx('Question Marks', 'marks') : (hasRecognizedHeaders ? -1 : 5);
  const optionsIdx = getColIdx('Options', 'choices') !== -1 ? getColIdx('Options', 'choices') : (hasRecognizedHeaders ? -1 : 6);
  const blankLinesIdx = getColIdx('Blank Lines') !== -1 ? getColIdx('Blank Lines') : (hasRecognizedHeaders ? -1 : 7);
  const matchPairsIdx = getColIdx('Match Pairs', 'match', 'matching', 'pairs') !== -1 ? getColIdx('Match Pairs', 'match', 'matching', 'pairs') : (hasRecognizedHeaders ? -1 : 8);
  const imageDataIdx = getColIdx('Image Data', 'image') !== -1 ? getColIdx('Image Data', 'image') : (hasRecognizedHeaders ? -1 : 9);
  const imageWidthIdx = getColIdx('Image Width') !== -1 ? getColIdx('Image Width') : (hasRecognizedHeaders ? -1 : 10);
  const imageHeightIdx = getColIdx('Image Height') !== -1 ? getColIdx('Image Height') : (hasRecognizedHeaders ? -1 : 11);
  const subQsIdx = getColIdx('Sub Questions', 'subquestions') !== -1 ? getColIdx('Sub Questions', 'subquestions') : (hasRecognizedHeaders ? -1 : 12);
  const shuffleColBIdx = getColIdx('Shuffle Column B', 'shuffle');
  const tableDataIdx = getColIdx('Table Data', 'table');
  const pageBreakBeforeIdx = getColIdx('Page Break Before');

  const importedSections = [];
  let importedBranding = null;
  let importedMetadata = null;
  let currentSection = null;

  for (let i = 1; i < parsedRows.length; i++) {
    const row = parsedRows[i];
    if (!row || row.length === 0) continue;

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

    if (!importedMetadata && (examTitleIdx !== -1 || subjectIdx !== -1 || classDivIdx !== -1 || maxMarksIdx !== -1)) {
      const titleVal = examTitleIdx !== -1 ? row[examTitleIdx] : undefined;
      const subjectVal = subjectIdx !== -1 ? row[subjectIdx] : undefined;
      const classDivVal = classDivIdx !== -1 ? row[classDivIdx] : undefined;
      const maxMarksVal = maxMarksIdx !== -1 ? row[maxMarksIdx] : undefined;
      const durationVal = durationIdx !== -1 ? row[durationIdx] : undefined;
      const sepAnsVal = separateAnswerSheetIdx !== -1 ? row[separateAnswerSheetIdx] : undefined;
      const langVal = languageIdx !== -1 ? row[languageIdx] : undefined;
      const generalInstrVal = generalInstructionsIdx !== -1 ? row[generalInstructionsIdx] : undefined;

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
        if (generalInstrVal !== undefined && generalInstrVal !== '') importedMetadata.instructions = generalInstrVal;
      }
    }

    const secTitle = secTitleIdx !== -1 ? (row[secTitleIdx] || '') : '';
    const secMarks = secMarksIdx !== -1 ? Math.max(0, Math.round((Number(row[secMarksIdx]) || 0) * 100) / 100) : 0;
    const secInstructions = secInstructionsIdx !== -1 ? (row[secInstructionsIdx] || '') : '';
    const secPageBreakVal = secPageBreakBeforeIdx !== -1 ? (row[secPageBreakBeforeIdx] || '') : '';
    const rawQType = qTypeIdx !== -1 ? (row[qTypeIdx] || '') : '';
    const qText = qTextIdx !== -1 ? (row[qTextIdx] || '') : '';
    const qMarks = qMarksIdx !== -1 ? Math.max(0, Math.round((Number(row[qMarksIdx]) || 0) * 100) / 100) : 1;
    const optionsStr = optionsIdx !== -1 ? (row[optionsIdx] || '') : '';
    const blankLinesVal = blankLinesIdx !== -1 ? (row[blankLinesIdx] || '') : '';
    const matchPairsStr = matchPairsIdx !== -1 ? (row[matchPairsIdx] || '') : '';
    const imageData = imageDataIdx !== -1 ? (row[imageDataIdx] || '') : '';
    const imageWidthVal = imageWidthIdx !== -1 ? (row[imageWidthIdx] || '') : '';
    const imageHeightVal = imageHeightIdx !== -1 ? (row[imageHeightIdx] || '') : '';
    const subQsStr = subQsIdx !== -1 ? (row[subQsIdx] || '') : '';
    const shuffleColBVal = shuffleColBIdx !== -1 ? row[shuffleColBIdx] : '';
    const tableDataStr = tableDataIdx !== -1 ? row[tableDataIdx] : '';
    const pageBreakBeforeVal = pageBreakBeforeIdx !== -1 ? row[pageBreakBeforeIdx] : '';

    const hasQuestionData = !!(qText || matchPairsStr || optionsStr || tableDataStr || subQsStr);

    if (!secTitle && !hasQuestionData) continue;

    if (secTitle && (!currentSection || currentSection.title !== secTitle)) {
      currentSection = {
        id: `sec-${Date.now()}-${i}`,
        title: secTitle,
        marks: secMarks,
        instructions: secInstructions,
        type: normalizeQType(rawQType, 'essay'),
        pageBreakBefore: secPageBreakVal === 'true',
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
        type: normalizeQType(rawQType, 'essay'),
        questions: []
      };
      importedSections.push(currentSection);
    }

    const effectiveQType = normalizeQType(rawQType || currentSection.type, 'essay');

    if (hasQuestionData || effectiveQType === 'match_following') {
      const q = {
        id: `q-${Date.now()}-${i}`,
        type: effectiveQType,
        text: qText || (effectiveQType === 'match_following' ? 'Match the following items:' : ''),
        marks: qMarks
      };

      if (subQsStr) {
        try {
          const parsedSub = typeof subQsStr === 'string' ? JSON.parse(subQsStr) : subQsStr;
          if (Array.isArray(parsedSub)) {
            q.subQuestions = parsedSub.map((sq, sqIdx) => ({
              id: sq?.id || `sq-${Date.now()}-${sqIdx}`,
              label: sq?.label || `(${String.fromCharCode(97 + (sqIdx % 26))})`,
              text: sq?.text || '',
              marks: Number(sq?.marks) || 1,
              blankLines: sq?.blankLines !== undefined && sq?.blankLines !== '' ? Math.max(0, Number(sq.blankLines)) : 2,
              image: sq?.image || '',
              imageWidth: Number(sq?.imageWidth) || 300,
              imageHeight: Number(sq?.imageHeight) || 200
            }));
            q.marks = q.subQuestions.reduce((sum, sq) => sum + (Number(sq.marks) || 0), 0);
          }
        } catch (e) {
          console.warn('Failed to parse subQuestions from CSV', e);
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
          const parsedTbl = JSON.parse(tableDataStr);
          if (parsedTbl) {
            q.tableData = parsedTbl;
            if (parsedTbl.rows) {
              q.tableRows = parsedTbl.rows.length + 1;
              q.tableCols = parsedTbl.headers ? parsedTbl.headers.length : 3;
            }
          }
        } catch (e) {
          console.warn('Failed to parse tableData from CSV', e);
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
        if (optionsStr) {
          const trimmedOpt = optionsStr.trim();
          if (trimmedOpt.startsWith('[') || trimmedOpt.startsWith('{')) {
            try {
              const parsedOpts = JSON.parse(trimmedOpt);
              q.options = Array.isArray(parsedOpts) ? parsedOpts : ['', '', '', ''];
            } catch {
              q.options = optionsStr.split(';');
            }
          } else {
            q.options = optionsStr.split(';');
          }
        } else {
          q.options = ['', '', '', ''];
        }
      } else if (effectiveQType === 'essay') {
        q.blankLines = (blankLinesVal !== '' && !isNaN(blankLinesVal)) ? Math.max(0, parseInt(blankLinesVal, 10)) : 5;
      } else if (effectiveQType === 'match_following') {
        if (matchPairsStr) {
          const trimmed = matchPairsStr.trim();
          if (trimmed.startsWith('[') || trimmed.startsWith('{')) {
            try {
              const parsed = JSON.parse(trimmed);
              q.matchPairs = Array.isArray(parsed) ? parsed.map(p => {
                if (typeof p === 'string') {
                  const parts = p.split('=');
                  return { premise: parts[0] || '', premiseImage: '', response: parts[1] || '', responseImage: '' };
                }
                return {
                  premise: p.premise !== undefined ? p.premise : (p.premiseText !== undefined ? p.premiseText : (p.left || '')),
                  premiseImage: p.premiseImage || p.leftImage || '',
                  response: p.response !== undefined ? p.response : (p.responseText !== undefined ? p.responseText : (p.right || '')),
                  responseImage: p.responseImage || p.rightImage || ''
                };
              }) : [];
            } catch {
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
          q.matchPairs = [
            { premise: 'Column A Item 1', premiseImage: '', response: 'Column B Match 1', responseImage: '' }
          ];
        }
      }

      if (imageData) {
        q.image = imageData;
        q.imageWidth = Number(imageWidthVal) || 300;
        q.imageHeight = Number(imageHeightVal) || 200;
      }

      currentSection.questions.push(q);
    }
  }

  return {
    branding: importedBranding,
    metadata: importedMetadata,
    sections: importedSections
  };
};
