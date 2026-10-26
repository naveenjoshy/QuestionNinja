import schoolLogo from '../assets/school_logo.png';

export const normalizeLogo = (val) => {
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

export const formatMarks = (val) => {
  const num = Number(val) || 0;
  if (Number.isInteger(num)) return String(num);
  return num.toFixed(2);
};

export const getQuestionMarks = (q) => {
  if (q && q.subQuestions && q.subQuestions.length > 0) {
    return q.subQuestions.reduce((sum, sq) => sum + (Number(sq.marks) || 0), 0);
  }
  return Number(q?.marks) || 0;
};

export const getSectionTotalMarks = (section) => {
  const rawTotal = section.questions.reduce((total, q) => total + getQuestionMarks(q), 0);
  return Math.round(rawTotal * 100) / 100;
};

export const getExamCurrentTotalMarks = (sections) => {
  const rawTotal = sections.reduce((total, s) => total + getSectionTotalMarks(s), 0);
  return Math.round(rawTotal * 100) / 100;
};

export const getExamTotalQuestionsCount = (sections) => {
  return sections.reduce((total, sec) => total + (sec.questions ? sec.questions.length : 0), 0);
};

export const hasBlankQuestions = (sections) => {
  return sections.some(s => s.questions.some(q => {
    if (q.subQuestions && q.subQuestions.length > 0) {
      return q.subQuestions.some(sq => !sq.text || !sq.text.trim());
    }
    return !q.text || !q.text.trim();
  }));
};

export const hasZeroMarkQuestions = (sections) => {
  return sections.some(s => s.questions.some(q => {
    if (q.subQuestions && q.subQuestions.length > 0) {
      return q.subQuestions.some(sq => Number(sq.marks) === 0);
    }
    return Number(q.marks) === 0;
  }));
};

export const getOptionText = (opt) => {
  if (typeof opt === 'object' && opt !== null) return opt.text || '';
  return opt || '';
};

export const getOptionImage = (opt) => {
  if (typeof opt === 'object' && opt !== null) return opt.image || '';
  return '';
};

export const getOptionImageWidth = (opt) => {
  if (typeof opt === 'object' && opt !== null) return opt.imageWidth || 150;
  return 150;
};

export const getOptionImageHeight = (opt) => {
  if (typeof opt === 'object' && opt !== null) return opt.imageHeight || 100;
  return 100;
};

export const canFitSingleLine = (options) => {
  if (!options || options.length === 0) return false;
  if (options.some(opt => getOptionImage(opt))) return false;
  if (options.some(opt => getOptionText(opt).includes('\n'))) return false;
  const totalChars = options.reduce((acc, opt) => acc + getOptionText(opt).length, 0);
  const maxOptLen = Math.max(...options.map(opt => getOptionText(opt).length));
  return options.length <= 4 && totalChars <= 36 && maxOptLen <= 12;
};
