const DEFAULT_BRANDING = {
  logo: schoolLogo,
  logoWidth: 65,
  logoHeight: 65,
  logoPos: { x: 0, y: 0 },
  schoolName: 'Girijyothi CMI Public School',
  schoolAddress: 'Vazhathope, Idukki',
  fontFamily: 'Inter',
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
  language: 'english',
  instructions: ''
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
