import schoolLogo from '../assets/school_logo.png';

export const FORMULA_BUTTONS = [
  {
    label: 'Basic', buttons: [
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
    ]
  },
  {
    label: 'Greek', buttons: [
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
    ]
  },
  {
    label: 'Operators', buttons: [
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
    ]
  },
  {
    label: 'Calculus', buttons: [
      { symbol: '∫', insert: '∫', hint: 'Integral' },
      { symbol: '∂', insert: '∂', hint: 'Partial derivative' },
      { symbol: '∑', insert: '∑', hint: 'Summation' },
      { symbol: '∏', insert: '∏', hint: 'Product' },
    ]
  },
  {
    label: 'Trig & Log', buttons: [
      { symbol: 'sin', insert: 'sin ', hint: 'Sine' },
      { symbol: 'cos', insert: 'cos ', hint: 'Cosine' },
      { symbol: 'tan', insert: 'tan ', hint: 'Tangent' },
      { symbol: 'log', insert: 'log ', hint: 'Logarithm' },
      { symbol: 'ln', insert: 'ln ', hint: 'Natural log' },
    ]
  }
];

export const DEFAULT_BRANDING = {
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

export const DEFAULT_METADATA = {
  title: 'FIRST TERM SUMMATIVE ASSESSMENT',
  subject: 'Computer Science & Programming',
  classDiv: 'Class X - Division A & B',
  maxMarks: 50,
  duration: '90 Minutes',
  separateAnswerSheet: true,
  language: 'english',
  instructions: ''
};

export const DEFAULT_SECTIONS = [
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

export const _PRESET_TEMPLATES = {
  cbse10: {
    name: 'Class 10 Science (CBSE 80 Marks)',
    branding: {
      schoolName: 'Greenwood International School',
      schoolAddress: 'Sector 14, Main Campus, New Delhi - 110001\nPhone: +91 11 2345 6789 | Web: www.greenwood.edu.in',
      fontFamily: 'Inter',
      hideSchoolLogo: false,
      headerLogoOnly: false,
      logoPos: { x: 0, y: 0 },
      logoWidth: 80,
      logoHeight: 80
    },
    metadata: {
      examTitle: 'MID-TERM ASSESSMENT 2026',
      subject: 'SCIENCE (PHYSICAL & LIFE SCIENCES)',
      grade: 'CLASS X',
      duration: '3 Hours',
      maxMarks: 80,
      date: '2026-09-15',
      instructions: '1. All questions are compulsory.\n2. Section A contains 16 MCQs of 1 mark each.\n3. Section B contains 6 Short Answer Questions of 2 marks each.\n4. Section C contains 7 Short Answer Questions of 3 marks each.\n5. Section D contains 3 Long Answer Questions of 5 marks each.\n6. Use of calculators is not permitted.',
      language: 'english',
      separateAnswerSheet: true
    },
    sections: [
      {
        id: 'sec-cbse-mcq',
        title: 'SECTION A - MULTIPLE CHOICE QUESTIONS',
        instructions: 'Select the single correct option for each question.',
        type: 'mcq',
        questions: [
          { id: 'q-cbse-1', text: 'Which of the following represents the correct balanced chemical equation for the reaction between hydrogen gas and chlorine gas?', marks: 1, options: ['H₂ + Cl₂ → 2HCl', '2H + 2Cl → 2HCl', 'H₂ + Cl₂ → HCl₂', 'H₂ + 2Cl → H₂Cl'] },
          { id: 'q-cbse-2', text: 'The focal length of a concave mirror is 20 cm. Its radius of curvature $R$ will be:', marks: 1, options: ['10 cm', '20 cm', '40 cm', '80 cm'] },
          { id: 'q-cbse-3', text: 'Which gas is evolved when dilute hydrochloric acid reacts with zinc granules?', marks: 1, options: ['Carbon dioxide', 'Hydrogen gas', 'Oxygen gas', 'Nitrogen dioxide'] },
          { id: 'q-cbse-4', text: 'The breakdown of pyruvate to give carbon dioxide, water and energy takes place in:', marks: 1, options: ['Cytoplasm', 'Mitochondria', 'Chloroplast', 'Nucleus'] }
        ]
      },
      {
        id: 'sec-cbse-sa',
        title: 'SECTION B - SHORT ANSWER QUESTIONS',
        instructions: 'Answer each question in about 30-50 words.',
        type: 'essay',
        questions: [
          { id: 'q-cbse-5', text: 'State Ohm\'s Law and write its mathematical formula $V = I \\times R$. Draw a circuit diagram to verify Ohm\'s Law.', marks: 2, blankLines: 4 },
          { id: 'q-cbse-6', text: 'What is a redox reaction? Identify the substance oxidized and reduced in the following reaction:\n$CuO + H_2 \\rightarrow Cu + H_2O$', marks: 2, blankLines: 4 },
          { id: 'q-cbse-7', text: 'Explain why the sky appears blue to an observer on Earth during daytime.', marks: 2, blankLines: 4 }
        ]
      }
    ]
  },
  physics12: {
    name: 'Class 12 Physics (70 Marks)',
    branding: {
      schoolName: 'St. Xavier Higher Secondary School',
      schoolAddress: 'Campus Avenue, Science Block, Bengaluru\nContact: exam.dep@stxaviers.edu',
      fontFamily: 'Montserrat',
      hideSchoolLogo: false,
      headerLogoOnly: false,
      logoPos: { x: 0, y: 0 },
      logoWidth: 80,
      logoHeight: 80
    },
    metadata: {
      examTitle: 'FIRST QUARTERLY EXAMINATION 2026',
      subject: 'PHYSICS (THEORY)',
      grade: 'CLASS XII',
      duration: '3 Hours',
      maxMarks: 70,
      date: '2026-10-01',
      instructions: '1. All questions are compulsory.\n2. Numerical constants: $c = 3 \\times 10^8\\text{ m/s}$, $e = 1.6 \\times 10^{-19}\\text{ C}$.',
      language: 'english',
      separateAnswerSheet: true
    },
    sections: [
      {
        id: 'sec-p12-mcq',
        title: 'SECTION A - CONCEPTUAL MCQs',
        instructions: 'Choose the correct option.',
        type: 'mcq',
        questions: [
          { id: 'q-p12-1', text: 'The electrostatic potential $V$ at any point $(x, y, z)$ in space is given by $V = 4x^2\\text{ volts}$. The electric field $\\vec{E}$ at point $(1, 0, 2)\\text{ m}$ is:', marks: 1, options: ['-8 î V/m', '8 î V/m', '-16 î V/m', '4 î V/m'] },
          { id: 'q-p12-2', text: 'Two thin convex lenses of focal lengths $f_1 = 15\\text{ cm}$ and $f_2 = 30\\text{ cm}$ are kept in contact. Equivalent power $P$ of the combination is:', marks: 1, options: ['+10 D', '+5 D', '+15 D', '+2.5 D'] }
        ]
      }
    ]
  },
  mathQuiz: {
    name: 'Mathematics Weekly Quiz (25 Marks)',
    branding: {
      schoolName: 'Model Public Academy',
      schoolAddress: 'Department of Mathematics',
      fontFamily: 'Inter',
      hideSchoolLogo: false,
      headerLogoOnly: false,
      logoPos: { x: 0, y: 0 },
      logoWidth: 80,
      logoHeight: 80
    },
    metadata: {
      examTitle: 'WEEKLY MATHEMATICS QUIZ',
      subject: 'ALGEBRA & GEOMETRY',
      grade: 'CLASS IX',
      duration: '45 Minutes',
      maxMarks: 25,
      date: '2026-08-20',
      instructions: 'Answer all questions clearly. Show formula steps.',
      language: 'english',
      separateAnswerSheet: true
    },
    sections: [
      {
        id: 'sec-mq-1',
        title: 'SECTION A - ALGEBRA & POLYNOMIALS',
        instructions: 'Solve the following algebraic equations.',
        type: 'essay',
        questions: [
          { id: 'q-mq-1', text: 'Factorise the quadratic polynomial: $x^2 - 7x + 12 = 0$. Find its roots.', marks: 3, blankLines: 4 },
          { id: 'q-mq-2', text: 'Find the value of $k$ if $(x - 1)$ is a factor of $p(x) = 2x^2 + kx + \\sqrt{2}$.', marks: 3, blankLines: 4 }
        ]
      }
    ]
  }
};
