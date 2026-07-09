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
  HelpCircle, 
  Image as ImageIcon,
  CheckCircle, 
  AlertTriangle,
  Move,
  Maximize2,
  ChevronDown,
  ChevronUp,
  Sun,
  Moon
} from 'lucide-react';
import * as docx from 'docx';

const DEFAULT_BRANDING = {
  logo: '',
  logoSize: 100,
  logoPos: { x: 0, y: 0 },
  schoolName: 'St. Andrew\'s International School',
  schoolAddress: 'Sector 4, Green Valley District, New Delhi - 110001\nPhone: +91 11 2345 6789 | Email: contact@standrews.edu.in',
  fontFamily: 'Cinzel'
};

const DEFAULT_METADATA = {
  title: 'FIRST TERM SUMMATIVE ASSESSMENT',
  subject: 'Computer Science & Programming',
  classDiv: 'Class X - Division A & B',
  maxMarks: 50,
  duration: '120 Minutes',
  separateAnswerSheet: false
};

const DEFAULT_SECTIONS = [
  {
    id: 'sec-1',
    title: 'SECTION A: OBJECTIVE TYPE QUESTIONS',
    marks: 10,
    instructions: 'Answer all the questions. Each question carries 1 mark. Select the most appropriate option.',
    questions: [
      {
        id: 'q-1',
        type: 'mcq',
        text: 'Which of the following is NOT a high-level programming language?',
        marks: 1,
        options: ['Python', 'Assembly Language', 'Java', 'C++']
      },
      {
        id: 'q-2',
        type: 'true_false',
        text: 'In Python, variables are statically typed and must be declared before use.',
        marks: 1
      }
    ]
  },
  {
    id: 'sec-2',
    title: 'SECTION B: SHORT ANSWER & FITB',
    marks: 15,
    instructions: 'Answer any 5 questions. Each question carries 3 marks. Fill in the blanks with the correct answer where applicable.',
    questions: [
      {
        id: 'q-3',
        type: 'fill_blank',
        text: 'The process of finding and resolving bugs or defects in a software program is called _______.',
        marks: 3
      },
      {
        id: 'q-4',
        type: 'match_following',
        text: 'Match the computer hardware components with their appropriate primary functions.',
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
    title: 'SECTION C: LONG ANSWER & ESSAYS',
    marks: 25,
    instructions: 'Answer all questions. Assign marks based on the depth and logic of your explanation.',
    questions: [
      {
        id: 'q-5',
        type: 'essay',
        text: 'Discuss the security implications of cloud computing and explain the key differences between Public, Private, and Hybrid Cloud architectures.',
        marks: 10,
        blankLines: 12
      },
      {
        id: 'q-6',
        type: 'normal',
        text: 'Write a Python program to generate the Fibonacci series up to a given number N, and analyze its time complexity using Big O notation.',
        marks: 15
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
        width: branding.logoSize
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
      const newSize = Math.max(40, Math.min(300, resizeStart.current.width + deltaX));
      setBranding(prev => ({
        ...prev,
        logoSize: newSize
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

  const getSectionTargetMarks = () => {
    return sections.reduce((total, s) => total + (Number(s.marks) || 0), 0);
  };

  // State update helpers for Sections & Questions
  const addSection = () => {
    const newSection = {
      id: `sec-${Date.now()}`,
      title: `SECTION ${String.fromCharCode(65 + sections.length)}: NEW SECTION`,
      marks: 10,
      instructions: 'Answer all questions. Each question carries equal marks.',
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

  const addQuestion = (secId, type = 'normal') => {
    const defaultQuestion = {
      id: `q-${Date.now()}`,
      type,
      text: type === 'fill_blank' ? 'Fill in the blank: The sun rises in the _______.' : 'New Question details here...',
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
    }

    setSections(sections.map(s => {
      if (s.id === secId) {
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
        logoSize: 100,
        logoPos: { x: 0, y: 0 },
        schoolName: '',
        schoolAddress: '',
        fontFamily: 'Inter'
      });
      setMetadata({
        title: '',
        subject: '',
        classDiv: '',
        maxMarks: 100,
        duration: '',
        separateAnswerSheet: false
      });
      setSections([]);
    }
  };

  // Print PDF Trigger
  const triggerPrint = () => {
    window.print();
  };

  // DOCX Export Implementation
  const triggerDocxExport = async () => {
    const docSections = [];

    // Create the School branding header
    const headerChildren = [];
    
    if (branding.schoolName) {
      headerChildren.push(
        new docx.Paragraph({
          alignment: docx.AlignmentType.CENTER,
          children: [
            new docx.TextRun({
              text: branding.schoolName.toUpperCase(),
              bold: true,
              size: 28,
              font: branding.fontFamily === 'Inter' ? 'Calibri' : 'Times New Roman'
            })
          ],
          spacing: { after: 120 }
        })
      );
    }

    if (branding.schoolAddress) {
      const addressLines = branding.schoolAddress.split('\n');
      addressLines.forEach(line => {
        headerChildren.push(
          new docx.Paragraph({
            alignment: docx.AlignmentType.CENTER,
            children: [
              new docx.TextRun({
                text: line,
                size: 20,
                font: branding.fontFamily === 'Inter' ? 'Calibri' : 'Times New Roman'
              })
            ],
            spacing: { after: 80 }
          })
        );
      });
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

    // Exam Metadata Info table
    const metaTable = new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      rows: [
        new docx.TableRow({
          children: [
            new docx.TableCell({
              width: { size: 50, type: docx.WidthType.PERCENTAGE },
              borders: {
                top: { style: docx.BorderStyle.NONE },
                bottom: { style: docx.BorderStyle.NONE },
                left: { style: docx.BorderStyle.NONE },
                right: { style: docx.BorderStyle.NONE }
              },
              children: [
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({ text: 'Examination: ', bold: true, size: 22 }),
                    new docx.TextRun({ text: metadata.title, size: 22 })
                  ]
                })
              ]
            }),
            new docx.TableCell({
              width: { size: 50, type: docx.WidthType.PERCENTAGE },
              borders: {
                top: { style: docx.BorderStyle.NONE },
                bottom: { style: docx.BorderStyle.NONE },
                left: { style: docx.BorderStyle.NONE },
                right: { style: docx.BorderStyle.NONE }
              },
              children: [
                new docx.Paragraph({
                  alignment: docx.AlignmentType.RIGHT,
                  children: [
                    new docx.TextRun({ text: 'Subject: ', bold: true, size: 22 }),
                    new docx.TextRun({ text: metadata.subject, size: 22 })
                  ]
                })
              ]
            })
          ]
        }),
        new docx.TableRow({
          children: [
            new docx.TableCell({
              borders: {
                top: { style: docx.BorderStyle.NONE },
                bottom: { style: docx.BorderStyle.NONE },
                left: { style: docx.BorderStyle.NONE },
                right: { style: docx.BorderStyle.NONE }
              },
              children: [
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({ text: 'Class & Div: ', bold: true, size: 22 }),
                    new docx.TextRun({ text: metadata.classDiv, size: 22 })
                  ]
                })
              ]
            }),
            new docx.TableCell({
              borders: {
                top: { style: docx.BorderStyle.NONE },
                bottom: { style: docx.BorderStyle.NONE },
                left: { style: docx.BorderStyle.NONE },
                right: { style: docx.BorderStyle.NONE }
              },
              children: [
                new docx.Paragraph({
                  alignment: docx.AlignmentType.RIGHT,
                  children: [
                    new docx.TextRun({ text: 'Max Marks: ', bold: true, size: 22 }),
                    new docx.TextRun({ text: `${metadata.maxMarks}`, size: 22 })
                  ]
                })
              ]
            })
          ]
        }),
        new docx.TableRow({
          children: [
            new docx.TableCell({
              borders: {
                top: { style: docx.BorderStyle.NONE },
                bottom: { style: docx.BorderStyle.NONE },
                left: { style: docx.BorderStyle.NONE },
                right: { style: docx.BorderStyle.NONE }
              },
              children: [
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({ text: 'Time Allowed: ', bold: true, size: 22 }),
                    new docx.TextRun({ text: metadata.duration, size: 22 })
                  ]
                })
              ]
            }),
            new docx.TableCell({
              borders: {
                top: { style: docx.BorderStyle.NONE },
                bottom: { style: docx.BorderStyle.NONE },
                left: { style: docx.BorderStyle.NONE },
                right: { style: docx.BorderStyle.NONE }
              },
              children: [
                new docx.Paragraph({
                  alignment: docx.AlignmentType.RIGHT,
                  children: [
                    new docx.TextRun({ text: 'Current Marks: ', bold: true, size: 22 }),
                    new docx.TextRun({ text: `${getExamCurrentTotalMarks()}`, size: 22 })
                  ]
                })
              ]
            })
          ]
        })
      ]
    });

    headerChildren.push(metaTable);
    
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
    
    sections.forEach((sec, sIdx) => {
      // Section header
      headerChildren.push(
        new docx.Paragraph({
          spacing: { before: 240, after: 80 },
          children: [
            new docx.TextRun({
              text: sec.title.toUpperCase(),
              bold: true,
              size: 24,
              font: branding.fontFamily === 'Inter' ? 'Calibri' : 'Times New Roman'
            }),
            new docx.TextRun({
              text: `\t(Total: ${sec.marks} Marks)`,
              bold: true,
              size: 22,
              font: branding.fontFamily === 'Inter' ? 'Calibri' : 'Times New Roman'
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
                text: sec.instructions,
                italic: true,
                size: 20,
                font: branding.fontFamily === 'Inter' ? 'Calibri' : 'Times New Roman'
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
              new docx.TextRun({
                text: q.text,
                size: 22
              }),
              new docx.TextRun({
                text: `\t[${q.marks} Marks]`,
                bold: true,
                size: 20
              })
            ]
          })
        );

        // Formatting specific question types
        if (q.type === 'mcq' && q.options) {
          // Render MCQ choices in a 2x2 style list or block list
          q.options.forEach((opt, oIdx) => {
            const letter = String.fromCharCode(65 + oIdx);
            headerChildren.push(
              new docx.Paragraph({
                indent: { left: 720 },
                spacing: { after: 40 },
                children: [
                  new docx.TextRun({
                    text: `(${letter})  ${opt}`,
                    size: 22
                  })
                ]
              })
            );
          });
        } 
        
        else if (q.type === 'essay') {
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
        
        else if (q.type === 'true_false') {
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
        
        else if (q.type === 'match_following' && q.matchPairs) {
          // Build match list
          const columnA = q.matchPairs.map(p => p.premise);
          let columnB = q.matchPairs.map(p => p.response);
          if (q.shuffleB) {
            columnB = [...columnB].sort(() => Math.random() - 0.5);
          }

          // Renders a simple side-by-side matches listing
          for (let index = 0; index < columnA.length; index++) {
            const letter = String.fromCharCode(97 + index); // a, b, c...
            const roman = String.fromCharCode(105 + index); // wait: roman numerals i, ii, iii...
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
                    text: `${index + 1}. ${columnA[index]}`,
                    size: 22
                  }),
                  new docx.TextRun({
                    text: `\t\t\t\t\t\t${romanNum(index)}. ${columnB[index]}`,
                    size: 22
                  })
                ]
              })
            );
          }
        }
      });
    });

    const doc = new docx.Document({
      sections: [{
        properties: {},
        children: headerChildren
      }]
    });

    docx.Packer.toBlob(doc).then(blob => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const safeName = (metadata.title || 'QuestionPaper').replace(/[^a-z0-9]/gi, '_').toLowerCase();
      a.download = `${safeName}.docx`;
      a.click();
    });
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
            <p>Local Question Paper Designer</p>
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
            <Layers size={14} /> Exam Layout
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

                <div className="form-group">
                  <label>Branding Font Family</label>
                  <select 
                    value={branding.fontFamily}
                    onChange={(e) => setBranding({ ...branding, fontFamily: e.target.value })}
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

              <div className="form-group" style={{ display: 'flex', alignItems: 'center', gap: '8px', marginTop: '12px' }}>
                <input 
                  type="checkbox" 
                  id="separateAnswerSheet"
                  checked={metadata.separateAnswerSheet || false}
                  onChange={(e) => setMetadata({ ...metadata, separateAnswerSheet: e.target.checked })}
                  style={{ width: 'auto', cursor: 'pointer', margin: 0 }}
                />
                <label htmlFor="separateAnswerSheet" style={{ cursor: 'pointer', marginBottom: 0, userSelect: 'none', fontSize: '13px', fontWeight: '500' }}>
                  Separate Answer Sheet (remove write spaces & answer slots)
                </label>
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
                                  Q{qIdx + 1} ({q.type.toUpperCase()})
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
                                  value={q.text}
                                  style={{ minHeight: '60px', fontSize: '13px' }}
                                  onChange={(e) => updateQuestion(sec.id, q.id, { text: e.target.value })}
                                />
                                {q.type === 'fill_blank' && (
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
                              {q.type === 'mcq' && q.options && (
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                  <label style={{ fontSize: '10px', fontWeight: 'bold' }}>Options</label>
                                  {q.options.map((opt, oIdx) => (
                                    <div key={oIdx} style={{ display: 'flex', gap: '6px' }}>
                                      <span style={{ fontSize: '13px', alignSelf: 'center' }}>{String.fromCharCode(65 + oIdx)}.</span>
                                      <input 
                                        type="text" 
                                        value={opt}
                                        style={{ padding: '4px 8px', fontSize: '12px' }}
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
                              {q.type === 'essay' && (
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
                              {q.type === 'match_following' && q.matchPairs && (
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
                                    <div key={pIdx} style={{ display: 'grid', gridTemplateColumns: '1fr 1fr auto', gap: '6px' }}>
                                      <input 
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
                            </div>
                          ))}

                          <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: '8px' }}>
                            <select 
                              style={{ width: 'auto', padding: '6px 12px', fontSize: '13px' }}
                              onChange={(e) => {
                                if (e.target.value) {
                                  addQuestion(sec.id, e.target.value);
                                  e.target.value = ''; // Reset select
                                }
                              }}
                            >
                              <option value="">+ Add Question...</option>
                              <option value="normal">Normal Question</option>
                              <option value="mcq">Multiple Choice (MCQ)</option>
                              <option value="essay">Essay / Long Form</option>
                              <option value="fill_blank">Fill in the Blanks</option>
                              <option value="true_false">True / False</option>
                              <option value="match_following">Match the Following</option>
                            </select>
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
            </div>
          )}
        </div>

        {/* Global Action Bar */}
        <div className="action-bar">
          <button className="btn btn-secondary btn-sm" onClick={loadDemo} style={{ flex: 1 }}>
            Demo Data
          </button>
          <button className="btn btn-danger btn-sm" onClick={resetAll} style={{ flex: 1 }}>
            Clear Draft
          </button>
        </div>
      </div>

      {/* Preview Modal */}
      {isPreviewOpen && (
        <div className="modal-overlay" onClick={() => setIsPreviewOpen(false)}>
          <div className="modal-content" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h3>Question Paper Live Preview</h3>
              <button 
                className="btn btn-secondary btn-sm" 
                onClick={() => setIsPreviewOpen(false)}
                style={{ padding: '6px 10px', fontSize: '12px' }}
              >
                Close
              </button>
            </div>
            
            <div className="modal-body">
              {/* Dynamic A4 Preview Sheet */}
              <div className={`paper-sheet font-${branding.fontFamily}`}>
                
                {/* Header Layout */}
                <div className="paper-header">
                  {branding.logo && (
                    <div 
                      ref={logoRef}
                      className={`brand-logo-container ${isDragging ? 'dragging' : ''}`}
                      style={{
                        width: `${branding.logoSize}px`,
                        height: `${branding.logoSize}px`,
                        left: `${branding.logoPos.x}px`,
                        top: `${branding.logoPos.y}px`
                      }}
                      onPointerDown={handleLogoPointerDown}
                      onPointerMove={handleLogoPointerMove}
                      onPointerUp={handleLogoPointerUp}
                    >
                      <div className="drag-indicator">Drag to move</div>
                      <img src={branding.logo} className="brand-logo-img" alt="School Logo" />
                      <div className="resize-handle"></div>
                    </div>
                  )}

                  <div className="school-details">
                    {branding.schoolName && <h1 className="school-name-render">{branding.schoolName}</h1>}
                    {branding.schoolAddress && <p className="school-address-render">{branding.schoolAddress}</p>}
                  </div>
                </div>

                {/* Exam Specs Row */}
                <div className="exam-meta-grid">
                  <div className="exam-meta-item">
                    <span className="exam-meta-label">Examination:</span>
                    <span>{metadata.title || '_______________________'}</span>
                  </div>
                  <div className="exam-meta-item">
                    <span className="exam-meta-label">Subject:</span>
                    <span>{metadata.subject || '_______________________'}</span>
                  </div>
                  <div className="exam-meta-item">
                    <span className="exam-meta-label">Class:</span>
                    <span>{metadata.classDiv || '_______________________'}</span>
                  </div>
                  <div className="exam-meta-item">
                    <span className="exam-meta-label">Max Marks:</span>
                    <span>{metadata.maxMarks}</span>
                  </div>
                  <div className="exam-meta-item">
                    <span className="exam-meta-label">Duration:</span>
                    <span>{metadata.duration || '_______________________'}</span>
                  </div>
                  <div className="exam-meta-item">
                    <span className="exam-meta-label">Total Marks:</span>
                    <span>{getExamCurrentTotalMarks()}</span>
                  </div>
                </div>

                {/* Render Sections & Questions */}
                <div className="paper-sections-container">
                  {sections.length === 0 ? (
                    <div style={{ textAlign: 'center', color: '#999', padding: '40px 0', fontSize: '14px', fontStyle: 'italic' }}>
                      No sections added yet. Use the "Exam Layout" tab in the editor to create sections and populate questions.
                    </div>
                  ) : (
                    sections.map((sec, sIdx) => (
                      <div key={sec.id} className="paper-section page-break-avoid">
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
                                  <p style={{ fontWeight: '500' }}>{q.text}</p>

                                  {/* MCQ Options */}
                                  {q.type === 'mcq' && q.options && (
                                    <div className="paper-mcq-options">
                                      {q.options.map((opt, oIdx) => (
                                        <div key={oIdx} className="paper-mcq-option">
                                          <span style={{ fontWeight: '600' }}>({String.fromCharCode(65 + oIdx)})</span>
                                          <span>{opt}</span>
                                        </div>
                                      ))}
                                    </div>
                                  )}

                                  {/* Essay spaces */}
                                  {q.type === 'essay' && !metadata.separateAnswerSheet && (
                                    <div style={{ display: 'flex', flexDirection: 'column', gap: '14px', marginTop: '8px' }}>
                                      {Array.from({ length: q.blankLines || 5 }).map((_, lineIdx) => (
                                        <div key={lineIdx} style={{ borderBottom: '1px dotted #ccc', height: '14px' }}></div>
                                      ))}
                                    </div>
                                  )}

                                  {/* True/False selection */}
                                  {q.type === 'true_false' && (
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
                                  {q.type === 'match_following' && q.matchPairs && (
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
                                              <td style={{ padding: '4px 0' }}>{pIdx + 1}. {pair.premise}</td>
                                              <td style={{ padding: '4px 0', paddingLeft: '20px' }}>
                                                {roman(pIdx)}. {shuffledList[pIdx] || pair.response}
                                              </td>
                                            </tr>
                                          );
                                        })}
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
              </div>
            </div>

            <div className="modal-footer">
              <button className="btn btn-secondary" onClick={triggerPrint}>
                <Download size={16} /> Export PDF / Print
              </button>
              <button className="btn btn-primary" onClick={triggerDocxExport}>
                <FileText size={16} /> Export DOCX (Word)
              </button>
              <button className="btn btn-secondary" onClick={() => setIsPreviewOpen(false)}>
                Close
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
