import * as docx from 'docx';
import html2pdf from 'html2pdf.js';
import { normalizeLogo } from './formatUtils';
import { formatMarks, getQuestionMarks, getOptionText, getOptionImage, getOptionImageWidth, getOptionImageHeight, canFitSingleLine } from './formatUtils';
import { latexToPlainText } from './mathUtils';
import { generateExportFilename } from './csvUtils';

export const docxTextRunsWithMath = (text, defaultOptions = {}) => {
  if (!text) return [];
  let processedText = text;

  if (!processedText.includes('$') && (processedText.includes('\\sqrt'))) {
    processedText = processedText.replace(/\\sqrt\{([^}]*)\}/g, '$\\sqrt{$1}$');
  }

  const parts = processedText.split('$');
  const runs = [];
  parts.forEach((part, index) => {
    if (index % 2 === 1) {
      const plainMath = latexToPlainText(part);
      runs.push(new docx.TextRun({
        text: plainMath,
        italics: true,
        font: 'Cambria Math',
        size: defaultOptions.size || 30,
        ...defaultOptions
      }));
    } else {
      const formattedPart = part.replace(/(\b[0-9a-zA-Z]+\b|\))\s*-\s*(\b[0-9a-zA-Z]+\b|\()/g, '$1 − $2');
      const lines = formattedPart.split('\n');
      lines.forEach((line, lIdx) => {
        const runProps = {
          text: line,
          size: defaultOptions.size || 30,
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

export const dataURLToUint8Array = (dataURL) => {
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

export const getFontFamily = (branding, metadata) => {
  if (metadata.language === 'malayalam') return 'Manjari';
  if (metadata.language === 'hindi') return 'Noto Sans Devanagari';
  switch (branding.fontFamily) {
    case 'Inter': return 'Inter';
    case 'Playfair': return 'Playfair Display';
    case 'Montserrat': return 'Montserrat';
    case 'Courier': return 'Courier Prime';
    case 'Merriweather': return 'Merriweather';
    case 'Cinzel': return 'Cinzel';
    default: return branding.fontFamily || 'Inter';
  }
};

export const imageToUint8Array = async (src) => {
  if (!src) return null;
  const resolvedSrc = normalizeLogo(src);
  if (!resolvedSrc) return null;

  return new Promise((resolve) => {
    const img = new Image();
    img.crossOrigin = 'Anonymous';
    img.onload = () => {
      try {
        const canvas = document.createElement('canvas');
        canvas.width = img.width;
        canvas.height = img.height;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0);
        const dataUrl = canvas.toDataURL('image/png');
        const parts = dataUrl.split(';base64,');
        if (parts.length < 2) {
          resolve(null);
          return;
        }
        const base64 = parts[1];
        const raw = window.atob(base64);
        const array = new Uint8Array(raw.length);
        for (let i = 0; i < raw.length; i++) {
          array[i] = raw.charCodeAt(i);
        }
        resolve(array);
      } catch (e) {
        console.error("Canvas PNG conversion failed", e);
        resolve(null);
      }
    };
    img.onerror = () => {
      console.error("Failed to load image for DOCX export");
      resolve(null);
    };
    img.src = resolvedSrc;
  });
};

export const getShuffledList = (q) => {
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

export const generatePdfExport = async (element, metadata) => {
  const originalTransform = element.style.transform;
  element.style.transform = 'none';

  const opt = {
    margin:       0,
    filename:     generateExportFilename('pdf', metadata),
    image:        { type: 'jpeg', quality: 0.98 },
    html2canvas:  { scale: 2, useCORS: true, logging: false },
    jsPDF:        { unit: 'in', format: 'a4', orientation: 'portrait' },
    pagebreak:    { mode: ['css', 'legacy'] }
  };

  try {
    await html2pdf().set(opt).from(element).save();
  } finally {
    element.style.transform = originalTransform;
  }
};

export const generateDocxBlob = async (branding, metadata, sections) => {
  const headerChildren = [];

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
            text: branding.schoolName || '',
            bold: true,
            size: 50,
            font: getFontFamily(branding, metadata)
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
              size: 32,
              font: getFontFamily(branding, metadata)
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
        width: Number(branding.logoWidth) || 100,
        height: Number(branding.logoHeight) || 100
      }
    });

    if (branding.headerLogoOnly) {
      headerChildren.push(
        new docx.Paragraph({
          alignment: docx.AlignmentType.CENTER,
          children: [logoRun],
          spacing: { before: 0, after: 200 }
        })
      );
    } else {
      const headerTable = new docx.Table({
        width: { size: 9000, type: docx.WidthType.DXA },
        columnWidths: [1800, 7200],
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
                width: { size: 1800, type: docx.WidthType.DXA },
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
                width: { size: 7200, type: docx.WidthType.DXA },
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

  const createMetaParagraph = (label, value) => {
    return new docx.Paragraph({
      tabStops: [
        {
          type: docx.TabStopType.RIGHT,
          position: docx.TabStopPosition.MAX
        }
      ],
      children: [
        new docx.TextRun({ text: label, bold: true, size: 30 }),
        new docx.TextRun({ text: `\t${value || ''}`, size: 30 })
      ],
      spacing: { after: 120 }
    });
  };

  headerChildren.push(createMetaParagraph('Examination: ', metadata.title));
  headerChildren.push(createMetaParagraph('Subject: ', metadata.subject));
  headerChildren.push(createMetaParagraph('Class: ', metadata.classDiv));
  headerChildren.push(createMetaParagraph('Max Marks: ', formatMarks(metadata.maxMarks)));
  headerChildren.push(createMetaParagraph('Duration: ', metadata.duration));

  if (metadata.instructions) {
    const instructionRuns = [new docx.TextRun({ text: "General Instructions:", bold: true, size: 28 })];
    const lines = metadata.instructions.split('\n');
    lines.forEach((line) => {
      instructionRuns.push(new docx.TextRun({ break: 1, text: line, size: 24 }));
    });
    headerChildren.push(
      new docx.Paragraph({
        spacing: { before: 120, after: 120 },
        children: instructionRuns
      })
    );
  }

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

  let absoluteQuestionCount = 1;

  for (const sec of sections) {
    const firstQHasBreak = sec.questions && sec.questions[0] && sec.questions[0].pageBreakBefore;
    if (sec.pageBreakBefore || firstQHasBreak) {
      headerChildren.push(
        new docx.Paragraph({
          children: [new docx.PageBreak()]
        })
      );
    }

    const hasInstructions = !!sec.instructions;
    const hasQuestionsInSec = sec.questions && sec.questions.length > 0;
    headerChildren.push(
      new docx.Paragraph({
        keepNext: hasInstructions || hasQuestionsInSec,
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
        spacing: { before: 140, after: hasInstructions ? 20 : 40 },
        children: [
          new docx.TextRun({
            text: (sec.title || '').toUpperCase(),
            bold: true,
            size: 36,
            font: getFontFamily(branding, metadata)
          }),
          new docx.TextRun({
            text: `\t[${formatMarks(sec.marks)} Marks]`,
            bold: true,
            size: 32,
            font: getFontFamily(branding, metadata)
          })
        ]
      })
    );

    if (sec.instructions) {
      const instructionLines = String(sec.instructions).split('\n');
      instructionLines.forEach((line, idx) => {
        const isLastInstructionLine = idx === instructionLines.length - 1;
        headerChildren.push(
          new docx.Paragraph({
            keepNext: isLastInstructionLine ? hasQuestionsInSec : true,
            alignment: docx.AlignmentType.JUSTIFY,
            border: isLastInstructionLine ? {
              bottom: {
                style: docx.BorderStyle.SINGLE,
                size: 8,
                color: '000000',
                space: 4
              }
            } : undefined,
            spacing: {
              before: idx === 0 ? 20 : 10,
              after: isLastInstructionLine ? 40 : 10
            },
            children: [
              new docx.TextRun({
                text: line,
                italic: true,
                size: 30,
                font: getFontFamily(branding, metadata)
              })
            ]
          })
        );
      });
    }

    for (let qIdx = 0; qIdx < sec.questions.length; qIdx++) {
      const q = sec.questions[qIdx];
      const qNum = `Q${absoluteQuestionCount}.`;
      absoluteQuestionCount++;

      if (q.pageBreakBefore && qIdx > 0) {
        headerChildren.push(
          new docx.Paragraph({
            children: [new docx.PageBreak()]
          })
        );
      }

      const isMCQ = Boolean(
        sec.type === 'mcq' ||
        q.type === 'mcq' ||
        (Array.isArray(q.options) && q.options.some(opt => {
          const txt = getOptionText(opt);
          return (txt && txt.trim() !== '') || Boolean(getOptionImage(opt));
        }))
      );
      const hasMCQ = Boolean(isMCQ && Array.isArray(q.options) && q.options.length > 0);
      const hasEssaySubs = Boolean(sec.type === 'essay' && !hasMCQ && q.type !== 'mcq' && q.subQuestions && q.subQuestions.length > 0);

      const linesCount = (!metadata.separateAnswerSheet && (!q.subQuestions || q.subQuestions.length === 0) && !hasMCQ && q.type !== 'mcq')
        ? ((q.blankLines !== undefined && q.blankLines !== '' && q.blankLines !== null)
          ? Math.max(0, parseInt(q.blankLines, 10) || 0)
          : (sec.type === 'essay' ? 4 : 0))
        : 0;

      const hasImage = Boolean(q.image && sec.type !== 'match_following');
      const hasTF = Boolean(sec.type === 'true_false' && !metadata.separateAnswerSheet);
      const hasMatch = Boolean((sec.type === 'match_following' || (q.matchPairs && q.matchPairs.length > 0)) && q.matchPairs && q.matchPairs.length > 0);
      const hasTable = Boolean(sec.type === 'table' && q.tableData);
      const hasBlankLines = linesCount > 0;

      const hasSubElementsAfterText = hasImage || hasMCQ || hasEssaySubs || hasTF || hasMatch || hasTable || hasBlankLines;

      const cleanText = (q.text || '').trimEnd();
      const qLines = cleanText ? cleanText.split('\n') : [''];
      headerChildren.push(
        new docx.Paragraph({
          keepNext: qLines.length > 1 ? true : hasSubElementsAfterText,
          alignment: docx.AlignmentType.JUSTIFY,
          spacing: { before: 100, after: 30 },
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
              size: 30
            }),
            ...docxTextRunsWithMath(qLines[0] || '', { size: 30 }),
            new docx.TextRun({
              text: `\t(${formatMarks(getQuestionMarks(q))} M)`,
              italic: true,
              size: 30
            })
          ]
        })
      );

      for (let lIdx = 1; lIdx < qLines.length; lIdx++) {
        const isLastQLine = lIdx === qLines.length - 1;
        headerChildren.push(
          new docx.Paragraph({
            keepNext: isLastQLine ? hasSubElementsAfterText : true,
            indent: { left: 450 },
            spacing: { after: 20 },
            children: docxTextRunsWithMath(qLines[lIdx], { size: 30 })
          })
        );
      }

      if (hasImage) {
        const imageBytes = await imageToUint8Array(q.image);
        if (imageBytes) {
          const hasSubElementsAfterImage = hasMCQ || hasEssaySubs || hasTF || hasMatch || hasTable || hasBlankLines;
          headerChildren.push(
            new docx.Paragraph({
              keepNext: hasSubElementsAfterImage,
              indent: { left: 450 },
              spacing: { before: 80, after: 120 },
              children: [
                new docx.ImageRun({
                  data: imageBytes,
                  transformation: {
                    width: Number(q.imageWidth) || 300,
                    height: Number(q.imageHeight) || 200
                  }
                })
              ]
            })
          );
        }
      }

      if (hasMCQ) {
        const createOptionParagraphs = (letter, opt) => {
          const text = getOptionText(opt);
          const imgData = getOptionImage(opt);
          const imgW = getOptionImageWidth(opt) || 150;
          const imgH = getOptionImageHeight(opt) || 100;

          const paragraphs = [];
          const lines = text ? text.split('\n') : [''];
          paragraphs.push(
            new docx.Paragraph({
              spacing: { before: 40, after: 40 },
              children: [
                new docx.TextRun({ text: `(${letter})  `, bold: true, size: 26 }),
                ...docxTextRunsWithMath(lines[0] || '', { size: 26 })
              ]
            })
          );
          for (let i = 1; i < lines.length; i++) {
            paragraphs.push(
              new docx.Paragraph({
                indent: { left: 450 },
                spacing: { before: 40, after: 40 },
                children: docxTextRunsWithMath(lines[i], { size: 26 })
              })
            );
          }
          if (imgData) {
            const imgBytes = dataURLToUint8Array(imgData);
            if (imgBytes) {
              paragraphs.push(
                new docx.Paragraph({
                  indent: { left: 450 },
                  spacing: { before: 60, after: 100 },
                  children: [
                    new docx.ImageRun({
                      data: imgBytes,
                      transformation: { width: Number(imgW) || 150, height: Number(imgH) || 100 }
                    })
                  ]
                })
              );
            }
          }
          return paragraphs;
        };

        const isSingle = canFitSingleLine(q.options);

        if (isSingle) {
          const numOpts = q.options.length;
          const availWidth = 8550;
          const colWidth = Math.floor(availWidth / numOpts);
          const cells = q.options.map((opt, oIdx) => {
            const letter = String.fromCharCode(65 + oIdx);
            return new docx.TableCell({
              width: { size: colWidth, type: docx.WidthType.DXA },
              borders: {
                top: { style: docx.BorderStyle.NONE },
                bottom: { style: docx.BorderStyle.NONE },
                left: { style: docx.BorderStyle.NONE },
                right: { style: docx.BorderStyle.NONE }
              },
              children: createOptionParagraphs(letter, opt)
            });
          });

          headerChildren.push(
            new docx.Table({
              width: { size: availWidth, type: docx.WidthType.DXA },
              indent: { size: 450, type: docx.WidthType.DXA },
              columnWidths: Array(numOpts).fill(colWidth),
              borders: {
                top: { style: docx.BorderStyle.NONE },
                bottom: { style: docx.BorderStyle.NONE },
                left: { style: docx.BorderStyle.NONE },
                right: { style: docx.BorderStyle.NONE },
                insideHorizontal: { style: docx.BorderStyle.NONE },
                insideVertical: { style: docx.BorderStyle.NONE }
              },
              rows: [new docx.TableRow({ cantSplit: true, children: cells })]
            })
          );
        } else {
          const optRows = [];
          for (let i = 0; i < q.options.length; i += 2) {
            const leftLetter = String.fromCharCode(65 + i);
            const leftOpt = q.options[i];
            const rightLetter = i + 1 < q.options.length ? String.fromCharCode(65 + i + 1) : '';
            const rightOpt = i + 1 < q.options.length ? q.options[i + 1] : null;

            const leftChildren = createOptionParagraphs(leftLetter, leftOpt);
            const rightChildren = rightLetter ? createOptionParagraphs(rightLetter, rightOpt) : [new docx.Paragraph({ children: [] })];

            optRows.push(
              new docx.TableRow({
                cantSplit: true,
                children: [
                  new docx.TableCell({
                    width: { size: 4275, type: docx.WidthType.DXA },
                    borders: {
                      top: { style: docx.BorderStyle.NONE },
                      bottom: { style: docx.BorderStyle.NONE },
                      left: { style: docx.BorderStyle.NONE },
                      right: { style: docx.BorderStyle.NONE }
                    },
                    children: leftChildren
                  }),
                  new docx.TableCell({
                    width: { size: 4275, type: docx.WidthType.DXA },
                    borders: {
                      top: { style: docx.BorderStyle.NONE },
                      bottom: { style: docx.BorderStyle.NONE },
                      left: { style: docx.BorderStyle.NONE },
                      right: { style: docx.BorderStyle.NONE }
                    },
                    children: rightChildren
                  })
                ]
              })
            );
          }

          headerChildren.push(
            new docx.Table({
              width: { size: 8550, type: docx.WidthType.DXA },
              indent: { size: 450, type: docx.WidthType.DXA },
              columnWidths: [4275, 4275],
              borders: {
                top: { style: docx.BorderStyle.NONE },
                bottom: { style: docx.BorderStyle.NONE },
                left: { style: docx.BorderStyle.NONE },
                right: { style: docx.BorderStyle.NONE },
                insideHorizontal: { style: docx.BorderStyle.NONE },
                insideVertical: { style: docx.BorderStyle.NONE }
              },
              rows: optRows
            })
          );
        }
      } else if (hasEssaySubs) {
          for (let sqIdx = 0; sqIdx < q.subQuestions.length; sqIdx++) {
            const sq = q.subQuestions[sqIdx];
            const sqLabel = sq.label || `(${String.fromCharCode(97 + (sqIdx % 26))})`;
            const sqCleanText = (sq.text || '').trimEnd();
            const sqLinesText = sqCleanText ? sqCleanText.split('\n') : [''];
            const sqLinesCount = (!metadata.separateAnswerSheet)
              ? ((sq.blankLines !== undefined && sq.blankLines !== '' && sq.blankLines !== null)
                ? Math.max(0, parseInt(sq.blankLines, 10) || 0)
                : 4)
              : 0;
            const sqHasImage = Boolean(sq.image);
            const sqHasFollowup = sqHasImage || sqLinesCount > 0;

            for (let slIdx = 0; slIdx < sqLinesText.length; slIdx++) {
              const isLastSqLine = slIdx === sqLinesText.length - 1;
              const keepNextSq = isLastSqLine ? sqHasFollowup : true;

              if (slIdx === 0) {
                headerChildren.push(
                  new docx.Paragraph({
                    keepNext: keepNextSq,
                    alignment: docx.AlignmentType.JUSTIFY,
                    indent: { left: 360 },
                    spacing: { before: 40, after: 20 },
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
                        size: 28
                      }),
                      ...docxTextRunsWithMath(sqLinesText[0] || '', { size: 28 }),
                      new docx.TextRun({
                        text: `\t(${formatMarks(sq.marks)} M)`,
                        italic: true,
                        size: 28
                      })
                    ]
                  })
                );
              } else {
                headerChildren.push(
                  new docx.Paragraph({
                    keepNext: keepNextSq,
                    indent: { left: 720 },
                    spacing: { after: 20 },
                    children: docxTextRunsWithMath(sqLinesText[slIdx], { size: 28 })
                  })
                );
              }
            }

            if (sqHasImage) {
              const sqImageBytes = await imageToUint8Array(sq.image);
              if (sqImageBytes) {
                headerChildren.push(
                  new docx.Paragraph({
                    keepNext: sqLinesCount > 0,
                    indent: { left: 720 },
                    spacing: { before: 80, after: 80 },
                    children: [
                      new docx.ImageRun({
                        data: sqImageBytes,
                        transformation: {
                          width: Number(sq.imageWidth) || 300,
                          height: Number(sq.imageHeight) || 200
                        }
                      })
                    ]
                  })
                );
              }
            }

            if (sqLinesCount > 0) {
              for (let i = 0; i < sqLinesCount; i++) {
                headerChildren.push(
                  new docx.Paragraph({
                    indent: { left: 360 },
                    tabStops: [
                      {
                        type: docx.TabStopType.RIGHT,
                        position: docx.TabStopPosition.MAX
                      }
                    ],
                    spacing: { before: 100, after: 100 },
                    children: [
                      new docx.TextRun({
                        text: '\t',
                        underline: {
                          type: docx.UnderlineType.SINGLE,
                          color: '777777'
                        }
                      })
                    ]
                  })
                );
              }
            }
          }
        }

        else if (hasTF) {
          headerChildren.push(
            new docx.Paragraph({
              keepNext: hasBlankLines,
              indent: { left: 450 },
              spacing: { before: 80, after: 120 },
              children: [
                new docx.TextRun({
                  text: '[    ] True        [    ] False',
                  bold: true,
                  size: 26
                })
              ]
            })
          );
        }

        else if (hasMatch) {
          const columnA = q.matchPairs.map(p => {
            if (typeof p === 'string') {
              const parts = p.split('=');
              return { text: parts[0] || p, image: '' };
            }
            return { text: p.premise || p.text || '', image: p.premiseImage || '' };
          });

          const shuffledList = getShuffledList(q);
          const columnB = q.matchPairs.map((pair, pIdx) => {
            const itemB = shuffledList[pIdx] || (typeof pair === 'string' ? { response: pair } : pair);
            if (typeof itemB === 'string') {
              const parts = itemB.split('=');
              return { text: parts.length > 1 ? parts[1] : parts[0], image: '' };
            }
            return {
              text: itemB.response || itemB.text || '',
              image: itemB.responseImage || ''
            };
          });

          const tableRows = [];

          tableRows.push(
            new docx.TableRow({
              tableHeader: true,
              cantSplit: true,
              children: [
                new docx.TableCell({
                  width: { size: 4275, type: docx.WidthType.DXA },
                  borders: {
                    top: { style: docx.BorderStyle.NONE },
                    bottom: { style: docx.BorderStyle.SINGLE, size: 4, color: '000000' },
                    left: { style: docx.BorderStyle.NONE },
                    right: { style: docx.BorderStyle.NONE }
                  },
                  children: [
                    new docx.Paragraph({
                      spacing: { before: 60, after: 60 },
                      children: [new docx.TextRun({ text: 'Column A', bold: true, size: 30 })]
                    })
                  ]
                }),
                new docx.TableCell({
                  width: { size: 4275, type: docx.WidthType.DXA },
                  borders: {
                    top: { style: docx.BorderStyle.NONE },
                    bottom: { style: docx.BorderStyle.SINGLE, size: 4, color: '000000' },
                    left: { style: docx.BorderStyle.NONE },
                    right: { style: docx.BorderStyle.NONE }
                  },
                  children: [
                    new docx.Paragraph({
                      spacing: { before: 60, after: 60 },
                      children: [new docx.TextRun({ text: 'Column B', bold: true, size: 30 })]
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
                  new docx.TextRun({ text: `${index + 1}. `, bold: true, size: 26 }),
                  ...docxTextRunsWithMath(itemA.text || '', { size: 26 })
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
                  new docx.TextRun({ text: `${romanNum(index)}. `, bold: true, size: 26 }),
                  ...docxTextRunsWithMath(itemB.text || '', { size: 26 })
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
                cantSplit: true,
                children: [
                  new docx.TableCell({
                    width: { size: 4275, type: docx.WidthType.DXA },
                    borders: {
                      top: { style: docx.BorderStyle.NONE },
                      bottom: { style: docx.BorderStyle.NONE },
                      left: { style: docx.BorderStyle.NONE },
                      right: { style: docx.BorderStyle.NONE }
                    },
                    children: cellAChildren
                  }),
                  new docx.TableCell({
                    width: { size: 4275, type: docx.WidthType.DXA },
                    borders: {
                      top: { style: docx.BorderStyle.NONE },
                      bottom: { style: docx.BorderStyle.NONE },
                      left: { style: docx.BorderStyle.NONE },
                      right: { style: docx.BorderStyle.NONE }
                    },
                    children: cellBChildren
                  })
                ]
              })
            );
          }

          headerChildren.push(
            new docx.Table({
              width: { size: 8550, type: docx.WidthType.DXA },
              indent: { size: 450, type: docx.WidthType.DXA },
              columnWidths: [4275, 4275],
              borders: {
                top: { style: docx.BorderStyle.NONE },
                bottom: { style: docx.BorderStyle.NONE },
                left: { style: docx.BorderStyle.NONE },
                right: { style: docx.BorderStyle.NONE },
                insideHorizontal: { style: docx.BorderStyle.NONE },
                insideVertical: { style: docx.BorderStyle.NONE }
              },
              rows: tableRows
            })
          );
        }

        if (hasTable) {
          const tblRows = [];
          const headerCells = q.tableData.headers.map(h =>
            new docx.TableCell({
              shading: { fill: 'E8E8E8' },
              children: [
                new docx.Paragraph({
                  spacing: { before: 40, after: 40 },
                  children: docxTextRunsWithMath(h || '', { bold: true, size: 28 })
                })
              ]
            })
          );
          tblRows.push(new docx.TableRow({ tableHeader: true, cantSplit: true, children: headerCells }));
          q.tableData.rows.forEach(row => {
            const bodyCells = row.map(cell =>
              new docx.TableCell({
                children: [
                  new docx.Paragraph({
                    spacing: { before: 40, after: 40 },
                    children: docxTextRunsWithMath(cell || '', { size: 28 })
                  })
                ]
              })
            );
            tblRows.push(new docx.TableRow({ cantSplit: true, children: bodyCells }));
          });
          headerChildren.push(
            new docx.Table({
              width: { size: 8550, type: docx.WidthType.DXA },
              indent: { size: 450, type: docx.WidthType.DXA },
              rows: tblRows
            })
          );
        }

        if (hasBlankLines) {
          for (let i = 0; i < linesCount; i++) {
            headerChildren.push(
              new docx.Paragraph({
                tabStops: [
                  {
                    type: docx.TabStopType.RIGHT,
                    position: docx.TabStopPosition.MAX
                  }
                ],
                spacing: { before: 100, after: 100 },
                children: [
                  new docx.TextRun({
                    text: '\t',
                    underline: {
                      type: docx.UnderlineType.SINGLE,
                      color: '777777'
                    }
                  })
                ]
              })
            );
          }
        }
      }
    }

    const doc = new docx.Document({
      styles: {
        default: {
          document: {
            run: {
              font: getFontFamily(branding, metadata)
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
                    size: 18
                  }),
                  new docx.SimpleField("PAGE"),
                  new docx.TextRun({
                    text: ' of ',
                    size: 18
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

  export const triggerDocxExport = async (branding, metadata, sections, hasQuestions, hasBlankQuestions) => {
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
      const blob = await generateDocxBlob(branding, metadata, sections);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = generateExportFilename('docx', metadata);
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Error exporting DOCX:', err);
      alert('Failed to generate DOCX file.');
    }
  };

  export const triggerPdfExport = async (targetElement, metadata, hasQuestions, hasBlankQuestions, setIsPdfExporting) => {
    if (hasQuestions && !hasQuestions()) {
      alert("No questions added to export.");
      return;
    }
    if (hasBlankQuestions && hasBlankQuestions()) {
      if (!window.confirm("Some questions have empty text. Are you sure you want to export?")) {
        return;
      }
    }
    if (!targetElement) {
      alert("Exam paper element not found.");
      return;
    }
    if (setIsPdfExporting) setIsPdfExporting(true);
    try {
      await generatePdfExport(targetElement, metadata);
    } catch (err) {
      console.error('Error exporting PDF:', err);
      alert('Failed to generate PDF file.');
    } finally {
      if (setIsPdfExporting) setIsPdfExporting(false);
    }
  };
