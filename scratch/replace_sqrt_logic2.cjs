const fs = require('fs');
let content = fs.readFileSync('src/App.jsx', 'utf8');

// Replace everything between `result = result.replace(/\\sqrt\[([^\]]*)\]\{([^}]*)\}/g` and `result = result.replace(/√/g, '<span class="math-sqrt-symbol">√</span>');`
const startIndex = content.indexOf('result = result.replace(/\\\\sqrt\\[([^\\]]*)\\]\\{([^}]*)\\}/g');
const endIndex = content.indexOf('// 4. Convert Unicode cube root ∛ to styled span');

if (startIndex !== -1 && endIndex !== -1) {
  const replacement = `result = result.replace(/\\\\sqrt\\[([^\\]]*)\\]\\{([^}]*)\\}/g, (_, degree, inner) => \`\${degree}√\${inner}\`);
  result = result.replace(/\\\\sqrt\\{([^}]*)\\}/g, (_, inner) => renderSquareRootHTML(inner));
  // Removed √ grouping replacements per user request to treat it as a regular character
  result = result.replace(/√/g, '<span class="math-sqrt-symbol">√</span>');

  `;
  content = content.substring(0, startIndex) + replacement + content.substring(endIndex);
}

const docxStartIndex = content.indexOf(`if (!processedText.includes('$') && (processedText.includes('√') || processedText.includes('\\\\sqrt'))) {`);
const docxEndIndex = content.indexOf(`const parts = processedText.split('$');`);

if (docxStartIndex !== -1 && docxEndIndex !== -1) {
  const replacement = `if (!processedText.includes('$') && (processedText.includes('\\\\sqrt'))) {
    processedText = processedText.replace(/\\\\sqrt\\{([^}]*)\\}/g, '$\\\\sqrt{$1}$');
  }
  // Removed √ grouping replacements for docx per user request

  `;
  content = content.substring(0, docxStartIndex) + replacement + content.substring(docxEndIndex);
}

// And also cube roots
content = content.replace(/result = result\.replace\(\/∛\\s\*\\(\(\[\^\)\]\+\)\\)\/g,[\s\S]*?result = result\.replace\(\/∛\/g, '<span class="math-sqrt-symbol">∛<\/span>'\);/g, `result = result.replace(/∛/g, '<span class="math-sqrt-symbol">∛</span>');`);


fs.writeFileSync('src/App.jsx', content);
