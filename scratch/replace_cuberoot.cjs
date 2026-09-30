const fs = require('fs');
let content = fs.readFileSync('src/App.jsx', 'utf8');

const startIndex = content.indexOf('// 4. Convert Unicode cube root ∛ to styled span');
const endIndex = content.indexOf('// 5. Restore rendered $...$ math blocks');

if (startIndex !== -1 && endIndex !== -1) {
  content = content.substring(0, startIndex) + `// 4. Convert Unicode cube root ∛ to styled span
  result = result.replace(/∛/g, '<span class="math-sqrt-symbol">∛</span>');

  ` + content.substring(endIndex);
}

fs.writeFileSync('src/App.jsx', content);
