const fs = require('fs');
let content = fs.readFileSync('src/App.jsx', 'utf8');

content = content.replace(/if \(dRect\.height > 0 && dRect\.width > 0\) \{/g, 'if (dRect.height > 0) {');

fs.writeFileSync('src/App.jsx', content);
console.log("Removed width check.");
