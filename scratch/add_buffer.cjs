const fs = require('fs');
let content = fs.readFileSync('src/App.jsx', 'utf8');

// Add a 15px safety buffer to all bottom checks to account for html2canvas rendering differences
content = content.replace(/sh\.bottom > limit/g, '(sh.bottom + 15) > limit');
content = content.replace(/q\.bottom > limit/g, '(q.bottom + 15) > limit');
content = content.replace(/si\.bottom > limit/g, '(si.bottom + 15) > limit');
content = content.replace(/b\.bottom > limit/g, '(b.bottom + 15) > limit');
content = content.replace(/el\.bottom > breakY/g, '(el.bottom + 15) > breakY');

fs.writeFileSync('src/App.jsx', content);
console.log("Replaced bounds checks.");
