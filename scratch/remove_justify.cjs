const fs = require('fs');
let content = fs.readFileSync('src/App.jsx', 'utf8');

content = content.replace(/textAlign: 'justify'/g, "textAlign: 'left'");

fs.writeFileSync('src/App.jsx', content);
console.log("Replaced justify with left.");
