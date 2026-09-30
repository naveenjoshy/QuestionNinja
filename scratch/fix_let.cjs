const fs = require('fs');
let content = fs.readFileSync('src/App.jsx', 'utf8');

content = content.replace(/const questionBoxes = /g, 'let questionBoxes = ');
content = content.replace(/const explicitBreakBoxes = /g, 'let explicitBreakBoxes = ');
content = content.replace(/const sectionHeaderBoxes = /g, 'let sectionHeaderBoxes = ');
content = content.replace(/const sectionInstructionBoxes = /g, 'let sectionInstructionBoxes = ');
content = content.replace(/const paragraphBoxes = /g, 'let paragraphBoxes = ');
content = content.replace(/const tableRowBoxes = /g, 'let tableRowBoxes = ');
content = content.replace(/const tableBoxes = /g, 'let tableBoxes = ');
content = content.replace(/const subItemBoxes = /g, 'let subItemBoxes = ');

fs.writeFileSync('src/App.jsx', content);
console.log("Changed const to let.");
