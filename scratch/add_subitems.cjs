const fs = require('fs');
let content = fs.readFileSync('src/App.jsx', 'utf8');

const oldSubItemBoxes = "const subItemBoxes = getBoxes('.katex-html, .katex-display, .paper-mcq-option, .math-line, .paper-mcq-options, .paper-image-container, .paper-formula-block, .paper-blank-line')";
const newSubItemBoxes = "const subItemBoxes = getBoxes('.katex-html, .katex-display, .paper-mcq-option, .math-line, .paper-mcq-options, .paper-image-container, .paper-formula-block, .paper-blank-line, .paper-subquestion-item, .paper-table-question tr, .paper-match-table tr, .paper-match-table .match-pair')";

content = content.replace(oldSubItemBoxes, newSubItemBoxes);

fs.writeFileSync('src/App.jsx', content);
console.log("Added sub-items to subItemBoxes.");
