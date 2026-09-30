const fs = require('fs');
let content = fs.readFileSync('src/App.jsx', 'utf8');

// Replace + 15 with + 30
content = content.replace(/\(sh\.bottom \+ 15\) > limit/g, '(sh.bottom + 30) > limit');
content = content.replace(/\(q\.bottom \+ 15\) > limit/g, '(q.bottom + 30) > limit');
content = content.replace(/\(si\.bottom \+ 15\) > limit/g, '(si.bottom + 30) > limit');
content = content.replace(/\(b\.bottom \+ 15\) > limit/g, '(b.bottom + 30) > limit');
content = content.replace(/\(el\.bottom \+ 15\) > breakY/g, '(el.bottom + 30) > breakY');

// Adjust Math.max fallbacks to leave more top margin
content = content.replace(/Math\.max\(afterY \+ 20, [^ ]+\.top - 10\)/g, (match) => {
    return match.replace('- 10', '- 25');
});
content = content.replace(/Math\.max\(currentY \+ 20, cutElem\.top - 10\)/g, 'Math.max(currentY + 20, cutElem.top - 25)');

// Add .katex-display and .katex-html to subItemBoxes
content = content.replace(/const subItemBoxes = getBoxes\('\.paper-mcq-option, /g, "const subItemBoxes = getBoxes('.katex-html, .katex-display, .paper-mcq-option, ");

fs.writeFileSync('src/App.jsx', content);
console.log("Increased buffers and margins.");
