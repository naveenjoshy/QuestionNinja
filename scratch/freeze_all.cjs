const fs = require('fs');
let content = fs.readFileSync('src/App.jsx', 'utf8');

// 1. Add paragraphBoxes and subItemBoxes to freezeBoxes
const oldFreezeBoxes = `      const freezeBoxes = [
        ...questionBoxes,
        ...sectionHeaderBoxes,
        ...sectionInstructionBoxes
      ];`;
const newFreezeBoxes = `      const freezeBoxes = [
        ...questionBoxes,
        ...sectionHeaderBoxes,
        ...sectionInstructionBoxes,
        ...paragraphBoxes,
        ...subItemBoxes
      ];`;
content = content.replace(oldFreezeBoxes, newFreezeBoxes);

// 2. Increase safety buffer from 30 to 45
content = content.replace(/\(sh\.bottom \+ 30\) > limit/g, '(sh.bottom + 45) > limit');
content = content.replace(/\(q\.bottom \+ 30\) > limit/g, '(q.bottom + 45) > limit');
content = content.replace(/\(si\.bottom \+ 30\) > limit/g, '(si.bottom + 45) > limit');
content = content.replace(/\(b\.bottom \+ 30\) > limit/g, '(b.bottom + 45) > limit');
content = content.replace(/\(el\.bottom \+ 30\) > breakY/g, '(el.bottom + 45) > breakY');

// 3. Make sure Math.max fallbacks also give a bit more top margin (change -25 to -35)
content = content.replace(/Math\.max\(afterY \+ 20, [^ ]+\.top - 25\)/g, (match) => {
    return match.replace('- 25', '- 35');
});
content = content.replace(/Math\.max\(currentY \+ 20, cutElem\.top - 25\)/g, 'Math.max(currentY + 20, cutElem.top - 35)');

fs.writeFileSync('src/App.jsx', content);
console.log("Applied ultimate freeze and buffer.");
