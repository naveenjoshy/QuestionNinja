const fs = require('fs');
let content = fs.readFileSync('src/App.jsx', 'utf8');

const replacement = `
      // Use getBoundingClientRect for absolute pixel-perfect accuracy
      const cloneRect = clone.getBoundingClientRect();
      
      const getTrueBounds = (elem) => {
        let rect = elem.getBoundingClientRect();
        let minTop = rect.top;
        let maxBottom = rect.bottom;
        
        const descendants = elem.getElementsByTagName('*');
        for (let i = 0; i < descendants.length; i++) {
          const dRect = descendants[i].getBoundingClientRect();
          if (dRect.height > 0 && dRect.width > 0) {
            if (dRect.top < minTop) minTop = dRect.top;
            if (dRect.bottom > maxBottom) maxBottom = dRect.bottom;
          }
        }
        return {
          top: minTop - cloneRect.top,
          bottom: maxBottom - cloneRect.top,
          height: maxBottom - minTop
        };
      };

      const getBoxes = (selector) => {
        return Array.from(clone.querySelectorAll(selector)).map(elem => {
          const bounds = getTrueBounds(elem);
          return { elem, top: bounds.top, bottom: bounds.bottom, height: bounds.height };
        }).filter(b => b.height > 0).sort((a, b) => a.top - b.top);
      };
`;

content = content.replace(
  /\/\/ Use getBoundingClientRect for absolute pixel-perfect accuracy[\s\S]*?\}\)\.filter\(b => b\.height > 0\)\.sort\(\(a, b\) => a\.top - b\.top\);\n\s*\};/,
  replacement.trim()
);

fs.writeFileSync('src/App.jsx', content);
