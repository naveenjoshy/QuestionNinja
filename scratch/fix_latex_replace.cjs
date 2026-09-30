const fs = require('fs');
let content = fs.readFileSync('src/App.jsx', 'utf8');

const oldGreeksLoop = `  for (const [k, v] of Object.entries(greeks)) {
    text = text.replaceAll(k, v);
  }`;

const newGreeksLoop = `  const sortedGreeks = Object.keys(greeks).sort((a, b) => b.length - a.length);
  for (const k of sortedGreeks) {
    text = text.replaceAll(k, greeks[k]);
  }`;

const oldOpsLoop = `  for (const [k, v] of Object.entries(ops)) {
    text = text.replaceAll(k, v);
  }`;

const newOpsLoop = `  const sortedOps = Object.keys(ops).sort((a, b) => b.length - a.length);
  for (const k of sortedOps) {
    text = text.replaceAll(k, ops[k]);
  }`;

content = content.replace(oldGreeksLoop, newGreeksLoop);
content = content.replace(oldOpsLoop, newOpsLoop);

fs.writeFileSync('src/App.jsx', content);
