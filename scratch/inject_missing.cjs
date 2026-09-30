const fs = require('fs');
const appContent = fs.readFileSync('src/App.jsx', 'utf8');
const missingContent = fs.readFileSync('scratch/missing_code.js', 'utf8');

const newAppContent = appContent.replace(
  'const _PRESET_TEMPLATES = {',
  missingContent + '\nconst _PRESET_TEMPLATES = {'
);

fs.writeFileSync('src/App.jsx', newAppContent);
console.log('Restored missing code!');
