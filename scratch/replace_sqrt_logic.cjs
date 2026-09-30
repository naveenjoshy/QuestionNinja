const fs = require('fs');
let content = fs.readFileSync('src/App.jsx', 'utf8');

content = content.replace(/result = result\.replace\(\/\\sqrt\\\{([^}]*)\\\}\/g, \(_, inner\) => renderSquareRootHTML\(inner\)\);\n\s*result = result\.replace\(\/√\\s\*\\(\(\[\^\)\]\+\)\\)\/g, \(_, inner\) => renderSquareRootHTML\(\`\(\$\{inner\}\)\`\)\);\n\s*result = result\.replace\(\/√\\s\*\\\{([^}]+)\\\}\/g, \(_, inner\) => renderSquareRootHTML\(inner\)\);\n\s*result = result\.replace\(\/√\\s\*\(\[0-9a-zA-Z\]\+\)\/g, \(_, inner\) => renderSquareRootHTML\(inner\)\);/g, `// Removed √ grouping replacements per user request to treat it as a regular character`);

// Also update docx export logic
content = content.replace(/processedText = processedText\.replace\(\/\\sqrt\\\{([^}]*)\\\}\/g, '\\$\\sqrt\{\\$1\}\\$'\);\n\s*processedText = processedText\.replace\(\/√\\s\*\\(\(\[\^\)\]\+\)\\)\/g, '\\$\\sqrt\{\\$1\}\\$'\);\n\s*processedText = processedText\.replace\(\/√\\s\*\\\{([^}]+)\\\}\/g, '\\$\\sqrt\{\\$1\}\\$'\);\n\s*processedText = processedText\.replace\(\/√\\s\*\(\[0-9a-zA-Z\]\+\)\/g, '\\$\\sqrt\{\\$1\}\\$'\);/g, `// Removed √ grouping replacements for docx per user request`);

fs.writeFileSync('src/App.jsx', content);
