/* eslint-disable @typescript-eslint/no-var-requires */
const fs = require('fs');
const path = require('path');

const filesToPatch = [
  path.join(__dirname, '..', 'node_modules', '@microsoft', 'sp-css-loader', 'node_modules', '@microsoft', 'load-themed-styles', 'lib-es6', 'index.js'),
  path.join(__dirname, '..', 'node_modules', '@microsoft', 'sp-css-loader', 'node_modules', '@microsoft', 'load-themed-styles', 'lib', 'index.js'),
  path.join(__dirname, '..', 'release', 'assets', 'custom-faq-web-part.js')
];

const safeRegexLine = "var _themeTokenRegex = /[\\'\\\"]\\[theme:\\s*([A-Za-z0-9_-]+)\\s*(?:,\\s*default:\\s*([^\\]\\r\\n]+?))?\\s*\\][\\'\\\"]/g;";
const themeTokenPattern = /var _themeTokenRegex = \/[^;]+;/;

filesToPatch.forEach((filePath) => {
  if (!fs.existsSync(filePath)) {
    return;
  }

  const original = fs.readFileSync(filePath, 'utf8');

  if (original.includes(safeRegexLine)) {
    return;
  }

  if (!themeTokenPattern.test(original)) {
    return;
  }

  const updated = original.replace(themeTokenPattern, safeRegexLine);

  if (updated !== original) {
    fs.writeFileSync(filePath, updated, 'utf8');
    console.log(`Patched theme token regex in ${path.relative(path.join(__dirname, '..'), filePath)}`);
  }
});
