/* eslint-disable @typescript-eslint/no-var-requires */
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');

const filesToPatch = [
  {
    type: 'esm',
    filePath: path.join(root, 'node_modules', '@microsoft', 'sp-css-loader', 'node_modules', '@microsoft', 'load-themed-styles', 'lib-es6', 'index.js')
  },
  {
    type: 'cjs',
    filePath: path.join(root, 'node_modules', '@microsoft', 'sp-css-loader', 'node_modules', '@microsoft', 'load-themed-styles', 'lib', 'index.js')
  },
  {
    type: 'cjs',
    filePath: path.join(root, 'release', 'assets', 'custom-faq-web-part.js')
  }
];

const regexPattern = /var _themeTokenRegex = \/[^;]+;/;

filesToPatch.forEach(({ type, filePath }) => {
  if (!fs.existsSync(filePath)) {
    return;
  }

  const originalContent = fs.readFileSync(filePath, 'utf8');
  let updatedContent = originalContent;

  if (regexPattern.test(updatedContent)) {
    updatedContent = updatedContent.replace(regexPattern, 'var _themeTokenRegex = null; // Replaced with safe parser implementation');
  }

  updatedContent = replaceSplitStyles(updatedContent, type === 'esm');

  if (updatedContent !== originalContent) {
    fs.writeFileSync(filePath, updatedContent, 'utf8');
    console.log(`Hardened theming parser in ${path.relative(root, filePath)}`);
  }
});

function replaceSplitStyles(content, isEsm) {
  const signature = isEsm ? 'export function splitStyles(styles)' : 'function splitStyles(styles)';
  const startIndex = content.indexOf(signature);
  if (startIndex === -1) {
    return content;
  }

  const endToken = '\nfunction registerStyles';
  const endIndex = content.indexOf(endToken, startIndex);
  if (endIndex === -1) {
    return content;
  }

  const safeFunction = buildSafeSplitFunction(isEsm);
  return content.slice(0, startIndex) + safeFunction + '\n' + content.slice(endIndex);
}

function buildSafeSplitFunction(isEsm) {
  const header = isEsm ? 'export function splitStyles(styles) {' : 'function splitStyles(styles) {';
  return `${header}
    return _safeSplitStyles(styles);
}
function _safeSplitStyles(styles) {
    var result = [];
    if (!styles) {
        return result;
    }
    var pos = 0;
    var len = styles.length;
    while (pos < len) {
        var token = _findNextThemeToken(styles, pos);
        if (!token) {
            if (pos < len) {
                result.push({
                    rawString: styles.substring(pos)
                });
            }
            break;
        }
        if (token.index > pos) {
            result.push({
                rawString: styles.substring(pos, token.index)
            });
        }
        result.push({
            theme: token.theme,
            defaultValue: token.defaultValue
        });
        pos = token.nextPos;
    }
    return result;
}
function _findNextThemeToken(styles, fromIndex) {
    var len = styles.length;
    var searchIndex = fromIndex;
    while (searchIndex < len) {
        var singleIdx = styles.indexOf("'[theme:", searchIndex);
        var doubleIdx = styles.indexOf('"[theme:', searchIndex);
        var start = -1;
        var quote = '';
        if (singleIdx === -1 && doubleIdx === -1) {
            return null;
        }
        if (singleIdx !== -1 && (doubleIdx === -1 || singleIdx <= doubleIdx)) {
            start = singleIdx;
            quote = "'";
        }
        else {
            start = doubleIdx;
            quote = '"';
        }
        var openBracketIndex = start + 1;
        var closeBracketIndex = styles.indexOf(']', openBracketIndex);
        if (closeBracketIndex === -1) {
            return null;
        }
        var closingQuoteIndex = closeBracketIndex + 1;
        if (closingQuoteIndex >= len || styles.charAt(closingQuoteIndex) !== quote) {
            searchIndex = openBracketIndex + 1;
            continue;
        }
        var tokenContent = styles.substring(openBracketIndex + 1, closeBracketIndex);
        var parsed = _parseThemeToken(tokenContent);
        if (!parsed) {
            searchIndex = closingQuoteIndex + 1;
            continue;
        }
        return {
            index: start,
            nextPos: closingQuoteIndex + 1,
            theme: parsed.theme,
            defaultValue: parsed.defaultValue
        };
    }
    return null;
}
function _parseThemeToken(content) {
    if (!content) {
        return null;
    }
    var themeMarker = 'theme:';
    var defaultMarker = 'default:';
    var themeIndex = content.indexOf(themeMarker);
    if (themeIndex === -1) {
        return null;
    }
    var afterTheme = content.substring(themeIndex + themeMarker.length);
    var commaIndex = afterTheme.indexOf(',');
    var themeSegment = commaIndex === -1 ? afterTheme : afterTheme.substring(0, commaIndex);
    var themeName = themeSegment.trim();
    if (!_isSafeThemeName(themeName)) {
        return null;
    }
    var defaultValue;
    var defaultIndex = content.indexOf(defaultMarker);
    if (defaultIndex !== -1) {
        defaultValue = _sanitizeDefaultValue(content.substring(defaultIndex + defaultMarker.length));
    }
    return {
        theme: themeName,
        defaultValue: defaultValue
    };
}
function _isSafeThemeName(name) {
    if (!name) {
        return false;
    }
    for (var i = 0; i < name.length; i++) {
        var ch = name.charAt(i);
        var isLetter = (ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z');
        var isDigit = ch >= '0' && ch <= '9';
        if (!isLetter && !isDigit && ch !== '_' && ch !== '-') {
            return false;
        }
    }
    return true;
}
function _sanitizeDefaultValue(value) {
    if (!value) {
        return undefined;
    }
    var trimmed = value.trim();
    if (!trimmed) {
        return undefined;
    }
    var firstChar = trimmed.charAt(0);
    var lastChar = trimmed.charAt(trimmed.length - 1);
    if ((firstChar === '"' && lastChar === '"') || (firstChar === "'" && lastChar === "'")) {
        trimmed = trimmed.substring(1, trimmed.length - 1);
    }
    if (trimmed.length > 512) {
        trimmed = trimmed.substring(0, 512);
    }
    return trimmed;
}
`;
}
