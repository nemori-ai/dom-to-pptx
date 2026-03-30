// src/utils/font-detector.js
// Font detection and auto-embedding utilities

/**
 * Traverses the target DOM and collects all unique font-family names used.
 */
export function getUsedFontFamilies(root) {
  const families = new Set();

  function scan(node) {
    if (node.nodeType === 1) {
      // Element
      const style = window.getComputedStyle(node);
      const fontList = style.fontFamily.split(',');
      // The first font in the stack is the primary one
      const primary = fontList[0].trim().replace(/['"]/g, '');
      if (primary) families.add(primary);
    }
    for (const child of node.childNodes) {
      scan(child);
    }
  }

  // Handle array of roots or single root
  const elements = Array.isArray(root) ? root : [root];
  elements.forEach((el) => {
    const node = typeof el === 'string' ? document.querySelector(el) : el;
    if (node) scan(node);
  });

  return families;
}

/**
 * Scans document.styleSheets to find @font-face URLs for the requested families.
 * Returns an array of { name, url } objects.
 */
export async function getAutoDetectedFonts(usedFamilies) {
  const foundFonts = [];
  const processedUrls = new Set();

  // Helper to extract clean URL from CSS src string
  const extractUrl = (srcStr) => {
    // Look for url("...") or url('...') or url(...)
    // Supports all formats: woff2, woff, ttf, otf
    const matches = srcStr.match(/url\((['"]?)(.*?)\1\)/g);
    if (!matches) return null;

    let chosenUrl = null;
    for (const match of matches) {
      const urlRaw = match.replace(/url\((['"]?)(.*?)\1\)/, '$2');
      if (urlRaw.startsWith('data:')) continue;

      const ext = urlRaw.split('.').pop().split(/[?#]/)[0].toLowerCase();
      if (['woff2', 'woff', 'ttf', 'otf'].includes(ext)) {
        chosenUrl = urlRaw;
        break;
      }
      if (!chosenUrl) chosenUrl = urlRaw;
    }
    return chosenUrl;
  };

  for (const sheet of Array.from(document.styleSheets)) {
    try {
      // Accessing cssRules on cross-origin sheets (like Google Fonts) might fail
      // if CORS headers aren't set. We wrap in try/catch.
      const rules = sheet.cssRules || sheet.rules;
      if (!rules) continue;

      for (const rule of Array.from(rules)) {
        if (rule.constructor.name === 'CSSFontFaceRule' || rule.type === 5) {
          const familyName = rule.style.getPropertyValue('font-family').replace(/['"]/g, '').trim();

          if (usedFamilies.has(familyName)) {
            const src = rule.style.getPropertyValue('src');
            const url = extractUrl(src);

            if (url && !processedUrls.has(url)) {
              processedUrls.add(url);
              foundFonts.push({ name: familyName, url: url });
            }
          }
        }
      }
    } catch (e) {
      // SecurityError is common for external stylesheets (CORS).
      // We cannot scan those automatically via CSSOM.
      console.warn('error:', e);
      console.warn('Cannot scan stylesheet for fonts (CORS restriction):', sheet.href);
    }
  }

  return foundFonts;
}
