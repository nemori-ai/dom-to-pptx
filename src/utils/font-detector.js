// src/utils/font-detector.js
// Font detection and auto-embedding utilities

/**
 * Traverses the target DOM and collects all unique font-family names used.
 */
export function getUsedFontFamilies(root) {
  const families = new Set();

  function addPrimary(fontFamily) {
    const primary = fontFamily.split(',')[0].trim().replace(/['"]/g, '');
    if (primary) families.add(primary);
  }

  function scan(node) {
    if (node.nodeType === 1) {
      // Element
      const style = window.getComputedStyle(node);
      addPrimary(style.fontFamily);

      // Pseudo-elements may use different fonts (e.g. CSS counters with Orbitron)
      for (const pseudo of ['::before', '::after']) {
        const ps = window.getComputedStyle(node, pseudo);
        if (ps.content && ps.content !== 'none' && ps.content !== 'normal') {
          addPrimary(ps.fontFamily);
        }
      }
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
 * For Google Fonts: extract family names from <link> tags and @import rules,
 * then download full TTF from the google/fonts GitHub repo (CORS-friendly).
 * Returns a Map of familyName -> { buffer: ArrayBuffer, type: 'ttf' }.
 */
async function resolveGoogleFontsFullTTF(usedFamilies) {
  const result = new Map();
  const gfFamilies = new Set();

  // Check <link> tags
  for (const link of document.querySelectorAll('link[href*="fonts.googleapis.com"]')) {
    const href = link.getAttribute('href');
    if (!href) continue;
    for (const m of href.matchAll(/family=([^&:]+)/g)) {
      const name = decodeURIComponent(m[1]).replace(/\+/g, ' ');
      if (usedFamilies.has(name)) gfFamilies.add(name);
    }
  }

  // Check @import in stylesheets
  for (const sheet of document.styleSheets) {
    try {
      const rules = sheet.cssRules || sheet.rules;
      if (!rules) continue;
      for (const rule of rules) {
        if (rule.type === 3 && rule.href && rule.href.includes('fonts.googleapis.com')) {
          for (const m of rule.href.matchAll(/family=([^&:]+)/g)) {
            const name = decodeURIComponent(m[1]).replace(/\+/g, ' ');
            if (usedFamilies.has(name)) gfFamilies.add(name);
          }
        }
      }
    } catch {
      // CORS
    }
  }

  // Download full TTF from Google Fonts GitHub repo in parallel.
  // URL pattern: https://raw.githubusercontent.com/google/fonts/main/{license}/{dir}/{Name}-Regular.ttf
  // Also tries variable font pattern: {Name}[wght].ttf for fonts without static Regular.
  // Directory is font name lowercased with spaces removed.
  // Falls back to apache/ and ufl/ directories if ofl/ fails.
  const GITHUB_BASE = 'https://raw.githubusercontent.com/google/fonts/main';
  const LICENSE_DIRS = ['ofl', 'apache', 'ufl'];
  const FETCH_TIMEOUT = 60000;

  const fetchWithTimeout = (url) => {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), FETCH_TIMEOUT);
    return fetch(url, { signal: controller.signal }).finally(() => clearTimeout(timer));
  };

  // Helper: check if a TTF buffer has an 'fvar' table (= variable font)
  const hasVarTable = (buf) => {
    try {
      const view = new DataView(buf);
      const n = view.getUint16(4);
      for (let i = 0; i < n; i++) {
        const off = 12 + i * 16;
        const tag =
          String.fromCharCode(view.getUint8(off), view.getUint8(off + 1),
            view.getUint8(off + 2), view.getUint8(off + 3));
        if (tag === 'fvar') return true;
      }
    } catch { /* ignore */ }
    return false;
  };

  await Promise.all(
    [...gfFamilies].map(async (family) => {
      const dirName = family.toLowerCase().replace(/\s+/g, '');
      const baseName = family.replace(/\s+/g, '');

      // 1. Try static Regular TTF from GitHub first
      for (const licDir of LICENSE_DIRS) {
        try {
          const url = `${GITHUB_BASE}/${licDir}/${dirName}/${baseName}-Regular.ttf`;
          const resp = await fetchWithTimeout(url);
          if (!resp.ok) continue;
          const buf = await resp.arrayBuffer();
          if (!hasVarTable(buf)) {
            result.set(family, { buffer: buf, type: 'ttf' });
            return;
          }
        } catch { /* next */ }
      }

      // 2. Variable TTF from GitHub (will be converted to static in fontToEot)
      for (const licDir of LICENSE_DIRS) {
        try {
          const url = `${GITHUB_BASE}/${licDir}/${dirName}/${baseName}[wght].ttf`;
          const resp = await fetchWithTimeout(url);
          if (!resp.ok) continue;
          const buf = await resp.arrayBuffer();
          result.set(family, { buffer: buf, type: 'ttf' });
          return;
        } catch { /* next */ }
      }

      console.warn(`Could not find font for "${family}"`);
    })
  );

  return result;
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

  // Skip library-internal fonts that are not useful in PPTX output.
  // Math renderers (MathJax, KaTeX) inject @font-face rules for their engines.
  // These get rasterized as images during conversion, so embedding the font files
  // is unnecessary — and some (e.g. MJXZERO) are CFF-outline WOFFs that
  // fonteditor-core cannot convert.
  const isLibraryInternalFont = (name) => /^MJX|^MathJax[_-]|^KaTeX[_-]/i.test(name);

  const processedFamilies = new Set();

  for (const sheet of Array.from(document.styleSheets)) {
    try {
      // Accessing cssRules on cross-origin sheets (like Google Fonts) might fail
      // if CORS headers aren't set. We wrap in try/catch.
      const rules = sheet.cssRules || sheet.rules;
      if (!rules) continue;

      for (const rule of Array.from(rules)) {
        if (rule.constructor.name === 'CSSFontFaceRule' || rule.type === 5) {
          const familyName = rule.style.getPropertyValue('font-family').replace(/['"]/g, '').trim();

          // Only embed one file per font family (Google Fonts returns multiple
          // unicode-range subsets — we need the full font, not subsets).
          if (usedFamilies.has(familyName) && !processedFamilies.has(familyName) && !isLibraryInternalFont(familyName)) {
            const src = rule.style.getPropertyValue('src');
            const url = extractUrl(src);

            if (url && !processedUrls.has(url)) {
              processedUrls.add(url);
              processedFamilies.add(familyName);
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

  // For Google Fonts: replace WOFF2 subset URL with full TTF from GitHub.
  const googleFullFonts = await resolveGoogleFontsFullTTF(usedFamilies);
  for (const font of foundFonts) {
    const fullFont = googleFullFonts.get(font.name);
    if (fullFont && fullFont.buffer) {
      font.buffer = fullFont.buffer;
      font.type = fullFont.type;
    }
  }

  // Add Google Fonts that weren't found via @font-face scanning (e.g. CORS blocked)
  for (const [name, data] of googleFullFonts) {
    if (data.buffer && !foundFonts.some((f) => f.name === name)) {
      foundFonts.push({ name, buffer: data.buffer, type: data.type });
    }
  }

  return foundFonts;
}
