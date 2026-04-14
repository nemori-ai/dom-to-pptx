// src/utils/detect-fonts.js
// Canvas-based font detection for PPTX output.
// Detects the best-matching font for each OOXML script category
// (latin, ea, cs, symbol) by measuring text rendering similarity.
// Uses dual-fallback glyph coverage check to reject fonts lacking target script glyphs.
//
// Ported from detectFontsForPPTX.js (UMD) to ES module.

/**
 * Detect fonts suitable for PPTX output for a given HTML element.
 *
 * Returns 4 script categories matching OOXML font slots:
 * - latin: Latin/Western main font
 * - ea: East Asian (CJK) main font
 * - cs: Complex Script (Arabic, Hebrew, Devanagari, Thai, etc.)
 * - symbol: Math/symbol/arrow/formula font
 *
 * Uses Canvas measureText to find the closest-matching font from a
 * 3-tier candidate pool: current element → ancestors → Office fallback.
 * For ea/cs/symbol, applies dual-fallback glyph coverage check to reject
 * candidates that don't actually contain the target script's glyphs.
 *
 * @param {HTMLElement} element - DOM element to detect fonts for
 * @param {string|null} [customText=null] - Optional text override; defaults to element.textContent
 * @param {object} [options={}] - Optional configuration
 * @returns {{ latin: string, ea: string, cs: string, symbol: string, details: object }}
 */
export function detectFontsForPPTX(element, customText = null, options = {}) {
  if (!(element instanceof HTMLElement)) {
    throw new TypeError("detectFontsForPPTX: element must be an HTMLElement");
  }

  const config = {
    maxSampleLength: options.maxSampleLength ?? 32,
    currentThreshold: options.currentThreshold ?? 0.35,
    ancestorThreshold: options.ancestorThreshold ?? 0.75,
    preferOfficeFonts: options.preferOfficeFonts ?? true,
    symbolRatioThreshold: options.symbolRatioThreshold ?? 0.12,
    forceSymbolByFontFamily: options.forceSymbolByFontFamily ?? true,
  };

  const style = window.getComputedStyle(element);
  const text = String(customText ?? element.textContent ?? "").trim();

  const OFFICE_LATIN_FONTS = getOfficeLatinFonts();
  const OFFICE_EA_FONTS = getOfficeEAFonts();
  const OFFICE_CS_FONTS = getOfficeCSFonts();
  const OFFICE_SYMBOL_FONTS = getOfficeSymbolFonts();

  if (!text) {
    return {
      latin: "Arial",
      ea: "Microsoft YaHei",
      cs: "Arial",
      symbol: "Cambria Math",
      details: {
        text: "",
        samples: { latin: "", ea: "", cs: "", symbol: "" },
        currentFontStack: normalizeFontList(style.fontFamily),
        ancestorFontStack: [],
        scores: { latin: 0, ea: 0, cs: 0, symbol: 0 },
        stages: { latin: "office", ea: "office", cs: "office", symbol: "office" },
        scriptCounts: { latin: 0, ea: 0, cs: 0, symbol: 0, total: 0 },
        hasSignificantSymbolContent: false,
      },
    };
  }

  const ctx = options.ctx || getSharedCanvasContext();

  if (!ctx) {
    return {
      latin: "Arial",
      ea: "Microsoft YaHei",
      cs: "Arial",
      symbol: "Cambria Math",
      details: { text, error: "Canvas 2D context unavailable" },
    };
  }

  const baseStyle = buildCanvasFontBase(style);
  const currentFontStack = normalizeFontList(style.fontFamily);
  const ancestorFontStack = collectAncestorFontFamilies(element);
  const currentPlusAncestors = uniqueFonts([...currentFontStack, ...ancestorFontStack]);

  const scriptCounts = countScripts(text);
  const totalCount = Math.max(scriptCounts.total, 1);

  const latinSample = collectSample(text, isLatinLikeChar, "The quick brown fox jumps 123 ABC xyz", config.maxSampleLength);
  const eaSample = collectSample(text, isEastAsianChar, "测试漢字かなカナ한글，。、（）", config.maxSampleLength);
  const csSample = collectSample(text, isComplexScriptChar, "العربية עברית हिन्दी ไทย", config.maxSampleLength);
  const symbolSample = collectSample(text, isSymbolLikeChar, "∑∫√∞≈≠≤≥→⇒αβγθ", config.maxSampleLength);

  const symbolRatio = scriptCounts.symbol / totalCount;
  const hasMathLikeFontFamily = config.forceSymbolByFontFamily && containsMathOrSymbolFontName(currentPlusAncestors);
  const hasSignificantSymbolContent = symbolRatio >= config.symbolRatioThreshold || hasMathLikeFontFamily;

  const latinResult = resolveScriptFont({
    ctx, baseStyle, fullFontFamily: style.fontFamily, sample: latinSample,
    currentFonts: currentFontStack, ancestorFonts: currentPlusAncestors,
    officeFonts: OFFICE_LATIN_FONTS, currentThreshold: config.currentThreshold,
    ancestorThreshold: config.ancestorThreshold,
    normalizeResult: (name) => normalizeToPPTLatinFont(name, OFFICE_LATIN_FONTS, config.preferOfficeFonts),
  });

  const eaResult = resolveScriptFont({
    ctx, baseStyle, fullFontFamily: style.fontFamily, sample: eaSample,
    currentFonts: currentFontStack, ancestorFonts: currentPlusAncestors,
    officeFonts: OFFICE_EA_FONTS, currentThreshold: config.currentThreshold,
    ancestorThreshold: config.ancestorThreshold,
    normalizeResult: (name) => normalizeToPPTEAFont(name, OFFICE_EA_FONTS, config.preferOfficeFonts),
    checkGlyphCoverage: true,
  });

  const csResult = resolveScriptFont({
    ctx, baseStyle, fullFontFamily: style.fontFamily, sample: csSample,
    currentFonts: currentFontStack, ancestorFonts: currentPlusAncestors,
    officeFonts: OFFICE_CS_FONTS, currentThreshold: config.currentThreshold,
    ancestorThreshold: config.ancestorThreshold,
    normalizeResult: (name) => normalizeToPPTCSFont(name, OFFICE_CS_FONTS, config.preferOfficeFonts),
    checkGlyphCoverage: true,
  });

  let symbolResult;
  if (hasSignificantSymbolContent) {
    symbolResult = resolveScriptFont({
      ctx, baseStyle, fullFontFamily: style.fontFamily, sample: symbolSample,
      currentFonts: currentFontStack, ancestorFonts: currentPlusAncestors,
      officeFonts: OFFICE_SYMBOL_FONTS, currentThreshold: config.currentThreshold,
      ancestorThreshold: config.ancestorThreshold,
      normalizeResult: (name) => normalizeToPPTSymbolFont(name, OFFICE_SYMBOL_FONTS, config.preferOfficeFonts),
      checkGlyphCoverage: true,
    });
  } else {
    symbolResult = {
      font: latinResult.font, rawFont: latinResult.rawFont,
      score: latinResult.score, stage: "latin-fallback",
      candidates: latinResult.candidates.slice(),
    };
  }

  return {
    latin: latinResult.font,
    ea: eaResult.font,
    cs: csResult.font,
    symbol: symbolResult.font,
    details: {
      text,
      samples: { latin: latinSample, ea: eaSample, cs: csSample, symbol: symbolSample },
      currentFontStack, ancestorFontStack,
      candidates: { latin: latinResult.candidates, ea: eaResult.candidates, cs: csResult.candidates, symbol: symbolResult.candidates },
      scores: { latin: latinResult.score, ea: eaResult.score, cs: csResult.score, symbol: symbolResult.score },
      stages: { latin: latinResult.stage, ea: eaResult.stage, cs: csResult.stage, symbol: symbolResult.stage },
      rawFonts: { latin: latinResult.rawFont, ea: eaResult.rawFont, cs: csResult.rawFont, symbol: symbolResult.rawFont },
      scriptCounts, symbolRatio, hasMathLikeFontFamily, hasSignificantSymbolContent,
    },
  };
}

/* =========================================================
 * Script font resolution (3-stage progressive search)
 * ========================================================= */

function resolveScriptFont({
  ctx, baseStyle, fullFontFamily, sample, currentFonts, ancestorFonts,
  officeFonts, currentThreshold, ancestorThreshold, normalizeResult,
  checkGlyphCoverage,
}) {
  const actual = measureWithStack(ctx, baseStyle, fullFontFamily, sample);
  const glyphCheck = checkGlyphCoverage;
  const currentSet = new Set(currentFonts.map((f) => f.toLowerCase()));

  // Stage 1: current element font stack
  const stage1 = detectBestMatch(ctx, baseStyle, sample, actual, currentFonts, glyphCheck);
  if (stage1.font && stage1.score <= currentThreshold) {
    return { font: normalizeResult(stage1.font), rawFont: stage1.font, score: stage1.score, stage: "current", candidates: currentFonts.slice() };
  }

  // Stage 2: ancestor fonts (incremental)
  const ancestorOnly = ancestorFonts.filter((f) => !currentSet.has(f.toLowerCase()));
  const stage2Incremental = detectBestMatch(ctx, baseStyle, sample, actual, ancestorOnly, glyphCheck);
  const stage2 = stage2Incremental.score < stage1.score ? stage2Incremental : stage1;

  if (stage2.font && stage2.score <= ancestorThreshold) {
    return { font: normalizeResult(stage2.font), rawFont: stage2.font, score: stage2.score, stage: "ancestor", candidates: uniqueFonts([...ancestorFonts]) };
  }

  // Stage 3: Office fallback pool (incremental)
  const knownSet = new Set(ancestorFonts.map((f) => f.toLowerCase()));
  const officeOnly = officeFonts.filter((f) => !knownSet.has(f.toLowerCase()));
  const stage3Incremental = detectBestMatch(ctx, baseStyle, sample, actual, officeOnly, glyphCheck);
  const stage3 = stage3Incremental.score < stage2.score ? stage3Incremental : stage2;

  return {
    font: normalizeResult(stage3.font || officeFonts[0]),
    rawFont: stage3.font || officeFonts[0],
    score: stage3.score, stage: "office",
    candidates: uniqueFonts([...ancestorFonts, ...officeFonts]),
  };
}

function detectBestMatch(ctx, baseStyle, sample, actual, fonts, glyphCheck) {
  const candidates = fonts.filter(Boolean);
  if (!candidates.length) return { font: "", score: Number.POSITIVE_INFINITY };

  let bestFont = candidates[candidates.length - 1];
  let bestScore = Number.POSITIVE_INFINITY;

  for (const fontName of candidates) {
    const test = measureSingleFont(ctx, baseStyle, fontName, sample);
    let score = calcMetricScore(test, actual);

    // Dual-fallback glyph coverage check:
    // Measures the candidate with two different fallbacks (monospace and serif).
    // If the font has the target glyphs, both measurements match (font renders its own glyphs).
    // If it lacks them, the browser falls back differently for each, producing divergent metrics.
    // A +1000 penalty ensures fonts without target glyphs never win.
    if (glyphCheck) {
      if (!fontHasGlyphForSample(ctx, baseStyle, fontName, sample)) {
        score += 1000;
      }
    }

    if (score < bestScore) { bestScore = score; bestFont = fontName; }
    if (score <= 0.0001) break;
  }

  return { font: bestFont, score: bestScore };
}

/* =========================================================
 * Text samples and script identification
 * ========================================================= */

function collectSample(text, predicate, fallback, maxLen) {
  let result = "";
  let count = 0;
  for (const ch of text) {
    if (/\s/.test(ch)) continue;
    if (predicate(ch)) {
      result += ch;
      count += 1;
      if (count >= maxLen) break;
    }
  }
  return result || fallback;
}

function countScripts(text) {
  const counts = { latin: 0, ea: 0, cs: 0, symbol: 0, other: 0, total: 0 };
  for (const ch of text) {
    if (/\s/.test(ch)) continue;
    counts.total += 1;
    if (isEastAsianChar(ch)) { counts.ea += 1; continue; }
    if (isComplexScriptChar(ch)) { counts.cs += 1; continue; }
    if (isLatinLikeChar(ch)) { counts.latin += 1; continue; }
    if (isSymbolLikeChar(ch)) { counts.symbol += 1; continue; }
    counts.other += 1;
  }
  return counts;
}

function isLatinLikeChar(ch) {
  return /[A-Za-z0-9\u00C0-\u024F\u1E00-\u1EFF\u2000-\u206F]/.test(ch);
}

function isEastAsianChar(ch) {
  if (/[\u3000-\u303F\uFF00-\uFFEF\u3040-\u30FF\uAC00-\uD7AF]/.test(ch)) return true;
  return /\p{Script=Han}/u.test(ch);
}

function isComplexScriptChar(ch) {
  return /[\u0530-\u058F\u0590-\u05FF\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\u0900-\u097F\u0980-\u09FF\u0A00-\u0A7F\u0A80-\u0AFF\u0B00-\u0B7F\u0B80-\u0BFF\u0C00-\u0C7F\u0C80-\u0CFF\u0D00-\u0D7F\u0D80-\u0DFF\u0E00-\u0E7F\u0E80-\u0EFF\u0F00-\u0FFF\u1000-\u109F\u10A0-\u10FF\u1200-\u137F]/.test(ch);
}

function isSymbolLikeChar(ch) {
  if (/[\u0370-\u03FF\u2100-\u214F\u2190-\u21FF\u2200-\u22FF\u2300-\u23FF\u2500-\u257F\u25A0-\u25FF\u2600-\u26FF\u2700-\u27BF\u27C0-\u27EF]/.test(ch)) return true;
  return /[\u{1D400}-\u{1D7FF}\u{1F600}-\u{1F64F}\u{1F300}-\u{1F5FF}\u{1F680}-\u{1F6FF}\u{1F900}-\u{1F9FF}]/u.test(ch);
}

function containsMathOrSymbolFontName(fonts) {
  const patterns = [/cambria math/i, /stix/i, /latin modern math/i, /xits/i, /\bmath\b/i, /\bsymbol\b/i, /wingdings/i, /webdings/i, /mt extra/i, /seg(oe)? ui symbol/i, /katex/i];
  return fonts.some((font) => patterns.some((p) => p.test(font)));
}

/* =========================================================
 * Font candidate collection
 * ========================================================= */

function collectAncestorFontFamilies(element) {
  const fonts = [];
  const seen = new Set();
  let node = element.parentElement;
  while (node && node.nodeType === 1) {
    const style = window.getComputedStyle(node);
    const list = normalizeFontList(style.fontFamily);
    for (const font of list) {
      const key = font.toLowerCase();
      if (!seen.has(key)) { seen.add(key); fonts.push(font); }
    }
    node = node.parentElement;
  }
  return fonts;
}

function normalizeFontList(fontFamily) {
  if (!fontFamily || typeof fontFamily !== "string") return [];
  return uniqueFonts(fontFamily.split(",").map((f) => f.trim()).map(stripQuotes).filter(Boolean));
}

function uniqueFonts(list) {
  const result = [];
  const seen = new Set();
  for (const item of list) {
    const value = String(item || "").trim();
    if (!value) continue;
    const key = value.toLowerCase();
    if (!seen.has(key)) { seen.add(key); result.push(value); }
  }
  return result;
}

function stripQuotes(name) {
  return String(name || "").replace(/^["']|["']$/g, "").trim();
}

/* =========================================================
 * Canvas text measurement
 * ========================================================= */

let _sharedCtx = null;
function getSharedCanvasContext() {
  if (_sharedCtx) return _sharedCtx;
  if (typeof document === "undefined") return null;
  const canvas = document.createElement("canvas");
  _sharedCtx = canvas.getContext("2d", { alpha: false });
  return _sharedCtx;
}

function buildCanvasFontBase(style) {
  const sizeFragment = style.lineHeight && style.lineHeight !== "normal"
    ? `${style.fontSize}/${style.lineHeight}` : style.fontSize;
  return [style.fontStyle, style.fontVariant, style.fontWeight, sizeFragment].filter(Boolean).join(" ");
}

function measureWithStack(ctx, baseStyle, fullFontFamily, sample) {
  ctx.font = `${baseStyle} ${fullFontFamily}`;
  return readTextMetrics(ctx.measureText(sample));
}

function measureSingleFont(ctx, baseStyle, fontName, sample) {
  const family = quoteFontFamily(fontName);
  ctx.font = `${baseStyle} ${family}, monospace`;
  return readTextMetrics(ctx.measureText(sample));
}

/**
 * Dual-fallback glyph coverage check.
 * Measures candidate font with two different fallbacks (monospace and serif).
 * If the font has the target glyphs, both produce identical metrics.
 * If it lacks them, the browser falls back differently, producing divergent metrics.
 */
function fontHasGlyphForSample(ctx, baseStyle, fontName, sample) {
  const family = quoteFontFamily(fontName);

  ctx.font = `${baseStyle} ${family}, monospace`;
  const withMono = readTextMetrics(ctx.measureText(sample));

  ctx.font = `${baseStyle} ${family}, serif`;
  const withSerif = readTextMetrics(ctx.measureText(sample));

  const dist = calcMetricScore(withMono, withSerif);
  return dist < 0.01;
}

function quoteFontFamily(name) {
  const n = stripQuotes(name);
  if (/^[a-zA-Z0-9_-]+$/.test(n)) return n;
  return `"${n.replace(/"/g, '\\"')}"`;
}

function readTextMetrics(m) {
  return {
    width: m.width || 0,
    ascent: m.actualBoundingBoxAscent || 0,
    descent: m.actualBoundingBoxDescent || 0,
    left: m.actualBoundingBoxLeft || 0,
    right: m.actualBoundingBoxRight || 0,
  };
}

function calcMetricScore(test, actual) {
  const safeDiv = (diff, base) => Math.abs(diff) / Math.max(Math.abs(base), 1);
  return (
    safeDiv(test.width - actual.width, actual.width) * 10 +
    safeDiv(test.ascent - actual.ascent, actual.ascent) * 4 +
    safeDiv(test.descent - actual.descent, actual.descent) * 4 +
    safeDiv(test.left - actual.left, actual.left) * 2 +
    safeDiv(test.right - actual.right, actual.right) * 2
  );
}

/* =========================================================
 * Office / PPT friendly fallback font pools
 * ========================================================= */

function getOfficeLatinFonts() {
  return ["Calibri", "Arial", "Times New Roman", "Verdana", "Tahoma", "Segoe UI"];
}

function getOfficeEAFonts() {
  return ["Microsoft YaHei", "SimSun", "SimHei", "DengXian", "KaiTi", "FangSong", "Meiryo", "MS Gothic", "MS Mincho", "Malgun Gothic", "Batang"];
}

function getOfficeCSFonts() {
  return ["Arial", "Times New Roman", "Tahoma", "Segoe UI", "Nirmala UI", "Mangal", "Aparajita", "Kokila", "Shruti", "Raavi", "Gautami", "Vrinda", "Leelawadee UI", "David"];
}

function getOfficeSymbolFonts() {
  return ["Cambria Math", "Segoe UI Symbol", "Arial Unicode MS", "Symbol", "MT Extra", "Times New Roman", "Arial", "Microsoft YaHei"];
}

/* =========================================================
 * Output normalization for PPTX
 * ========================================================= */

function normalizeToPPTLatinFont(name, officeFonts, preferOfficeFonts = true) {
  const font = stripQuotes(String(name || "").trim());
  if (!font) return officeFonts[0] || "Arial";
  const lower = font.toLowerCase();
  const aliasMap = {
    "system-ui": "Arial", "sans-serif": "Arial", serif: "Times New Roman",
    monospace: "Courier New", "ui-sans-serif": "Arial", "ui-serif": "Times New Roman",
    "ui-monospace": "Courier New", "helvetica neue": "Arial", helvetica: "Arial",
    arialmt: "Arial", "sf pro text": "Arial", "sf pro display": "Arial", "segoe ui": "Segoe UI",
  };
  if (aliasMap[lower]) return aliasMap[lower];
  return font;
}

function normalizeToPPTEAFont(name, officeFonts, preferOfficeFonts = true) {
  const font = stripQuotes(String(name || "").trim());
  if (!font) return officeFonts[0] || "Microsoft YaHei";
  const lower = font.toLowerCase();
  const aliasMap = {
    "system-ui": officeFonts[0] || "Microsoft YaHei", "sans-serif": officeFonts[0] || "Microsoft YaHei",
    serif: "SimSun", "ui-sans-serif": officeFonts[0] || "Microsoft YaHei",
    "pingfang sc": "Microsoft YaHei", "pingfang tc": "Microsoft JhengHei",
    ".pingfang sc": "Microsoft YaHei", "hiragino sans gb": "Microsoft YaHei",
    "hiragino sans": "Meiryo", "heiti sc": "SimHei", "heiti tc": "Microsoft JhengHei",
    "songti sc": "SimSun", "songti tc": "PMingLiU", stheiti: "SimHei",
    stsong: "SimSun", stkaiti: "KaiTi", stfangsong: "FangSong",
    "apple lihei pro": "Microsoft YaHei",
    "microsoft yahei ui": "Microsoft YaHei", "ms yahei": "Microsoft YaHei",
    "ms gothic": "MS Gothic", "ms mincho": "MS Mincho",
    "yu gothic": "Meiryo", "yu gothic ui": "Meiryo", "yu mincho": "MS Mincho",
    "noto sans cjk sc": "Microsoft YaHei", "noto sans cjk tc": "Microsoft JhengHei",
    "noto sans cjk jp": "Meiryo", "noto sans cjk kr": "Malgun Gothic",
    "noto serif cjk sc": "SimSun", "source han sans sc": "Microsoft YaHei",
    "source han sans tc": "Microsoft JhengHei", "source han serif sc": "SimSun",
    "wenquanyi micro hei": "Microsoft YaHei", "wenquanyi zen hei": "Microsoft YaHei",
    "wenquanyi micro hei mono": "Microsoft YaHei",
    "apple sd gothic neo": "Malgun Gothic", "nanum gothic": "Malgun Gothic",
    nanumgothic: "Malgun Gothic", batang: "Batang", gulim: "Malgun Gothic",
  };
  if (aliasMap[lower]) return aliasMap[lower];
  return font;
}

function normalizeToPPTCSFont(name, officeFonts, preferOfficeFonts = true) {
  const font = stripQuotes(String(name || "").trim());
  if (!font) return officeFonts[0] || "Arial";
  const lower = font.toLowerCase();
  const aliasMap = {
    "system-ui": officeFonts[0] || "Arial", "sans-serif": officeFonts[0] || "Arial",
    serif: "Times New Roman", "ui-sans-serif": officeFonts[0] || "Arial", "ui-serif": "Times New Roman",
  };
  if (aliasMap[lower]) return aliasMap[lower];
  return font;
}

function normalizeToPPTSymbolFont(name, officeFonts, preferOfficeFonts = true) {
  const font = stripQuotes(String(name || "").trim());
  if (!font) return officeFonts[0] || "Cambria Math";
  const lower = font.toLowerCase();
  const aliasMap = {
    "system-ui": "Segoe UI Symbol", "sans-serif": "Segoe UI Symbol", serif: "Times New Roman",
    "cambria math": "Cambria Math", "segoe ui symbol": "Segoe UI Symbol",
    "stix two math": "STIX Two Math", symbol: "Symbol", "mt extra": "MT Extra",
    "arial unicode ms": "Arial Unicode MS",
  };
  if (aliasMap[lower]) return aliasMap[lower];
  return font;
}
