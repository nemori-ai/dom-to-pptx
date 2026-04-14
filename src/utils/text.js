// src/utils/text.js
// Text styling and font processing utilities

import {
  parseColor,
  getGradientFallbackColor,
  blendColors,
  getEffectiveBackground,
} from './color.js';
import { PX_TO_PT, FONT_SIZE_FACTOR } from './constants.js';
import detectFontsForPPTX from './detect-fonts.js';

// CJK font indicators in font names (lowercase for matching)
const CJK_FONT_INDICATORS = [
  // Chinese
  'sc',
  'tc',
  'hk',
  'cn',
  'gb',
  'chinese',
  'pingfang',
  'yahei',
  'heiti',
  'songti',
  'kaiti',
  'fangsong',
  'simhei',
  'simsun',
  'nsimsun',
  'mingliu',
  'pmingliu',
  'microsoft jhenghei',
  'microsoft yahei',
  'stheitisc',
  'stheiti',
  'hiragino sans gb',
  'wenquanyi',
  'fzlangtingheiti',
  'noto sans cjk',
  'noto serif cjk',
  'source han',
  '思源',
  '黑体',
  '宋体',
  '楷体',
  '仿宋',
  // Japanese
  'jp',
  'japanese',
  'hiragino',
  'meiryo',
  'yu gothic',
  'yu mincho',
  'ms gothic',
  'ms mincho',
  'kozuka',
  'morisawa',
  // Korean
  'kr',
  'korean',
  'malgun',
  'gulim',
  'dotum',
  'batang',
  'gungsuh',
  'nanum',
  'apple sd gothic',
  // Generic CJK
  'cjk',
];

/**
 * Check if a font name indicates CJK support based on naming conventions.
 */
function isCJKFont(fontName) {
  const lower = fontName.toLowerCase();
  return CJK_FONT_INDICATORS.some((indicator) => lower.includes(indicator));
}

/**
 * Check if text contains CJK characters.
 */
function containsCJK(text) {
  if (!text) return false;
  return /[\u4e00-\u9fff\u3040-\u309f\u30a0-\u30ff\uac00-\ud7af]/.test(text);
}

/**
 * Detect language for PptxGenJS lang property.
 * Returns language code for CJK text to ensure proper rendering.
 */
function detectLanguage(text) {
  if (!text) return undefined;
  if (/[\u4e00-\u9fff]/.test(text)) return 'zh-CN';
  if (/[\u3040-\u309f\u30a0-\u30ff]/.test(text)) return 'ja-JP';
  if (/[\uac00-\ud7af]/.test(text)) return 'ko-KR';
  return undefined;
}

// Platform-specific CJK fonts that don't work reliably in PowerPoint.
// Maps to Office-safe cross-platform alternatives.
const CJK_FONT_COMPAT_MAP = {
  // macOS fonts → Office-safe
  'pingfang sc': 'Microsoft YaHei',
  'pingfang tc': 'Microsoft JhengHei',
  'pingfang hk': 'Microsoft JhengHei',
  'heiti sc': 'Microsoft YaHei',
  'heiti tc': 'Microsoft JhengHei',
  'songti sc': 'SimSun',
  'songti tc': 'MingLiU',
  'kaiti sc': 'KaiTi',
  'stheitisc-light': 'Microsoft YaHei',
  stheiti: 'Microsoft YaHei',
  'hiragino sans gb': 'Microsoft YaHei',
  // Linux/Android fonts → Office-safe
  'noto sans sc': 'Microsoft YaHei',
  'noto sans tc': 'Microsoft JhengHei',
  'noto sans cjk sc': 'Microsoft YaHei',
  'noto sans cjk tc': 'Microsoft JhengHei',
  'noto sans cjk jp': 'Yu Gothic',
  'noto sans cjk kr': 'Malgun Gothic',
  'noto serif sc': 'SimSun',
  'noto serif cjk sc': 'SimSun',
  'source han sans sc': 'Microsoft YaHei',
  'source han sans tc': 'Microsoft JhengHei',
  'source han serif sc': 'SimSun',
  'wenquanyi micro hei': 'Microsoft YaHei',
  'wenquanyi zen hei': 'Microsoft YaHei',
  'droid sans fallback': 'Microsoft YaHei',
};

/**
 * Map a CJK font to an Office-compatible equivalent.
 * Returns the original font if no mapping is needed.
 */
function toOfficeCJKFont(fontName) {
  return CJK_FONT_COMPAT_MAP[fontName.toLowerCase()] || fontName;
}

/**
 * Return the default CJK font for PPTX output.
 * Always returns an Office-safe font regardless of platform.
 */
function getSystemCJKFont() {
  return 'Microsoft YaHei';
}

/**
 * Cross-platform safe generic font mappings.
 * Uses fonts that exist on both Windows and macOS with good Office compatibility.
 *
 * Strategy:
 * - Latin: Use universally available fonts (Arial, Times New Roman, Courier New)
 * - CJK: Use Microsoft YaHei (available on Mac Office, fallbacks to PingFang if not installed)
 */
const GENERIC_FONT_MAP = {
  // === Sans-serif (default body text) ===
  // Arial is universal. Microsoft YaHei works on Mac Office, fallbacks to PingFang.
  'system-ui': { latin: 'Arial', ea: 'Microsoft YaHei' },
  '-apple-system': { latin: 'Arial', ea: 'Microsoft YaHei' },
  blinkmacsystemfont: { latin: 'Arial', ea: 'Microsoft YaHei' },
  'ui-sans-serif': { latin: 'Arial', ea: 'Microsoft YaHei' },
  'sans-serif': { latin: 'Arial', ea: 'Microsoft YaHei' },

  // === Serif ===
  // Times New Roman is cross-platform. SimSun is the only cross-platform CJK serif.
  'ui-serif': { latin: 'Times New Roman', ea: 'SimSun' },
  serif: { latin: 'Times New Roman', ea: 'SimSun' },

  // === Monospace ===
  // Courier New is universal. Microsoft YaHei for CJK (no true monospace CJK).
  'ui-monospace': { latin: 'Courier New', ea: 'Microsoft YaHei' },
  monospace: { latin: 'Courier New', ea: 'Microsoft YaHei' },

  // === Decorative ===
  'ui-rounded': { latin: 'Arial Rounded MT Bold', ea: 'Microsoft YaHei' },
  cursive: { latin: 'Comic Sans MS', ea: 'KaiTi' },
  fantasy: { latin: 'Impact', ea: 'Microsoft YaHei' },
};

/**
 * Map CSS generic font families to actual font names.
 * @param {string} fontName - Generic font name (e.g., 'sans-serif')
 * @param {boolean} [forCJK=false] - Whether to return CJK font
 * @returns {string|null} Actual font name or null
 */
function resolveGenericFont(fontName, forCJK = false) {
  const lower = fontName.toLowerCase();
  const mapping = GENERIC_FONT_MAP[lower];

  if (mapping) {
    return forCJK ? mapping.ea : mapping.latin;
  }

  return null;
}

/**
 * Resolve a font name - if it's a generic family, map to actual font.
 * @param {string} fontName - Font name (possibly generic)
 * @param {boolean} [forCJK=false] - Whether to resolve for CJK text
 */
function resolveFontName(fontName, forCJK = false) {
  const resolved = resolveGenericFont(fontName, forCJK);
  return resolved || fontName;
}

/**
 * Check if a font name is a generic font family.
 */
function isGenericFont(fontName) {
  const generics = [
    'system-ui',
    '-apple-system',
    'blinkmacsystemfont',
    'ui-sans-serif',
    'ui-serif',
    'ui-monospace',
    'ui-rounded',
    'sans-serif',
    'serif',
    'monospace',
    'cursive',
    'fantasy',
  ];
  return generics.includes(fontName.toLowerCase());
}

/**
 * Font weight compensation for text box width.
 * Bold text in PPTX tends to render slightly wider than in browsers,
 * causing text to overflow and wrap unexpectedly.
 *
 * @param {number} fontWeight - CSS font-weight value (100-900)
 * @param {number} fontSizePx - Font size in pixels
 * @returns {number} Width compensation per character in points
 */
export function getWeightCompensation(fontWeight, fontSizePx = 16) {
  // Only compensate for bold text (weight >= 700)
  if (fontWeight < 700) return 0;

  // Compensation factor: bold text renders wider in PPTX than browsers
  // Base: 0.5pt per character at 16px, scales with font size
  // This accounts for font rendering differences between browser and PPTX
  const basePt = 0.5;
  const scaleFactor = fontSizePx / 16;

  return basePt * scaleFactor;
}

/**
 * Calculate text box width compensation for charSpacing adjustments.
 * This accounts for the difference between browser rendering and PPTX rendering.
 *
 * @param {number} fontWeight - CSS font-weight value (100-900)
 * @param {number} letterSpacingPx - CSS letter-spacing in pixels (0 if normal)
 * @param {number} charCount - Number of characters in the text
 * @param {number} scale - Layout scale factor
 * @param {number} [fontSizePx=16] - CSS font-size in pixels
 * @returns {number} Width compensation in inches
 */
export function getCharSpacingWidthCompensation(
  fontWeight,
  letterSpacingPx,
  charCount,
  scale,
  fontSizePx = 16
) {
  if (charCount <= 1) return 0;

  // 1. Font weight compensation (additional spacing we add)
  const weightCompPt = getWeightCompensation(fontWeight, fontSizePx) * scale;

  // 2. CSS letter-spacing: browser already includes this in width,
  //    but PPTX rendering may differ slightly. We calculate the
  //    PPTX charSpacing value for reference.
  const cssLetterSpacingPt = letterSpacingPx * PX_TO_PT * scale;

  // The browser's getBoundingClientRect already includes CSS letter-spacing effect.
  // We only need to compensate for the ADDITIONAL spacing we add (font-weight compensation).
  // However, if there's CSS letter-spacing, there might be rendering differences
  // between browser and PPTX, so we apply a small correction factor.

  // Only compensate for font-weight adjustment (what we add extra)
  let compensationPt = weightCompPt;

  // If CSS letter-spacing exists, add a small correction for potential rendering differences
  // (PPTX tends to render slightly differently than browsers)
  if (cssLetterSpacingPt !== 0) {
    // Small correction factor (5%) for letter-spacing rendering differences
    compensationPt += cssLetterSpacingPt * 0.05;
  }

  if (compensationPt === 0) return 0;

  // Convert points to inches: 1pt = 1/72 inch
  // Apply to (charCount - 1) gaps between characters
  return (charCount - 1) * compensationPt * (1 / 72);
}

// Backward compatibility alias
export function getWeightWidthCompensation(fontWeight, charCount, scale) {
  return getCharSpacingWidthCompensation(fontWeight, 0, charCount, scale);
}

/**
 * Collect font-family from element and all ancestors (simulating CSS inheritance).
 * Returns a deduplicated list of all fonts in the inheritance chain.
 */
function collectInheritedFonts(node) {
  const allFonts = [];
  const seen = new Set();

  let current = node;
  while (current && current.nodeType === 1) {
    const style = window.getComputedStyle(current);
    const fonts = style.fontFamily.split(',').map((f) => f.trim().replace(/['"]/g, ''));

    for (const font of fonts) {
      if (font && !seen.has(font.toLowerCase())) {
        seen.add(font.toLowerCase());
        allFonts.push(font);
      }
    }

    current = current.parentElement;
  }

  return allFonts;
}

/**
 * Get the appropriate font from font-family list based on text content.
 * If text contains CJK characters and different fonts are detected for Latin/CJK,
 * returns JSON string: '{"latin":"Arial","ea":"Microsoft YaHei"}'
 *
 * @param {string} fontFamilyStr - The element's computed font-family
 * @param {string} text - The text content to analyze
 * @param {HTMLElement} [node] - Optional: the DOM node for inheritance chain lookup
 */
function selectFontForText(fontFamilyStr, text, node = null) {
  let fonts = fontFamilyStr.split(',').map((f) => f.trim().replace(/['"]/g, ''));

  // No CJK characters - use first font (resolved from generic if needed)
  if (!containsCJK(text)) {
    return resolveFontName(fonts[0], false);
  }

  // Text contains CJK - if node provided, collect fonts from inheritance chain
  if (node) {
    fonts = collectInheritedFonts(node);
  }

  // Find Latin font and CJK font from the collected list
  let latinFont = null;
  let eaFont = null;

  for (const font of fonts) {
    if (isGenericFont(font)) {
      if (!latinFont) latinFont = resolveFontName(font, false);
      if (!eaFont) eaFont = resolveFontName(font, true);
    } else {
      // Use the user-specified font for both latin and EA slots.
      // We don't try to guess whether a font is "CJK" by name — that's unreliable
      // (misses fonts like ZCOOL KuaiLe, and wrongly replaces embedded fonts like
      // Noto Sans SC). The user's explicit font choice should be respected.
      if (!latinFont) latinFont = font;
      if (!eaFont) eaFont = font;
    }
    if (latinFont && eaFont) break;
  }

  // Fallback to first font if not found
  latinFont = latinFont || resolveFontName(fonts[0], false);

  // For CJK fallback: try to resolve first generic font for CJK, or use system font
  if (!eaFont) {
    const firstGeneric = fonts.find(isGenericFont);
    if (firstGeneric) {
      eaFont = resolveFontName(firstGeneric, true);
    } else {
      eaFont = getSystemCJKFont();
    }
  }

  // If both fonts are the same, just return the font name
  if (latinFont === eaFont) {
    return latinFont;
  }

  // Return JSON string for post-processing
  return JSON.stringify({ latin: latinFont, ea: eaFont });
}

/**
 * Parse CSS text-shadow to PptxGenJS shadow format.
 * CSS: "2px 2px 4px rgba(0,0,0,0.5)"
 * PPTX: { type: 'outer', color, blur, offset, angle }
 */
function parseTextShadow(shadowStr, scale) {
  if (!shadowStr || shadowStr === 'none') return null;

  // Match: [color] x y blur [color]
  // CSS allows color at start or end
  const colorRegex = /(rgba?\([^)]+\)|hsla?\([^)]+\)|#[0-9a-fA-F]{3,8}|\b[a-z]+\b)/gi;
  const colors = shadowStr.match(colorRegex) || [];
  const nums = shadowStr.replace(colorRegex, '').match(/-?[\d.]+/g) || [];

  if (nums.length < 2) return null;

  const xPx = parseFloat(nums[0]) || 0;
  const yPx = parseFloat(nums[1]) || 0;
  const blurPx = parseFloat(nums[2]) || 0;

  // Convert Cartesian (x, y) to Polar (angle, distance)
  const distancePx = Math.sqrt(xPx * xPx + yPx * yPx);
  let angleDeg = Math.atan2(yPx, xPx) * (180 / Math.PI);
  if (angleDeg < 0) angleDeg += 360;

  // Parse color
  let colorHex = '000000';
  let opacity = 1;
  if (colors.length > 0) {
    const colorObj = parseColor(colors[0]);
    if (colorObj.hex) colorHex = colorObj.hex;
    opacity = colorObj.opacity;
  }

  // Skip if shadow is invisible
  if (distancePx < 0.5 && blurPx < 0.5) return null;
  if (opacity === 0) return null;

  return {
    type: 'outer',
    color: colorHex,
    blur: blurPx * PX_TO_PT * scale,
    offset: distancePx * PX_TO_PT * scale,
    angle: Math.round(angleDeg),
    opacity: opacity,
  };
}

/**
 * Extracts text styling from computed style for PptxGenJS.
 * Supports: color, fontFace, fontSize, bold, italic, underline, strike,
 * superscript, subscript, lineSpacing, charSpacing, shadow, highlight,
 * transparency, rtlMode, vertical text.
 */
/**
 * @param {CSSStyleDeclaration} style - Computed style
 * @param {number} scale - Layout scale factor
 * @param {string} [text=''] - Text content for CJK detection
 * @param {HTMLElement} [node=null] - DOM node for font inheritance lookup
 * @param {Object} [options={}] - Additional options (e.g., backgroundSnapshot)
 */
export function getTextStyle(style, scale, text = '', node = null, options = {}) {
  let colorObj = parseColor(style.color);

  const bgClip = style.webkitBackgroundClip || style.backgroundClip;
  if (colorObj.opacity === 0 && bgClip === 'text') {
    const fallback = getGradientFallbackColor(style.backgroundImage);
    if (fallback) colorObj = parseColor(fallback);
  }

  // --- Line Height → lineSpacing (absolute pt) ---
  // Use absolute "Exactly" line spacing (pt) instead of "Multiple" to avoid
  // coupling with the fontSize empirical shrink factor (FONT_SIZE_FACTOR).
  // CSS line-height in px → convert to pt (×PX_TO_PT) → apply layout scale.
  const fontSizePx = parseFloat(style.fontSize);
  const lhStr = style.lineHeight;
  let lineSpacingPt = null;

  if (fontSizePx > 0 && lhStr && lhStr !== 'normal') {
    const lhValue = parseFloat(lhStr);

    if (!isNaN(lhValue) && lhValue > 0) {
      let lineHeightPx;
      if (/^[0-9.]+$/.test(lhStr)) {
        lineHeightPx = lhValue * fontSizePx;
      } else if (lhStr.includes('%')) {
        lineHeightPx = (lhValue / 100) * fontSizePx;
      } else if (lhStr.includes('em')) {
        lineHeightPx = lhValue * fontSizePx;
      } else {
        lineHeightPx = lhValue;
      }
      lineSpacingPt = Math.round(lineHeightPx * PX_TO_PT * scale * 100) / 100;
    }
  } else if (fontSizePx > 0) {
    // line-height: normal → use 1.2x as default (matches browser default and previous behavior)
    lineSpacingPt = Math.round(fontSizePx * 1.2 * PX_TO_PT * scale * 100) / 100;
  }

  // --- Letter Spacing ---
  let charSpacing = null;
  const lsStr = style.letterSpacing;
  if (lsStr && lsStr !== 'normal') {
    const lsPx = parseFloat(lsStr);
    if (!isNaN(lsPx) && lsPx !== 0) {
      charSpacing = lsPx * PX_TO_PT * scale;
    }
  }

  // --- Font Weight Compensation ---
  // CSS has 100-900 weight levels, PPTX only has bold (true/false).
  // Compensate intermediate weights by adjusting character spacing to
  // visually approximate the weight difference.
  const fontWeight = parseInt(style.fontWeight) || 400;
  const isBold = fontWeight >= 600;

  // Weight compensation mapping (in points, before scale):
  // - Light weights (< 500): no adjustment
  // - Medium (500-599): slight positive spacing for heavier look
  // - Bold (600+): -1pt only for fontSize >= 18px
  const weightCompensation = getWeightCompensation(fontWeight, fontSizePx);

  if (weightCompensation !== 0) {
    const adjustment = weightCompensation * scale;
    charSpacing = charSpacing ? charSpacing + adjustment : adjustment;
  }

  // Select appropriate font(s) for each OOXML script category.
  // Uses Canvas measureText to detect actual rendering fonts when a DOM node is available.
  // Falls back to heuristic-based selectFontForText when no node (theoretical safety net).
  let fontFace;
  if (node) {
    const detected = detectFontsForPPTX(node, text);
    const allSame = detected.latin === detected.ea
      && detected.latin === detected.cs
      && detected.latin === detected.symbol;
    if (allSame) {
      fontFace = detected.latin;
    } else {
      fontFace = JSON.stringify({
        latin: detected.latin,
        ea: detected.ea,
        cs: detected.cs,
        sym: detected.symbol,
      });
    }
  } else {
    fontFace = selectFontForText(style.fontFamily, text, node);
  }

  // --- Strikethrough ---
  // CSS text-decoration can include: underline, line-through, overline
  const textDecoration = style.textDecoration || style.textDecorationLine || '';
  const hasStrike = textDecoration.includes('line-through');

  // --- Superscript / Subscript ---
  // CSS vertical-align: super, sub, baseline, etc.
  const vertAlign = style.verticalAlign;
  const isSuperscript = vertAlign === 'super';
  const isSubscript = vertAlign === 'sub';

  // --- Text Shadow ---
  const shadowObj = parseTextShadow(style.textShadow, scale);

  // --- RTL Mode ---
  // CSS direction: rtl | ltr
  const isRtl = style.direction === 'rtl';

  // --- Vertical Text ---
  // CSS writing-mode: vertical-rl, vertical-lr, horizontal-tb
  const writingMode = style.writingMode;
  let vert = null;
  if (writingMode === 'vertical-rl' || writingMode === 'vertical-lr') {
    // For East Asian vertical text
    vert = 'eaVert';
  }

  // --- Text Transparency ---
  // Extracted from color alpha channel
  const transparency = colorObj.opacity < 1 ? Math.round((1 - colorObj.opacity) * 100) : null;

  // --- Text Outline (Stroke) ---
  // CSS: -webkit-text-stroke: 1px #000; or text-stroke
  let outline = null;
  const strokeWidth = style.webkitTextStrokeWidth || style.textStrokeWidth;
  const strokeColor = style.webkitTextStrokeColor || style.textStrokeColor;
  if (strokeWidth && strokeColor) {
    const widthPx = parseFloat(strokeWidth);
    const colorParsed = parseColor(strokeColor);
    if (widthPx > 0 && colorParsed.hex) {
      outline = {
        color: colorParsed.hex,
        size: widthPx * PX_TO_PT * scale,
      };
    }
  }

  // Build result object
  const result = {
    color: colorObj.hex || '000000',
    fontFace,
    fontSize: Math.round(fontSizePx * PX_TO_PT * FONT_SIZE_FACTOR * scale * 10) / 10,
    bold: isBold,
    italic: style.fontStyle === 'italic',
    underline:
      textDecoration.includes('underline') ||
      (parseFloat(style.borderBottomWidth) > 0 &&
        style.borderBottomStyle !== 'none' &&
        style.borderBottomStyle !== 'hidden' &&
        parseColor(style.borderBottomColor).opacity > 0 &&
        !(parseFloat(style.borderTopWidth) > 0) &&
        !(parseFloat(style.borderLeftWidth) > 0) &&
        !(parseFloat(style.borderRightWidth) > 0) &&
        style.display &&
        style.display.includes('inline')),
  };

  // Conditional properties (only add if truthy/valid)
  if (hasStrike) result.strike = 'sngStrike';
  if (isSuperscript) result.superscript = true;
  if (isSubscript) result.subscript = true;
  if (lineSpacingPt) result.lineSpacing = lineSpacingPt;
  if (charSpacing) result.charSpacing = charSpacing;
  // Force paragraph spacing to 0 (layout uses absolute positioning)
  result.paraSpaceBefore = 0;
  result.paraSpaceAfter = 0;
  if (shadowObj) result.shadow = shadowObj;
  if (isRtl) result.rtlMode = true;
  if (vert) result.vert = vert;
  if (transparency) result.transparency = transparency;
  if (outline) result.outline = outline;

  const lang = detectLanguage(text);
  if (lang) result.lang = lang;

  const bgColorObj = parseColor(style.backgroundColor);
  if (bgColorObj.hex && bgColorObj.opacity > 0.05) {
    if (bgColorObj.opacity < 1 && node) {
      // Semi-transparent: blend with effective background behind this element
      // Use snapshot if available for accurate sampling (handles images, gradients, etc.)
      const parentBg = getEffectiveBackground(node, options.backgroundSnapshot);
      const blended = blendColors(bgColorObj.hex, bgColorObj.opacity, parentBg.hex);
      result.highlight = blended.hex;
    } else {
      // Fully opaque: use directly
      result.highlight = bgColorObj.hex;
    }
  }

  return result;
}

/**
 * Determines if a given DOM node is primarily a text container.
 * Updated to correctly reject Icon elements so they are rendered as images.
 */
export function isTextContainer(node) {
  const hasText = node.textContent.trim().length > 0;
  if (!hasText) return false;

  const children = Array.from(node.children);
  if (children.length === 0) return true;

  const isSafeInline = (el) => {
    // 1. Reject Web Components / Custom Elements
    if (el.tagName.includes('-')) return false;
    // 2. Reject Explicit Images/SVGs
    const tag = el.tagName.toUpperCase();
    if (tag === 'IMG' || tag === 'SVG') return false;

    // 3. Reject Class-based Icons (FontAwesome, Material, Bootstrap, etc.)
    // If an <i> or <span> has icon classes, it is a visual object, not text.
    if (el.tagName === 'I' || el.tagName === 'SPAN') {
      const cls = el.getAttribute('class') || '';
      if (
        cls.includes('fa-') ||
        cls.includes('fas') ||
        cls.includes('far') ||
        cls.includes('fab') ||
        cls.includes('material-icons') ||
        cls.includes('bi-') ||
        cls.includes('icon')
      ) {
        return false;
      }
    }

    const style = window.getComputedStyle(el);
    const display = style.display;

    // 4. Display check - block/flex/grid elements are NOT inline, even if they are SPAN tags
    // This handles Tailwind's "block" class on spans
    const isBlockDisplay = ['block', 'flex', 'grid', 'table', 'list-item'].some(
      (d) => display === d || display.startsWith(d + ' ')
    );
    if (isBlockDisplay) return false;

    // 5. Standard Inline Tag Check
    const isInlineTag = ['SPAN', 'B', 'STRONG', 'EM', 'I', 'A', 'SMALL', 'MARK'].includes(
      el.tagName
    );
    const isInlineDisplay = display.includes('inline');

    if (!isInlineTag && !isInlineDisplay) return false;

    // 5. Structural Styling Check
    // If a child has a background or border, check if it's a styled inline text span
    // (like code tags with bg+padding+radius) vs a standalone UI element (button/badge).
    const bgColor = parseColor(style.backgroundColor);
    const hasVisibleBg = bgColor.hex && bgColor.opacity > 0;
    const hasBorder =
      parseFloat(style.borderWidth) > 0 && parseColor(style.borderColor).opacity > 0;

    if (hasVisibleBg || hasBorder) {
      const hasBorderRadius = parseFloat(style.borderRadius) > 0;
      const hasPadding =
        parseFloat(style.paddingLeft) > 2 ||
        parseFloat(style.paddingRight) > 2 ||
        parseFloat(style.paddingTop) > 2 ||
        parseFloat(style.paddingBottom) > 2;

      // Elements with visible border are rendered as separate shapes to preserve
      // border/radius fidelity. TextNodeRenderer handles the surrounding text fragments
      // using per-line getClientRects() for correct multi-line positioning.
      if (hasBorder) return false;

      if (hasBorderRadius || hasPadding) {
        // Allow border-less styled inline text spans (bg+padding+radius).
        // PptxGenJS highlight handles the background color on the text run.
        const hasTextContent = el.textContent.trim().length > 0;
        const isKnownInlineTag = [
          'SPAN',
          'B',
          'STRONG',
          'EM',
          'I',
          'A',
          'SMALL',
          'MARK',
          'CODE',
          'ABBR',
          'SUB',
          'SUP',
        ].includes(el.tagName);
        const maxPad = Math.max(
          parseFloat(style.paddingLeft) || 0,
          parseFloat(style.paddingRight) || 0,
          parseFloat(style.paddingTop) || 0,
          parseFloat(style.paddingBottom) || 0
        );
        if (!hasTextContent || !isKnownInlineTag || maxPad >= 16) {
          return false;
        }
      }
    }

    // 4. Check for empty elements (visual objects without text, like dots)
    // Empty elements that take up space should not be treated as inline text
    const hasContent = el.textContent.trim().length > 0;
    if (!hasContent) {
      // Empty element - reject if it has any visual presence or fixed dimensions
      const hasFixedSize =
        (parseFloat(style.width) > 0 || parseFloat(style.minWidth) > 0) &&
        (parseFloat(style.height) > 0 || parseFloat(style.minHeight) > 0);
      if (hasVisibleBg || hasBorder || hasFixedSize) {
        return false;
      }
    }

    // Reject inline containers that contain media or icon descendants
    if (el.querySelector('svg, img, canvas, video')) return false;
    if (el.querySelector('i.fa, i.fas, i.far, i.fab, i.material-icons, [class*="bi-"]')) {
      return false;
    }

    return true;
  };

  return children.every(isSafeInline);
}
