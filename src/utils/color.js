// src/utils/color.js
// Color parsing utilities - core dependency for other modules

// Canvas context for color normalization (singleton)
let _ctx;
function getCtx() {
  if (!_ctx) _ctx = document.createElement('canvas').getContext('2d', { willReadFrequently: true });
  return _ctx;
}

/** Sentinel value for invalid/unparseable colors */
export const INVALID_COLOR = { hex: null, opacity: 0, invalid: true };

/**
 * Parses any CSS color string into { hex, opacity } format.
 * Handles modern CSS colors (oklch, lab, display-p3) via canvas-based normalization.
 */
export function parseColor(str) {
  if (!str || str === 'transparent' || str.trim() === 'rgba(0, 0, 0, 0)') {
    return { hex: null, opacity: 0 };
  }

  // Fast-path validation: check if browser recognizes this color syntax
  if (typeof CSS !== 'undefined' && CSS.supports && !CSS.supports('color', str)) {
    console.warn(`[dom-to-pptx] Unsupported color syntax: "${str}"`);
    return INVALID_COLOR;
  }

  // Fast path: direct rgb/rgba parsing (avoids canvas roundtrip and black-detection bug)
  const rgbMatch = str.match(
    /^rgba?\s*\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)(?:\s*,\s*([\d.]+))?\s*\)$/
  );
  if (rgbMatch) {
    const r = parseInt(rgbMatch[1]);
    const g = parseInt(rgbMatch[2]);
    const b = parseInt(rgbMatch[3]);
    const a = rgbMatch[4] !== undefined ? parseFloat(rgbMatch[4]) : 1;
    if (a === 0) return { hex: null, opacity: 0 };
    const hex = ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
    return { hex, opacity: a };
  }

  const ctx = getCtx();
  ctx.fillStyle = '#000000'; // Reset to known value
  ctx.fillStyle = str;
  const computed = ctx.fillStyle;

  // If fillStyle didn't change from reset value, the color was invalid
  if (computed === '#000000' && str.toLowerCase() !== '#000000' && str.toLowerCase() !== 'black') {
    console.warn(`[dom-to-pptx] Failed to parse color: "${str}"`);
    return INVALID_COLOR;
  }

  // 1. Handle Hex Output (e.g. #ff0000) - Fast Path
  if (computed.startsWith('#')) {
    let hex = computed.slice(1);
    let opacity = 1;
    if (hex.length === 3)
      hex = hex
        .split('')
        .map((c) => c + c)
        .join('');
    if (hex.length === 4)
      hex = hex
        .split('')
        .map((c) => c + c)
        .join('');
    if (hex.length === 8) {
      opacity = parseInt(hex.slice(6), 16) / 255;
      hex = hex.slice(0, 6);
    }
    return { hex: hex.toUpperCase(), opacity };
  }

  // 2. Handle RGB/RGBA Output (standard) - Fast Path
  if (computed.startsWith('rgb')) {
    const match = computed.match(/[\d.]+/g);
    if (match && match.length >= 3) {
      const r = parseInt(match[0]);
      const g = parseInt(match[1]);
      const b = parseInt(match[2]);
      const a = match.length > 3 ? parseFloat(match[3]) : 1;
      const hex = ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
      return { hex, opacity: a };
    }
  }

  // 3. Fallback: Browser returned a format we don't parse (oklch, lab, color(srgb...), etc.)
  // Use Canvas API to convert to sRGB
  ctx.clearRect(0, 0, 1, 1);
  ctx.fillRect(0, 0, 1, 1);
  const data = ctx.getImageData(0, 0, 1, 1).data;
  // data = [r, g, b, a]
  const r = data[0];
  const g = data[1];
  const b = data[2];
  const a = data[3] / 255;

  if (a === 0) return { hex: null, opacity: 0 };

  const hex = ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
  return { hex, opacity: a };
}

/**
 * Extracts the first color from a gradient string as fallback.
 * Used for gradient text where we need a solid color approximation.
 */
export function getGradientFallbackColor(bgImage) {
  if (!bgImage || bgImage === 'none') return null;

  // 1. Extract content inside function(...)
  // Handles linear-gradient(...), radial-gradient(...), repeating-linear-gradient(...)
  const match = bgImage.match(/gradient\((.*)\)/);
  if (!match) return null;

  const content = match[1];

  // 2. Split by comma, respecting parentheses (to avoid splitting inside rgb(), oklch(), etc.)
  const parts = [];
  let current = '';
  let parenDepth = 0;

  for (const char of content) {
    if (char === '(') parenDepth++;
    if (char === ')') parenDepth--;
    if (char === ',' && parenDepth === 0) {
      parts.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  if (current) parts.push(current.trim());

  // 3. Find first part that is a color (skip angle/direction)
  for (const part of parts) {
    // Ignore directions (to right) or angles (90deg, 0.5turn)
    if (/^(to\s|[\d.]+(deg|rad|turn|grad))/.test(part)) continue;

    // Extract color: Remove trailing position (e.g. "red 50%" -> "red")
    // Regex matches whitespace + number + unit at end of string
    const colorPart = part.replace(/\s+(-?[\d.]+(%|px|em|rem|ch|vh|vw)?)$/, '');

    // Check if it's not just a number (some gradients might have bare numbers? unlikely in standard syntax)
    if (colorPart) return colorPart;
  }

  return null;
}

/**
 * Converts a hex color string to RGB components.
 * @param {string} hex - 6-character hex string (without #)
 * @returns {{ r: number, g: number, b: number }}
 */
function hexToRgb(hex) {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  return { r, g, b };
}

/**
 * Converts RGB components to a hex string.
 * @param {number} r
 * @param {number} g
 * @param {number} b
 * @returns {string} 6-character uppercase hex string
 */
function rgbToHex(r, g, b) {
  return ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
}

/**
 * Blends a semi-transparent foreground color with a background color.
 * Uses standard alpha compositing: result = fg * alpha + bg * (1 - alpha)
 *
 * @param {string} fgHex - Foreground color (6-char hex)
 * @param {number} fgOpacity - Foreground opacity (0-1)
 * @param {string} bgHex - Background color (6-char hex)
 * @returns {{ hex: string, opacity: number }} Blended color (always opaque)
 */
export function blendColors(fgHex, fgOpacity, bgHex) {
  const fg = hexToRgb(fgHex);
  const bg = hexToRgb(bgHex);

  const r = Math.round(fg.r * fgOpacity + bg.r * (1 - fgOpacity));
  const g = Math.round(fg.g * fgOpacity + bg.g * (1 - fgOpacity));
  const b = Math.round(fg.b * fgOpacity + bg.b * (1 - fgOpacity));

  return { hex: rgbToHex(r, g, b), opacity: 1 };
}

/**
 * Gets the effective background color behind an element.
 * If a snapshot is provided, samples directly from the pre-captured image.
 * Otherwise falls back to DOM traversal.
 *
 * @param {Element} node - DOM node to sample behind
 * @param {Object} [snapshot] - Pre-captured background snapshot from captureBackgroundSnapshot()
 * @returns {{ hex: string, opacity: number }} Sampled background color
 */
export function getEffectiveBackground(node, snapshot = null) {
  if (!node) return { hex: 'FFFFFF', opacity: 1 };

  try {
    const rect = node.getBoundingClientRect();
    if (rect.width === 0 || rect.height === 0) {
      return { hex: 'FFFFFF', opacity: 1 };
    }

    // Sample center point of the element
    const x = rect.left + rect.width / 2;
    const y = rect.top + rect.height / 2;

    // Priority 1: Use snapshot if available (most accurate - handles images, gradients, etc.)
    if (snapshot?.imageData) {
      // Sample directly from snapshot ImageData (inline to avoid circular dependency)
      // Account for DPI scaling: snapshot canvas may be larger than CSS rect
      const scaleX = snapshot.width / snapshot.rect.width;
      const scaleY = snapshot.height / snapshot.rect.height;
      const localX = Math.max(
        0,
        Math.min(Math.round((x - snapshot.rect.left) * scaleX), snapshot.width - 1)
      );
      const localY = Math.max(
        0,
        Math.min(Math.round((y - snapshot.rect.top) * scaleY), snapshot.height - 1)
      );

      if (localX >= 0 && localX < snapshot.width && localY >= 0 && localY < snapshot.height) {
        const idx = (localY * snapshot.width + localX) * 4;
        const r = snapshot.imageData.data[idx];
        const g = snapshot.imageData.data[idx + 1];
        const b = snapshot.imageData.data[idx + 2];
        const a = snapshot.imageData.data[idx + 3] / 255;

        if (a > 0) {
          const hex = ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
          return { hex, opacity: a };
        }
      }
    }

    // Priority 2: Fall back to DOM traversal for solid colors
    let current = node.parentElement;
    while (current) {
      const style = window.getComputedStyle(current);
      const color = parseColor(style.backgroundColor);
      if (color.hex && color.opacity > 0.5) {
        return color;
      }
      current = current.parentElement;
    }

    return { hex: 'FFFFFF', opacity: 1 };
  } catch (e) {
    return { hex: 'FFFFFF', opacity: 1 };
  }
}
