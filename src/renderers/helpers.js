// src/renderers/helpers.js
// Shared helper functions for renderers

import { snapdom } from '@zumer/snapdom';
import { getTextStyle } from '../utils/text.js';
import { parseColor } from '../utils/color.js';
import { generateGradientSVG } from '../utils/svg.js';
import { getBorderInfo, generateCompositeBorderSVG } from '../utils/border.js';
import { LAYER } from '../utils/stacking.js';
import { PX_TO_INCH, PX_TO_PT } from '../utils/constants.js';

// ---------------------------------------------------------------------------
// CSS Counter resolution
// ---------------------------------------------------------------------------

/**
 * Parse a counter-reset or counter-increment CSS property value into pairs.
 * E.g. "steps 0 list-item 0" → [{ name: 'steps', value: 0 }, { name: 'list-item', value: 0 }]
 * @param {string} prop - The computed CSS counter-reset or counter-increment value
 * @param {number} defaultValue - Default value when not specified (0 for reset, 1 for increment)
 * @returns {Array<{ name: string, value: number }>}
 */
function parseCounterProp(prop, defaultValue) {
  if (!prop || prop === 'none') return [];
  const result = [];
  const parts = prop.trim().split(/\s+/);
  let i = 0;
  while (i < parts.length) {
    const name = parts[i];
    if (name === 'none') { i++; continue; }
    const next = parts[i + 1];
    if (next !== undefined && /^-?\d+$/.test(next)) {
      result.push({ name, value: parseInt(next) });
      i += 2;
    } else {
      result.push({ name, value: defaultValue });
      i++;
    }
  }
  return result;
}

/**
 * Resolve a CSS counter value for a specific element by walking the DOM in
 * document order (DFS), tracking counter-reset and counter-increment.
 *
 * @param {HTMLElement} targetElement - The element whose counter value we want
 * @param {string} counterName - Name of the counter (e.g. 'steps')
 * @returns {number} The resolved counter value
 */
function resolveCounterValue(targetElement, counterName) {
  let value = 0;
  let found = false;

  function walk(el) {
    if (found) return;

    const style = window.getComputedStyle(el);

    // counter-reset
    for (const { name, value: resetVal } of parseCounterProp(style.counterReset, 0)) {
      if (name === counterName) value = resetVal;
    }

    // counter-increment (applied after reset on the same element)
    for (const { name, value: incVal } of parseCounterProp(style.counterIncrement, 1)) {
      if (name === counterName) value += incVal;
    }

    if (el === targetElement) { found = true; return; }

    for (const child of el.children) {
      if (found) return;
      walk(child);
    }
  }

  walk(targetElement.ownerDocument.documentElement);
  return value;
}

/**
 * Resolve counter() functions in a CSS content value.
 * E.g. 'counter(steps)' → '"1"', '"Step " counter(steps)' → '"Step " "1"'
 *
 * @param {string} contentValue - The raw CSS content property value from getComputedStyle
 * @param {HTMLElement} element - The element the pseudo-element belongs to
 * @returns {string} The content value with counter() functions replaced
 */
function resolveContentCounters(contentValue, element) {
  if (!contentValue || !contentValue.includes('counter(')) return contentValue;

  return contentValue.replace(
    /counter\(\s*([\w-]+)(?:\s*,\s*[\w-]+)?\s*\)/g,
    (_, counterName) => {
      const val = resolveCounterValue(element, counterName);
      return `"${val}"`;
    }
  );
}

// ---------------------------------------------------------------------------

let concurrencyLimit = 10;
let activeCount = 0;
const waitQueue = [];

/** Update the concurrency limit for canvas/snapshot operations. */
export function setConcurrencyLimit(n) {
  if (n > 0) concurrencyLimit = n;
}

/** Reset concurrency state between exports to prevent stale leakage. */
export function resetConcurrency() {
  activeCount = 0;
  waitQueue.length = 0;
}

export async function limitConcurrency(fn) {
  if (activeCount >= concurrencyLimit) {
    // Wait for a slot to open
    await new Promise((resolve) => waitQueue.push(resolve));
  }

  activeCount++;
  try {
    return await fn();
  } finally {
    activeCount--;
    // Release next waiting task
    if (waitQueue.length > 0) {
      const next = waitQueue.shift();
      next();
    }
  }
}

/**
 * Captures a snapshot of the entire root element for background color sampling.
 * Returns an object with ImageData and bounding rect for pixel lookup.
 */
export async function captureBackgroundSnapshot(root) {
  return limitConcurrency(async () => {
    try {
      const rect = root.getBoundingClientRect();
      if (rect.width === 0 || rect.height === 0) return null;

      const result = await snapdom(root, {
        scale: 1,
        backgroundColor: 'transparent',
        embedFonts: true,
        embedImages: true,
      });
      const canvas = await result.toCanvas();

      const ctx = canvas.getContext('2d', { willReadFrequently: true });
      const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);

      return { imageData, rect, width: canvas.width, height: canvas.height };
    } catch (e) {
      console.warn('[dom-to-pptx] Background snapshot failed:', e);
      return null;
    }
  });
}

/**
 * Samples a pixel color from a background snapshot.
 *
 * @param {{ imageData: ImageData, rect: DOMRect, width: number, height: number }} snapshot
 * @param {number} pageX - X coordinate relative to viewport
 * @param {number} pageY - Y coordinate relative to viewport
 * @returns {{ hex: string, opacity: number } | null}
 */
export function sampleFromSnapshot(snapshot, pageX, pageY) {
  if (!snapshot?.imageData) return null;

  // Convert page (CSS) coordinates to snapshot pixel coordinates.
  // The snapshot canvas may be larger than the CSS rect (e.g. 2x on Retina displays).
  const scaleX = snapshot.width / snapshot.rect.width;
  const scaleY = snapshot.height / snapshot.rect.height;
  const localX = Math.max(
    0,
    Math.min(Math.round((pageX - snapshot.rect.left) * scaleX), snapshot.width - 1)
  );
  const localY = Math.max(
    0,
    Math.min(Math.round((pageY - snapshot.rect.top) * scaleY), snapshot.height - 1)
  );

  const idx = (localY * snapshot.width + localX) * 4;
  const r = snapshot.imageData.data[idx];
  const g = snapshot.imageData.data[idx + 1];
  const b = snapshot.imageData.data[idx + 2];
  const a = snapshot.imageData.data[idx + 3] / 255;

  if (a === 0) return null;

  const hex = ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
  return { hex, opacity: a };
}

/**
 * Samples the average color from a rectangular region of a snapshot.
 * Computes alpha-weighted mean of all non-transparent pixels in the region.
 * Used for backdrop-filter: produces a color that approximates the real blur effect.
 *
 * @param {{ imageData: ImageData, rect: DOMRect, width: number, height: number }} snapshot
 * @param {DOMRect} elementRect - The bounding rect of the element to sample behind
 * @returns {{ hex: string, opacity: number } | null}
 */
export function sampleAverageColor(snapshot, elementRect) {
  if (!snapshot?.imageData) return null;

  const { imageData, rect: snapshotRect, width, height } = snapshot;
  const scaleX = width / snapshotRect.width;
  const scaleY = height / snapshotRect.height;

  // Convert element rect to snapshot pixel coordinates
  const startX = Math.max(0, Math.round((elementRect.left - snapshotRect.left) * scaleX));
  const startY = Math.max(0, Math.round((elementRect.top - snapshotRect.top) * scaleY));
  const endX = Math.min(width, Math.round((elementRect.right - snapshotRect.left) * scaleX));
  const endY = Math.min(height, Math.round((elementRect.bottom - snapshotRect.top) * scaleY));

  if (startX >= endX || startY >= endY) return null;

  let totalR = 0,
    totalG = 0,
    totalB = 0,
    totalA = 0;
  let count = 0;

  for (let y = startY; y < endY; y++) {
    for (let x = startX; x < endX; x++) {
      const idx = (y * width + x) * 4;
      const a = imageData.data[idx + 3];
      if (a > 0) {
        const weight = a / 255;
        totalR += imageData.data[idx] * weight;
        totalG += imageData.data[idx + 1] * weight;
        totalB += imageData.data[idx + 2] * weight;
        totalA += weight;
        count++;
      }
    }
  }

  if (totalA === 0 || count === 0) return null;

  const r = Math.round(totalR / totalA);
  const g = Math.round(totalG / totalA);
  const b = Math.round(totalB / totalA);
  const avgOpacity = totalA / count;

  const hex = ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
  return { hex, opacity: avgOpacity };
}

export async function captureGradientText(node, widthPx, heightPx, imageScale = 3) {
  return limitConcurrency(async () => {
    try {
      const result = await snapdom(node, {
        scale: imageScale,
        backgroundColor: 'transparent',
        embedFonts: true,
        embedImages: true,
      });
      const canvas = await result.toCanvas();
      return canvas.toDataURL('image/png');
    } catch (e) {
      console.warn('[dom-to-pptx] snapdom gradient capture failed:', e);
      return null;
    }
  });
}

/**
 * Captures an element as PNG using snapdom.
 * Handles border-radius clipping in post-processing.
 */
export async function elementToCanvasImage(node, widthPx, heightPx, imageScale = 3) {
  return limitConcurrency(async () => {
    try {
      const width = Math.max(Math.ceil(widthPx), 1);
      const height = Math.max(Math.ceil(heightPx), 1);
      const style = window.getComputedStyle(node);

      const result = await snapdom(node, {
        scale: imageScale,
        backgroundColor: 'transparent',
        embedFonts: true,
        embedImages: true,
      });
      const sourceCanvas = await result.toCanvas();

      // Use the full resolution from snapdom (imageScale × CSS pixels)
      const scaledW = sourceCanvas.width;
      const scaledH = sourceCanvas.height;

      const destCanvas = document.createElement('canvas');
      destCanvas.width = scaledW;
      destCanvas.height = scaledH;
      const ctx = destCanvas.getContext('2d');

      ctx.drawImage(sourceCanvas, 0, 0);

      // Border radius clipping (scale radii to match canvas resolution)
      let tl = (parseFloat(style.borderTopLeftRadius) || 0) * imageScale;
      let tr = (parseFloat(style.borderTopRightRadius) || 0) * imageScale;
      let br = (parseFloat(style.borderBottomRightRadius) || 0) * imageScale;
      let bl = (parseFloat(style.borderBottomLeftRadius) || 0) * imageScale;

      const f = Math.min(
        scaledW / (tl + tr) || Infinity,
        scaledH / (tr + br) || Infinity,
        scaledW / (br + bl) || Infinity,
        scaledH / (bl + tl) || Infinity
      );

      if (f < 1) {
        tl *= f;
        tr *= f;
        br *= f;
        bl *= f;
      }

      if (tl + tr + br + bl > 0) {
        ctx.globalCompositeOperation = 'destination-in';
        ctx.beginPath();
        ctx.moveTo(tl, 0);
        ctx.lineTo(scaledW - tr, 0);
        ctx.arcTo(scaledW, 0, scaledW, tr, tr);
        ctx.lineTo(scaledW, scaledH - br);
        ctx.arcTo(scaledW, scaledH, scaledW - br, scaledH, br);
        ctx.lineTo(bl, scaledH);
        ctx.arcTo(0, scaledH, 0, scaledH - bl, bl);
        ctx.lineTo(0, tl);
        ctx.arcTo(0, 0, tl, 0, tl);
        ctx.closePath();
        ctx.fill();
      }

      return destCanvas.toDataURL('image/png');
    } catch (e) {
      console.warn('[dom-to-pptx] snapdom capture failed:', e);
      return null;
    }
  });
}

/**
 * Helper to identify elements that should be rendered as icons (Images).
 * Detects Custom Elements AND generic tags (<i>, <span>) with icon classes/pseudo-elements.
 */
export function isIconElement(node) {
  // 1. Custom Elements (hyphenated tags) or Explicit Library Tags
  const tag = node.tagName.toUpperCase();
  if (
    tag.includes('-') ||
    [
      'MATERIAL-ICON',
      'ICONIFY-ICON',
      'REMIX-ICON',
      'ION-ICON',
      'EVA-ICON',
      'BOX-ICON',
      'FA-ICON',
    ].includes(tag)
  ) {
    return true;
  }

  // 2. Class-based Icons (FontAwesome, Bootstrap, Material symbols) on <i> or <span>
  if (tag === 'I' || tag === 'SPAN') {
    const cls = node.getAttribute('class') || '';
    if (
      typeof cls === 'string' &&
      (cls.includes('fa-') ||
        cls.includes('fas') ||
        cls.includes('far') ||
        cls.includes('fab') ||
        cls.includes('bi-') ||
        cls.includes('material-icons') ||
        cls.includes('icon'))
    ) {
      // Double-check: Must have pseudo-element content to be a CSS icon
      const before = window.getComputedStyle(node, '::before').content;
      const after = window.getComputedStyle(node, '::after').content;
      const hasContent = (c) => c && c !== 'none' && c !== 'normal' && c !== '""';

      if (hasContent(before) || hasContent(after)) return true;
    }
  }

  return false;
}

/**
 * Check if a list (UL/OL) has complex hierarchy that cannot be rendered natively.
 */
export function isComplexHierarchy(root) {
  // Use a simple tree traversal to find forbidden elements in the list structure
  const stack = [root];
  while (stack.length > 0) {
    const el = stack.pop();

    // 1. Layouts or visual styling on LIs that would be lost by flattening
    if (el.tagName === 'LI') {
      const s = window.getComputedStyle(el);
      if (s.display === 'flex' || s.display === 'grid' || s.display === 'inline-flex') return true;
      const liBg = s.backgroundColor;
      const hasBg = liBg && liBg !== 'rgba(0, 0, 0, 0)' && liBg !== 'transparent';
      const hasBorder =
        parseFloat(s.borderLeftWidth) > 0 ||
        parseFloat(s.borderTopWidth) > 0 ||
        parseFloat(s.borderRightWidth) > 0 ||
        parseFloat(s.borderBottomWidth) > 0;
      const hasRadius = parseFloat(s.borderRadius) > 0;
      if (hasBg || hasBorder || hasRadius) return true;

      // Positioned pseudo-elements (CSS counter styling, custom bullets, decorative lines)
      // cannot be represented by flat list text — require ShapeRenderer path
      for (const pseudo of ['::before', '::after']) {
        const ps = window.getComputedStyle(el, pseudo);
        if (!ps.content || ps.content === 'none' || ps.content === 'normal') continue;
        if (ps.display === 'none') continue;
        if (ps.position === 'absolute' || ps.position === 'relative') return true;
      }
    }

    // 2. Media / Icons
    if (['IMG', 'SVG', 'CANVAS', 'VIDEO', 'IFRAME'].includes(el.tagName)) return true;
    if (isIconElement(el)) return true;

    // 3. Nested Lists (Flattening logic doesn't support nested bullets well yet)
    if (el !== root && (el.tagName === 'UL' || el.tagName === 'OL')) return true;

    // Recurse, but don't go too deep if not needed
    for (let i = 0; i < el.children.length; i++) {
      stack.push(el.children[i]);
    }
  }
  return false;
}

/**
 * Extract hyperlink info from an anchor element.
 * Returns { url, tooltip } or null if not a valid link.
 */
function getHyperlinkInfo(node) {
  if (node.tagName !== 'A') return null;
  const href = node.getAttribute('href');
  if (!href || href.startsWith('#') || href.startsWith('javascript:')) return null;

  return {
    url: href,
    tooltip: node.getAttribute('title') || null,
  };
}

/**
 * Collect text parts from a list item, handling nested elements.
 * Now supports hyperlinks (<a> tags).
 * @param {Node} node - DOM node to collect from
 * @param {CSSStyleDeclaration} parentStyle - Parent style for inheritance
 * @param {number} scale - Layout scale factor
 * @param {Object} [globalOptions={}] - Global options (backgroundSnapshot, etc.)
 */
export function collectListParts(node, parentStyle, scale, globalOptions = {}) {
  const parts = [];

  // Check for CSS Content (::before) - often used for icons
  if (node.nodeType === 1) {
    const beforeStyle = window.getComputedStyle(node, '::before');
    const content = resolveContentCounters(beforeStyle.content, node);
    if (content && content !== 'none' && content !== 'normal' && content !== '""') {
      // Strip quotes
      const cleanContent = content.replace(/^['"]|['"]$/g, '');
      if (cleanContent.trim()) {
        parts.push({
          text: cleanContent + ' ', // Add space after icon
          options: getTextStyle(
            window.getComputedStyle(node),
            scale,
            cleanContent,
            node,
            globalOptions
          ),
        });
      }
    }
  }

  node.childNodes.forEach((child) => {
    if (child.nodeType === 3) {
      // Text
      const val = child.nodeValue.replace(/[\n\r\t]+/g, ' ').replace(/\s{2,}/g, ' ');
      if (val) {
        // Use parent style if child is text node, otherwise current style
        const styleToUse = node.nodeType === 1 ? window.getComputedStyle(node) : parentStyle;
        // For font inheritance, use the element node (text node's parent)
        const fontNode = node.nodeType === 1 ? node : child.parentElement;
        parts.push({
          text: val,
          options: getTextStyle(styleToUse, scale, val, fontNode, globalOptions),
        });
      }
    } else if (child.nodeType === 1) {
      // Element (span, i, b, a)
      // Check for hyperlink
      const linkInfo = getHyperlinkInfo(child);
      if (linkInfo) {
        // For <a> tags, collect inner text with hyperlink attached
        const innerParts = collectListParts(child, parentStyle, scale, globalOptions);
        innerParts.forEach((part) => {
          part.options.hyperlink = linkInfo;
        });
        parts.push(...innerParts);
      } else {
        // Recurse normally
        parts.push(...collectListParts(child, parentStyle, scale, globalOptions));
      }
    }
  });

  return parts;
}

export function createCompositeBorderItems(sides, x, y, w, h, scale, domOrder) {
  const items = [];
  const common = { type: 'shape', layer: LAYER.BORDER, domOrder, shapeType: 'rect' };

  const fillFor = (side) => {
    const f = { color: side.color };
    if (side.opacity < 1) f.transparency = Math.round((1 - side.opacity) * 100);
    return f;
  };

  if (sides.top.width > 0)
    items.push({
      ...common,
      options: { x, y, w, h: sides.top.width * PX_TO_INCH * scale, fill: fillFor(sides.top) },
    });
  if (sides.right.width > 0)
    items.push({
      ...common,
      options: {
        x: x + w - sides.right.width * PX_TO_INCH * scale,
        y,
        w: sides.right.width * PX_TO_INCH * scale,
        h,
        fill: fillFor(sides.right),
      },
    });
  if (sides.bottom.width > 0)
    items.push({
      ...common,
      options: {
        x,
        y: y + h - sides.bottom.width * PX_TO_INCH * scale,
        w,
        h: sides.bottom.width * PX_TO_INCH * scale,
        fill: fillFor(sides.bottom),
      },
    });
  if (sides.left.width > 0)
    items.push({
      ...common,
      options: {
        x,
        y,
        w: sides.left.width * PX_TO_INCH * scale,
        h,
        fill: fillFor(sides.left),
      },
    });

  return items;
}

/**
 * Detect and render visible ::before / ::after pseudo-elements as PPTX items.
 *
 * Handles two common patterns:
 *   A) Solid/gradient background bars (e.g. accent stripes, decorative bars)
 *   B) Border-triangles (CSS arrow trick using transparent borders)
 *
 * @param {HTMLElement} node - The parent DOM element
 * @param {Object} config - Layout config (rootX, rootY, offX, offY, scale)
 * @param {number} domOrder
 * @param {number} parentX - Parent element's PPTX x position (inches)
 * @param {number} parentY - Parent element's PPTX y position (inches)
 * @param {number} parentWidthPx
 * @param {number} parentHeightPx
 * @returns {Array} PPTX render items
 */
export function collectPseudoElementItems(
  node,
  config,
  domOrder,
  parentX,
  parentY,
  parentWidthPx,
  parentHeightPx
) {
  const items = [];
  const jobs = [];
  const pseudos = ['::before', '::after'];

  for (const pseudo of pseudos) {
    const ps = window.getComputedStyle(node, pseudo);

    if (!ps.content || ps.content === 'none' || ps.content === 'normal') continue;
    if (ps.display === 'none' || ps.visibility === 'hidden') continue;

    const position = ps.position;
    const isPositioned = position === 'absolute' || position === 'relative';
    const rawContent = resolveContentCounters(ps.content, node);
    const isInlineDecor =
      !isPositioned &&
      (ps.display === 'inline-block' || ps.display === 'block' || ps.display === 'flex') &&
      (rawContent === '""' || rawContent === "''");

    if (!isPositioned && !isInlineDecor) continue;

    const bTop = parseFloat(ps.borderTopWidth) || 0;
    const bRight = parseFloat(ps.borderRightWidth) || 0;
    const bBottom = parseFloat(ps.borderBottomWidth) || 0;
    const bLeft = parseFloat(ps.borderLeftWidth) || 0;

    const boxSizing = ps.boxSizing || 'content-box';
    let pseudoW =
      ps.width && ps.width.includes('%')
        ? (parseFloat(ps.width) / 100) * parentWidthPx
        : parseFloat(ps.width) || 0;
    let pseudoH =
      ps.height && ps.height.includes('%')
        ? (parseFloat(ps.height) / 100) * parentHeightPx
        : parseFloat(ps.height) || 0;
    if (pseudoW === 0) pseudoW = bLeft + bRight;
    if (pseudoH === 0) pseudoH = bTop + bBottom;
    if (pseudoW <= 0 && pseudoH <= 0) continue;

    if (boxSizing === 'content-box') {
      pseudoW += bLeft + bRight;
      pseudoH += bTop + bBottom;
    }

    let pxLeft = 0;
    let pxTop = 0;

    if (isPositioned) {
      const resolvePos = (val, parentDim) => {
        if (!val || val === 'auto') return null;
        if (val.includes('%')) return (parseFloat(val) / 100) * parentDim;
        return parseFloat(val) || 0;
      };
      const topResolved = resolvePos(ps.top, parentHeightPx);
      const leftResolved = resolvePos(ps.left, parentWidthPx);
      const rightResolved = resolvePos(ps.right, parentWidthPx);
      const bottomResolved = resolvePos(ps.bottom, parentHeightPx);

      pxLeft = leftResolved !== null ? leftResolved : 0;
      pxTop = topResolved !== null ? topResolved : 0;
      if (leftResolved === null && rightResolved !== null) {
        pxLeft = parentWidthPx - pseudoW - rightResolved;
      }
      if (topResolved === null && bottomResolved !== null) {
        pxTop = parentHeightPx - pseudoH - bottomResolved;
      }

      const transform = ps.transform;
      if (transform && transform !== 'none') {
        const matMatch = transform.match(/matrix\(([^)]+)\)/);
        if (matMatch) {
          const vals = matMatch[1].split(',').map((v) => parseFloat(v.trim()));
          if (vals.length >= 6) {
            pxLeft += vals[4];
            pxTop += vals[5];
          }
        }
      }
    } else if (isInlineDecor) {
      const parentStyle = window.getComputedStyle(node);
      const padLeft = parseFloat(parentStyle.paddingLeft) || 0;
      const padTop = parseFloat(parentStyle.paddingTop) || 0;
      const isFlex = parentStyle.display.includes('flex');
      const isRow = !parentStyle.flexDirection || parentStyle.flexDirection === 'row';
      const alignItems = parentStyle.alignItems || 'stretch';

      if (pseudo === '::before') {
        pxLeft = padLeft;
        if (isFlex && isRow) {
          if (alignItems === 'center') pxTop = (parentHeightPx - pseudoH) / 2;
          else if (alignItems === 'flex-end') pxTop = parentHeightPx - padTop - pseudoH;
          else pxTop = padTop;
        } else {
          pxTop = padTop;
        }
      } else {
        pxLeft = parentWidthPx - (parseFloat(parentStyle.paddingRight) || 0) - pseudoW;
        pxTop = padTop;
      }
    }

    const pX = parentX + pxLeft * PX_TO_INCH * config.scale;
    const pY = parentY + pxTop * PX_TO_INCH * config.scale;
    const pW = pseudoW * PX_TO_INCH * config.scale;
    const pH = pseudoH * PX_TO_INCH * config.scale;

    const bgColor = parseColor(ps.backgroundColor);
    const bgImage = ps.backgroundImage || '';
    const hasBg = (bgColor.hex && bgColor.opacity > 0) || bgImage.includes('gradient');
    const maskImg = ps.webkitMaskImage || ps.maskImage || 'none';
    const hasMask = maskImg !== 'none';

    if (hasMask && hasBg && pseudoW > 0 && pseudoH > 0) {
      // Extract the data URI from url("...") or url(...).
      // Chrome wraps in double quotes and escapes inner quotes as \".
      let svgUrl = '';
      const quotedMatch = maskImg.match(/url\("(.+?)"\)/);
      if (quotedMatch) {
        svgUrl = quotedMatch[1].replace(/\\"/g, '"');
      } else {
        const unquotedMatch = maskImg.match(/url\(([^)]+)\)/);
        if (unquotedMatch) svgUrl = unquotedMatch[1];
      }

      if (svgUrl && svgUrl.startsWith('data:image/svg+xml')) {
        let svgContent = '';
        if (svgUrl.includes(';utf8,'))
          svgContent = decodeURIComponent(svgUrl.split(';utf8,')[1]);
        else if (svgUrl.includes(';charset=utf-8,'))
          svgContent = decodeURIComponent(svgUrl.split(';charset=utf-8,')[1]);
        else if (svgUrl.includes(';base64,'))
          svgContent = atob(svgUrl.split(';base64,')[1]);
        else {
          const commaIdx = svgUrl.indexOf(',');
          if (commaIdx > -1) svgContent = decodeURIComponent(svgUrl.substring(commaIdx + 1));
        }

        if (svgContent) {
          const hexColor = bgColor.hex ? '#' + bgColor.hex : '#FFFFFF';
          const s = 4;
          const cw = pseudoW * s,
            ch = pseudoH * s;
          // Replace currentColor with the actual background color, set explicit size
          const maskSvg = svgContent
            .replace(/<svg\b/, `<svg width="${cw}" height="${ch}"`)
            .replace(/stroke="currentColor"/g, `stroke="${hexColor}"`)
            .replace(/fill="currentColor"/g, `fill="${hexColor}"`);

          const item = {
            type: 'image',
            layer: LAYER.CONTENT,
            domOrder: domOrder + 0.1,
            options: { data: null, x: pX, y: pY, w: pW, h: pH },
          };
          items.push(item);
          const capturedSvg = maskSvg;
          jobs.push(async () => {
            const img = new Image();
            const svgDataUrl =
              'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(capturedSvg);
            await new Promise((r) => {
              img.onload = r;
              img.onerror = r;
              img.src = svgDataUrl;
            });
            if (img.naturalWidth > 0) {
              const cvs = document.createElement('canvas');
              cvs.width = cw;
              cvs.height = ch;
              const ctx = cvs.getContext('2d');
              ctx.drawImage(img, 0, 0, cw, ch);
              item.options.data = cvs.toDataURL('image/png');
            }
            if (!item.options.data) item.skip = true;
          });
          continue;
        }
      }
      continue;
    }

    if (hasBg && pseudoW > 0 && pseudoH > 0) {
      const hasGradientBg = bgImage.includes('gradient');
      const hasTextContent = (() => {
        let t = rawContent.replace(/^['"]|['"]$/g, '');
        t = t.replace(/\\([\da-fA-F]{1,6})\s?/g, (_, h) => String.fromCodePoint(parseInt(h, 16)));
        t = t.replace(/\\(.)/g, '$1');
        return t && t !== '""' && t !== "''";
      })();

      if (hasGradientBg && !bgColor.hex && !hasTextContent) {
        const gradSvg = generateGradientSVG(pseudoW, pseudoH, bgImage, 0, null);
        if (gradSvg) {
          items.push({
            type: 'image',
            layer: LAYER.CONTENT,
            domOrder: domOrder + 0.1,
            options: { data: gradSvg, x: pX, y: pY, w: pW, h: pH },
          });
        }
        continue;
      }

      if (!hasTextContent) {
        const transparency = bgColor.hex ? (1 - bgColor.opacity) * 100 : 0;
        const brStr = ps.borderRadius || '';
        const isCircle = brStr === '50%' && Math.abs(pseudoW - pseudoH) < 1;
        const brPx = brStr.includes('%')
          ? (parseFloat(brStr) / 100) * Math.min(pseudoW, pseudoH)
          : parseFloat(brStr) || 0;

        const borderInfo = getBorderInfo(ps, config.scale);
        let lineOpt = { type: 'none' };
        if (borderInfo.type === 'uniform') {
          lineOpt = borderInfo.options;
        }

        let shapeType = 'rect';
        const shapeOpts = {
          x: pX,
          y: pY,
          w: pW,
          h: pH,
          fill: bgColor.hex ? { color: bgColor.hex, transparency } : { type: 'none' },
          line: lineOpt,
        };

        if (
          isCircle ||
          (brPx >= Math.min(pseudoW, pseudoH) / 2 - 0.5 && Math.abs(pseudoW - pseudoH) < 1)
        ) {
          shapeType = 'ellipse';
        } else if (brPx > 0) {
          shapeType = 'roundRect';
          shapeOpts.rectRadius = brPx * PX_TO_INCH * config.scale;
        }

        items.push({
          type: 'shape',
          layer: LAYER.CONTENT,
          domOrder: domOrder + 0.1,
          shapeType,
          options: shapeOpts,
        });

        if (borderInfo.type === 'composite') {
          const borderSvg = generateCompositeBorderSVG(pseudoW, pseudoH, brPx, borderInfo.sides);
          if (borderSvg) {
            items.push({
              type: 'image',
              layer: LAYER.OVERLAY,
              domOrder: domOrder + 0.3,
              options: { data: borderSvg, x: pX, y: pY, w: pW, h: pH },
            });
          }
        }
        continue;
      }
      // hasTextContent && hasBg: handled below as unified text element with fill/line
    }

    const hasBorderTriangle = bTop + bRight + bBottom + bLeft > 0 && !hasBg;
    if (hasBorderTriangle) {
      const cTop = parseColor(ps.borderTopColor);
      const cRight = parseColor(ps.borderRightColor);
      const cBottom = parseColor(ps.borderBottomColor);
      const cLeft = parseColor(ps.borderLeftColor);

      const sides = [
        { w: bTop, color: cTop, dir: 'top' },
        { w: bRight, color: cRight, dir: 'right' },
        { w: bBottom, color: cBottom, dir: 'bottom' },
        { w: bLeft, color: cLeft, dir: 'left' },
      ];
      const visibleSide = sides.find((s) => s.w > 0 && s.color.hex && s.color.opacity > 0);
      if (!visibleSide) continue;

      const totalW = bLeft + bRight;
      const totalH = bTop + bBottom;
      const hex = visibleSide.color.hex;
      const opacity = visibleSide.color.opacity;

      let points;
      switch (visibleSide.dir) {
        case 'right':
          points = `${totalW},0 ${bLeft},${bTop} ${totalW},${totalH}`;
          break;
        case 'left':
          points = `0,0 ${bLeft},${bTop} 0,${totalH}`;
          break;
        case 'top':
          points = `0,0 ${bLeft},${bTop} ${totalW},0`;
          break;
        case 'bottom':
          points = `0,${totalH} ${bLeft},${bTop} ${totalW},${totalH}`;
          break;
      }

      const svg =
        `data:image/svg+xml;base64,` +
        btoa(
          `<svg xmlns="http://www.w3.org/2000/svg" width="${totalW}" height="${totalH}">` +
            `<polygon points="${points}" fill="#${hex}" fill-opacity="${opacity}"/>` +
            `</svg>`
        );

      const svgW = totalW * PX_TO_INCH * config.scale;
      const svgH = totalH * PX_TO_INCH * config.scale;

      items.push({
        type: 'image',
        layer: LAYER.CONTENT,
        domOrder: domOrder + 0.1,
        options: { data: svg, x: pX, y: pY, w: svgW, h: svgH },
      });
      continue;
    }

    let text = rawContent.replace(/^['"]|['"]$/g, '');
    text = text.replace(/\\([\da-fA-F]{1,6})\s?/g, (_, hex) =>
      String.fromCodePoint(parseInt(hex, 16))
    );
    text = text.replace(/\\(.)/g, '$1');
    if (text && text !== '""' && text !== "''") {
      const textColor = parseColor(ps.color);
      const fontSize = parseFloat(ps.fontSize) || 14;

      // Bullet characters (•, ◦, ●, etc.) have unpredictable vertical
      // positioning within a text box. Replace single-character bullets
      // with a geometric ellipse shape for pixel-accurate alignment.
      const BULLET_CHARS = '\u2022\u25CF\u25CB\u25E6\u25AA\u25AB\u25A0\u25A1\u2023\u25B8\u25B9\u25B6\u25C0\u2013\u2014';
      if (text.length === 1 && BULLET_CHARS.includes(text) && isPositioned) {
        const parentLH = parseFloat(window.getComputedStyle(node).lineHeight);
        const lineH = !isNaN(parentLH) ? parentLH : fontSize * 1.4;
        // Size the dot relative to font size (~30%)
        const dotSize = fontSize * 0.3;
        const dotIn = dotSize * PX_TO_INCH * config.scale;
        // Centre the dot vertically within the parent's first line
        const dotX = pX + (pW - dotIn) / 2;
        const dotY = pY + (lineH * PX_TO_INCH * config.scale - dotIn) / 2;

        const hex = textColor.hex || 'FFFFFF';
        const transparency = textColor.opacity < 1 ? (1 - textColor.opacity) * 100 : 0;

        items.push({
          type: 'shape',
          layer: LAYER.CONTENT,
          domOrder: domOrder + 0.1,
          shapeType: 'ellipse',
          options: {
            x: dotX,
            y: dotY,
            w: dotIn,
            h: dotIn,
            fill: { color: hex, transparency },
            line: { type: 'none' },
          },
        });
      } else {
        // General text pseudo-element
        const textStyle = getTextStyle(ps, config.scale, text, node, {});
        if (textColor.hex) textStyle.color = textColor.hex;
        if (textColor.opacity < 1) textStyle.transparency = Math.round((1 - textColor.opacity) * 100);

        const padL = parseFloat(ps.paddingLeft) || 0;
        const padR = parseFloat(ps.paddingRight) || 0;
        const padT = parseFloat(ps.paddingTop) || 0;
        const padB = parseFloat(ps.paddingBottom) || 0;

        const estTextW = fontSize * text.length * PX_TO_INCH * config.scale;
        const estTextH = fontSize * 1.4 * PX_TO_INCH * config.scale;

        let textAlign = ps.textAlign || 'center';
        if (textAlign === 'start') textAlign = 'left';

        // Detect flex centering (common in counter-styled pseudo-elements)
        if (ps.display === 'flex' || ps.display === 'inline-flex') {
          const jc = ps.justifyContent;
          if (jc === 'center' || jc === 'space-around' || jc === 'space-evenly') {
            textAlign = 'center';
          } else if (jc === 'flex-end' || jc === 'end') {
            textAlign = 'right';
          }
        }

        const brStr = ps.borderRadius || '';
        const brPx = brStr.includes('%')
          ? (parseFloat(brStr) / 100) * Math.min(pseudoW, pseudoH)
          : parseFloat(brStr) || 0;

        const borderInfo = getBorderInfo(ps, config.scale);

        // Circle detection: border-radius creates a circle when w ≈ h and radius ≥ half side
        const isCircle =
          (brStr === '50%' || brPx >= Math.min(pseudoW, pseudoH) / 2 - 0.5) &&
          Math.abs(pseudoW - pseudoH) < 1;

        if (isCircle) {
          // Split into ellipse shape (bg + border) + transparent text overlay
          const bgTransparency = bgColor.hex ? (1 - bgColor.opacity) * 100 : 0;
          let lineOpt = { type: 'none' };
          if (borderInfo.type === 'uniform') lineOpt = borderInfo.options;

          items.push({
            type: 'shape',
            layer: LAYER.CONTENT,
            domOrder: domOrder + 0.09,
            shapeType: 'ellipse',
            options: {
              x: pX,
              y: pY,
              w: pW,
              h: pH,
              fill: bgColor.hex ? { color: bgColor.hex, transparency: bgTransparency } : { type: 'none' },
              line: lineOpt,
            },
          });

          items.push({
            type: 'text',
            layer: LAYER.CONTENT,
            domOrder: domOrder + 0.1,
            textParts: [{ text, options: textStyle }],
            options: {
              x: pX,
              y: pY,
              w: pW > 0 ? pW : estTextW,
              h: pH > 0 ? pH : estTextH,
              align: textAlign,
              valign: 'middle',
              margin: 0,
              wrap: false,
              autoFit: false,
              fill: { type: 'none' },
              line: { type: 'none' },
            },
          });
        } else {
          // Non-circle: unified text element with fill/border/rectRadius
          const textOpts = {
            x: pX,
            y: pY,
            w: pW > 0 ? pW : estTextW,
            h: pH > 0 ? pH : estTextH,
            align: textAlign,
            valign: 'middle',
            margin: [
              padT * PX_TO_PT * config.scale,
              padR * PX_TO_PT * config.scale,
              padB * PX_TO_PT * config.scale,
              padL * PX_TO_PT * config.scale,
            ],
            wrap: false,
            autoFit: false,
          };

          if (hasBg && bgColor.hex) {
            const transparency = (1 - bgColor.opacity) * 100;
            textOpts.fill = { color: bgColor.hex, transparency };
          }

          if (borderInfo.type === 'uniform') {
            textOpts.line = borderInfo.options;
          }

          if (brPx > 0) {
            textOpts.rectRadius = brPx * PX_TO_INCH * config.scale;
          }

          items.push({
            type: 'text',
            layer: LAYER.CONTENT,
            domOrder: domOrder + 0.1,
            textParts: [{ text, options: textStyle }],
            options: textOpts,
          });
        }

        if (borderInfo.type === 'composite') {
          const borderSvg = generateCompositeBorderSVG(pseudoW, pseudoH, brPx, borderInfo.sides);
          if (borderSvg) {
            items.push({
              type: 'image',
              layer: LAYER.OVERLAY,
              domOrder: domOrder + 0.3,
              options: { data: borderSvg, x: pX, y: pY, w: pW, h: pH },
            });
          }
        }
      }
    }
  }

  return { items, jobs };
}
