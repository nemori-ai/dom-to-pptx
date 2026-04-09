// src/utils/style.js
// General style parsing utilities

import { parseColor } from './color.js';
import { PX_TO_PT, PX_TO_INCH } from './constants.js';

/**
 * Extracts padding values from computed style.
 * Returns both individual values (for position/size adjustment) and uniform inset (for PptxGenJS).
 * PptxGenJS only supports a single inset value, so we return both formats.
 * @returns {Object} { top, right, bottom, left, inset } in inches (scaled)
 */
export function getPadding(style, scale) {
  const top = (parseFloat(style.paddingTop) || 0) * PX_TO_INCH * scale;
  const right = (parseFloat(style.paddingRight) || 0) * PX_TO_INCH * scale;
  const bottom = (parseFloat(style.paddingBottom) || 0) * PX_TO_INCH * scale;
  const left = (parseFloat(style.paddingLeft) || 0) * PX_TO_INCH * scale;

  // PptxGenJS only supports uniform inset, use minimum to avoid text overflow
  // Position/size will be adjusted to compensate for non-uniform padding
  const inset = Math.min(top, right, bottom, left);

  return { top, right, bottom, left, inset };
}

/**
 * Extracts blur value from CSS filter for soft edges.
 */
export function getSoftEdges(filterStr, scale) {
  if (!filterStr || filterStr === 'none') return null;
  const match = filterStr.match(/blur\(([\d.]+)px\)/);
  if (match) return parseFloat(match[1]) * PX_TO_PT * scale;
  return null;
}

/**
 * Detects if a CSS transform contains a skew component.
 * PPTX has no native skew support, so skewed elements need canvas capture.
 *
 * In a 2D matrix(a, b, c, d, e, f):
 *   Pure rotation: b = sin(θ), c = -sin(θ)  → b + c = 0
 *   With skew:     b + c ≠ 0
 * In a 3D matrix3d: check m21 (values[4]) vs m12 (values[1]) similarly.
 */
export function hasSkewTransform(transformStr) {
  if (!transformStr || transformStr === 'none') return false;

  const matrix3dMatch = transformStr.match(/matrix3d\(([^)]+)\)/);
  if (matrix3dMatch) {
    const v = matrix3dMatch[1].split(',').map((s) => parseFloat(s.trim()));
    if (v.length >= 6) {
      // m12 = v[1], m21 = v[4]; for pure rotation+scale b+c ≈ 0
      return Math.abs(v[1] + v[4]) > 0.001;
    }
    return false;
  }

  const matrixMatch = transformStr.match(/matrix\(([^)]+)\)/);
  if (matrixMatch) {
    const v = matrixMatch[1].split(',').map((s) => parseFloat(s.trim()));
    if (v.length >= 4) {
      const b = v[1];
      const c = v[2];
      // For pure rotation: b = sin(θ), c = -sin(θ) → b + c = 0
      // For skew: b and c are independent → b + c ≠ 0
      return Math.abs(b + c) > 0.001;
    }
  }

  return false;
}

/**
 * Extracts rotation angle from CSS transform matrix (2D or 3D).
 * For matrix3d, extracts Z-axis rotation from the 3D rotation matrix.
 */
export function getRotation(transformStr) {
  if (!transformStr || transformStr === 'none') return 0;

  const matrix3dMatch = transformStr.match(/matrix3d\(([^)]+)\)/);
  if (matrix3dMatch) {
    const values = matrix3dMatch[1].split(',').map((v) => parseFloat(v.trim()));
    if (values.length >= 6) {
      // matrix3d: m11=values[0], m12=values[1], m21=values[4], m22=values[5]
      // Z-axis rotation: atan2(m12, m11)
      const m11 = values[0];
      const m12 = values[1];
      return Math.round(Math.atan2(m12, m11) * (180 / Math.PI));
    }
    return 0;
  }

  const matrixMatch = transformStr.match(/matrix\(([^)]+)\)/);
  if (matrixMatch) {
    const values = matrixMatch[1].split(',').map((v) => parseFloat(v.trim()));
    if (values.length >= 4) {
      // matrix(a, b, c, d, e, f) where:
      // a = scaleX * cos(θ), b = scaleX * sin(θ)
      // c = -scaleY * sin(θ), d = scaleY * cos(θ)
      const a = values[0];
      const b = values[1];
      const c = values[2];
      const d = values[3];

      // Detect pure scale (no rotation): b=0 and c=0
      // For flip transforms, a or d will be negative but b and c are 0
      if (b === 0 && c === 0) {
        return 0;
      }

      // Extract rotation, accounting for negative scale
      const scaleX = Math.sign(a) || 1;
      return Math.round(Math.atan2(b * scaleX, a * scaleX) * (180 / Math.PI));
    }
  }

  return 0;
}

/**
 * Extracts visible shadow from box-shadow CSS.
 * Converts CSS Cartesian (x, y, blur) to PowerPoint Polar (angle, distance).
 */
export function getVisibleShadow(shadowStr, scale) {
  if (!shadowStr || shadowStr === 'none') return null;
  const shadows = shadowStr.split(/,(?![^()]*\))/);
  for (let s of shadows) {
    s = s.trim();
    if (s.startsWith('rgba(0, 0, 0, 0)')) continue;
    const match = s.match(
      /(rgba?\([^)]+\)|#[0-9a-fA-F]+)\s+(-?[\d.]+)px\s+(-?[\d.]+)px\s+([\d.]+)px/
    );
    if (match) {
      const colorStr = match[1];
      const x = parseFloat(match[2]);
      const y = parseFloat(match[3]);
      const blur = parseFloat(match[4]);
      const distance = Math.sqrt(x * x + y * y);
      // Skip zero-offset zero-blur shadows (e.g., Tailwind ring utilities)
      // These are invisible in PPTX and would prevent real shadows from being found
      if (distance === 0 && blur === 0) continue;
      let angle = Math.atan2(y, x) * (180 / Math.PI);
      if (angle < 0) angle += 360;
      const colorObj = parseColor(colorStr);
      return {
        type: 'outer',
        // PptxGenJS uses `|| 270` / `|| 4` fallbacks that clobber 0 values.
        // Use tiny non-zero values to prevent falsy-coercion while keeping visual fidelity.
        angle: angle || 0.01,
        blur: blur * PX_TO_PT * scale,
        offset: distance * PX_TO_PT * scale || 0.01,
        color: colorObj.hex || '000000',
        opacity: colorObj.opacity,
      };
    }
  }
  return null;
}

/**
 * Extracts CSS ring (box-shadow with spread-only, no offset/blur) for rendering
 * as a separate outline shape. Returns { color, opacity, spread } or null.
 */
export function getRingShadow(shadowStr) {
  if (!shadowStr || shadowStr === 'none') return null;
  const shadows = shadowStr.split(/,(?![^()]*\))/);
  for (let s of shadows) {
    s = s.trim();
    if (s.startsWith('rgba(0, 0, 0, 0)') || s.startsWith('rgb(255, 255, 255) 0px 0px 0px 0px'))
      continue;
    // Match: color 0px 0px 0px Npx (zero offset, zero blur, non-zero spread)
    const match = s.match(/(rgba?\([^)]+\)|#[0-9a-fA-F]+)\s+0px\s+0px\s+0px\s+([\d.]+)px/);
    if (match) {
      const spread = parseFloat(match[2]);
      if (spread <= 0) continue;
      const colorObj = parseColor(match[1]);
      if (!colorObj.hex || colorObj.opacity === 0) continue;
      return { color: colorObj.hex, opacity: colorObj.opacity, spread };
    }
  }
  return null;
}

/**
 * Checks if any parent element has overflow: hidden which would clip this element.
 */
export function isClippedByParent(node) {
  let parent = node.parentElement;
  while (parent && parent !== document.body) {
    const style = window.getComputedStyle(parent);
    const overflow = style.overflow;
    if (overflow === 'hidden' || overflow === 'clip') {
      return true;
    }
    parent = parent.parentElement;
  }
  return false;
}

/**
 * Find the nearest ancestor with overflow: hidden/clip.
 * Returns the ancestor element or null if none found.
 */
export function getClippingAncestor(node) {
  let parent = node.parentElement;
  while (parent && parent !== document.body) {
    const style = window.getComputedStyle(parent);
    const overflow = style.overflow;
    if (overflow === 'hidden' || overflow === 'clip') {
      return parent;
    }
    parent = parent.parentElement;
  }
  return null;
}

/**
 * Check if an element is completely outside the bounds of its clipping ancestor.
 * Returns true if the element should be skipped (fully clipped).
 */
export function isFullyClipped(node) {
  const clipAncestor = getClippingAncestor(node);
  if (!clipAncestor) return false;

  const nodeRect = node.getBoundingClientRect();
  const clipRect = clipAncestor.getBoundingClientRect();

  // Check if completely outside
  if (
    nodeRect.right <= clipRect.left ||
    nodeRect.left >= clipRect.right ||
    nodeRect.bottom <= clipRect.top ||
    nodeRect.top >= clipRect.bottom
  ) {
    return true;
  }

  return false;
}

/**
 * Check if an element is partially clipped by its ancestor.
 * Returns clipping info or null if not clipped.
 */
export function getClipInfo(node) {
  const clipAncestor = getClippingAncestor(node);
  if (!clipAncestor) return null;

  const nodeRect = node.getBoundingClientRect();
  const clipRect = clipAncestor.getBoundingClientRect();

  // Check if any part is outside
  const isClippedLeft = nodeRect.left < clipRect.left;
  const isClippedRight = nodeRect.right > clipRect.right;
  const isClippedTop = nodeRect.top < clipRect.top;
  const isClippedBottom = nodeRect.bottom > clipRect.bottom;

  if (isClippedLeft || isClippedRight || isClippedTop || isClippedBottom) {
    return {
      ancestor: clipAncestor,
      clipRect,
      nodeRect,
      isClippedLeft,
      isClippedRight,
      isClippedTop,
      isClippedBottom,
    };
  }

  return null;
}

/**
 * Resolves user imageScale config into a normalized object.
 * Supports both simple number and detailed object config.
 * @param {number|Object} userConfig - User's imageScale configuration
 * @returns {Object} - Normalized scale config { html2canvas, svg, img, canvas }
 */
export function resolveImageScale(userConfig) {
  const DEFAULT = { html2canvas: 3, svg: 3, img: 2, canvas: 2, maxScale: 4 };

  // Simple number mode: apply same scale to all
  if (typeof userConfig === 'number') {
    const clamped = Math.min(Math.max(userConfig, 1), DEFAULT.maxScale);
    return { html2canvas: clamped, svg: clamped, img: clamped, canvas: clamped };
  }

  // Object mode: merge with defaults and apply clamping
  if (typeof userConfig === 'object' && userConfig !== null) {
    const merged = { ...DEFAULT, ...userConfig };
    const clamp = (val) => Math.min(Math.max(val, 1), merged.maxScale);
    return {
      html2canvas: clamp(merged.html2canvas ?? merged.default ?? DEFAULT.html2canvas),
      svg: clamp(merged.svg ?? merged.default ?? DEFAULT.svg),
      img: clamp(merged.img ?? merged.default ?? DEFAULT.img),
      canvas: clamp(merged.canvas ?? merged.default ?? DEFAULT.canvas),
    };
  }

  // No config provided: return defaults
  return DEFAULT;
}

/**
 * Detect flip transforms from CSS transform string.
 * Handles scaleX(-1), scaleY(-1), scale(-1, 1), rotateY(180deg), rotateX(180deg),
 * and computed matrix forms like matrix(-1, 0, 0, 1, 0, 0).
 */
export function getFlip(transformStr) {
  if (!transformStr || transformStr === 'none') {
    return { flipH: false, flipV: false };
  }

  let flipH = false;
  let flipV = false;

  if (/scaleX\(\s*-1\s*\)/.test(transformStr) || /rotateY\(\s*180deg\s*\)/.test(transformStr)) {
    flipH = true;
  }

  if (/scaleY\(\s*-1\s*\)/.test(transformStr) || /rotateX\(\s*180deg\s*\)/.test(transformStr)) {
    flipV = true;
  }

  const scaleMatch = transformStr.match(/scale\(\s*(-?[\d.]+)\s*,\s*(-?[\d.]+)\s*\)/);
  if (scaleMatch) {
    if (parseFloat(scaleMatch[1]) < 0) flipH = true;
    if (parseFloat(scaleMatch[2]) < 0) flipV = true;
  }

  // Handle computed matrix form: matrix(a, b, c, d, e, f) where a=scaleX, d=scaleY
  const matrixMatch = transformStr.match(
    /matrix\(\s*(-?[\d.]+)\s*,\s*[\d.-]+\s*,\s*[\d.-]+\s*,\s*(-?[\d.]+)\s*,/
  );
  if (matrixMatch) {
    const scaleX = parseFloat(matrixMatch[1]);
    const scaleY = parseFloat(matrixMatch[2]);
    if (scaleX < 0) flipH = true;
    if (scaleY < 0) flipV = true;
  }

  return { flipH, flipV };
}
