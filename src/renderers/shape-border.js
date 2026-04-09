// src/renderers/shape-border.js
// Border-radius computation and gradient border detection/rendering for ShapeRenderer

import { parseColor, getClippingAncestor } from '../utils/index.js';

/**
 * Compute border-radius values for an element, including inheritance
 * from clipping ancestors when the element has no own radius.
 *
 * @param {HTMLElement} node
 * @param {CSSStyleDeclaration} style
 * @param {number} widthPx
 * @param {number} heightPx
 * @returns {Object} { borderTopLeftRadius, borderTopRightRadius, borderBottomRightRadius, borderBottomLeftRadius, allCornersEqual, borderRadiusValue, hasPartialBorderRadius }
 */
export function computeBorderRadii(node, style, widthPx, heightPx) {
  const resolveRadius = (val, dimPx) => {
    if (!val) return 0;
    if (val.includes('%')) return (parseFloat(val) / 100) * dimPx;
    return parseFloat(val) || 0;
  };

  let borderTopLeftRadius = resolveRadius(style.borderTopLeftRadius, widthPx);
  let borderTopRightRadius = resolveRadius(style.borderTopRightRadius, widthPx);
  let borderBottomRightRadius = resolveRadius(style.borderBottomRightRadius, widthPx);
  let borderBottomLeftRadius = resolveRadius(style.borderBottomLeftRadius, widthPx);

  // CSS spec §5.5: when the sum of adjacent radii exceeds a side's length,
  // all radii are scaled down by the same factor so every side fits.
  // Only consider sides where both adjacent corners have a radius (sum > 0).
  const sums = [
    borderTopLeftRadius + borderTopRightRadius,     // top edge
    borderBottomLeftRadius + borderBottomRightRadius, // bottom edge
    borderTopLeftRadius + borderBottomLeftRadius,     // left edge
    borderTopRightRadius + borderBottomRightRadius,   // right edge
  ];
  const dims = [widthPx, widthPx, heightPx, heightPx];
  let f = 1;
  for (let i = 0; i < 4; i++) {
    if (sums[i] > 0) f = Math.min(f, dims[i] / sums[i]);
  }
  if (f < 1) {
    borderTopLeftRadius *= f;
    borderTopRightRadius *= f;
    borderBottomRightRadius *= f;
    borderBottomLeftRadius *= f;
  }

  // Helper for inherited radius clamping below (simple per-dimension limit)
  const clampRadius = (r) => Math.min(r, widthPx, heightPx);

  let allCornersEqual =
    borderTopLeftRadius === borderTopRightRadius &&
    borderTopRightRadius === borderBottomRightRadius &&
    borderBottomRightRadius === borderBottomLeftRadius;
  let borderRadiusValue = allCornersEqual ? borderTopLeftRadius : 0;

  const hasOwnRadius =
    borderTopLeftRadius > 0 ||
    borderTopRightRadius > 0 ||
    borderBottomRightRadius > 0 ||
    borderBottomLeftRadius > 0;

  // Inherit clipping ancestor's border-radius for edges that align
  // (PPTX has no parent-child clipping, so children must self-clip).
  // Walk past intermediate clipping ancestors that have no radius
  // (e.g. an overflow:hidden wrapper between the element and the
  // rounded-corner card) to find the one that actually clips corners.
  if (borderRadiusValue === 0 && !hasOwnRadius) {
    let clipAnc = getClippingAncestor(node);
    while (clipAnc) {
      const ps = window.getComputedStyle(clipAnc);
      const pTL = parseFloat(ps.borderTopLeftRadius) || 0;
      const pTR = parseFloat(ps.borderTopRightRadius) || 0;
      const pBR = parseFloat(ps.borderBottomRightRadius) || 0;
      const pBL = parseFloat(ps.borderBottomLeftRadius) || 0;

      if (pTL > 0 || pTR > 0 || pBR > 0 || pBL > 0) {
        const pRect = clipAnc.getBoundingClientRect();
        const nRect = node.getBoundingClientRect();
        const T = 2; // edge alignment tolerance in px

        const atLeft = Math.abs(nRect.left - pRect.left) < T;
        const atRight = Math.abs(nRect.right - pRect.right) < T;
        const atTop = Math.abs(nRect.top - pRect.top) < T;
        const atBottom = Math.abs(nRect.bottom - pRect.bottom) < T;

        // Only inherit if element is large enough to accommodate the radius
        // without extreme clamping (e.g. a 4px bar with 16px radius -> skip)
        const canFit = (r) => widthPx >= r && heightPx >= r;
        const iTL = atLeft && atTop && canFit(pTL) ? clampRadius(pTL) : 0;
        const iTR = atRight && atTop && canFit(pTR) ? clampRadius(pTR) : 0;
        const iBR = atRight && atBottom && canFit(pBR) ? clampRadius(pBR) : 0;
        const iBL = atLeft && atBottom && canFit(pBL) ? clampRadius(pBL) : 0;

        if (iTL > 0 || iTR > 0 || iBR > 0 || iBL > 0) {
          borderTopLeftRadius = iTL;
          borderTopRightRadius = iTR;
          borderBottomRightRadius = iBR;
          borderBottomLeftRadius = iBL;

          allCornersEqual = iTL === iTR && iTR === iBR && iBR === iBL;
          borderRadiusValue = allCornersEqual ? iTL : 0;
        }
        break; // Found ancestor with radius, stop searching
      }
      // No radius on this clipping ancestor, continue up
      clipAnc = getClippingAncestor(clipAnc);
    }
  }

  const hasPartialBorderRadius =
    !allCornersEqual &&
    (borderTopLeftRadius > 0 ||
      borderTopRightRadius > 0 ||
      borderBottomRightRadius > 0 ||
      borderBottomLeftRadius > 0);

  return {
    borderTopLeftRadius,
    borderTopRightRadius,
    borderBottomRightRadius,
    borderBottomLeftRadius,
    allCornersEqual,
    borderRadiusValue,
    hasPartialBorderRadius,
  };
}

/**
 * Detect CSS gradient border technique:
 * - backgroundImage has TWO linear-gradients (fill + border)
 * - border is transparent
 * - background shorthand contains both padding-box and border-box
 *
 * @param {CSSStyleDeclaration} style
 * @param {number} gradientCount
 * @returns {boolean}
 */
export function detectGradientBorder(style, gradientCount) {
  const bgShorthand = style.background || '';
  const hasTwoGradients = gradientCount >= 2;
  const hasBorderBoxTechnique =
    bgShorthand.includes('padding-box') && bgShorthand.includes('border-box');
  const hasTransparentBorder =
    style.borderStyle !== 'none' &&
    parseFloat(style.borderWidth) > 0 &&
    (style.borderColor === 'transparent' ||
      style.borderColor === 'rgba(0, 0, 0, 0)' ||
      parseColor(style.borderColor).opacity === 0);

  return hasTwoGradients && hasBorderBoxTechnique && hasTransparentBorder;
}
