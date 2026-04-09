// src/renderers/shape-background.js
// Background-related helpers for ShapeRenderer:
// - Canvas capture necessity detection

import { getClipInfo } from '../utils/index.js';

/**
 * Check if any children have complex visual features that require
 * canvas capture when the parent has clipping or partial border radius.
 *
 * @param {HTMLElement} node
 * @returns {boolean}
 */
export function checkComplexChildren(node) {
  return Array.from(node.children).some((child) => {
    const cs = window.getComputedStyle(child);
    const childBgImage = cs.backgroundImage || '';
    const childBgSize = cs.backgroundSize || '';
    const childHasTiled =
      childBgImage.includes('gradient') &&
      childBgSize !== '' &&
      childBgSize !== 'auto' &&
      childBgSize !== 'cover' &&
      childBgSize !== 'contain';
    // Note: mix-blend-mode is intentionally excluded from triggering canvas
    // capture. PPTX has no native blend mode support, so capturing via canvas
    // only loses fidelity (especially when children contain cross-origin images
    // that taint the canvas). Rendering children individually is more reliable.
    return childHasTiled || childBgImage.includes('repeating-');
  });
}

/**
 * Determine whether the element needs canvas capture (html2canvas)
 * for accurate rendering. This is needed for elements with complex
 * visual features that PPTX cannot natively represent.
 *
 * @param {Object} params
 * @param {HTMLElement} params.node
 * @param {CSSStyleDeclaration} params.style
 * @param {boolean} params.isRootElement
 * @param {boolean} params.hasPartialBorderRadius
 * @param {string} params.bgImageStr
 * @returns {boolean}
 */
export function checkNeedsCanvasCapture({
  node,
  style,
  isRootElement,
  hasPartialBorderRadius,
  bgImageStr,
}) {
  // Count gradients in background-image
  const gradientCount = (bgImageStr.match(/linear-gradient|radial-gradient/g) || []).length;
  const hasMultipleGradients = gradientCount > 1;
  const hasRepeatingGradient = bgImageStr.includes('repeating-');

  // Detect tiled/pattern backgrounds (background-size creates repeating patterns)
  const bgSize = style.backgroundSize || '';
  const hasTiledBackground =
    bgImageStr.includes('gradient') &&
    bgSize !== '' &&
    bgSize !== 'auto' &&
    bgSize !== 'cover' &&
    bgSize !== 'contain' &&
    !bgSize.includes('100%');

  // Only capture parent with children when parent has clipping/overflow behavior
  // Skip root element (body) to avoid capturing entire slide as one image
  const hasClippingBehavior =
    !isRootElement &&
    (style.overflow === 'hidden' || style.overflow === 'clip' || getClipInfo(node));

  // Skip canvas capture when the container has <img> children with external
  // sources — cross-origin images taint the canvas, making toDataURL() fail.
  // Rendering children individually is more reliable in this case.
  const hasExternalImages = node.querySelector && node.querySelector('img[src^="http"]');

  return (
    (hasPartialBorderRadius && getClipInfo(node) && !hasExternalImages) ||
    (hasPartialBorderRadius && checkComplexChildren(node) && !hasExternalImages) ||
    (hasClippingBehavior && checkComplexChildren(node) && !hasExternalImages) ||
    (!isRootElement && hasMultipleGradients) ||
    (!isRootElement && hasRepeatingGradient) ||
    (!isRootElement && hasTiledBackground)
  );
}
