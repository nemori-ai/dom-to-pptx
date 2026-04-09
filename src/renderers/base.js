// src/renderers/base.js
// Base class for all element renderers

import { getRotation } from '../utils/style.js';
import { PX_TO_INCH } from '../utils/constants.js';

/**
 * Base class that provides shared functionality for all renderers.
 * Each renderer handles a specific type of DOM element.
 */
export class ElementRenderer {
  constructor({ node, config, domOrder, pptx, stackChain, style, globalOptions }) {
    this.node = node;
    this.config = config;
    this.domOrder = domOrder;
    this.pptx = pptx;
    this.stackChain = stackChain;
    this.style = style;
    this.globalOptions = globalOptions || {};
  }

  /**
   * Calculate dimensions and position for the element.
   * @returns {Object} { rect, widthPx, heightPx, x, y, w, h, rotation, centerX, centerY }
   */
  getDimensions() {
    const { node, config, style } = this;
    const rect = node.getBoundingClientRect();
    const rotation = style ? getRotation(style.transform) : 0;

    // offsetWidth/Height give pre-transform layout dimensions.
    // getBoundingClientRect() returns the axis-aligned bounding box AFTER rotation,
    // which is larger than the actual element for non-90° rotations.
    // SVG elements lack offsetWidth/Height (they are not HTMLElement), so use
    // getComputedStyle as intermediate fallback for pre-transform dimensions.
    const computed = style || window.getComputedStyle(node);
    const widthPx = node.offsetWidth || parseFloat(computed.width) || rect.width;
    const heightPx = node.offsetHeight || parseFloat(computed.height) || rect.height;
    const unrotatedW = widthPx * PX_TO_INCH * config.scale;
    const unrotatedH = heightPx * PX_TO_INCH * config.scale;

    // The center of the bounding rect is the rotation center.
    // PPTX rotates shapes around their center, so we derive the unrotated
    // top-left from the bounding rect center minus half the original size.
    const centerX = rect.left + rect.width / 2;
    const centerY = rect.top + rect.height / 2;
    const x = config.offX + (centerX - widthPx / 2 - config.rootX) * PX_TO_INCH * config.scale;
    const y = config.offY + (centerY - heightPx / 2 - config.rootY) * PX_TO_INCH * config.scale;

    return {
      rect,
      widthPx,
      heightPx,
      x,
      y,
      w: unrotatedW,
      h: unrotatedH,
      rotation,
      centerX,
      centerY,
    };
  }

  /**
   * Get border-radius values, including inherited from parent.
   * @returns {Object} { tl, tr, br, bl } in pixels
   */
  getBorderRadii() {
    const { node, style } = this;
    const width = node.offsetWidth || node.getBoundingClientRect().width;
    const height = node.offsetHeight || node.getBoundingClientRect().height;

    const parseRadius = (value, dimension) => {
      if (!value) return 0;
      if (value.includes('%')) {
        return (parseFloat(value) / 100) * dimension;
      }
      return parseFloat(value) || 0;
    };

    let radii = {
      tl: parseRadius(style?.borderTopLeftRadius, Math.min(width, height)),
      tr: parseRadius(style?.borderTopRightRadius, Math.min(width, height)),
      br: parseRadius(style?.borderBottomRightRadius, Math.min(width, height)),
      bl: parseRadius(style?.borderBottomLeftRadius, Math.min(width, height)),
    };

    const hasAnyRadius = radii.tl > 0 || radii.tr > 0 || radii.br > 0 || radii.bl > 0;

    // Walk up ancestors to find a clipping ancestor with border-radius.
    // PPTX has no parent-child clipping, so children must self-clip to
    // match the browser's overflow:hidden + border-radius behaviour.
    if (!hasAnyRadius) {
      const nodeRect = node.getBoundingClientRect();
      const clampR = (r) => Math.min(r, width, height);
      let ancestor = node.parentElement;

      while (ancestor && ancestor !== document.body) {
        const aStyle = window.getComputedStyle(ancestor);
        if (aStyle.overflow === 'hidden' || aStyle.overflow === 'clip') {
          const aRect = ancestor.getBoundingClientRect();
          const aDim = Math.min(aRect.width, aRect.height);
          const aTL = parseRadius(aStyle.borderTopLeftRadius, aDim);
          const aTR = parseRadius(aStyle.borderTopRightRadius, aDim);
          const aBR = parseRadius(aStyle.borderBottomRightRadius, aDim);
          const aBL = parseRadius(aStyle.borderBottomLeftRadius, aDim);

          if (aTL > 0 || aTR > 0 || aBR > 0 || aBL > 0) {
            const T = 3; // edge-alignment tolerance in px
            const atLeft = Math.abs(nodeRect.left - aRect.left) < T;
            const atRight = Math.abs(nodeRect.right - aRect.right) < T;
            const atTop = Math.abs(nodeRect.top - aRect.top) < T;
            const atBottom = Math.abs(nodeRect.bottom - aRect.bottom) < T;

            const canFit = (r) => width >= r && height >= r;
            if (atTop && atLeft && canFit(aTL)) radii.tl = clampR(aTL);
            if (atTop && atRight && canFit(aTR)) radii.tr = clampR(aTR);
            if (atBottom && atRight && canFit(aBR)) radii.br = clampR(aBR);
            if (atBottom && atLeft && canFit(aBL)) radii.bl = clampR(aBL);

            if (radii.tl > 0 || radii.tr > 0 || radii.br > 0 || radii.bl > 0) break;
          }
          // No radius on this clipping ancestor, continue searching up
        }
        ancestor = ancestor.parentElement;
      }
    }

    return radii;
  }

  /**
   * Get element's effective opacity (inherited from ancestors).
   * CSS opacity is not inherited via getComputedStyle, but visually affects children.
   * Uses pre-computed accumulated opacity from DOM traversal when available,
   * otherwise falls back to walking up the DOM tree.
   * @returns {number} effective opacity value between 0 and 1
   */
  getOpacity() {
    if (this.config.accumulatedOpacity !== undefined) {
      return this.config.accumulatedOpacity;
    }

    // Fallback: traverse up the DOM tree and multiply all opacity values
    let effectiveOpacity = 1;
    let current = this.node;

    while (current && current.nodeType === 1) {
      const style = window.getComputedStyle(current);
      const opacity = parseFloat(style.opacity);
      if (!isNaN(opacity)) {
        effectiveOpacity *= opacity;
      }
      current = current.parentElement;
    }

    return effectiveOpacity;
  }

  /**
   * Render the element to PPTX items.
   * Must be implemented by subclasses.
   * @returns {Object|null} { items: Array, job?: Function, stopRecursion?: boolean }
   */
  render() {
    throw new Error('render() must be implemented by subclass');
  }
}

// Re-export PX_TO_INCH for use in renderers
export { PX_TO_INCH };
