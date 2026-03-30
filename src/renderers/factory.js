// src/renderers/factory.js
// Factory for creating appropriate renderers based on DOM node type

import { TextNodeRenderer } from './text-node.js';
import { TableRenderer } from './table.js';
import { ListRenderer } from './list.js';
import { CanvasRenderer } from './canvas.js';
import { SVGRenderer } from './svg-element.js';
import { ImgRenderer } from './img.js';
import { IconRenderer } from './icon.js';
import { ShapeRenderer } from './shape.js';
import { isIconElement, isComplexHierarchy } from './helpers.js';

/**
 * Factory class that creates the appropriate renderer for a given DOM node.
 */
export class RendererFactory {
  /**
   * Create a renderer for the given input.
   * @param {Object} input - Renderer initialization parameters
   * @param {Node} input.node - The DOM node to render
   * @param {Object} input.config - Layout configuration
   * @param {number} input.domOrder - DOM traversal order
   * @param {Object} input.pptx - PptxGenJS instance
   * @param {Array} input.stackChain - Stacking context chain
   * @param {CSSStyleDeclaration} input.style - Pre-computed style
   * @param {Object} input.globalOptions - Global options
   * @returns {Object|null} Renderer instance or null
   */
  static create(input) {
    const { node, style } = input;

    // 1. Text Node
    if (node.nodeType === 3) {
      return new TextNodeRenderer(input);
    }

    // Not an element node
    if (node.nodeType !== 1) {
      return null;
    }

    const tagName = node.tagName.toUpperCase();

    // 2. Table
    if (tagName === 'TABLE') {
      return new TableRenderer(input);
    }

    // 3. List (UL/OL) - only if not complex
    if ((tagName === 'UL' || tagName === 'OL') && !isComplexHierarchy(node)) {
      return new ListRenderer(input);
    }

    // 4. Canvas
    if (tagName === 'CANVAS') {
      return new CanvasRenderer(input);
    }

    // 5. SVG
    if (node.nodeName.toUpperCase() === 'SVG') {
      return new SVGRenderer(input);
    }

    // 6. IMG
    if (tagName === 'IMG') {
      return new ImgRenderer(input);
    }

    // 7. Icon elements (FontAwesome, Material, custom elements, etc.)
    if (isIconElement(node)) {
      return new IconRenderer(input);
    }

    // 8. Default: Shape renderer for all other elements
    return new ShapeRenderer(input);
  }
}
