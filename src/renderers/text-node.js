// src/renderers/text-node.js
// Renderer for text nodes (nodeType === 3)

import { ElementRenderer, PX_TO_INCH } from './base.js';
import {
  getTextStyle,
  isTextContainer,
  getCharSpacingWidthCompensation,
  LAYER,
} from '../utils/index.js';

export class TextNodeRenderer extends ElementRenderer {
  render() {
    const { node, config } = this;
    const textContent = node.nodeValue.trim();
    if (!textContent) return null;

    const parent = node.parentElement;
    if (!parent) return null;

    // If parent is a text container, it handles this text node
    if (isTextContainer(parent)) return null;

    const style = window.getComputedStyle(parent);
    const textOpts = getTextStyle(style, config.scale, textContent, parent, this.globalOptions);

    // Don't add highlight if parent has visible background (parent's shape already provides the background)
    // This prevents duplicate background rendering for text inside styled containers
    if (textOpts.highlight) {
      delete textOpts.highlight;
    }

    const anchor = parent.closest('a');
    if (anchor) {
      const href = anchor.getAttribute('href');
      if (href && !href.startsWith('#') && !href.startsWith('javascript:')) {
        textOpts.hyperlink = { url: href, tooltip: anchor.getAttribute('title') || undefined };
      }
    }

    const range = document.createRange();
    range.selectNode(node);
    const rects = range.getClientRects();
    range.detach();

    // Single-line text: fast path (original behavior)
    if (rects.length <= 1) {
      const rect = rects.length === 1 ? rects[0] : { left: 0, top: 0, width: 0, height: 0 };
      return this._createSingleItem(textContent, rect, style, textOpts);
    }

    // Multi-line text: split into per-line text boxes for accurate positioning.
    // This happens when a text node wraps across lines — getBoundingClientRect()
    // returns the union bounding box which mispositions the text.
    // Instead, use getClientRects() which returns one rect per line fragment.
    const items = [];
    const lineTexts = this._splitTextByRects(node, rects);

    for (let i = 0; i < lineTexts.length; i++) {
      const lineText = lineTexts[i].text;
      if (!lineText) continue;

      const rect = lineTexts[i].rect;
      const lineOpts = { ...textOpts };
      const result = this._createSingleItem(lineText, rect, style, lineOpts);
      if (result) {
        items.push(...result.items);
      }
    }

    if (items.length === 0) return null;
    return { items, stopRecursion: false };
  }

  /**
   * Split a text node's content into per-line segments using Range and getClientRects.
   * Walks character by character, grouping characters that share the same line rect.
   */
  _splitTextByRects(node, rects) {
    const text = node.nodeValue;
    const lines = [];

    // Build a map of per-line rects by their Y position
    const rectList = [];
    for (let i = 0; i < rects.length; i++) {
      rectList.push(rects[i]);
    }

    // Walk characters and assign each to its line rect
    const range = document.createRange();
    let currentLine = '';
    let currentRect = null;
    const tolerance = 2; // px tolerance for same-line detection

    for (let i = 0; i < text.length; i++) {
      range.setStart(node, i);
      range.setEnd(node, i + 1);
      const charRects = range.getClientRects();
      if (charRects.length === 0) {
        currentLine += text[i];
        continue;
      }

      const charRect = charRects[0];

      if (!currentRect || Math.abs(charRect.top - currentRect.top) > tolerance) {
        // New line
        if (currentLine && currentRect) {
          lines.push({ text: currentLine, rect: currentRect });
        }
        currentLine = text[i];
        currentRect = charRect;
      } else {
        currentLine += text[i];
        // Extend the current rect width
        currentRect = {
          left: currentRect.left,
          top: currentRect.top,
          width: charRect.right - currentRect.left,
          height: Math.max(currentRect.height, charRect.height),
          right: charRect.right,
          bottom: Math.max(currentRect.bottom, charRect.bottom),
        };
      }
    }

    if (currentLine && currentRect) {
      lines.push({ text: currentLine, rect: currentRect });
    }

    range.detach();

    // Trim lines and filter empty
    return lines
      .map((l) => ({ text: l.text.replace(/\s+/g, ' ').trim(), rect: l.rect }))
      .filter((l) => l.text.length > 0);
  }

  _createSingleItem(textContent, rect, style, textOpts) {
    const { config, domOrder } = this;

    const widthPx = rect.width;
    const heightPx = rect.height;
    let unrotatedW = widthPx * PX_TO_INCH * config.scale;
    const unrotatedH = heightPx * PX_TO_INCH * config.scale;

    if (unrotatedW <= 0 || unrotatedH <= 0) return null;

    // Compensate width for charSpacing adjustments
    const fontWeight = parseInt(style.fontWeight) || 400;
    const fontSizePx = parseFloat(style.fontSize) || 16;
    const letterSpacingPx = parseFloat(style.letterSpacing) || 0;
    const widthCompensation = getCharSpacingWidthCompensation(
      fontWeight,
      letterSpacingPx,
      textContent.length,
      config.scale,
      fontSizePx
    );
    if (widthCompensation !== 0) {
      unrotatedW += widthCompensation;
    }

    const x = config.offX + (rect.left - config.rootX) * PX_TO_INCH * config.scale;
    const y = config.offY + (rect.top - config.rootY) * PX_TO_INCH * config.scale;

    return {
      items: [
        {
          type: 'text',
          layer: LAYER.CONTENT,
          domOrder,
          textParts: [
            {
              text: textContent,
              options: textOpts,
            },
          ],
          options: { x, y, w: unrotatedW, h: unrotatedH, margin: 0, inset: 0, autoFit: false },
        },
      ],
      stopRecursion: false,
    };
  }
}
