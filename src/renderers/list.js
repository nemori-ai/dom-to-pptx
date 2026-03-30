// src/renderers/list.js
// Renderer for UL/OL list elements

import { ElementRenderer } from './base.js';
import { parseColor, LAYER } from '../utils/index.js';
import { PX_TO_PT } from '../utils/constants.js';
import { isComplexHierarchy, collectListParts } from './helpers.js';

export class ListRenderer extends ElementRenderer {
  canHandle() {
    const { node } = this;
    return (node.tagName === 'UL' || node.tagName === 'OL') && !isComplexHierarchy(node);
  }

  render() {
    const { node, config, style, domOrder, globalOptions } = this;

    // Check if this list can be handled natively
    if (!this.canHandle()) {
      return null; // Let fallback handling occur
    }

    const dims = this.getDimensions();
    const items = [];
    const listItems = [];
    const liChildren = Array.from(node.children).filter((c) => c.tagName === 'LI');

    liChildren.forEach((child, index) => {
      const liStyle = window.getComputedStyle(child);
      const liRect = child.getBoundingClientRect();
      const parentRect = node.getBoundingClientRect(); // node is UL/OL

      // 1. Determine Bullet Config
      let bullet = { type: 'bullet' };
      const listStyleType = liStyle.listStyleType || 'disc';

      if (node.tagName === 'OL' || listStyleType === 'decimal') {
        bullet = { type: 'number' };
      } else if (listStyleType === 'none') {
        bullet = false;
      } else {
        let code = '2022'; // disc
        if (listStyleType === 'circle') code = '25CB';
        if (listStyleType === 'square') code = '25A0';

        // --- Color & Size Logic (Option > ::marker > CSS color) ---
        let finalHex = '000000';
        let markerFontSize = null;

        // A. Check Global Option override
        if (globalOptions?.listConfig?.color) {
          finalHex = parseColor(globalOptions.listConfig.color).hex || '000000';
        }
        // B. Check ::marker pseudo element (supported in modern browsers)
        else {
          const markerStyle = window.getComputedStyle(child, '::marker');
          const markerColor = parseColor(markerStyle.color);
          if (markerColor.hex) {
            finalHex = markerColor.hex;
          } else {
            // C. Fallback to LI text color
            const colorObj = parseColor(liStyle.color);
            if (colorObj.hex) finalHex = colorObj.hex;
          }

          // Check ::marker font-size
          const markerFs = parseFloat(markerStyle.fontSize);
          if (!isNaN(markerFs) && markerFs > 0) {
            // Convert px->pt for PPTX
            markerFontSize = markerFs * PX_TO_PT * config.scale;
          }
        }

        bullet = { code, color: finalHex };
        if (markerFontSize) {
          bullet.fontSize = markerFontSize;
        }
      }

      // 2. Calculate Dynamic Indent (Respects padding-left)
      const visualIndentPx = liRect.left - parentRect.left;
      const computedIndentPt = visualIndentPx * PX_TO_PT * config.scale;

      if (bullet && computedIndentPt > 0) {
        bullet.indent = computedIndentPt;
      }

      // 3. Extract Text Parts
      const parts = collectListParts(child, liStyle, config.scale, globalOptions);

      if (parts.length > 0) {
        parts.forEach((p) => {
          if (!p.options) p.options = {};
        });

        // A. Apply Bullet
        // Workaround: pptxgenjs bullets inherit the style of the text run they are attached to.
        // To support ::marker styles (color, size) that differ from the text, we create
        // a "dummy" text run at the start of the list item that carries the bullet configuration.
        if (bullet) {
          const firstPartInfo = parts[0].options;

          // Create a dummy run. We use a Zero Width Space to ensure it's rendered but invisible.
          // This "run" will hold the bullet and its specific color/size.
          const bulletRun = {
            text: '\u200B',
            options: {
              ...firstPartInfo, // Inherit base props (fontFace, etc.)
              color: bullet.color || firstPartInfo.color,
              fontSize: bullet.fontSize || firstPartInfo.fontSize,
              bullet: bullet,
            },
          };

          // Don't duplicate transparent or empty color from firstPart if bullet has one
          if (bullet.color) bulletRun.options.color = bullet.color;
          if (bullet.fontSize) bulletRun.options.fontSize = bullet.fontSize;

          // Prepend
          parts.unshift(bulletRun);
        }

        // Force paragraph spacing to 0 (layout uses absolute positioning)
        parts[0].options.paraSpaceBefore = 0;
        parts[0].options.paraSpaceAfter = 0;

        if (index < liChildren.length - 1) {
          parts[parts.length - 1].options.breakLine = true;
        }

        listItems.push(...parts);
      }
    });

    if (listItems.length > 0) {
      const bgColorObj = parseColor(style.backgroundColor);
      if (bgColorObj.hex && bgColorObj.opacity > 0) {
        items.push({
          type: 'shape',
          layer: LAYER.BACKGROUND,
          domOrder,
          shapeType: 'rect',
          options: { x: dims.x, y: dims.y, w: dims.w, h: dims.h, fill: { color: bgColorObj.hex } },
        });
      }

      items.push({
        type: 'text',
        layer: LAYER.CONTENT,
        domOrder,
        textParts: listItems,
        options: {
          x: dims.x,
          y: dims.y,
          w: dims.w,
          h: dims.h,
          align: 'left',
          valign: 'top',
          margin: 0,
          autoFit: false,
          wrap: true,
        },
      });

      return { items, stopRecursion: true };
    }

    return null;
  }
}
