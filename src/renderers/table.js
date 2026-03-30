// src/renderers/table.js
// Renderer for TABLE elements

import { ElementRenderer, PX_TO_INCH } from './base.js';
import { extractTableData, parseColor, getTextStyle, LAYER } from '../utils/index.js';
import { PX_TO_PT } from '../utils/constants.js';

function isStyledTable(tableNode) {
  const tableStyle = window.getComputedStyle(tableNode);
  const isCollapse = tableStyle.borderCollapse === 'collapse';
  if (!isCollapse) {
    const spacing = tableStyle.borderSpacing || '';
    const spacingValues = spacing.split(/\s+/).map((v) => parseFloat(v) || 0);
    if (spacingValues.some((v) => v > 0)) return true;
  }

  const firstTd = tableNode.querySelector('td, th');
  if (firstTd) {
    const s = window.getComputedStyle(firstTd);
    if (parseFloat(s.borderRadius) > 0) return true;
  }
  return false;
}

function renderStyledTable(node, config, domOrder, dims, globalOptions) {
  const items = [];
  const scale = config.scale;
  const tableRect = node.getBoundingClientRect();
  const baseX = dims.x;
  const baseY = dims.y;

  const trList = node.querySelectorAll('tr');
  trList.forEach((tr) => {
    const cells = Array.from(tr.children).filter((c) => ['TD', 'TH'].includes(c.tagName));

    cells.forEach((cell) => {
      const cellRect = cell.getBoundingClientRect();
      const style = window.getComputedStyle(cell);

      const cx = baseX + (cellRect.left - tableRect.left) * PX_TO_INCH * scale;
      const cy = baseY + (cellRect.top - tableRect.top) * PX_TO_INCH * scale;
      const cw = cellRect.width * PX_TO_INCH * scale;
      const ch = cellRect.height * PX_TO_INCH * scale;

      const bgColor = parseColor(style.backgroundColor);
      const hasBg = bgColor.hex && bgColor.opacity > 0;

      const radius = parseFloat(style.borderRadius) || 0;
      const radiusIn = radius * PX_TO_INCH * scale;

      const borderLeft = parseFloat(style.borderLeftWidth) || 0;
      const borderLeftColor = parseColor(style.borderLeftColor);
      const hasBorderLeft = borderLeft > 0 && borderLeftColor.hex && borderLeftColor.opacity > 0;

      const borderBottom = parseFloat(style.borderBottomWidth) || 0;
      const borderBottomColor = parseColor(style.borderBottomColor);
      const hasBorderBottom =
        borderBottom > 0 && borderBottomColor.hex && borderBottomColor.opacity > 0;

      if (hasBg || hasBorderLeft || hasBorderBottom || radius > 0) {
        const shapeOpts = {
          x: cx,
          y: cy,
          w: cw,
          h: ch,
          fill: hasBg
            ? { color: bgColor.hex, transparency: Math.round((1 - bgColor.opacity) * 100) }
            : { type: 'none' },
          line: { type: 'none' },
        };

        if (radiusIn > 0) shapeOpts.rectRadius = radiusIn;

        items.push({
          type: 'shape',
          layer: LAYER.BACKGROUND,
          domOrder,
          shapeType: radiusIn > 0 ? 'roundRect' : 'rect',
          options: shapeOpts,
        });
      }

      if (hasBorderLeft) {
        const blW = borderLeft * PX_TO_INCH * scale;
        items.push({
          type: 'shape',
          layer: LAYER.BORDER,
          domOrder,
          shapeType: 'rect',
          options: {
            x: cx,
            y: cy,
            w: blW,
            h: ch,
            fill: {
              color: borderLeftColor.hex,
              transparency: Math.round((1 - borderLeftColor.opacity) * 100),
            },
            line: { type: 'none' },
          },
        });
      }

      if (hasBorderBottom) {
        const bbH = borderBottom * PX_TO_INCH * scale;
        items.push({
          type: 'shape',
          layer: LAYER.BORDER,
          domOrder,
          shapeType: 'rect',
          options: {
            x: cx,
            y: cy + ch - bbH,
            w: cw,
            h: bbH,
            fill: {
              color: borderBottomColor.hex,
              transparency: Math.round((1 - borderBottomColor.opacity) * 100),
            },
            line: { type: 'none' },
          },
        });
      }

      const cellText = cell.innerText.replace(/[\n\r\t]+/g, ' ').trim();
      if (cellText) {
        const textStyle = getTextStyle(style, scale, cellText, cell, globalOptions);
        const padding = {
          top: (parseFloat(style.paddingTop) || 0) * PX_TO_INCH * scale,
          right: (parseFloat(style.paddingRight) || 0) * PX_TO_INCH * scale,
          bottom: (parseFloat(style.paddingBottom) || 0) * PX_TO_INCH * scale,
          left: (parseFloat(style.paddingLeft) || 0) * PX_TO_INCH * scale,
        };

        let align = 'left';
        if (style.textAlign === 'center') align = 'center';
        if (style.textAlign === 'right' || style.textAlign === 'end') align = 'right';

        const textOpts = {
          x: cx + padding.left,
          y: cy + padding.top,
          w: cw - padding.left - padding.right,
          h: ch - padding.top - padding.bottom,
          align,
          valign: 'middle',
          margin: 0,
          wrap: true,
          autoFit: false,
        };

        items.push({
          type: 'text',
          layer: LAYER.CONTENT,
          domOrder,
          textParts: [{ text: cellText, options: textStyle }],
          options: textOpts,
        });
      }
    });
  });

  return items;
}

export class TableRenderer extends ElementRenderer {
  render() {
    const { node, config, domOrder, globalOptions } = this;
    const dims = this.getDimensions();

    if (isStyledTable(node)) {
      const items = renderStyledTable(node, config, domOrder, dims, globalOptions);
      return { items, stopRecursion: true };
    }

    const tableData = extractTableData(node, config.scale, globalOptions);
    const items = [];

    const tableStyle = window.getComputedStyle(node);
    const tableBg = parseColor(tableStyle.backgroundColor);
    const tableBorderW = parseFloat(tableStyle.borderTopWidth) || 0;
    const tableBorderColor = parseColor(tableStyle.borderTopColor);
    const tableRadius = parseFloat(tableStyle.borderTopLeftRadius) || 0;
    const hasBg = tableBg.hex && tableBg.opacity > 0;
    const hasBorder = tableBorderW > 0 && tableBorderColor.hex && tableBorderColor.opacity > 0;

    if (hasBg || hasBorder) {
      const radiusIn =
        tableRadius > 0
          ? tableRadius.toString().includes('%')
            ? (tableRadius / 100) * dims.w
            : tableRadius * PX_TO_INCH * config.scale
          : 0;
      const shapeOpts = {
        x: dims.x,
        y: dims.y,
        w: dims.w,
        h: dims.h,
        fill: hasBg
          ? { color: tableBg.hex, transparency: Math.round((1 - tableBg.opacity) * 100) }
          : { type: 'none' },
        line: hasBorder
          ? {
              color: tableBorderColor.hex,
              width: tableBorderW * PX_TO_PT * config.scale,
              transparency: Math.round((1 - tableBorderColor.opacity) * 100),
            }
          : { type: 'none' },
      };
      if (radiusIn > 0) shapeOpts.rectRadius = radiusIn;
      items.push({
        type: 'shape',
        layer: LAYER.BACKGROUND,
        domOrder: domOrder - 0.1,
        shapeType: radiusIn > 0 ? 'roundRect' : 'rect',
        options: shapeOpts,
      });
    }

    items.push({
      type: 'table',
      layer: LAYER.CONTENT,
      domOrder,
      tableData: tableData,
      options: { x: dims.x, y: dims.y, w: dims.w, h: dims.h },
    });

    return { items, stopRecursion: true };
  }
}
