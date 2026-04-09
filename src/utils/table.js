// src/utils/table.js
// Table data extraction utilities

import { parseColor } from './color.js';
import { getTextStyle } from './text.js';
import { getPadding } from './style.js';
import { PX_TO_PT, PX_TO_INCH } from './constants.js';

/**
 * Gets border info for a specific side of a table cell.
 */
function getTableBorder(style, side, scale) {
  const widthStr = style[`border${side}Width`];
  const styleStr = style[`border${side}Style`];
  const colorStr = style[`border${side}Color`];

  const width = parseFloat(widthStr) || 0;
  if (width === 0 || styleStr === 'none' || styleStr === 'hidden') {
    return null;
  }

  const color = parseColor(colorStr);
  if (!color.hex || color.opacity === 0) return null;

  let dash = 'solid';
  if (styleStr === 'dashed') dash = 'dash';
  if (styleStr === 'dotted') dash = 'dot';

  return {
    pt: width * PX_TO_PT * scale,
    color: color.hex,
    style: dash,
  };
}

/**
 * Extracts native table data for PptxGenJS.
 * @param {HTMLTableElement} node - The table element
 * @param {number} scale - Layout scale factor
 * @param {Object} [globalOptions={}] - Global options (backgroundSnapshot, etc.)
 */
export function extractTableData(node, scale, globalOptions = {}) {
  const rows = [];
  const colWidths = [];
  const rowHeights = [];

  // 1. Calculate Column Widths based on the first row of cells
  // We look at the first <tr>'s children to determine visual column widths.
  // Note: This assumes a fixed grid. Complex colspan/rowspan on the first row
  // might skew widths, but getBoundingClientRect captures the rendered result.
  const firstRow = node.querySelector('tr');
  if (firstRow) {
    const cells = Array.from(firstRow.children);
    cells.forEach((cell) => {
      const rect = cell.getBoundingClientRect();
      const wIn = rect.width * PX_TO_INCH * scale;
      colWidths.push(wIn);
    });
  }

  // 2. Iterate Rows
  const trList = node.querySelectorAll('tr');
  trList.forEach((tr) => {
    const rowData = [];
    const trRect = tr.getBoundingClientRect();
    rowHeights.push(trRect.height * PX_TO_INCH * scale);

    const cellList = Array.from(tr.children).filter((c) => ['TD', 'TH'].includes(c.tagName));

    cellList.forEach((cell) => {
      const style = window.getComputedStyle(cell);
      const cellText = cell.innerText.replace(/[\n\r\t]+/g, ' ').trim();

      // A. Cell Background first (preserve zebra striping, etc.)
      //    Also check parent tr/thead/tbody for inherited background
      let fill = null;
      let cellBg = parseColor(style.backgroundColor);
      if ((!cellBg.hex || cellBg.opacity === 0) && cell.parentElement) {
        const trStyle = window.getComputedStyle(cell.parentElement);
        const trBg = parseColor(trStyle.backgroundColor);
        if (trBg.hex && trBg.opacity > 0) cellBg = trBg;
      }
      if (cellBg.hex && cellBg.opacity > 0.05) {
        const transparency = Math.round((1 - cellBg.opacity) * 100);
        fill = { color: cellBg.hex, transparency };
      } else if (style.backgroundImage && style.backgroundImage.includes('gradient')) {
        const gradientMatch = style.backgroundImage.match(
          /rgba?\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/
        );
        if (gradientMatch) {
          const [, r, g, b] = gradientMatch;
          const hex = ((1 << 24) + (parseInt(r) << 16) + (parseInt(g) << 8) + parseInt(b))
            .toString(16)
            .slice(1)
            .toUpperCase();
          fill = { color: hex };
        }
      }

      // B. Badge detection - only if cell has no background and has single styled child
      let badgeStyle = null;
      const cellHasNoBg = !fill;
      const firstChild = cell.firstElementChild;
      const isSingleChild = cell.children.length === 1;
      const isBadgeTag = firstChild && ['SPAN', 'DIV'].includes(firstChild.tagName);

      if (cellHasNoBg && isSingleChild && isBadgeTag) {
        const childText = firstChild.innerText.replace(/[\n\r\t]+/g, ' ').trim();
        if (childText === cellText) {
          const childStyle = window.getComputedStyle(firstChild);
          const childBg = parseColor(childStyle.backgroundColor);
          if (childBg.hex && childBg.opacity > 0) {
            // Use the badge's text styling (color, font) but don't fill the
            // entire cell with the badge background — native PPTX table cells
            // can't render a small rounded pill inside a cell.
            badgeStyle = { bg: childBg, textStyle: childStyle, node: firstChild };
          }
        }
      }

      // C. Text Style - use badge style if present
      const effectiveNode = badgeStyle ? badgeStyle.node : cell;
      const effectiveStyle = badgeStyle ? badgeStyle.textStyle : style;
      const textStyle = getTextStyle(effectiveStyle, scale, cellText, effectiveNode, globalOptions);

      // C. Alignment
      let align = 'left';
      if (style.textAlign === 'center') align = 'center';
      if (style.textAlign === 'right' || style.textAlign === 'end') align = 'right';

      let valign = 'top';
      if (style.verticalAlign === 'middle') valign = 'middle';
      if (style.verticalAlign === 'bottom') valign = 'bottom';

      // D. Padding (Margins in PPTX)
      const padding = getPadding(style, scale);
      // getPadding returns { top, right, bottom, left, inset } in inches (scaled).
      // PptxGenJS margin uses a heuristic: if margin[0] >= 1, values are treated as
      // points (via valToPts); otherwise as inches (via inch2Emu).
      // Pass values in inches to avoid the heuristic misinterpreting zero top-padding
      // as "all values are inches" when they were meant to be points.
      const margin = [
        padding.top, // top (inches)
        padding.right, // right (inches)
        padding.bottom, // bottom (inches)
        padding.left, // left (inches)
      ];

      // E. Borders - use { type: 'none' } to explicitly disable borders
      const noBorder = { type: 'none' };
      const borderTop = getTableBorder(style, 'Top', scale) || noBorder;
      const borderRight = getTableBorder(style, 'Right', scale) || noBorder;
      const borderBottom = getTableBorder(style, 'Bottom', scale) || noBorder;
      const borderLeft = getTableBorder(style, 'Left', scale) || noBorder;

      // F. Construct Cell Object
      const cellOptions = {
        color: textStyle.color,
        fontFace: textStyle.fontFace,
        fontSize: textStyle.fontSize,
        bold: textStyle.bold,
        italic: textStyle.italic,
        underline: textStyle.underline,

        fill: fill,
        align: align,
        valign: valign,
        margin: margin,

        rowspan: parseInt(cell.getAttribute('rowspan')) || null,
        colspan: parseInt(cell.getAttribute('colspan')) || null,

        border: [borderTop, borderRight, borderBottom, borderLeft],
      };
      if (textStyle.lang) cellOptions.lang = textStyle.lang;
      if (textStyle.transparency) cellOptions.transparency = textStyle.transparency;

      rowData.push({ text: cellText, options: cellOptions });
    });

    if (rowData.length > 0) {
      rows.push(rowData);
    }
  });

  return { rows, colWidths, rowHeights };
}
