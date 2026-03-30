// src/utils/border.js
// Border processing utilities

import { parseColor } from './color.js';
import { PX_TO_PT } from './constants.js';

/**
 * Maps CSS border-style to PptxGenJS dashType.
 */
function mapDashType(style) {
  if (style === 'dashed') return 'dash';
  if (style === 'dotted') return 'dot';
  return 'solid';
}

/**
 * Analyzes computed border styles and determines the rendering strategy.
 * Returns { type: 'none' | 'uniform' | 'composite', options?, sides? }
 */
export function getBorderInfo(style, scale) {
  const parseSide = (widthProp, styleProp, colorProp) => {
    const parsed = parseColor(style[colorProp]);
    return {
      width: parseFloat(style[widthProp]) || 0,
      style: style[styleProp],
      color: parsed.hex,
      opacity: parsed.opacity,
    };
  };
  const top = parseSide('borderTopWidth', 'borderTopStyle', 'borderTopColor');
  const right = parseSide('borderRightWidth', 'borderRightStyle', 'borderRightColor');
  const bottom = parseSide('borderBottomWidth', 'borderBottomStyle', 'borderBottomColor');
  const left = parseSide('borderLeftWidth', 'borderLeftStyle', 'borderLeftColor');

  const hasAnyBorder = top.width > 0 || right.width > 0 || bottom.width > 0 || left.width > 0;
  if (!hasAnyBorder) return { type: 'none' };

  // Check if all sides are uniform
  const isUniform =
    top.width === right.width &&
    top.width === bottom.width &&
    top.width === left.width &&
    top.style === right.style &&
    top.style === bottom.style &&
    top.style === left.style &&
    top.color === right.color &&
    top.color === bottom.color &&
    top.color === left.color;

  if (isUniform) {
    return {
      type: 'uniform',
      options: {
        width: top.width * PX_TO_PT * scale,
        color: top.color,
        transparency: (1 - parseColor(style.borderTopColor).opacity) * 100,
        dashType: mapDashType(top.style),
      },
    };
  } else {
    return {
      type: 'composite',
      sides: { top, right, bottom, left },
    };
  }
}

/**
 * Generates an SVG image for composite borders that respects border-radius.
 * Uses stroke-based rendering for each border side to properly follow rounded corners.
 * @param {number} w - Width in pixels
 * @param {number} h - Height in pixels
 * @param {number|object} radius - Uniform radius or {tl, tr, br, bl} object
 * @param {object} sides - Border sides info
 */
export function generateCompositeBorderSVG(w, h, radius, sides) {
  // Normalize radius to object form
  let tl, tr, br, bl;
  if (typeof radius === 'object') {
    ({ tl, tr, br, bl } = radius);
  } else {
    tl = tr = br = bl = radius || 0;
  }

  // Clamp radii to half of dimension to avoid overlap
  const maxR = Math.min(w, h) / 2;
  tl = Math.min(tl, maxR);
  tr = Math.min(tr, maxR);
  br = Math.min(br, maxR);
  bl = Math.min(bl, maxR);

  let paths = '';

  // Helper: render a straight border as a filled rect (avoids stroke rendering quirks in PPT)
  // or as a stroked path when corner arcs are needed.
  const opacityFill = (side) => (side.opacity < 1 ? ` fill-opacity="${side.opacity}"` : '');
  const opacityStroke = (side) => (side.opacity < 1 ? ` stroke-opacity="${side.opacity}"` : '');

  // Top border
  if (sides.top.width > 0 && sides.top.color) {
    const sw = sides.top.width;
    if (tl === 0 && tr === 0) {
      paths += `<rect x="0" y="0" width="${w}" height="${sw}" fill="#${sides.top.color}"${opacityFill(sides.top)} />`;
    } else {
      const inset = sw / 2;
      let d = '';
      if (tl > 0) {
        const r = Math.max(0, tl - inset);
        d += `M ${inset} ${tl} A ${r} ${r} 0 0 1 ${tl} ${inset}`;
      } else {
        d += `M 0 ${inset}`;
      }
      d += ` L ${tr > 0 ? w - tr : w} ${inset}`;
      if (tr > 0) {
        const r = Math.max(0, tr - inset);
        d += ` A ${r} ${r} 0 0 1 ${w - inset} ${tr}`;
      }
      paths += `<path d="${d}" stroke="#${sides.top.color}" stroke-width="${sw}" fill="none" stroke-linecap="butt"${opacityStroke(sides.top)} />`;
    }
  }

  // Right border
  if (sides.right.width > 0 && sides.right.color) {
    const sw = sides.right.width;
    if (tr === 0 && br === 0) {
      paths += `<rect x="${w - sw}" y="0" width="${sw}" height="${h}" fill="#${sides.right.color}"${opacityFill(sides.right)} />`;
    } else {
      const inset = sw / 2;
      const x = w - inset;
      let d = '';
      if (tr > 0) {
        const r = Math.max(0, tr - inset);
        d += `M ${w - tr} ${inset} A ${r} ${r} 0 0 1 ${x} ${tr}`;
      } else {
        d += `M ${x} 0`;
      }
      d += ` L ${x} ${br > 0 ? h - br : h}`;
      if (br > 0) {
        const r = Math.max(0, br - inset);
        d += ` A ${r} ${r} 0 0 1 ${w - br} ${h - inset}`;
      }
      paths += `<path d="${d}" stroke="#${sides.right.color}" stroke-width="${sw}" fill="none" stroke-linecap="butt"${opacityStroke(sides.right)} />`;
    }
  }

  // Bottom border
  if (sides.bottom.width > 0 && sides.bottom.color) {
    const sw = sides.bottom.width;
    if (br === 0 && bl === 0) {
      paths += `<rect x="0" y="${h - sw}" width="${w}" height="${sw}" fill="#${sides.bottom.color}"${opacityFill(sides.bottom)} />`;
    } else {
      const inset = sw / 2;
      const y = h - inset;
      let d = '';
      if (br > 0) {
        const r = Math.max(0, br - inset);
        d += `M ${w - inset} ${h - br} A ${r} ${r} 0 0 1 ${w - br} ${y}`;
      } else {
        d += `M ${w} ${y}`;
      }
      d += ` L ${bl > 0 ? bl : 0} ${y}`;
      if (bl > 0) {
        const r = Math.max(0, bl - inset);
        d += ` A ${r} ${r} 0 0 1 ${inset} ${h - bl}`;
      }
      paths += `<path d="${d}" stroke="#${sides.bottom.color}" stroke-width="${sw}" fill="none" stroke-linecap="butt"${opacityStroke(sides.bottom)} />`;
    }
  }

  // Left border
  if (sides.left.width > 0 && sides.left.color) {
    const sw = sides.left.width;
    if (bl === 0 && tl === 0) {
      paths += `<rect x="0" y="0" width="${sw}" height="${h}" fill="#${sides.left.color}"${opacityFill(sides.left)} />`;
    } else {
      const inset = sw / 2;
      const x = inset;
      let d = '';
      if (bl > 0) {
        const r = Math.max(0, bl - inset);
        d += `M ${bl} ${h - inset} A ${r} ${r} 0 0 1 ${x} ${h - bl}`;
      } else {
        d += `M ${x} ${h}`;
      }
      d += ` L ${x} ${tl > 0 ? tl : 0}`;
      if (tl > 0) {
        const r = Math.max(0, tl - inset);
        d += ` A ${r} ${r} 0 0 1 ${tl} ${inset}`;
      }
      paths += `<path d="${d}" stroke="#${sides.left.color}" stroke-width="${sw}" fill="none" stroke-linecap="butt"${opacityStroke(sides.left)} />`;
    }
  }

  if (!paths) return null;

  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">${paths}</svg>`;

  return 'data:image/svg+xml;base64,' + btoa(svg);
}
