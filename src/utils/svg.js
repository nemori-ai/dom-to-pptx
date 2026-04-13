// src/utils/svg.js
// SVG generation utilities for gradients, custom shapes, and blurred effects

/** Unicode-safe btoa: handles non-ASCII characters (e.g. CJK text in SVG). */
function utoa(str) {
  return btoa(unescape(encodeURIComponent(str)));
}

/**
 * Generates an SVG data URL for a solid shape with non-uniform corner radii.
 * Handles special case: quarter-circle when only one corner has full radius.
 */
export function generateCustomShapeSVG(w, h, color, opacity, radii) {
  let { tl, tr, br, bl } = radii;

  // Clamp radii using CSS spec logic (avoid overlap)
  const factor = Math.min(
    w / (tl + tr) || Infinity,
    h / (tr + br) || Infinity,
    w / (br + bl) || Infinity,
    h / (bl + tl) || Infinity
  );

  if (factor < 1) {
    tl *= factor;
    tr *= factor;
    br *= factor;
    bl *= factor;
  }

  // Detect quarter-circle case: only one corner has radius >= half the min dimension
  // This happens with Tailwind's rounded-{corner}-full on square/rectangular elements
  const minDim = Math.min(w, h);
  const fullRadius = minDim / 2;
  const isFullTL = tl >= fullRadius - 0.5 && tr === 0 && br === 0 && bl === 0;
  const isFullTR = tr >= fullRadius - 0.5 && tl === 0 && br === 0 && bl === 0;
  const isFullBR = br >= fullRadius - 0.5 && tl === 0 && tr === 0 && bl === 0;
  const isFullBL = bl >= fullRadius - 0.5 && tl === 0 && tr === 0 && br === 0;

  let path;

  if (isFullTL) {
    // Quarter circle in top-left: arc from (0, h) to (w, 0)
    path = `M 0 ${h} A ${w} ${h} 0 0 1 ${w} 0 L ${w} ${h} Z`;
  } else if (isFullTR) {
    // Quarter circle in top-right: arc from (0, 0) to (w, h)
    path = `M 0 0 A ${w} ${h} 0 0 1 ${w} ${h} L 0 ${h} Z`;
  } else if (isFullBR) {
    // Quarter circle in bottom-right: arc from (w, 0) to (0, h)
    path = `M ${w} 0 A ${w} ${h} 0 0 1 0 ${h} L 0 0 Z`;
  } else if (isFullBL) {
    // Quarter circle in bottom-left: arc from (0, 0) to (w, h), going the other way
    // Start at top-right, line to bottom-right, arc to top-left
    path = `M ${w} 0 L ${w} ${h} A ${w} ${h} 0 0 1 0 0 Z`;
  } else {
    // Standard rounded rectangle with non-uniform corners
    path = `
      M ${tl} 0
      L ${w - tr} 0
      A ${tr} ${tr} 0 0 1 ${w} ${tr}
      L ${w} ${h - br}
      A ${br} ${br} 0 0 1 ${w - br} ${h}
      L ${bl} ${h}
      A ${bl} ${bl} 0 0 1 0 ${h - bl}
      L 0 ${tl}
      A ${tl} ${tl} 0 0 1 ${tl} 0
      Z
    `;
  }

  const svg = `
    <svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
      <path d="${path}" fill="#${color}" fill-opacity="${opacity}" />
    </svg>`;

  return 'data:image/svg+xml;base64,' + utoa(svg);
}

function parseGradientStops(parts, gradientLength) {
  const stops = [];
  parts.forEach((part, idx) => {
    let color = part;
    let offset = idx / (parts.length - 1);
    const posMatch = part.match(/^(.*?)\s+(-?[\d.]+)(%|px)?$/);
    if (posMatch) {
      color = posMatch[1];
      const value = parseFloat(posMatch[2]);
      const unit = posMatch[3];
      if (unit === 'px' && gradientLength > 0) {
        offset = value / gradientLength;
      } else {
        offset = value / 100;
      }
    }
    stops.push({ color: color.trim(), offset: Math.max(0, Math.min(1, offset)) });
  });
  return stops;
}

function generateRadialGradientSVG(w, h, content, radius, blurPx, renderScale) {
  const parts = content.split(/,(?![^()]*\))/).map((p) => p.trim());
  if (parts.length < 2) return null;

  let stopsStartIndex = 0;
  const firstPart = parts[0].toLowerCase();
  if (firstPart.includes('circle') || firstPart.includes('ellipse') || firstPart.includes('at ')) {
    stopsStartIndex = 1;
  }

  const stopParts = parts.slice(stopsStartIndex);
  const r = Math.max(w, h) / 2;
  const stops = parseGradientStops(stopParts, r);

  const isCircle = radius >= Math.min(w, h) / 2 - 1 && Math.abs(w - h) < 2;
  const pad = blurPx ? blurPx * 3 : 0;
  const scale = renderScale || 1;
  const fullW = (w + pad * 2) * scale;
  const fullH = (h + pad * 2) * scale;

  const canvas = document.createElement('canvas');
  canvas.width = fullW;
  canvas.height = fullH;
  const ctx = canvas.getContext('2d');
  if (scale !== 1) ctx.scale(scale, scale);

  if (blurPx) ctx.filter = `blur(${blurPx}px)`;

  const cx = (w + pad * 2) / 2,
    cy = (h + pad * 2) / 2;
  const grad = ctx.createRadialGradient(cx, cy, 0, cx, cy, r);
  stops.forEach((s) => grad.addColorStop(s.offset, s.color));
  ctx.fillStyle = grad;

  if (isCircle) {
    ctx.beginPath();
    ctx.arc(cx, cy, w / 2, 0, Math.PI * 2);
    ctx.fill();
  } else if (radius > 0) {
    const rr = Math.min(radius, w / 2, h / 2);
    ctx.beginPath();
    ctx.moveTo(pad + rr, pad);
    ctx.lineTo(pad + w - rr, pad);
    ctx.quadraticCurveTo(pad + w, pad, pad + w, pad + rr);
    ctx.lineTo(pad + w, pad + h - rr);
    ctx.quadraticCurveTo(pad + w, pad + h, pad + w - rr, pad + h);
    ctx.lineTo(pad + rr, pad + h);
    ctx.quadraticCurveTo(pad, pad + h, pad, pad + h - rr);
    ctx.lineTo(pad, pad + rr);
    ctx.quadraticCurveTo(pad, pad, pad + rr, pad);
    ctx.fill();
  } else {
    ctx.fillRect(pad, pad, w, h);
  }

  const dataUrl = canvas.toDataURL('image/png');
  return blurPx ? { data: dataUrl, padding: pad } : dataUrl;
}

function generateConicGradientPNG(w, h, content, radius) {
  const parts = content.split(/,(?![^()]*\))/).map((p) => p.trim());
  if (parts.length < 2) return null;

  let startAngle = -Math.PI / 2;
  let stopsStartIndex = 0;
  const first = parts[0];
  if (first.startsWith('from ')) {
    const degMatch = first.match(/from\s+([\d.]+)deg/);
    if (degMatch) startAngle = -Math.PI / 2 + (parseFloat(degMatch[1]) * Math.PI) / 180;
    stopsStartIndex = 1;
  }

  const segments = [];
  const stopParts = parts.slice(stopsStartIndex);
  for (const part of stopParts) {
    const tokens = part.match(/^(.*?)\s+([\d.]+)(%)\s+([\d.]+)(%)$/);
    if (tokens) {
      segments.push({
        color: tokens[1].trim(),
        from: parseFloat(tokens[2]) / 100,
        to: parseFloat(tokens[4]) / 100,
      });
    } else {
      const single = part.match(/^(.*?)\s+([\d.]+)(%)$/);
      if (single) {
        const pct = parseFloat(single[2]) / 100;
        const prevEnd = segments.length > 0 ? segments[segments.length - 1].to : 0;
        segments.push({ color: single[1].trim(), from: prevEnd, to: pct });
      }
    }
  }
  if (segments.length === 0) return null;

  const canvas = document.createElement('canvas');
  canvas.width = w;
  canvas.height = h;
  const ctx = canvas.getContext('2d');
  const cx = w / 2,
    cy = h / 2;
  const r = Math.max(w, h) / 2;
  const isCircle = radius >= Math.min(w, h) / 2 - 1 && Math.abs(w - h) < 2;

  for (const seg of segments) {
    const a1 = startAngle + seg.from * Math.PI * 2;
    const a2 = startAngle + seg.to * Math.PI * 2;
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, r, a1, a2);
    ctx.closePath();
    ctx.fillStyle = seg.color;
    ctx.fill();
  }

  if (!isCircle && radius > 0) {
    ctx.globalCompositeOperation = 'destination-in';
    const rr = Math.min(radius, w / 2, h / 2);
    ctx.beginPath();
    ctx.moveTo(rr, 0);
    ctx.lineTo(w - rr, 0);
    ctx.quadraticCurveTo(w, 0, w, rr);
    ctx.lineTo(w, h - rr);
    ctx.quadraticCurveTo(w, h, w - rr, h);
    ctx.lineTo(rr, h);
    ctx.quadraticCurveTo(0, h, 0, h - rr);
    ctx.lineTo(0, rr);
    ctx.quadraticCurveTo(0, 0, rr, 0);
    ctx.fill();
  } else if (isCircle) {
    ctx.globalCompositeOperation = 'destination-in';
    ctx.beginPath();
    ctx.arc(cx, cy, w / 2, 0, Math.PI * 2);
    ctx.fill();
  }

  return canvas.toDataURL('image/png');
}

export function generateGradientSVG(w, h, bgString, radius, border, blurPx, bgSize) {
  try {
    const conicMatch = bgString.match(/conic-gradient\((.*)\)/);
    if (conicMatch) {
      return generateConicGradientPNG(w, h, conicMatch[1], radius);
    }

    const radialMatch = bgString.match(/radial-gradient\((.*)\)/);
    if (radialMatch) {
      // Check if bgSize requires tiling
      let patW = w, patH = h;
      let needsTiling = false;
      if (bgSize && bgSize !== 'auto' && bgSize !== 'cover' && bgSize !== 'contain') {
        const sizeParts = bgSize.split(/\s+/);
        if (sizeParts[0] && sizeParts[0] !== 'auto') {
          patW = sizeParts[0].includes('%') ? (parseFloat(sizeParts[0]) / 100) * w : parseFloat(sizeParts[0]);
        }
        if (sizeParts[1] && sizeParts[1] !== 'auto') {
          patH = sizeParts[1].includes('%') ? (parseFloat(sizeParts[1]) / 100) * h : parseFloat(sizeParts[1]);
        }
        needsTiling = patW < w || patH < h;
      }

      if (!needsTiling) {
        return generateRadialGradientSVG(w, h, radialMatch[1], radius, blurPx);
      }

      // Render one tile as PNG at 3x resolution for smooth anti-aliasing
      const tilePng = generateRadialGradientSVG(patW, patH, radialMatch[1], 0, 0, 3);
      const tileData = typeof tilePng === 'string' ? tilePng : (tilePng && tilePng.data);
      if (!tileData) return null;

      // Build shape tag with border radius support
      let shapeTag;
      if (typeof radius === 'object' && radius !== null) {
        const { tl = 0, tr = 0, br = 0, bl = 0 } = radius;
        shapeTag =
          `<path d="M${tl},0 L${w - tr},0` +
          (tr > 0 ? ` A${tr},${tr} 0 0 1 ${w},${tr}` : ` L${w},0`) +
          ` L${w},${h - br}` +
          (br > 0 ? ` A${br},${br} 0 0 1 ${w - br},${h}` : ` L${w},${h}`) +
          ` L${bl},${h}` +
          (bl > 0 ? ` A${bl},${bl} 0 0 1 0,${h - bl}` : ` L0,${h}`) +
          ` L0,${tl}` +
          (tl > 0 ? ` A${tl},${tl} 0 0 1 ${tl},0` : ` L0,0`) +
          ` Z" fill="url(#pat)"/>`;
      } else {
        shapeTag = `<rect x="0" y="0" width="${w}" height="${h}" rx="${radius}" ry="${radius}" fill="url(#pat)"/>`;
      }

      const svg = `
        <svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
          <defs>
            <pattern id="pat" width="${patW}" height="${patH}" patternUnits="userSpaceOnUse">
              <image href="${tileData}" width="${patW}" height="${patH}"/>
            </pattern>
          </defs>
          ${shapeTag}
        </svg>`;

      return 'data:image/svg+xml;base64,' + utoa(svg);
    }

    const match = bgString.match(/linear-gradient\((.*)\)/);
    if (!match) return null;
    const content = match[1];

    // Split by comma, ignoring commas inside parentheses (e.g. rgba())
    const parts = content.split(/,(?![^()]*\))/).map((p) => p.trim());
    if (parts.length < 2) return null;

    let x1 = '0%',
      y1 = '0%',
      x2 = '0%',
      y2 = '100%';
    let stopsStartIndex = 0;
    const firstPart = parts[0].toLowerCase();

    // 1. Check for Keywords (to right, etc.)
    if (firstPart.startsWith('to ')) {
      stopsStartIndex = 1;
      const direction = firstPart.replace('to ', '').trim();
      switch (direction) {
        case 'top':
          y1 = '100%';
          y2 = '0%';
          break;
        case 'bottom':
          y1 = '0%';
          y2 = '100%';
          break;
        case 'left':
          x1 = '100%';
          y1 = '0%';
          x2 = '0%';
          y2 = '0%';
          break;
        case 'right':
          x2 = '100%';
          y2 = '0%';
          break;
        case 'top right':
          x1 = '0%';
          y1 = '100%';
          x2 = '100%';
          y2 = '0%';
          break;
        case 'top left':
          x1 = '100%';
          y1 = '100%';
          x2 = '0%';
          y2 = '0%';
          break;
        case 'bottom right':
          x1 = '0%';
          y1 = '0%';
          x2 = '100%';
          y2 = '100%';
          break;
        case 'bottom left':
          x1 = '100%';
          y1 = '0%';
          x2 = '0%';
          y2 = '100%';
          break;
      }
    }
    // 2. Check for Degrees (45deg, 90deg, etc.)
    else if (firstPart.match(/^-?[\d.]+(deg|rad|turn|grad)$/)) {
      stopsStartIndex = 1;
      const val = parseFloat(firstPart);
      // CSS 0deg is Top (North), 90deg is Right (East), 180deg is Bottom (South)
      // We convert this to SVG coordinates on a unit square (0-100%).
      // Formula: Map angle to perimeter coordinates.
      if (!isNaN(val)) {
        const deg = firstPart.includes('rad') ? val * (180 / Math.PI) : val;
        // CSS gradients: 0deg = bottom-to-top, 90deg = left-to-right (clockwise from north)
        // SVG linearGradient: x1,y1 = start, x2,y2 = end, Y-axis 0=top 100=bottom
        // Direction vector: (sin(θ), -cos(θ)) in screen coords
        const rad = (deg * Math.PI) / 180;
        const s = Math.sin(rad);
        const c = Math.cos(rad);

        x1 = (50 - s * 50).toFixed(1) + '%';
        y1 = (50 + c * 50).toFixed(1) + '%';
        x2 = (50 + s * 50).toFixed(1) + '%';
        y2 = (50 - c * 50).toFixed(1) + '%';
      }
    }

    // 3. Process Color Stops
    let stopsXML = '';
    const stopParts = parts.slice(stopsStartIndex);

    stopParts.forEach((part, idx) => {
      // Parse "Color Position" (e.g., "red 50%")
      // Regex looks for optional space + number + unit at the end of the string
      let color = part;
      let offset = Math.round((idx / (stopParts.length - 1)) * 100) + '%';

      const posMatch = part.match(/^(.*?)\s+(-?[\d.]+(?:%|px)?)$/);
      if (posMatch) {
        color = posMatch[1];
        offset = posMatch[2];
      }

      // Handle RGBA/RGB for SVG compatibility
      let opacity = 1;
      if (color.includes('rgba')) {
        const rgbaMatch = color.match(/[\d.]+/g);
        if (rgbaMatch && rgbaMatch.length >= 4) {
          opacity = rgbaMatch[3];
          color = `rgb(${rgbaMatch[0]},${rgbaMatch[1]},${rgbaMatch[2]})`;
        }
      }

      stopsXML += `<stop offset="${offset}" stop-color="${color.trim()}" stop-opacity="${opacity}"/>`;
    });

    let strokeAttr = '';
    if (border) {
      strokeAttr = `stroke="#${border.color}" stroke-width="${border.width}"`;
    }

    // Support both uniform radius (number) and per-corner radius (object {tl,tr,br,bl})
    let shapeTag;
    if (typeof radius === 'object' && radius !== null) {
      const { tl = 0, tr = 0, br = 0, bl = 0 } = radius;
      // SVG path with true circular arcs for each corner
      shapeTag =
        `<path d="M${tl},0 L${w - tr},0` +
        (tr > 0 ? ` A${tr},${tr} 0 0 1 ${w},${tr}` : ` L${w},0`) +
        ` L${w},${h - br}` +
        (br > 0 ? ` A${br},${br} 0 0 1 ${w - br},${h}` : ` L${w},${h}`) +
        ` L${bl},${h}` +
        (bl > 0 ? ` A${bl},${bl} 0 0 1 0,${h - bl}` : ` L0,${h}`) +
        ` L0,${tl}` +
        (tl > 0 ? ` A${tl},${tl} 0 0 1 ${tl},0` : ` L0,0`) +
        ` Z" fill="url(#grad)" ${strokeAttr}/>`;
    } else {
      shapeTag = `<rect x="0" y="0" width="${w}" height="${h}" rx="${radius}" ry="${radius}" fill="url(#grad)" ${strokeAttr} />`;
    }

    // When bgSize is specified (e.g. "100% 4px"), wrap gradient in a <pattern> for tiling
    let patternDefs = '';
    let fillRef = 'url(#grad)';
    if (bgSize && bgSize !== 'auto' && bgSize !== 'cover' && bgSize !== 'contain') {
      const sizeParts = bgSize.split(/\s+/);
      let patW = w, patH = h;
      if (sizeParts[0] && sizeParts[0] !== 'auto') {
        patW = sizeParts[0].includes('%') ? (parseFloat(sizeParts[0]) / 100) * w : parseFloat(sizeParts[0]);
      }
      if (sizeParts[1] && sizeParts[1] !== 'auto') {
        patH = sizeParts[1].includes('%') ? (parseFloat(sizeParts[1]) / 100) * h : parseFloat(sizeParts[1]);
      }
      if (patW < w || patH < h) {
        patternDefs = `<pattern id="pat" width="${patW}" height="${patH}" patternUnits="userSpaceOnUse">
              <rect width="${patW}" height="${patH}" fill="url(#grad)"/>
            </pattern>`;
        fillRef = 'url(#pat)';
      }
    }

    // Replace fill reference in shape tag
    if (fillRef !== 'url(#grad)') {
      shapeTag = shapeTag.replace(/fill="url\(#grad\)"/g, `fill="${fillRef}"`);
    }

    const svg = `
      <svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
          <defs>
            <linearGradient id="grad" x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}">
              ${stopsXML}
            </linearGradient>
            ${patternDefs}
          </defs>
          ${shapeTag}
      </svg>`;

    return 'data:image/svg+xml;base64,' + utoa(svg);
  } catch (e) {
    console.warn('Gradient generation failed:', e);
    return null;
  }
}

/**
 * Generates an SVG with Gaussian blur effect for soft-edge shapes.
 * @param {number} opacity - Fill opacity (0-1), defaults to 1
 */
export function generateBlurredSVG(w, h, color, radius, blurPx, opacity = 1) {
  const padding = blurPx * 3;
  const fullW = w + padding * 2;
  const fullH = h + padding * 2;
  const x = padding;
  const y = padding;
  let shapeTag = '';
  const isCircle = radius >= Math.min(w, h) / 2 - 1 && Math.abs(w - h) < 2;
  const opacityAttr = opacity < 1 ? ` fill-opacity="${opacity}"` : '';

  if (isCircle) {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const rx = w / 2;
    const ry = h / 2;
    shapeTag = `<ellipse cx="${cx}" cy="${cy}" rx="${rx}" ry="${ry}" fill="#${color}"${opacityAttr} filter="url(#f1)" />`;
  } else {
    shapeTag = `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="${radius}" ry="${radius}" fill="#${color}"${opacityAttr} filter="url(#f1)" />`;
  }

  const svg = `
  <svg xmlns="http://www.w3.org/2000/svg" width="${fullW}" height="${fullH}" viewBox="0 0 ${fullW} ${fullH}">
    <defs>
      <filter id="f1" x="-50%" y="-50%" width="200%" height="200%">
        <feGaussianBlur in="SourceGraphic" stdDeviation="${blurPx}" />
      </filter>
    </defs>
    ${shapeTag}
  </svg>`;

  return {
    data: 'data:image/svg+xml;base64,' + utoa(svg),
    padding: padding,
  };
}

/**
 * Resolves <use href="#id"> elements whose targets live in a different SVG.
 * Clones the referenced element into the target SVG's own <defs> so that the
 * serialized standalone SVG remains self-contained.
 */
function resolveUseReferences(sourceNode, cloneSvg) {
  const useEls = cloneSvg.querySelectorAll('use');
  if (!useEls.length) return;

  let cloneDefs = cloneSvg.querySelector('defs');

  for (const use of useEls) {
    const href =
      use.getAttribute('href') || use.getAttributeNS('http://www.w3.org/1999/xlink', 'href');
    if (!href || !href.startsWith('#')) continue;

    const id = href.slice(1);
    // Already defined inside this SVG clone — nothing to do
    if (cloneSvg.getElementById(id)) continue;

    // Look up the referenced element in the full document
    const referenced = document.getElementById(id);
    if (!referenced) continue;

    if (!cloneDefs) {
      cloneDefs = document.createElementNS('http://www.w3.org/2000/svg', 'defs');
      cloneSvg.insertBefore(cloneDefs, cloneSvg.firstChild);
    }

    const refClone = referenced.cloneNode(true);
    cloneDefs.appendChild(refClone);
  }
}

/**
 * Shared helper: clone an SVG element, inline computed styles, resolve
 * cross-SVG <use> references, and handle overflow:visible viewBox expansion.
 * Returns the serialized SVG data URL and geometry metadata.
 *
 * @param {SVGElement} node - The live SVG element in the DOM
 * @returns {{ svgUrl: string, fullWidth: number, fullHeight: number, minX: number, minY: number }}
 */
function prepareSvgClone(node) {
  const clone = node.cloneNode(true);
  const rect = node.getBoundingClientRect();
  // Use pre-transform CSS layout dimensions. getBoundingClientRect() returns
  // the axis-aligned bounding box AFTER CSS transforms (rotation), which would
  // distort the serialized SVG clone (it has no CSS transform applied).
  const computed = window.getComputedStyle(node);
  const cssWidth = parseFloat(computed.width) || rect.width || 300;
  const cssHeight = parseFloat(computed.height) || rect.height || 150;

  // Resolve <use href="#id"> references that point to <defs> in other SVGs.
  // When the clone is serialized standalone, cross-SVG references break,
  // so we inline the referenced content into the clone's own <defs>.
  resolveUseReferences(node, clone);

  function inlineStyles(source, target) {
    const computed = window.getComputedStyle(source);
    const properties = [
      'fill',
      'stroke',
      'stroke-width',
      'stroke-linecap',
      'stroke-linejoin',
      'opacity',
      'font-family',
      'font-size',
      'font-weight',
    ];

    if (computed.fill === 'none') target.setAttribute('fill', 'none');
    else if (computed.fill) target.style.fill = computed.fill;

    if (computed.stroke === 'none') target.setAttribute('stroke', 'none');
    else if (computed.stroke) target.style.stroke = computed.stroke;

    properties.forEach((prop) => {
      if (prop !== 'fill' && prop !== 'stroke') {
        const val = computed[prop];
        if (val && val !== 'auto') target.style[prop] = val;
      }
    });

    for (let i = 0; i < source.children.length; i++) {
      if (target.children[i]) inlineStyles(source.children[i], target.children[i]);
    }
  }

  inlineStyles(node, clone);
  clone.style.opacity = '1';

  // For SVGs with overflow:visible, expand the viewBox to include
  // all content (e.g. rotated text, negative-coordinate paths).
  // For normal SVGs (icons etc.), preserve the original viewBox.
  let fullWidth = cssWidth,
    fullHeight = cssHeight;
  let minX = 0,
    minY = 0;
  let viewBoxOverridden = false;

  const overflow = window.getComputedStyle(node).overflow;
  const svgOverflow = node.getAttribute('style')?.includes('overflow')
    ? node.style.overflow
    : node.getAttribute('overflow');
  const isOverflowVisible = overflow === 'visible' || svgOverflow === 'visible';

  if (isOverflowVisible) {
    try {
      const bbox = node.getBBox();
      const bboxMinX = Math.min(0, bbox.x);
      const bboxMinY = Math.min(0, bbox.y);
      const bboxMaxX = Math.max(cssWidth, bbox.x + bbox.width);
      const bboxMaxY = Math.max(cssHeight, bbox.y + bbox.height);
      if (bboxMinX < 0 || bboxMinY < 0 || bboxMaxX > cssWidth || bboxMaxY > cssHeight) {
        minX = bboxMinX;
        minY = bboxMinY;
        fullWidth = bboxMaxX - bboxMinX;
        fullHeight = bboxMaxY - bboxMinY;
        viewBoxOverridden = true;
      }
    } catch (_) {
      /* getBBox may fail on empty SVGs */
    }
  }

  if (viewBoxOverridden) {
    clone.setAttribute('viewBox', `${minX} ${minY} ${fullWidth} ${fullHeight}`);
  }
  clone.setAttribute('width', fullWidth);
  clone.setAttribute('height', fullHeight);
  clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');

  const xml = new XMLSerializer().serializeToString(clone);
  const svgUrl = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(xml)}`;

  return { svgUrl, xml, fullWidth, fullHeight, minX, minY };
}

/**
 * Converts an SVG element to a data:image/svg+xml URL (vector, no rasterization).
 * PptxGenJS natively supports SVG images and will embed them as vector graphics
 * with an automatic PNG fallback for older PowerPoint versions.
 *
 * @param {SVGElement} node - The SVG element to convert
 * @returns {{ data: string, offsetX: number, offsetY: number, fullWidth: number, fullHeight: number } | null}
 */
export function svgToDataUrl(node) {
  try {
    const { xml, fullWidth, fullHeight, minX, minY } = prepareSvgClone(node);
    return {
      data: 'data:image/svg+xml;base64,' + utoa(xml),
      offsetX: minX,
      offsetY: minY,
      fullWidth,
      fullHeight,
    };
  } catch (_) {
    return null;
  }
}

/**
 * Converts an SVG element to a PNG data URL (rasterized).
 * Used as fallback for SVGs that need cropping or when rasterization is preferred.
 *
 * @param {SVGElement} node - The SVG element to convert
 * @param {number} imageScale - Scale factor for the output image
 * @param {Object} [cropRect] - Optional crop rectangle in element pixels { x, y, w, h }
 */
export function svgToPng(node, imageScale = 3, cropRect = null) {
  return new Promise((resolve) => {
    const prepared = prepareSvgClone(node);
    if (!prepared) {
      resolve(null);
      return;
    }
    const { svgUrl, fullWidth, fullHeight, minX, minY } = prepared;

    const img = new Image();
    img.crossOrigin = 'Anonymous';
    img.onload = () => {
      const scale = imageScale;
      if (cropRect) {
        const canvas = document.createElement('canvas');
        canvas.width = cropRect.w * scale;
        canvas.height = cropRect.h * scale;
        const ctx = canvas.getContext('2d');
        ctx.scale(scale, scale);
        ctx.drawImage(
          img,
          cropRect.x,
          cropRect.y,
          cropRect.w,
          cropRect.h,
          0,
          0,
          cropRect.w,
          cropRect.h
        );
        resolve({ data: canvas.toDataURL('image/png') });
      } else {
        const canvas = document.createElement('canvas');
        canvas.width = fullWidth * scale;
        canvas.height = fullHeight * scale;
        const ctx = canvas.getContext('2d');
        ctx.scale(scale, scale);
        ctx.drawImage(img, 0, 0, fullWidth, fullHeight);
        resolve({
          data: canvas.toDataURL('image/png'),
          offsetX: minX,
          offsetY: minY,
          fullWidth,
          fullHeight,
        });
      }
    };
    img.onerror = () => resolve(null);
    img.src = svgUrl;
  });
}

/**
 * Generates an SVG for gradient border effect.
 * CSS gradient borders use: background: linear-gradient(...) padding-box, linear-gradient(...) border-box
 */
export function generateGradientBorderSVG(w, h, borderWidth, gradientString, radius, fillColor) {
  try {
    const match = gradientString.match(/linear-gradient\((.*)\)/);
    if (!match) return null;
    const content = match[1];

    const parts = content.split(/,(?![^()]*\))/).map((p) => p.trim());
    if (parts.length < 2) return null;

    let x1 = '0%',
      y1 = '0%',
      x2 = '0%',
      y2 = '100%';

    let colorParts = parts;
    const first = parts[0].toLowerCase();

    if (first.includes('deg')) {
      const deg = parseFloat(first);
      const rad = ((90 - deg) * Math.PI) / 180;
      x1 = `${50 - 50 * Math.cos(rad)}%`;
      y1 = `${50 + 50 * Math.sin(rad)}%`;
      x2 = `${50 + 50 * Math.cos(rad)}%`;
      y2 = `${50 - 50 * Math.sin(rad)}%`;
      colorParts = parts.slice(1);
    } else if (first.startsWith('to ')) {
      const dirMap = {
        'to right': ['0%', '50%', '100%', '50%'],
        'to left': ['100%', '50%', '0%', '50%'],
        'to bottom': ['50%', '0%', '50%', '100%'],
        'to top': ['50%', '100%', '50%', '0%'],
        'to bottom right': ['0%', '0%', '100%', '100%'],
        'to top right': ['0%', '100%', '100%', '0%'],
        'to bottom left': ['100%', '0%', '0%', '100%'],
        'to top left': ['100%', '100%', '0%', '0%'],
      };
      const coords = dirMap[first];
      if (coords) {
        [x1, y1, x2, y2] = coords;
      }
      colorParts = parts.slice(1);
    }

    const parseColorStop = (s) => {
      const str = s.trim();
      let color = str;
      let offset = null;

      const rgbaMatch = str.match(
        /rgba?\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d.]+))?\s*\)/
      );
      if (rgbaMatch) {
        color =
          '#' +
          [parseInt(rgbaMatch[1]), parseInt(rgbaMatch[2]), parseInt(rgbaMatch[3])]
            .map((n) => n.toString(16).padStart(2, '0'))
            .join('');
        const opacity = rgbaMatch[4] ? parseFloat(rgbaMatch[4]) : 1;
        const afterColor = str.slice(str.indexOf(')') + 1).trim();
        if (afterColor && afterColor.includes('%')) {
          offset = afterColor;
        }
        return { color, offset, opacity };
      }

      const hexMatch = str.match(/(#[0-9a-fA-F]{3,8})/);
      if (hexMatch) {
        color = hexMatch[1];
        const afterColor = str.slice(str.indexOf(color) + color.length).trim();
        if (afterColor && afterColor.includes('%')) {
          offset = afterColor;
        }
        return { color, offset, opacity: 1 };
      }

      const tokens = str.split(/\s+/);
      if (tokens.length > 1 && tokens[tokens.length - 1].includes('%')) {
        offset = tokens[tokens.length - 1];
        color = tokens.slice(0, -1).join(' ');
      }

      return { color, offset, opacity: 1 };
    };

    const stops = colorParts.map((p, i, arr) => {
      const { color, offset, opacity } = parseColorStop(p);
      const off = offset || `${(i / (arr.length - 1)) * 100}%`;
      const opacityAttr = opacity < 1 ? ` stop-opacity="${opacity}"` : '';
      return `<stop offset="${off}" stop-color="${color}"${opacityAttr} />`;
    });

    const stopsXML = stops.join('\n');
    const gradId = 'grad_' + Math.random().toString(36).substr(2, 9);
    const halfBorder = borderWidth / 2;
    const innerRadius = Math.max(0, radius - halfBorder);

    const svg = `
      <svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
        <defs>
          <linearGradient id="${gradId}" x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}">
            ${stopsXML}
          </linearGradient>
        </defs>
        <rect x="${halfBorder}" y="${halfBorder}" width="${w - borderWidth}" height="${h - borderWidth}" 
              rx="${innerRadius}" ry="${innerRadius}" 
              fill="${fillColor}" stroke="url(#${gradId})" stroke-width="${borderWidth}" />
      </svg>`;

    return 'data:image/svg+xml;base64,' + utoa(svg);
  } catch (e) {
    console.warn('Gradient border generation failed:', e);
    return null;
  }
}
