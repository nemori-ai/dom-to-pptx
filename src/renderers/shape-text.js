// src/renderers/shape-text.js
// Text extraction and payload assembly for ShapeRenderer

import {
  isTextContainer,
  getTextStyle,
  getPadding,
  getCharSpacingWidthCompensation,
  parseColor,
} from '../utils/index.js';
import { PX_TO_PT, PX_TO_INCH } from '../utils/constants.js';

/**
 * Extract hyperlink info from an <a> element.
 * @param {HTMLElement} el
 * @returns {Object|null} { url, tooltip } or null
 */
function getPseudoTextRun(node, pseudo, scale, globalOptions) {
  const ps = window.getComputedStyle(node, pseudo);
  if (!ps.content || ps.content === 'none' || ps.content === 'normal') return null;
  if (ps.display === 'none' || ps.visibility === 'hidden') return null;
  if (ps.position === 'absolute') return null;

  let text = ps.content.replace(/^['"]|['"]$/g, '');
  text = text.replace(/\\([\da-fA-F]{1,6})\s?/g, (_, hex) =>
    String.fromCodePoint(parseInt(hex, 16))
  );
  text = text.replace(/\\(.)/g, '$1');
  if (text === '""' || text === "''") text = '';

  if (!text) return null;

  const parentStyle = window.getComputedStyle(node);
  const hasGap = parseFloat(parentStyle.gap) > 0 || parseFloat(parentStyle.columnGap) > 0;
  if (hasGap) {
    if (pseudo === '::before') text = text + ' ';
    else text = ' ' + text;
  }

  const textOpts = getTextStyle(ps, scale, text, node, globalOptions);
  const pseudoColor = parseColor(ps.color);
  if (pseudoColor.hex) textOpts.color = pseudoColor.hex;
  if (pseudoColor.opacity < 1) textOpts.transparency = Math.round((1 - pseudoColor.opacity) * 100);

  return { text, options: textOpts };
}

function getHyperlinkInfo(el) {
  if (el.tagName !== 'A') return null;
  const href = el.getAttribute('href');
  if (!href || href.startsWith('#') || href.startsWith('javascript:')) return null;
  return { url: href, tooltip: el.getAttribute('title') || null };
}

/**
 * Add preformatted text to textParts, splitting on newlines and preserving indentation.
 *
 * @param {string} rawText
 * @param {CSSStyleDeclaration} nodeStyle
 * @param {HTMLElement} fontNode
 * @param {Object|null} hyperlink
 * @param {Array} textParts - mutated in place
 * @param {CSSStyleDeclaration} containerStyle - the parent container's style (for highlight comparison)
 * @param {number} scale
 * @param {Object} globalOptions
 */
function addPreformattedText(
  rawText,
  nodeStyle,
  fontNode,
  hyperlink,
  textParts,
  containerStyle,
  scale,
  globalOptions
) {
  const lines = rawText.split('\n');
  for (let li = 0; li < lines.length; li++) {
    if (li > 0) {
      textParts.push({ text: '', options: { breakLine: true } });
    }
    const lineText = lines[li];
    if (lineText.length > 0) {
      const textOpts = getTextStyle(nodeStyle, scale, lineText, fontNode, globalOptions);
      // Only keep highlight if the span itself has a distinct background-color
      if (textOpts.highlight) {
        const spanBg = nodeStyle.backgroundColor;
        const parentBg = containerStyle.backgroundColor;
        if (spanBg === parentBg || spanBg === 'transparent' || spanBg === 'rgba(0, 0, 0, 0)') {
          delete textOpts.highlight;
        }
      }
      if (hyperlink) textOpts.hyperlink = hyperlink;
      textParts.push({ text: lineText, options: textOpts });
    }
  }
}

/**
 * Recursive walker for preformatted content: walks ALL descendant text nodes
 * preserving the style of the closest element ancestor.
 *
 * @param {Node} parentNode
 * @param {CSSStyleDeclaration} parentStyle
 * @param {Object|null} hyperlink
 * @param {Array} textParts - mutated in place
 * @param {CSSStyleDeclaration} containerStyle
 * @param {number} scale
 * @param {Object} globalOptions
 */
function walkPreformatted(
  parentNode,
  parentStyle,
  hyperlink,
  textParts,
  containerStyle,
  scale,
  globalOptions
) {
  for (const child of parentNode.childNodes) {
    if (child.nodeType === 3) {
      const rawText = child.nodeValue;
      if (rawText.length === 0) continue;
      const fontNode = child.parentElement || parentNode;
      const nodeStyle = child.parentElement
        ? window.getComputedStyle(child.parentElement)
        : parentStyle;
      addPreformattedText(
        rawText,
        nodeStyle,
        fontNode,
        hyperlink,
        textParts,
        containerStyle,
        scale,
        globalOptions
      );
    } else if (child.nodeType === 1) {
      if (child.tagName === 'BR') {
        textParts.push({ text: '', options: { breakLine: true } });
        continue;
      }
      const childStyle = window.getComputedStyle(child);
      const linkInfo = getHyperlinkInfo(child);
      walkPreformatted(
        child,
        childStyle,
        linkInfo || hyperlink,
        textParts,
        containerStyle,
        scale,
        globalOptions
      );
    }
  }
}

/**
 * Extract the full text payload from a text container element.
 * Returns null if the element is not a text container or has no text parts.
 *
 * @param {Object} params
 * @param {HTMLElement} params.node
 * @param {CSSStyleDeclaration} params.style
 * @param {number} params.scale
 * @param {Object} params.globalOptions
 * @param {number} params.widthPx
 * @param {number} params.heightPx
 * @param {number} params.w - width in inches (may be modified for charSpacing compensation)
 * @param {Object} params.bgColorObj - { hex, opacity }
 * @returns {{ textPayload: Object, wAdjustment: number } | null}
 */
export function extractTextPayload({
  node,
  style,
  scale,
  globalOptions,
  widthPx,
  heightPx,
  w,
  bgColorObj,
}) {
  if (!isTextContainer(node)) return null;

  const textParts = [];
  let trimNextLeading = false;

  // Detect preformatted context (white-space: pre/pre-wrap/pre-line)
  const whiteSpace = style.whiteSpace || '';
  const isPreformatted =
    whiteSpace === 'pre' ||
    whiteSpace === 'pre-wrap' ||
    whiteSpace === 'pre-line' ||
    node.closest('pre') !== null;

  // Recursive helper to process child nodes (for nested elements within <a>)
  const processTextNode = (child, parentStyle, isFirst, isLast, hyperlink = null) => {
    let textVal = child.nodeType === 3 ? child.nodeValue : child.textContent;
    let nodeStyle = child.nodeType === 1 ? window.getComputedStyle(child) : parentStyle;
    const fontNode = child.nodeType === 1 ? child : child.parentElement;
    textVal = textVal.replace(/[\n\r\t]+/g, ' ').replace(/\s{2,}/g, ' ');

    if (isFirst) textVal = textVal.trimStart();
    if (trimNextLeading) {
      textVal = textVal.trimStart();
      trimNextLeading = false;
    }
    if (isLast) textVal = textVal.trimEnd();
    if (nodeStyle.textTransform === 'uppercase') textVal = textVal.toUpperCase();
    if (nodeStyle.textTransform === 'lowercase') textVal = textVal.toLowerCase();

    if (textVal.length > 0) {
      const textOpts = getTextStyle(nodeStyle, scale, textVal, fontNode, globalOptions);

      if (child.nodeType === 3 && textOpts.highlight) {
        delete textOpts.highlight;
      }

      if (hyperlink) {
        textOpts.hyperlink = hyperlink;
      }

      textParts.push({ text: textVal, options: textOpts });
    }
  };

  if (isPreformatted) {
    walkPreformatted(node, style, null, textParts, style, scale, globalOptions);

    // Trim leading/trailing empty break lines from the <pre> block
    while (textParts.length > 0 && textParts[0].options?.breakLine) textParts.shift();
    while (textParts.length > 0 && textParts[textParts.length - 1].options?.breakLine)
      textParts.pop();
  } else {
    const beforeRun = getPseudoTextRun(node, '::before', scale, globalOptions);
    if (beforeRun) textParts.push(beforeRun);

    node.childNodes.forEach((child, index) => {
      const isFirst = index === 0 && !beforeRun;
      const isLast = index === node.childNodes.length - 1;

      if (child.tagName === 'BR') {
        if (textParts.length > 0) {
          const lastPart = textParts[textParts.length - 1];
          if (lastPart.text && typeof lastPart.text === 'string') {
            lastPart.text = lastPart.text.trimEnd();
          }
        }
        textParts.push({ text: '', options: { breakLine: true } });
        trimNextLeading = true;
        return;
      }

      if (child.nodeType === 1) {
        const linkInfo = getHyperlinkInfo(child);
        if (linkInfo) {
          const linkStyle = window.getComputedStyle(child);
          child.childNodes.forEach((linkChild, linkIndex) => {
            const linkFirst = isFirst && linkIndex === 0;
            const linkLast = isLast && linkIndex === child.childNodes.length - 1;
            processTextNode(linkChild, linkStyle, linkFirst, linkLast, linkInfo);
          });
          return;
        }
      }

      processTextNode(child, style, isFirst, isLast);
    });

    const afterRun = getPseudoTextRun(node, '::after', scale, globalOptions);
    if (afterRun) textParts.push(afterRun);
  }

  if (textParts.length === 0) return null;

  let align = style.textAlign || 'left';
  if (align === 'start') align = 'left';
  if (align === 'end') align = 'right';
  let valign = 'top';
  if (style.alignItems === 'center') valign = 'middle';
  if (style.justifyContent === 'center' && style.display.includes('flex')) align = 'center';

  // Auto-center text when padding is symmetric (common UI pattern)
  // Skip for preformatted text which should preserve left/top alignment
  const pt = parseFloat(style.paddingTop) || 0;
  const pb = parseFloat(style.paddingBottom) || 0;
  const pl = parseFloat(style.paddingLeft) || 0;
  const pr = parseFloat(style.paddingRight) || 0;

  // Detect styled containers: solid bg, gradient bg, or visible borders
  const hasGradientBg =
    style.backgroundImage &&
    style.backgroundImage !== 'none' &&
    style.backgroundImage.includes('gradient');
  const hasBorderTop = parseFloat(style.borderTopWidth) > 0;
  const hasBorderBottom = parseFloat(style.borderBottomWidth) > 0;
  const isStyledContainer = !!(
    bgColorObj.hex ||
    hasGradientBg ||
    (hasBorderTop && hasBorderBottom)
  );

  const hasBorderLeft = parseFloat(style.borderLeftWidth) > 0;
  const hasBorderRight = parseFloat(style.borderRightWidth) > 0;
  const hasAsymmetricHBorder = hasBorderLeft !== hasBorderRight;

  if (!isPreformatted) {
    if (Math.abs(pt - pb) < 2 && isStyledContainer) valign = 'middle';
    if (Math.abs(pl - pr) < 2 && isStyledContainer && align === 'left' && !hasAsymmetricHBorder)
      align = 'center';
  }

  const padding = getPadding(style, scale);

  const fontWeight = parseInt(style.fontWeight) || 400;
  const fontSizePx = parseFloat(style.fontSize) || 16;
  const letterSpacingPx = parseFloat(style.letterSpacing) || 0;
  const maxLineChars = Math.ceil(widthPx / (fontSizePx * 0.6));
  const totalChars = textParts.reduce((sum, part) => sum + (part.text?.length || 0), 0);
  const charsForCompensation = Math.min(totalChars, maxLineChars);
  const widthCompensation = getCharSpacingWidthCompensation(
    fontWeight,
    letterSpacingPx,
    charsForCompensation,
    scale,
    fontSizePx
  );

  // PptxGenJS only supports uniform inset, so we adjust position/size for non-uniform padding
  let extraLeft = padding.left - padding.inset;
  const extraTop = padding.top - padding.inset;
  const extraRight = padding.right - padding.inset;
  const extraBottom = padding.bottom - padding.inset;

  for (const pseudo of ['::before', '::after']) {
    const ps = window.getComputedStyle(node, pseudo);
    if (!ps.content || ps.content === 'none' || ps.content === 'normal') continue;
    if (ps.display === 'none' || ps.visibility === 'hidden') continue;
    if (ps.position === 'absolute' || ps.position === 'relative') continue;
    const cnt = ps.content.replace(/^['"]|['"]$/g, '').replace(/\\(.)/g, '$1');
    if (cnt && cnt !== '""' && cnt !== "''") continue;
    const bgC = parseColor(ps.backgroundColor);
    if (!(bgC.hex && bgC.opacity > 0)) continue;
    const pW = parseFloat(ps.width) || 0;
    if (pW <= 0) continue;
    const pMargin = parseFloat(ps.marginRight) || 0;
    const offsetPx = (pW + pMargin) * PX_TO_INCH * scale;
    if (pseudo === '::before') extraLeft += offsetPx;
  }

  // Calculate line-height for single-line detection and half-leading
  const lineHeightStr = style.lineHeight;
  let lineHeightPx = fontSizePx * 1.2; // Default 1.2x
  if (lineHeightStr && lineHeightStr !== 'normal') {
    const lhValue = parseFloat(lineHeightStr);
    if (!isNaN(lhValue) && lhValue > 0) {
      if (/^[0-9.]+$/.test(lineHeightStr)) {
        lineHeightPx = lhValue * fontSizePx;
      } else if (lineHeightStr.includes('%')) {
        lineHeightPx = (lhValue / 100) * fontSizePx;
      } else if (lineHeightStr.includes('em')) {
        lineHeightPx = lhValue * fontSizePx;
      } else {
        lineHeightPx = lhValue;
      }
    }
  }

  // Detect single-line text by comparing content height to line height
  const paddingTopPx = parseFloat(style.paddingTop) || 0;
  const paddingBottomPx = parseFloat(style.paddingBottom) || 0;
  const contentHeightPx = heightPx - paddingTopPx - paddingBottomPx;
  const isSingleLine = !isPreformatted && contentHeightPx < lineHeightPx * 1.8;

  // For single-line text: reset line spacing and calculate half-leading for position adjustment
  let halfLeadingPt = 0;
  if (isSingleLine) {
    const fontSizePt = fontSizePx * PX_TO_PT * scale;
    const lineHeightPt = lineHeightPx * PX_TO_PT * scale;
    halfLeadingPt = (lineHeightPt - fontSizePt) / 2;

    textParts.forEach((part) => {
      if (part.options) {
        delete part.options.lineSpacingMultiple;
        delete part.options.lineSpacing;
      }
    });
  }

  return {
    textPayload: {
      text: textParts,
      align,
      valign,
      inset: padding.inset,
      extraPadding: { left: extraLeft, top: extraTop, right: extraRight, bottom: extraBottom },
      isSingleLine,
      halfLeadingPt,
    },
    wAdjustment: widthCompensation,
  };
}
