// src/renderers/shape.js
// Renderer for generic shape elements (divs with backgrounds, borders, text, etc.)

import { ElementRenderer, PX_TO_INCH } from './base.js';
import {
  parseColor,
  getTextStyle,
  isTextContainer,
  getVisibleShadow,
  getRingShadow,
  generateGradientSVG,
  generateGradientBorderSVG,
  getPadding,
  getSoftEdges,
  getBorderInfo,
  generateCustomShapeSVG,
  generateCompositeBorderSVG,
  getFlip,
  generateBlurredSVG,
  getClipInfo,
  LAYER,
} from '../utils/index.js';
import { elementToCanvasImage, captureGradientText, collectPseudoElementItems } from './helpers.js';
import { computeBorderRadii, detectGradientBorder } from './shape-border.js';
import { checkNeedsCanvasCapture } from './shape-background.js';
import { extractTextPayload } from './shape-text.js';

export class ShapeRenderer extends ElementRenderer {
  render() {
    const { node, config, style, domOrder, pptx, globalOptions } = this;
    let dims = this.getDimensions();
    let { x, y, w, h } = dims;
    let { widthPx, heightPx } = dims;
    const { rotation } = dims;

    // --- Overflow clipping: clamp to parent's visible bounds ---
    const clipInfo = getClipInfo(node);
    if (clipInfo) {
      const nr = clipInfo.nodeRect;
      const cr = clipInfo.clipRect;
      const visLeft = Math.max(nr.left, cr.left);
      const visTop = Math.max(nr.top, cr.top);
      const visRight = Math.min(nr.right, cr.right);
      const visBottom = Math.min(nr.bottom, cr.bottom);
      const clippedW = visRight - visLeft;
      const clippedH = visBottom - visTop;
      if (clippedW <= 0 || clippedH <= 0) return { items: [], stopRecursion: true };

      widthPx = clippedW;
      heightPx = clippedH;
      x = config.offX + (visLeft - config.rootX) * PX_TO_INCH * config.scale;
      y = config.offY + (visTop - config.rootY) * PX_TO_INCH * config.scale;
      w = clippedW * PX_TO_INCH * config.scale;
      h = clippedH * PX_TO_INCH * config.scale;
    }

    const items = [];
    const safeOpacity = this.getOpacity();

    const { flipH, flipV } = getFlip(style.transform);

    let shapeHyperlink = undefined;
    const anchorWrapper = node.closest('a');
    if (anchorWrapper && anchorWrapper !== node) {
      const href = anchorWrapper.getAttribute('href');
      if (href && !href.startsWith('#') && !href.startsWith('javascript:')) {
        shapeHyperlink = {
          url: href,
          tooltip: anchorWrapper.getAttribute('title') || undefined,
        };
      }
    }

    // --- Border Radius ---
    const radii = computeBorderRadii(node, style, widthPx, heightPx);

    if (clipInfo) {
      if (clipInfo.isClippedTop || clipInfo.isClippedLeft) radii.borderTopLeftRadius = 0;
      if (clipInfo.isClippedTop || clipInfo.isClippedRight) radii.borderTopRightRadius = 0;
      if (clipInfo.isClippedBottom || clipInfo.isClippedRight) radii.borderBottomRightRadius = 0;
      if (clipInfo.isClippedBottom || clipInfo.isClippedLeft) radii.borderBottomLeftRadius = 0;

      const vals = [
        radii.borderTopLeftRadius,
        radii.borderTopRightRadius,
        radii.borderBottomRightRadius,
        radii.borderBottomLeftRadius,
      ];
      const maxVal = Math.max(...vals);
      const allSame = vals.every((v) => v === maxVal);
      radii.borderRadiusValue = allSame
        ? maxVal
        : Math.min(...vals.filter((v) => v > 0), maxVal) || 0;
      radii.hasPartialBorderRadius = !allSame && vals.some((v) => v > 0);
    }

    const {
      borderTopLeftRadius,
      borderTopRightRadius,
      borderBottomRightRadius,
      borderBottomLeftRadius,
      borderRadiusValue,
      hasPartialBorderRadius,
    } = radii;

    // --- PRIORITY SVG: Solid Fill with Partial Border Radius (Vector Cone/Tab) ---
    const tempBg = parseColor(style.backgroundColor);
    const isTxt = isTextContainer(node);

    // Check if this element has meaningful child content
    const hasMeaningfulChildren =
      node.children.length > 0 || (node.textContent && node.textContent.trim().length > 0);

    let partialRadiusBgRendered = false;

    if (hasPartialBorderRadius && tempBg.hex && !isTxt) {
      const finalOpacity = safeOpacity * tempBg.opacity;
      const shapeSvg = generateCustomShapeSVG(widthPx, heightPx, tempBg.hex, finalOpacity, {
        tl: borderTopLeftRadius,
        tr: borderTopRightRadius,
        br: borderBottomRightRadius,
        bl: borderBottomLeftRadius,
      });

      items.push({
        type: 'image',
        layer: LAYER.BACKGROUND,
        domOrder,
        options: {
          data: shapeSvg,
          x,
          y,
          w,
          h,
          rotate: rotation,
          ...(flipH && { flipH: true }),
          ...(flipV && { flipV: true }),
          ...(shapeHyperlink && { hyperlink: shapeHyperlink }),
        },
      });

      partialRadiusBgRendered = true;

      // Only stop recursion if this is a pure decorative shape with no children
      if (!hasMeaningfulChildren) {
        return { items, stopRecursion: true };
      }
      // Otherwise, fall through to continue processing children
    }

    // --- ASYNC JOB: Complex Visual Elements via Canvas ---
    const bgImageStr = style.backgroundImage || '';
    const filterStr = style.filter || '';

    // Count gradients in background-image
    const gradientCount = (bgImageStr.match(/linear-gradient|radial-gradient/g) || []).length;

    // Check for large blur (glow effects)
    const blurMatch = filterStr.match(/blur\(([\d.]+)px\)/);
    const blurValue = blurMatch ? parseFloat(blurMatch[1]) : 0;
    const hasLargeBlur = blurValue > 20;

    // Large blur glow effects: use SVG with feGaussianBlur (better quality than canvas)
    if (hasLargeBlur && !hasMeaningfulChildren) {
      const bgColorObj = parseColor(style.backgroundColor);
      if (bgColorObj.hex) {
        const finalOpacity = safeOpacity * bgColorObj.opacity;
        const isCircle = borderRadiusValue >= Math.min(widthPx, heightPx) / 2 - 1;
        const svgInfo = generateBlurredSVG(
          widthPx,
          heightPx,
          bgColorObj.hex,
          isCircle ? widthPx / 2 : borderRadiusValue,
          blurValue,
          finalOpacity
        );

        const padIn = svgInfo.padding * PX_TO_INCH * config.scale;
        items.push({
          type: 'image',
          layer: LAYER.BACKGROUND,
          domOrder,
          options: {
            data: svgInfo.data,
            x: x - padIn,
            y: y - padIn,
            w: w + padIn * 2,
            h: h + padIn * 2,
            rotate: rotation,
          },
        });
        return { items, stopRecursion: true };
      }
    }

    // --- SYNC: Standard CSS Extraction ---
    const bgColorObj = parseColor(style.backgroundColor);

    const bgClip = style.webkitBackgroundClip || style.backgroundClip;
    const isBgClipText = bgClip === 'text';

    // --- Gradient Border (before canvas capture to avoid snapdom transform issues) ---
    const isGradientBorder = detectGradientBorder(style, gradientCount);

    const isRootElement = node === config.root;
    const needsCanvasCapture =
      !isGradientBorder &&
      checkNeedsCanvasCapture({
        node,
        style,
        isRootElement,
        hasPartialBorderRadius,
        bgImageStr,
      });

    if (needsCanvasCapture) {
      const marginLeft = parseFloat(style.marginLeft) || 0;
      const marginTop = parseFloat(style.marginTop) || 0;
      x += marginLeft * PX_TO_INCH * config.scale;
      y += marginTop * PX_TO_INCH * config.scale;

      const item = {
        type: 'image',
        layer: LAYER.BACKGROUND,
        domOrder,
        options: {
          x,
          y,
          w,
          h,
          rotate: rotation,
          data: null,
          ...(flipH && { flipH: true }),
          ...(flipV && { flipV: true }),
          ...(shapeHyperlink && { hyperlink: shapeHyperlink }),
        },
      };

      const job = async () => {
        const canvasImageData = await elementToCanvasImage(
          node,
          widthPx,
          heightPx,
          globalOptions.imageScaleConfig?.html2canvas ?? 3
        );
        if (canvasImageData) item.options.data = canvasImageData;
        else item.skip = true;
      };

      return { items: [item], job, stopRecursion: true };
    }

    if (isGradientBorder) {
      const borderWidth = parseFloat(style.borderWidth) || 3;
      const gradients = bgImageStr.split(/,\s*(?=linear-gradient)/);
      const borderGradient = gradients.length > 1 ? gradients[1] : gradients[0];
      let fillColor = '#' + (bgColorObj.hex || 'FFFFFF');
      if (gradients.length > 1) {
        // Extract fill color from first gradient: linear-gradient(color1, color2)
        // Use a regex that handles rgb()/rgba() commas inside parentheses
        const fillMatch = gradients[0].match(/linear-gradient\(([^)]*\([^)]*\)[^,]*),/);
        const colorStr = fillMatch ? fillMatch[1].trim() : null;
        if (colorStr) {
          const parsed = parseColor(colorStr);
          if (parsed.hex) fillColor = '#' + parsed.hex;
        }
      }

      const svgData = generateGradientBorderSVG(
        widthPx,
        heightPx,
        borderWidth,
        borderGradient,
        borderRadiusValue,
        fillColor
      );

      if (svgData) {
        items.push({
          type: 'image',
          layer: LAYER.BACKGROUND,
          domOrder,
          options: {
            data: svgData,
            x,
            y,
            w,
            h,
            rotate: rotation,
            ...(flipH && { flipH: true }),
            ...(flipV && { flipV: true }),
            ...(shapeHyperlink && { hyperlink: shapeHyperlink }),
          },
        });

        if (isTextContainer(node)) {
          const textParts = [];
          node.childNodes.forEach((child) => {
            let textVal = child.nodeType === 3 ? child.nodeValue : child.textContent;
            textVal = textVal.replace(/[\n\r\t]+/g, ' ').trim();
            if (textVal) {
              textParts.push({
                text: textVal,
                options: getTextStyle(style, config.scale, textVal, node, globalOptions),
              });
            }
          });

          if (textParts.length > 0) {
            const padding = getPadding(style, config.scale);
            let align = style.textAlign || 'left';
            if (align === 'start') align = 'left';
            if (align === 'end') align = 'right';
            if (style.justifyContent === 'center' && style.display.includes('flex')) {
              align = 'center';
            }
            let valign = 'top';
            if (style.alignItems === 'center') valign = 'middle';
            if (style.display.includes('flex') && style.alignItems === 'center') valign = 'middle';
            items.push({
              type: 'text',
              layer: LAYER.CONTENT,
              domOrder,
              textParts,
              options: {
                x: x + padding.left,
                y: y + padding.top,
                w: w - padding.left - padding.right,
                h: h - padding.top - padding.bottom,
                align,
                valign,
                margin: 0,
              },
            });
          }
        }

        return { items, stopRecursion: false };
      }
    }

    const hasGradient =
      !isBgClipText &&
      !isGradientBorder &&
      !isRootElement &&
      style.backgroundImage &&
      (style.backgroundImage.includes('linear-gradient') ||
        style.backgroundImage.includes('radial-gradient') ||
        style.backgroundImage.includes('conic-gradient'));

    // Gradient Text: Render as image (PowerPoint doesn't support native gradient text)
    if (isBgClipText && style.backgroundImage && style.backgroundImage.includes('gradient')) {
      const item = {
        type: 'image',
        layer: LAYER.CONTENT,
        domOrder,
        options: {
          x,
          y,
          w,
          h,
          data: null,
          ...(shapeHyperlink && { hyperlink: shapeHyperlink }),
        },
      };

      const job = async () => {
        const scale = globalOptions.imageScaleConfig?.html2canvas ?? 3;
        let imageData = await captureGradientText(node, widthPx, heightPx, scale);
        if (!imageData) {
          imageData = await elementToCanvasImage(node, widthPx, heightPx, scale);
        }
        if (imageData) item.options.data = imageData;
        else item.skip = true;
      };

      return { items: [item], job, stopRecursion: true };
    }

    const borderColorObj = parseColor(style.borderColor);
    const borderWidth = parseFloat(style.borderWidth);
    const hasBorder = borderWidth > 0 && borderColorObj.hex;

    const borderInfo = getBorderInfo(style, config.scale);
    const hasUniformBorder = borderInfo.type === 'uniform';
    const hasCompositeBorder = borderInfo.type === 'composite';

    const shadowStr = style.boxShadow;
    const hasShadow = shadowStr && shadowStr !== 'none';
    const softEdge = getSoftEdges(style.filter, config.scale);

    let isImageWrapper = false;
    const imgChild = Array.from(node.children).find((c) => c.tagName === 'IMG');
    if (imgChild) {
      const childW = imgChild.offsetWidth || imgChild.getBoundingClientRect().width;
      const childH = imgChild.offsetHeight || imgChild.getBoundingClientRect().height;
      if (childW >= widthPx - 2 && childH >= heightPx - 2) isImageWrapper = true;
    }

    // --- Text Extraction ---
    let textPayload = null;
    const textResult = extractTextPayload({
      node,
      style,
      scale: config.scale,
      globalOptions,
      widthPx,
      heightPx,
      w,
      bgColorObj,
    });
    let wAdjustment = 0;
    if (textResult) {
      textPayload = textResult.textPayload;
      wAdjustment = textResult.wAdjustment || 0;
    }

    const isCircularElement =
      borderRadiusValue >= Math.min(widthPx, heightPx) / 2 - 0.5 &&
      Math.abs(widthPx - heightPx) < 1;

    if (wAdjustment !== 0 && !isCircularElement) {
      if (textPayload && textPayload.align === 'center') {
        x -= wAdjustment / 2;
        w += wAdjustment;
      } else if (textPayload && textPayload.align === 'right') {
        x -= wAdjustment;
        w += wAdjustment;
      } else {
        w += wAdjustment;
      }
    }
    const hasRadialGradientOnly =
      style.backgroundImage &&
      style.backgroundImage.includes('radial-gradient') &&
      !style.backgroundImage.includes('linear-gradient');

    if (hasGradient && isCircularElement && hasRadialGradientOnly) {
      const gradMatch = style.backgroundImage.match(/radial-gradient\((.*)\)/);
      if (gradMatch) {
        if (softEdge) {
          const blurPx = softEdge / config.scale;
          const result = generateGradientSVG(
            widthPx,
            heightPx,
            style.backgroundImage,
            borderRadiusValue,
            null,
            blurPx
          );
          if (result && result.data) {
            const padIn = result.padding * PX_TO_INCH * config.scale;
            items.push({
              type: 'image',
              layer: LAYER.BACKGROUND,
              domOrder,
              options: {
                data: result.data,
                x: x - padIn,
                y: y - padIn,
                w: w + padIn * 2,
                h: h + padIn * 2,
                rotate: rotation,
              },
            });
          }
        } else {
          const result = generateGradientSVG(
            widthPx,
            heightPx,
            style.backgroundImage,
            borderRadiusValue,
            null,
            0
          );
          const imgData = typeof result === 'string' ? result : result && result.data;
          if (imgData) {
            items.push({
              type: 'image',
              layer: LAYER.BACKGROUND,
              domOrder,
              options: { data: imgData, x, y, w, h, rotate: rotation },
            });
          }
          if (hasBorder && hasUniformBorder) {
            items.push({
              type: 'shape',
              layer: LAYER.BORDER,
              domOrder,
              shapeType: pptx.ShapeType.ellipse,
              options: {
                x,
                y,
                w,
                h,
                fill: { type: 'none' },
                line: borderInfo.options,
                rotate: rotation,
              },
            });
          }
        }
      }
      if (textPayload) {
        const ep = textPayload.extraPadding || { left: 0, top: 0, right: 0, bottom: 0 };
        // Use inches directly — avoids PptxGenJS margin[0]>=1 heuristic bug
        const insetIn = textPayload.inset;
        items.push({
          type: 'text',
          layer: LAYER.CONTENT,
          domOrder,
          textParts: textPayload.text,
          options: {
            x: x + ep.left,
            y: y + ep.top,
            w: w - ep.left - ep.right,
            h: h - ep.top - ep.bottom,
            align: textPayload.align,
            valign: textPayload.valign,
            margin: [insetIn, insetIn, insetIn, insetIn],
            rotate: rotation,
            wrap: true,
            autoFit: false,
          },
        });
      }
    } else if ((hasGradient && !softEdge) || (softEdge && bgColorObj.hex && !isImageWrapper)) {
      let bgData = null;
      let padIn = 0;
      if (softEdge) {
        const finalOpacity = safeOpacity * bgColorObj.opacity;
        const svgInfo = generateBlurredSVG(
          widthPx,
          heightPx,
          bgColorObj.hex,
          borderRadiusValue,
          softEdge,
          finalOpacity
        );
        bgData = svgInfo.data;
        padIn = svgInfo.padding * PX_TO_INCH * config.scale;
      } else {
        bgData = generateGradientSVG(
          widthPx,
          heightPx,
          style.backgroundImage,
          borderRadiusValue,
          null
        );
      }

      if (bgData) {
        const imageOpts = {
          data: bgData,
          x: x - padIn,
          y: y - padIn,
          w: w + padIn * 2,
          h: h + padIn * 2,
          rotate: rotation,
        };
        // Apply shadow to gradient background image
        if (hasShadow) {
          const shadowObj = getVisibleShadow(shadowStr, config.scale);
          if (shadowObj) imageOpts.shadow = shadowObj;
        }
        items.push({
          type: 'image',
          layer: LAYER.BACKGROUND,
          domOrder,
          options: imageOpts,
        });
      }

      if (hasBorder && hasUniformBorder) {
        const isCircleShape =
          Math.abs(widthPx - heightPx) < 1 &&
          borderRadiusValue >= Math.min(widthPx, heightPx) / 2 - 0.5;
        items.push({
          type: 'shape',
          layer: LAYER.BORDER,
          domOrder,
          shapeType: isCircleShape ? pptx.ShapeType.ellipse : pptx.ShapeType.roundRect,
          options: {
            x,
            y,
            w,
            h,
            fill: { type: 'none' },
            line: borderInfo.options,
            rotate: rotation,
            ...(!isCircleShape &&
              borderRadiusValue > 0 && {
                rectRadius: borderRadiusValue * PX_TO_INCH * config.scale,
              }),
          },
        });
      }

      if (textPayload) {
        textPayload.text[0].options.fontSize = textPayload.text[0]?.options?.fontSize || 12;

        const ep = textPayload.extraPadding || { left: 0, top: 0, right: 0, bottom: 0 };
        let textX = x + ep.left;
        let textY = y + ep.top;
        let textW = w - ep.left - ep.right;
        let textH = h - ep.top - ep.bottom;

        const isSingleLine = textPayload.isSingleLine;
        if (isSingleLine && textPayload.halfLeadingPt > 0 && textPayload.valign !== 'middle') {
          const halfLeadingIn = textPayload.halfLeadingPt / 72;
          textY += halfLeadingIn;
          textH -= halfLeadingIn;
        }

        // Use inches directly — avoids PptxGenJS margin[0]>=1 heuristic bug
        const insetIn = textPayload.inset;
        items.push({
          type: 'text',
          layer: LAYER.CONTENT,
          domOrder,
          textParts: textPayload.text,
          options: {
            x: textX,
            y: textY,
            w: textW,
            h: textH,
            align: textPayload.align,
            valign: textPayload.valign,
            margin: [insetIn, insetIn, insetIn, insetIn],
            rotate: rotation,
            wrap: !isSingleLine,
            autoFit: false,
          },
        });
      }
      if (hasCompositeBorder) {
        const borderRadii = hasPartialBorderRadius
          ? {
              tl: borderTopLeftRadius,
              tr: borderTopRightRadius,
              br: borderBottomRightRadius,
              bl: borderBottomLeftRadius,
            }
          : borderRadiusValue;
        const borderSvgData = generateCompositeBorderSVG(
          widthPx,
          heightPx,
          borderRadii,
          borderInfo.sides
        );
        if (borderSvgData) {
          items.push({
            type: 'image',
            layer: LAYER.BORDER,
            domOrder,
            options: { data: borderSvgData, x, y, w, h, rotate: rotation },
          });
        }
      }
    } else if (
      (bgColorObj.hex && !isImageWrapper) ||
      hasUniformBorder ||
      hasCompositeBorder ||
      hasShadow ||
      textPayload
    ) {
      const finalAlpha = safeOpacity * bgColorObj.opacity;
      const transparency = (1 - finalAlpha) * 100;
      const useSolidFill = bgColorObj.hex && !isImageWrapper;

      if (hasPartialBorderRadius && useSolidFill && !textPayload && !partialRadiusBgRendered) {
        const shapeSvg = generateCustomShapeSVG(
          widthPx,
          heightPx,
          bgColorObj.hex,
          bgColorObj.opacity,
          {
            tl: borderTopLeftRadius,
            tr: borderTopRightRadius,
            br: borderBottomRightRadius,
            bl: borderBottomLeftRadius,
          }
        );

        items.push({
          type: 'image',
          layer: LAYER.BACKGROUND,
          domOrder,
          options: {
            data: shapeSvg,
            x,
            y,
            w,
            h,
            rotate: rotation,
            ...(flipH && { flipH: true }),
            ...(flipV && { flipV: true }),
            ...(shapeHyperlink && { hyperlink: shapeHyperlink }),
          },
        });
      } else if (!partialRadiusBgRendered) {
        const shapeOpts = {
          x,
          y,
          w,
          h,
          rotate: rotation,
          fill: useSolidFill
            ? { color: bgColorObj.hex, transparency: transparency }
            : { type: 'none' },
          line: hasUniformBorder ? borderInfo.options : null,
          ...(flipH && { flipH: true }),
          ...(flipV && { flipV: true }),
          ...(shapeHyperlink && { hyperlink: shapeHyperlink }),
        };

        if (hasShadow) shapeOpts.shadow = getVisibleShadow(shadowStr, config.scale);

        const minDimension = Math.min(widthPx, heightPx);
        const radiusPx = borderRadiusValue;

        let shapeType = pptx.ShapeType.rect;

        const isSquare = Math.abs(widthPx - heightPx) < 1;
        const isFullyRound = radiusPx >= minDimension / 2 - 0.5;
        const isCssBorderRadius50 =
          style.borderRadius === '50%' ||
          (style.borderTopLeftRadius?.includes('50%') &&
            style.borderTopRightRadius?.includes('50%') &&
            style.borderBottomRightRadius?.includes('50%') &&
            style.borderBottomLeftRadius?.includes('50%'));
        const cssW = parseFloat(style.width) || 0;
        const cssH = parseFloat(style.height) || 0;
        const isCssSquare = cssW > 0 && cssH > 0 && Math.abs(cssW - cssH) < 1;

        if ((isFullyRound && isSquare) || (isCssBorderRadius50 && isCssSquare)) {
          shapeType = pptx.ShapeType.ellipse;
          if (isCssSquare && !isSquare) {
            w = cssW * PX_TO_INCH * config.scale;
            h = cssH * PX_TO_INCH * config.scale;
          }
        } else if (radiusPx > 0) {
          shapeType = pptx.ShapeType.roundRect;
          shapeOpts.rectRadius = radiusPx * PX_TO_INCH * config.scale;
        }

        // Ring: box-shadow with spread-only (e.g. Tailwind ring-4) → outline shape behind element
        const ringInfo = hasShadow ? getRingShadow(shadowStr) : null;
        if (ringInfo) {
          const spreadIn = ringInfo.spread * PX_TO_INCH * config.scale;
          items.push({
            type: 'shape',
            layer: LAYER.BACKGROUND,
            domOrder: domOrder - 0.5,
            shapeType,
            options: {
              x: x - spreadIn,
              y: y - spreadIn,
              w: w + spreadIn * 2,
              h: h + spreadIn * 2,
              fill: { color: ringInfo.color, transparency: (1 - ringInfo.opacity) * 100 },
              line: { type: 'none' },
              ...(shapeType === pptx.ShapeType.roundRect && {
                rectRadius: (radiusPx + ringInfo.spread) * PX_TO_INCH * config.scale,
              }),
            },
          });
        }

        if (textPayload) {
          textPayload.text[0].options.fontSize = textPayload.text[0]?.options?.fontSize || 12;

          const ep = textPayload.extraPadding || { left: 0, top: 0, right: 0, bottom: 0 };
          const hasNonUniformPadding = ep.left > 0 || ep.top > 0 || ep.right > 0 || ep.bottom > 0;
          const isSingleLine = textPayload.isSingleLine;

          let halfLeadingIn = 0;
          if (isSingleLine && textPayload.halfLeadingPt > 0 && textPayload.valign !== 'middle') {
            halfLeadingIn = textPayload.halfLeadingPt / 72;
          }

          if (hasNonUniformPadding || isSingleLine) {
            if (useSolidFill || hasUniformBorder) {
              items.push({
                type: 'shape',
                layer: LAYER.BACKGROUND,
                domOrder,
                shapeType,
                options: shapeOpts,
              });
            }

            let textX = x + ep.left;
            let textY = y + ep.top + halfLeadingIn;
            let textW = w - ep.left - ep.right;
            let textH = h - ep.top - ep.bottom - halfLeadingIn;

            // Use inches directly — avoids PptxGenJS margin[0]>=1 heuristic bug
        const insetIn = textPayload.inset;
            items.push({
              type: 'text',
              layer: LAYER.CONTENT,
              domOrder,
              textParts: textPayload.text,
              options: {
                x: textX,
                y: textY,
                w: textW,
                h: textH,
                align: textPayload.align,
                valign: textPayload.valign,
                margin: [insetIn, insetIn, insetIn, insetIn],
                rotate: rotation,
                wrap: !isSingleLine,
                autoFit: false,
              },
            });
          } else {
            // Use inches directly — avoids PptxGenJS margin[0]>=1 heuristic bug
            const uniformInsetIn = textPayload.inset;
            const textOptions = {
              shape: shapeType,
              ...shapeOpts,
              rotate: rotation,
              align: textPayload.align,
              valign: textPayload.valign,
              margin: [uniformInsetIn, uniformInsetIn, uniformInsetIn, uniformInsetIn],
              wrap: true,
              autoFit: false,
            };
            items.push({
              type: 'text',
              layer: LAYER.CONTENT,
              domOrder,
              textParts: textPayload.text,
              options: textOptions,
            });
          }
        } else if (!hasPartialBorderRadius) {
          items.push({
            type: 'shape',
            layer: LAYER.BACKGROUND,
            domOrder,
            shapeType,
            options: shapeOpts,
          });
        }
      }

      if (hasCompositeBorder) {
        const borderRadii = hasPartialBorderRadius
          ? {
              tl: borderTopLeftRadius,
              tr: borderTopRightRadius,
              br: borderBottomRightRadius,
              bl: borderBottomLeftRadius,
            }
          : borderRadiusValue;
        const borderSvgData = generateCompositeBorderSVG(
          widthPx,
          heightPx,
          borderRadii,
          borderInfo.sides
        );
        if (borderSvgData) {
          items.push({
            type: 'image',
            layer: LAYER.BORDER,
            domOrder,
            options: { data: borderSvgData, x, y, w, h, rotate: rotation },
          });
        }
      }
    }

    const pseudoResult = collectPseudoElementItems(node, config, domOrder, x, y, widthPx, heightPx);
    items.push(...pseudoResult.items);

    const result = { items, stopRecursion: !!textPayload };
    if (pseudoResult.jobs.length > 0) {
      result.job = async () => {
        await Promise.all(pseudoResult.jobs.map((j) => j()));
      };
    }
    return result;
  }
}
