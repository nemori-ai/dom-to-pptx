// src/renderers/svg-element.js
// Renderer for SVG elements

import { ElementRenderer, PX_TO_INCH } from './base.js';
import { svgToDataUrl, svgToPng, getClipInfo, LAYER } from '../utils/index.js';

export class SVGRenderer extends ElementRenderer {
  render() {
    const { node, config, domOrder, globalOptions } = this;
    const dims = this.getDimensions();

    const opacity = this.getOpacity();
    const transparency = opacity < 1 ? Math.round((1 - opacity) * 100) : undefined;

    let hyperlink = undefined;
    const parentAnchor = node.closest('a');
    if (parentAnchor) {
      const href = parentAnchor.getAttribute('href');
      if (href && !href.startsWith('#') && !href.startsWith('javascript:')) {
        hyperlink = {
          url: href,
          tooltip: parentAnchor.getAttribute('title') || undefined,
        };
      }
    }

    const clipInfo = getClipInfo(node);
    let cropRect = null;
    let x = dims.x,
      y = dims.y,
      w = dims.w,
      h = dims.h;
    const hasRotation = dims.rotation !== 0;

    // Rotated + clipped SVG: bake rotation into the rasterised image so the
    // bounding-box-space crop rect clips correctly, then place with rotate=0.
    if (clipInfo && hasRotation) {
      const nr = clipInfo.nodeRect;
      const cr = clipInfo.clipRect;
      const visLeft = Math.max(nr.left, cr.left);
      const visTop = Math.max(nr.top, cr.top);
      const visRight = Math.min(nr.right, cr.right);
      const visBottom = Math.min(nr.bottom, cr.bottom);
      const cropW = visRight - visLeft;
      const cropH = visBottom - visTop;
      if (cropW <= 0 || cropH <= 0) return { items: [], stopRecursion: true };

      const cropX = visLeft - nr.left;
      const cropY = visTop - nr.top;

      x = config.offX + (visLeft - config.rootX) * PX_TO_INCH * config.scale;
      y = config.offY + (visTop - config.rootY) * PX_TO_INCH * config.scale;
      w = cropW * PX_TO_INCH * config.scale;
      h = cropH * PX_TO_INCH * config.scale;

      const item = {
        type: 'image',
        layer: LAYER.CONTENT,
        domOrder,
        options: {
          data: null,
          x,
          y,
          w,
          h,
          rotate: 0, // rotation is baked into the rasterised image
          ...(transparency !== undefined && { transparency }),
          ...(hyperlink && { hyperlink }),
        },
      };

      const job = async () => {
        const scale = globalOptions.imageScaleConfig?.svg ?? 3;
        // Step 1: rasterise SVG at its pre-rotation layout size (unrotated)
        const pngResult = await svgToPng(node, scale);
        if (!pngResult) {
          item.skip = true;
          return;
        }

        const layoutW = dims.widthPx;
        const layoutH = dims.heightPx;
        const boundW = dims.rect.width;
        const boundH = dims.rect.height;
        const angleRad = (dims.rotation * Math.PI) / 180;

        // Step 2: draw the unrotated image onto a crop-sized canvas with
        // the CSS rotation applied, so the result matches on-screen appearance
        const img = new Image();
        img.src = pngResult.data;
        await new Promise((resolve) => {
          img.onload = resolve;
          img.onerror = () => resolve();
        });

        const canvas = document.createElement('canvas');
        canvas.width = Math.ceil(cropW * scale);
        canvas.height = Math.ceil(cropH * scale);
        const ctx = canvas.getContext('2d');
        ctx.scale(scale, scale);

        // Bounding-rect centre relative to the crop area
        const cx = boundW / 2 - cropX;
        const cy = boundH / 2 - cropY;
        ctx.translate(cx, cy);
        ctx.rotate(angleRad);
        ctx.drawImage(img, -layoutW / 2, -layoutH / 2, layoutW, layoutH);

        item.options.data = canvas.toDataURL('image/png');
      };

      return { items: [item], job, stopRecursion: true };
    }

    if (clipInfo) {
      const nr = clipInfo.nodeRect;
      const cr = clipInfo.clipRect;
      const visLeft = Math.max(nr.left, cr.left);
      const visTop = Math.max(nr.top, cr.top);
      const visRight = Math.min(nr.right, cr.right);
      const visBottom = Math.min(nr.bottom, cr.bottom);

      cropRect = {
        x: visLeft - nr.left,
        y: visTop - nr.top,
        w: visRight - visLeft,
        h: visBottom - visTop,
      };

      x = config.offX + (visLeft - config.rootX) * PX_TO_INCH * config.scale;
      y = config.offY + (visTop - config.rootY) * PX_TO_INCH * config.scale;
      w = cropRect.w * PX_TO_INCH * config.scale;
      h = cropRect.h * PX_TO_INCH * config.scale;
    }

    const item = {
      type: 'image',
      layer: LAYER.CONTENT,
      domOrder,
      options: {
        data: null,
        x,
        y,
        w,
        h,
        rotate: dims.rotation,
        ...(transparency !== undefined && { transparency }),
        ...(hyperlink && { hyperlink }),
      },
    };

    // When not clipped, embed SVG directly as vector (PptxGenJS supports image/svg+xml).
    // Fall back to PNG rasterization only when cropping is needed.
    if (!cropRect) {
      const result = svgToDataUrl(node);
      if (result) {
        item.options.data = result.data;
        if (result.offsetX || result.offsetY) {
          item.options.x += result.offsetX * PX_TO_INCH * config.scale;
          item.options.y += result.offsetY * PX_TO_INCH * config.scale;
          item.options.w = result.fullWidth * PX_TO_INCH * config.scale;
          item.options.h = result.fullHeight * PX_TO_INCH * config.scale;
        }
      } else {
        item.skip = true;
      }
      return { items: [item], stopRecursion: true };
    }

    // Clipped SVG: must rasterize to apply the crop
    const job = async () => {
      const result = await svgToPng(node, globalOptions.imageScaleConfig?.svg ?? 3, cropRect);
      if (result) {
        item.options.data = result.data;
      } else {
        item.skip = true;
      }
    };

    return { items: [item], job, stopRecursion: true };
  }
}
