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
