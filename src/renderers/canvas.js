// src/renderers/canvas.js
// Renderer for CANVAS elements

import { ElementRenderer } from './base.js';
import { applyCanvasMask } from '../image-processor.js';
import { LAYER } from '../utils/index.js';

export class CanvasRenderer extends ElementRenderer {
  render() {
    const { node, domOrder } = this;
    const dims = this.getDimensions();
    const radii = this.getBorderRadii();

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

    const item = {
      type: 'image',
      layer: LAYER.CONTENT,
      domOrder,
      options: {
        x: dims.x,
        y: dims.y,
        w: dims.w,
        h: dims.h,
        rotate: dims.rotation,
        data: null,
        ...(transparency !== undefined && { transparency }),
        ...(hyperlink && { hyperlink }),
      },
    };

    const job = async () => {
      try {
        // Directly capture canvas at its native resolution - no re-drawing!
        // Canvas may already be high-DPI (e.g., ECharts sets canvas.width = cssWidth * dpr)
        const dataUrl = node.toDataURL('image/png');

        // Basic validation
        if (!dataUrl || dataUrl.length <= 10) {
          item.skip = true;
          return;
        }

        // Check if we need to apply border-radius mask
        const needsMask = radii.tl > 0 || radii.tr > 0 || radii.br > 0 || radii.bl > 0;

        if (needsMask) {
          // Get the actual pixel dimensions of the source canvas
          const sourceWidth = node.width;
          const sourceHeight = node.height;

          // Calculate the scale ratio between canvas pixels and CSS pixels
          const scaleRatio = sourceWidth / dims.widthPx;

          // Scale the border-radius to match the source resolution
          const scaledRadii = {
            tl: radii.tl * scaleRatio,
            tr: radii.tr * scaleRatio,
            br: radii.br * scaleRatio,
            bl: radii.bl * scaleRatio,
          };

          // Apply mask at source resolution (scale=1 since we're already at native res)
          const maskedData = await applyCanvasMask(
            dataUrl,
            sourceWidth,
            sourceHeight,
            scaledRadii,
            1
          );
          item.options.data = maskedData || dataUrl;
        } else {
          item.options.data = dataUrl;
        }
      } catch (e) {
        // Tainted canvas (CORS issues) will throw here
        console.warn('Failed to capture canvas content:', e);
        item.skip = true;
      }
    };

    return { items: [item], job, stopRecursion: true };
  }
}
