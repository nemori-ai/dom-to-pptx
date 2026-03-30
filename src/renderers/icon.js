// src/renderers/icon.js
// Renderer for icon elements (FontAwesome, Material Icons, custom elements, etc.)

import { ElementRenderer } from './base.js';
import { elementToCanvasImage } from './helpers.js';
import { LAYER } from '../utils/index.js';

export class IconRenderer extends ElementRenderer {
  render() {
    const { node, domOrder, globalOptions } = this;
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
      const pngData = await elementToCanvasImage(
        node,
        dims.widthPx,
        dims.heightPx,
        globalOptions.imageScaleConfig?.html2canvas ?? 3
      );
      if (pngData) item.options.data = pngData;
      else item.skip = true;
    };

    return { items: [item], job, stopRecursion: true };
  }
}
