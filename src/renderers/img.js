// src/renderers/img.js
// Renderer for IMG elements

import { ElementRenderer } from './base.js';
import { getProcessedImage } from '../image-processor.js';
import { getFlip, getVisibleShadow, LAYER } from '../utils/index.js';

export class ImgRenderer extends ElementRenderer {
  render() {
    const { node, config, style, domOrder, globalOptions } = this;
    const dims = this.getDimensions();
    const radii = this.getBorderRadii();

    const objectFit = style.objectFit || 'fill';
    const objectPosition = style.objectPosition || '50% 50%';

    const opacity = this.getOpacity();
    const transparency = opacity < 1 ? Math.round((1 - opacity) * 100) : undefined;

    const altText = node.getAttribute('alt') || undefined;

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

    const { flipH, flipV } = getFlip(style.transform);
    const shadowInfo = getVisibleShadow(style.boxShadow, config.scale);

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
        ...(altText && { altText }),
        ...(hyperlink && { hyperlink }),
        ...(flipH && { flipH: true }),
        ...(flipV && { flipV: true }),
        ...(shadowInfo && { shadow: shadowInfo }),
      },
    };

    const job = async () => {
      const processed = await getProcessedImage(
        node.src,
        dims.widthPx,
        dims.heightPx,
        radii,
        objectFit,
        objectPosition,
        globalOptions.imageScaleConfig?.img ?? 2
      );
      if (processed) item.options.data = processed;
      else item.skip = true;
    };

    return { items: [item], job, stopRecursion: true };
  }
}
