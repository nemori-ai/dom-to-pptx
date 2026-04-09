// src/renderers/img.js
// Renderer for IMG elements

import { ElementRenderer, PX_TO_INCH } from './base.js';
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

    // Account for padding on <img> elements.
    // CSS object-fit applies to the content box (inside padding), but
    // offsetWidth/Height and getBoundingClientRect include padding.
    // Shrink PPTX dimensions to the content box so the image matches
    // the browser's visible rendering.
    let imgX = dims.x;
    let imgY = dims.y;
    let imgW = dims.w;
    let imgH = dims.h;
    let imgWidthPx = dims.widthPx;
    let imgHeightPx = dims.heightPx;

    const padTop = parseFloat(style.paddingTop) || 0;
    const padRight = parseFloat(style.paddingRight) || 0;
    const padBottom = parseFloat(style.paddingBottom) || 0;
    const padLeft = parseFloat(style.paddingLeft) || 0;

    if (padTop + padRight + padBottom + padLeft > 0) {
      imgWidthPx -= padLeft + padRight;
      imgHeightPx -= padTop + padBottom;
      imgX += padLeft * PX_TO_INCH * config.scale;
      imgY += padTop * PX_TO_INCH * config.scale;
      imgW = imgWidthPx * PX_TO_INCH * config.scale;
      imgH = imgHeightPx * PX_TO_INCH * config.scale;
    }

    const item = {
      type: 'image',
      layer: LAYER.CONTENT,
      domOrder,
      options: {
        x: imgX,
        y: imgY,
        w: imgW,
        h: imgH,
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
        imgWidthPx,
        imgHeightPx,
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
