// src/utils/index.js
// Unified export for all utility modules

// Color utilities (core dependency)
export {
  parseColor,
  getGradientFallbackColor,
  INVALID_COLOR,
  blendColors,
  getEffectiveBackground,
} from './color.js';

// Text and font utilities
export {
  getTextStyle,
  isTextContainer,
  getWeightCompensation,
  getWeightWidthCompensation,
  getCharSpacingWidthCompensation,
} from './text.js';

// Table utilities
export { extractTableData } from './table.js';

// Border utilities
export { getBorderInfo, generateCompositeBorderSVG } from './border.js';

// SVG generation utilities
export {
  generateCustomShapeSVG,
  generateGradientSVG,
  generateBlurredSVG,
  generateGradientBorderSVG,
  svgToDataUrl,
  svgToPng,
} from './svg.js';

// Style utilities
export {
  getPadding,
  getSoftEdges,
  getRotation,
  getVisibleShadow,
  getRingShadow,
  isClippedByParent,
  isFullyClipped,
  getClipInfo,
  getClippingAncestor,
  resolveImageScale,
  getFlip,
} from './style.js';

// Font detection utilities
export { getUsedFontFamilies, getAutoDetectedFonts } from './font-detector.js';

// Stacking context utilities
export { LAYER, createsStackingContext, buildStackChain, compareItems } from './stacking.js';

// Shared constants
export {
  PX_TO_PT,
  FONT_SIZE_FACTOR,
  PPI,
  PX_TO_INCH,
  SLIDE_WIDTH_INCHES,
  SLIDE_HEIGHT_INCHES,
} from './constants.js';
