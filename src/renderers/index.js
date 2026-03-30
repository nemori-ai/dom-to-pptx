// src/renderers/index.js
// Unified export for all renderer modules

export { ElementRenderer, PX_TO_INCH } from './base.js';
export { RendererFactory } from './factory.js';

// Individual renderers (for direct use if needed)
export { TextNodeRenderer } from './text-node.js';
export { TableRenderer } from './table.js';
export { ListRenderer } from './list.js';
export { CanvasRenderer } from './canvas.js';
export { SVGRenderer } from './svg-element.js';
export { ImgRenderer } from './img.js';
export { IconRenderer } from './icon.js';
export { ShapeRenderer } from './shape.js';

// Helper functions
export {
  elementToCanvasImage,
  isIconElement,
  isComplexHierarchy,
  collectListParts,
  createCompositeBorderItems,
  captureBackgroundSnapshot,
  sampleFromSnapshot,
  sampleAverageColor,
  setConcurrencyLimit,
  resetConcurrency,
} from './helpers.js';
