// src/utils/stacking.js
// CSS Stacking Context modeling for correct z-index handling

export const LAYER = {
  BACKGROUND: 0,
  BORDER: 1,
  CONTENT: 2,
  OVERLAY: 3,
};

export function createsStackingContext(style) {
  if (!style) return false;

  const position = style.position;
  const zIndex = style.zIndex;

  // positioned element with z-index !== auto
  if (
    (position === 'absolute' ||
      position === 'relative' ||
      position === 'fixed' ||
      position === 'sticky') &&
    zIndex !== 'auto'
  ) {
    return true;
  }

  // opacity < 1
  const opacity = parseFloat(style.opacity);
  if (!isNaN(opacity) && opacity < 1) {
    return true;
  }

  // transform !== none
  if (style.transform && style.transform !== 'none') {
    return true;
  }

  // filter !== none
  if (style.filter && style.filter !== 'none') {
    return true;
  }

  // isolation: isolate
  if (style.isolation === 'isolate') {
    return true;
  }

  // mix-blend-mode !== normal
  if (style.mixBlendMode && style.mixBlendMode !== 'normal') {
    return true;
  }

  // will-change with certain values
  const willChange = style.willChange;
  if (
    willChange &&
    (willChange.includes('transform') ||
      willChange.includes('opacity') ||
      willChange.includes('filter'))
  ) {
    return true;
  }

  // contain: layout, paint, strict, content
  const contain = style.contain;
  if (
    contain &&
    (contain.includes('layout') ||
      contain.includes('paint') ||
      contain.includes('strict') ||
      contain.includes('content'))
  ) {
    return true;
  }

  // flex/grid item with z-index !== auto
  const parentDisplay = style._parentDisplay;
  if (
    parentDisplay &&
    (parentDisplay.includes('flex') || parentDisplay.includes('grid')) &&
    zIndex !== 'auto'
  ) {
    return true;
  }

  return false;
}

export function buildStackChain(parentChain, style, domOrder) {
  if (!createsStackingContext(style)) {
    return parentChain;
  }

  const zIndex = style.zIndex === 'auto' ? 0 : parseInt(style.zIndex) || 0;

  // Each entry: [zIndex, domOrder] - domOrder breaks ties at same z-index level
  return [...parentChain, [zIndex, domOrder]];
}

export function compareStackChains(chainA, chainB) {
  const len = Math.max(chainA.length, chainB.length);

  for (let i = 0; i < len; i++) {
    const hasA = i < chainA.length;
    const hasB = i < chainB.length;

    if (hasA && hasB) {
      // Both have entries at this level - compare normally
      if (chainA[i][0] !== chainB[i][0]) return chainA[i][0] - chainB[i][0];
      if (chainA[i][1] !== chainB[i][1]) return chainA[i][1] - chainB[i][1];
    } else if (hasA && !hasB) {
      // Only A creates a stacking context at this level
      // Non-zero z-index: decisive (positive = A on top, negative = A below)
      // Zero z-index: same paint layer as auto, let domOrder decide
      if (chainA[i][0] !== 0) return chainA[i][0];
      return 0;
    } else {
      // Only B creates a stacking context at this level
      if (chainB[i][0] !== 0) return -chainB[i][0];
      return 0;
    }
  }

  return 0;
}

export function compareItems(a, b) {
  // 1. Compare stacking context chains
  const chainCmp = compareStackChains(a.stackChain || [], b.stackChain || []);
  if (chainCmp !== 0) return chainCmp;

  // 2. Compare DOM order (later in DOM = on top)
  if (a.domOrder !== b.domOrder) return a.domOrder - b.domOrder;

  // 3. Compare layer within same element (background < border < content < overlay)
  const layerA = a.layer ?? LAYER.CONTENT;
  const layerB = b.layer ?? LAYER.CONTENT;
  return layerA - layerB;
}
