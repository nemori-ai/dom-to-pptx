// src/utils/constants.js
// Shared numeric constants used across the codebase

/** Conversion factor: CSS pixels to PowerPoint points (1pt = 1/72in, 1px = 1/96in → 96/72 = 0.75) */
export const PX_TO_PT = 0.75;

/** Empirical font size shrink factor for PowerPoint text rendering */
export const FONT_SIZE_FACTOR = 0.95;

/** Pixels per inch (CSS standard) */
export const PPI = 96;

/** Conversion factor: CSS pixels to inches (1 / PPI) */
export const PX_TO_INCH = 1 / PPI;

/** Standard 16:9 slide width in inches (Modern PowerPoint 2013+) */
export const SLIDE_WIDTH_INCHES = 13.333;

/** Standard 16:9 slide height in inches (Modern PowerPoint 2013+) */
export const SLIDE_HEIGHT_INCHES = 7.5;
