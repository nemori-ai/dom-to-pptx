# CLAUDE.md

## Project Overview

dom-to-pptx is a client-side library that converts HTML DOM elements into fully editable PowerPoint (.pptx) slides. It uses `@zumer/snapdom` for canvas capture, `pptxgenjs` for PPTX generation, and `fonteditor-core`/`pako` for font embedding.

## Build

```bash
npx rollup -c
```

Outputs: `dist/dom-to-pptx.mjs`, `dist/dom-to-pptx.cjs`, `dist/dom-to-pptx.bundle.js`

## Debugging PPTX Conversion Issues

### Workflow: Convert → Keynote Screenshot → Compare

When fixing rendering issues (files in `need-fix/`), use this workflow:

1. **Create a Puppeteer script** to convert the HTML file:
   ```js
   // Open HTML in headless browser, inject the bundle, call exportToPptx
   await page.goto(`file://${htmlPath}`, { waitUntil: 'networkidle0' });
   await page.evaluate(bundleContent); // dist/dom-to-pptx.bundle.js
   const blob = await page.evaluate(async () => {
     return await domToPptx.exportToPptx('body', { skipDownload: true });
   });
   ```

2. **Export via Keynote AppleScript** to get a PNG of the rendered PPTX:
   ```applescript
   tell application "Keynote"
     open POSIX file "/path/to/output.pptx"
     delay 5
     export front document as slide images to POSIX file "/path/to/export" with properties {image format:PNG}
   end tell
   ```

3. **Read both the HTML screenshot and Keynote export** to visually compare and identify rendering discrepancies.

4. **Iterate**: fix source → `npx rollup -c` → re-convert → re-export → compare.

### Inspecting PPTX Internals

PPTX files are ZIP archives. To inspect the XML:
```bash
mkdir pptx-unzip && cd pptx-unzip && unzip ../output.pptx
# Slide XML: ppt/slides/slide1.xml
# Images: ppt/media/
# Relationships: ppt/slides/_rels/slide1.xml.rels
```

## Key Gotchas

### PptxGenJS Margin Units

PptxGenJS `margin` is processed by `valToPts(val)` → `Math.round(val * 12700)` EMU. Despite the name, the input is treated as **points** (1pt = 12700 EMU). **Always pass margin values in points** (inches × 72). Passing raw inches (e.g. 0.104) results in near-zero inset because `0.104 * 12700 ≈ 1323 EMU` instead of the expected `95250 EMU`.

For table cells, PptxGenJS has a separate heuristic: `margin[0] >= 1` decides points vs inches. But for text boxes, `margin` always goes through `valToPts` — use points.

### Canvas Capture and Cross-Origin Images

`checkNeedsCanvasCapture` triggers snapdom to rasterize a container when it has complex children (tiled gradients, etc.) + `overflow: hidden`. But cross-origin `<img>` elements taint the canvas, making `toDataURL()` fail. When a container has external images (`img[src^="http"]`), canvas capture is skipped and children are rendered individually instead.
