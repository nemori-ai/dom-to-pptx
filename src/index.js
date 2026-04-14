// src/index.js
import * as PptxGenJSImport from 'pptxgenjs';
import { PPTXEmbedFonts } from './font-embedder.js';
import JSZip from 'jszip';
import { postProcessFontSlots } from './font-slot-processor.js';

// Normalize import
const PptxGenJS = PptxGenJSImport?.default ?? PptxGenJSImport;

import {
  getUsedFontFamilies,
  getAutoDetectedFonts,
  resolveImageScale,
  isFullyClipped,
  parseColor,
  generateGradientSVG,
  buildStackChain,
  compareItems,
} from './utils/index.js';
import {
  RendererFactory,
  captureBackgroundSnapshot,
  setConcurrencyLimit,
  resetConcurrency,
} from './renderers/index.js';
import { PX_TO_INCH, SLIDE_WIDTH_INCHES, SLIDE_HEIGHT_INCHES } from './utils/constants.js';

/**
 * Main export function.
 * @param {HTMLElement | string | Array<HTMLElement | string>} target
 * @param {Object} options
 * @param {string} [options.fileName]
 * @param {boolean} [options.skipDownload=false] - If true, prevents automatic download
 * @param {Object} [options.listConfig] - Config for bullets
 * @param {number|Object} [options.imageScale] - Scale factor for image quality
 * @returns {Promise<Blob>} - Returns the generated PPTX Blob
 */
export async function exportToPptx(target, options = {}) {
  const opts = { autoEmbedFonts: true, ...options };
  const imageScaleConfig = resolveImageScale(opts.imageScale);
  const resolvePptxConstructor = (pkg) => {
    if (!pkg) return null;
    if (typeof pkg === 'function') return pkg;
    if (pkg && typeof pkg.default === 'function') return pkg.default;
    if (pkg && typeof pkg.PptxGenJS === 'function') return pkg.PptxGenJS;
    if (pkg && pkg.PptxGenJS && typeof pkg.PptxGenJS.default === 'function')
      return pkg.PptxGenJS.default;
    return null;
  };

  const PptxConstructor = resolvePptxConstructor(PptxGenJS);
  if (!PptxConstructor) throw new Error('PptxGenJS constructor not found.');
  const pptx = new PptxConstructor();
  pptx.defineLayout({
    name: 'LAYOUT_16x9_MODERN',
    width: SLIDE_WIDTH_INCHES,
    height: SLIDE_HEIGHT_INCHES,
  });
  pptx.layout = 'LAYOUT_16x9_MODERN';

  const elements = Array.isArray(target) ? target : [target];

  for (const el of elements) {
    const root = typeof el === 'string' ? document.querySelector(el) : el;
    if (!root) {
      console.warn('Element not found, skipping slide:', el);
      continue;
    }
    const slide = pptx.addSlide();
    await processSlide(root, slide, pptx, { ...opts, imageScaleConfig });
  }

  let finalBlob;
  let fontsToEmbed = opts.fonts || [];

  if (opts.autoEmbedFonts) {
    // A. Scan DOM for used font families
    const usedFamilies = getUsedFontFamilies(elements);

    // B. Scan CSS for URLs matches
    const detectedFonts = await getAutoDetectedFonts(usedFamilies);

    // C. Merge (Avoid duplicates)
    const explicitNames = new Set(fontsToEmbed.map((f) => f.name));
    for (const autoFont of detectedFonts) {
      if (!explicitNames.has(autoFont.name)) {
        fontsToEmbed.push(autoFont);
      }
    }

    if (detectedFonts.length > 0) {
      console.log(
        'Auto-detected fonts:',
        detectedFonts.map((f) => f.name)
      );
    }
  }

  if (fontsToEmbed.length > 0) {
    // Generate initial PPTX
    const initialBlob = await pptx.write({ outputType: 'blob' });

    // Load into Embedder
    const zip = await JSZip.loadAsync(initialBlob);
    const embedder = new PPTXEmbedFonts();
    await embedder.loadZip(zip);

    // Fetch all fonts in parallel, then embed sequentially (XML manipulation is not concurrent-safe)
    const fontDataArr = await Promise.all(
      fontsToEmbed.map(async (fontCfg) => {
        try {
          // Pre-resolved buffer (e.g. full TTF from GitHub)
          if (fontCfg.buffer) {
            return { name: fontCfg.name, buffer: fontCfg.buffer, type: fontCfg.type || 'ttf' };
          }
          const ext = fontCfg.url.split('.').pop().split(/[?#]/)[0].toLowerCase();
          const response = await fetch(fontCfg.url);
          if (!response.ok) throw new Error(`Failed to fetch ${fontCfg.url}`);
          const buffer = await response.arrayBuffer();
          let type = 'ttf';
          if (['woff', 'woff2', 'otf'].includes(ext)) type = ext;
          return { name: fontCfg.name, buffer, type };
        } catch (e) {
          console.warn(`Failed to fetch font: ${fontCfg.name} (${fontCfg.url})`, e);
          return null;
        }
      })
    );
    for (const fontData of fontDataArr) {
      if (fontData) {
        await embedder.addFont(fontData.name, fontData.buffer, fontData.type, {
          woff2WasmUrl: opts.woff2WasmUrl,
          hbSubsetWasmUrl: opts.hbSubsetWasmUrl,
        });
      }
    }

    await embedder.updateFiles();
    // Get actual embedded font names (after pruning unreferenced fonts)
    var embeddedFontNames = embedder.getEmbeddedFontNames();
    finalBlob = await embedder.generateBlob();
  } else {
    // No fonts to embed
    var embeddedFontNames = [];
    finalBlob = await pptx.write({ outputType: 'blob' });
  }

  // 5. Post-process: Expand font slot JSON into proper OOXML elements
  finalBlob = await postProcessFontSlots(finalBlob, { embeddedFontNames });

  // 4. Output Handling
  // If skipDownload is NOT true, proceed with browser download
  if (!opts.skipDownload) {
    const fileName = opts.fileName || 'export.pptx';
    const url = URL.createObjectURL(finalBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // Always return the blob so the caller can use it (e.g. upload to server)
  return finalBlob;
}

async function fetchImageAsBase64(url, targetW, targetH) {
  return new Promise((resolve) => {
    const img = new Image();
    img.crossOrigin = 'Anonymous';

    img.onload = () => {
      const canvas = document.createElement('canvas');
      canvas.width = targetW;
      canvas.height = targetH;
      const ctx = canvas.getContext('2d');

      const imgAspect = img.naturalWidth / img.naturalHeight;
      const targetAspect = targetW / targetH;

      let sx = 0,
        sy = 0,
        sw = img.naturalWidth,
        sh = img.naturalHeight;

      if (imgAspect > targetAspect) {
        sw = img.naturalHeight * targetAspect;
        sx = (img.naturalWidth - sw) / 2;
      } else {
        sh = img.naturalWidth / targetAspect;
        sy = (img.naturalHeight - sh) / 2;
      }

      ctx.drawImage(img, sx, sy, sw, sh, 0, 0, targetW, targetH);
      resolve(canvas.toDataURL('image/png'));
    };

    img.onerror = () => {
      console.warn('[dom-to-pptx] Failed to load background image:', url);
      resolve(null);
    };

    img.src = url;
  });
}

/**
 * Worker function to process a single DOM element into a single PPTX slide.
 * @param {HTMLElement} root - The root element for this slide.
 * @param {PptxGenJS.Slide} slide - The PPTX slide object to add content to.
 * @param {PptxGenJS} pptx - The main PPTX instance.
 */
async function processSlide(root, slide, pptx, globalOptions = {}) {
  // Reset concurrency state from any previous (possibly failed) export
  resetConcurrency();
  // Apply user-configured concurrency limit (default: 4)
  if (globalOptions.concurrency) setConcurrencyLimit(globalOptions.concurrency);

  // Ensure root acts as containing block for absolute children.
  // Without this, absolute elements may reference the viewport instead of root,
  // causing size mismatches when body width differs from viewport width.
  const origPosition = root.style.position;
  const rootPosition = window.getComputedStyle(root).position;
  if (rootPosition === 'static') {
    root.style.position = 'relative';
  }

  const rootRect = root.getBoundingClientRect();
  const PPTX_WIDTH_IN = SLIDE_WIDTH_INCHES;
  const PPTX_HEIGHT_IN = SLIDE_HEIGHT_INCHES;

  // Capture background snapshot for accurate color sampling (handles images, gradients, etc.)
  const backgroundSnapshot = await captureBackgroundSnapshot(root);
  globalOptions.backgroundSnapshot = backgroundSnapshot;

  // --- Set slide background from root element ---
  const rootStyle = window.getComputedStyle(root);
  const bgImage = rootStyle.backgroundImage;
  const bgColor = parseColor(rootStyle.backgroundColor);

  // Check if background is a multi-gradient tiled pattern (e.g., grid lines with custom backgroundSize)
  // that generateGradientSVG cannot handle. Single tiled gradients are now rendered as SVG <pattern>.
  const rootBgSize = rootStyle.backgroundSize || '';
  const rootGradientCount = (bgImage ? bgImage.match(/linear-gradient|radial-gradient/g) || [] : []).length;
  // Check if bgSize indicates actual tiling (any dimension < 100%). 'auto', '100%', 'cover' etc. are full-size.
  const rootBgParts = rootBgSize.split(/\s+/);
  const isFullDim = (v) => !v || v === 'auto' || v === '100%' || v === 'cover' || v === 'contain';
  const isRootTiledSize = rootBgSize && !rootBgParts.every(isFullDim);
  const isRootComplexTiledGradient =
    bgImage && bgImage.includes('gradient') && isRootTiledSize && rootGradientCount > 1;

  if (bgImage && bgImage !== 'none' && bgImage.includes('gradient') && !isRootComplexTiledGradient) {
    // Gradient background → convert to SVG image (with pattern tiling if bgSize is set)
    const svgData = generateGradientSVG(rootRect.width, rootRect.height, bgImage, 0, null, 0, rootBgSize || undefined);
    if (svgData) {
      slide.background = { data: svgData };
    }
  } else if (bgImage && bgImage !== 'none' && bgImage.includes('url(')) {
    // URL background image → fetch and convert to base64
    const urlMatch = bgImage.match(/url\(["']?([^"')]+)["']?\)/);
    if (urlMatch && urlMatch[1]) {
      try {
        const imgUrl = urlMatch[1];
        const bgImageData = await fetchImageAsBase64(imgUrl, rootRect.width, rootRect.height);
        if (bgImageData) {
          slide.background = { data: bgImageData };
        }
      } catch (e) {
        console.warn('[dom-to-pptx] Failed to load background image:', e);
      }
    }
  }
  // Fallback to solid background color if no background was set
  // (e.g., multi-gradient grid patterns that generateGradientSVG cannot handle)
  if (!slide.background && bgColor.hex && bgColor.opacity > 0) {
    slide.background = { color: bgColor.hex };
  }

  const contentWidthIn = rootRect.width * PX_TO_INCH;
  const contentHeightIn = rootRect.height * PX_TO_INCH;
  const scale = Math.min(PPTX_WIDTH_IN / contentWidthIn, PPTX_HEIGHT_IN / contentHeightIn);

  const layoutConfig = {
    rootX: rootRect.x,
    rootY: rootRect.y,
    scale: scale,
    offX: (PPTX_WIDTH_IN - contentWidthIn * scale) / 2,
    offY: (PPTX_HEIGHT_IN - contentHeightIn * scale) / 2,
  };

  const renderQueue = [];
  const asyncTasks = []; // Queue for heavy operations (Images, Canvas)
  let domOrderCounter = 0;

  async function collect(node, parentStackChain, parentDisplay, parentOpacity) {
    const order = domOrderCounter++;

    let currentStackChain = parentStackChain;
    let nodeStyle = null;
    const nodeType = node.nodeType;
    let accumulatedOpacity = parentOpacity;

    if (nodeType === 1) {
      nodeStyle = window.getComputedStyle(node);
      if (
        nodeStyle.display === 'none' ||
        nodeStyle.visibility === 'hidden' ||
        nodeStyle.opacity === '0'
      ) {
        return;
      }
      if (isFullyClipped(node)) {
        return;
      }

      const ownOpacity = parseFloat(nodeStyle.opacity);
      if (!isNaN(ownOpacity) && ownOpacity !== 1) {
        accumulatedOpacity *= ownOpacity;
      }

      const styleWithParent = { ...nodeStyle, _parentDisplay: parentDisplay };
      currentStackChain = buildStackChain(parentStackChain, styleWithParent, order);
    }

    const renderer = RendererFactory.create({
      node,
      config: { ...layoutConfig, root, accumulatedOpacity },
      domOrder: order,
      pptx,
      stackChain: currentStackChain,
      style: nodeStyle,
      globalOptions,
    });

    const result = renderer ? await renderer.render() : null;

    if (result) {
      if (result.items) {
        for (const item of result.items) {
          if (!item.stackChain) item.stackChain = currentStackChain;
        }
        renderQueue.push(...result.items);
      }
      if (result.job) {
        asyncTasks.push(result.job);
      }
      if (result.stopRecursion) return;
    }

    const childNodes = node.childNodes;
    const currentDisplay = nodeStyle ? nodeStyle.display : parentDisplay;
    for (let i = 0; i < childNodes.length; i++) {
      await collect(childNodes[i], currentStackChain, currentDisplay, accumulatedOpacity);
    }
  }

  await collect(root, [], null, 1);

  // 2. Execute all heavy tasks in parallel (Fast)
  if (asyncTasks.length > 0) {
    await Promise.all(asyncTasks.map((task) => task()));
  }

  // 3. Cleanup and Sort
  // Remove items that failed to generate data (marked with skip)
  const finalQueue = renderQueue.filter(
    (item) => !item.skip && (item.type !== 'image' || item.options.data)
  );

  finalQueue.sort(compareItems);

  // 4. Add to Slide
  for (const item of finalQueue) {
    if (item.type === 'shape') slide.addShape(item.shapeType, item.options);
    if (item.type === 'image') slide.addImage(item.options);
    if (item.type === 'text') slide.addText(item.textParts, item.options);
    if (item.type === 'table') {
      slide.addTable(item.tableData.rows, {
        x: item.options.x,
        y: item.options.y,
        w: item.options.w,
        colW: item.tableData.colWidths,
        rowH: item.tableData.rowHeights,
        autoPage: false,
        border: { type: 'none' },
        fill: { color: 'FFFFFF', transparency: 100 },
      });
    }
  }

  if (rootPosition === 'static') {
    root.style.position = origPosition;
  }
}
