// src/font-utils.js
import { Font, woff2 } from 'fonteditor-core';
import pako from 'pako';

let woff2InitPromise = null;

/**
 * Resolves the woff2.wasm URL.
 * Priority: explicit wasmUrl > auto-detect from current script location.
 * For the browser bundle, Rollup inlines the WASM as a base64 data URL
 * via the __WOFF2_WASM_INLINE__ placeholder.
 */
function resolveWasmUrl(wasmUrl) {
  if (wasmUrl) return wasmUrl;

  // In browser bundle builds, __WOFF2_WASM_INLINE__ is replaced with a base64 data URL
  // by the Rollup inlineWoff2Wasm plugin. In non-bundle builds it stays as-is (undefined).
  try {
    // eslint-disable-next-line no-undef
    if (typeof __WOFF2_WASM_INLINE__ === 'string') {
      // eslint-disable-next-line no-undef
      return __WOFF2_WASM_INLINE__;
    }
  } catch {
    // __WOFF2_WASM_INLINE__ not defined — non-bundle build, fall through
  }

  // Fallback: try to derive from current script src (CDN / self-hosted)
  if (typeof document !== 'undefined' && document.currentScript && document.currentScript.src) {
    const scriptDir = document.currentScript.src.replace(/\/[^/]*$/, '/');
    return scriptDir + 'woff2.wasm';
  }

  return undefined; // Node.js — fonteditor-core resolves it automatically
}

/**
 * Initializes the WOFF2 WASM module (idempotent).
 * Must be called before decoding WOFF2 fonts.
 * @param {string} [wasmUrl] - URL to the woff2.wasm file
 */
async function ensureWoff2Ready(wasmUrl) {
  if (woff2.isInited()) return;
  if (!woff2InitPromise) {
    const resolvedUrl = resolveWasmUrl(wasmUrl);
    woff2InitPromise = woff2.init(resolvedUrl).catch((err) => {
      woff2InitPromise = null;
      throw err;
    });
  }
  await woff2InitPromise;
}

/**
 * Converts various font formats to EOT (Embedded OpenType),
 * which is highly compatible with PowerPoint embedding.
 * @param {string} type - 'ttf', 'woff', 'woff2', or 'otf'
 * @param {ArrayBuffer} fontBuffer - The raw font data
 * @param {object} [opts] - Options
 * @param {string} [opts.woff2WasmUrl] - URL to the woff2.wasm file (required for WOFF2 in browser)
 */
export async function fontToEot(type, fontBuffer, opts = {}) {
  let actualType = type;
  let actualBuffer = fontBuffer;

  if (type === 'woff2') {
    await ensureWoff2Ready(opts.woff2WasmUrl);
    const ttfBytes = woff2.decode(fontBuffer);
    actualBuffer = ttfBytes.buffer.slice(
      ttfBytes.byteOffset,
      ttfBytes.byteOffset + ttfBytes.byteLength
    );
    actualType = 'ttf';
  }

  const options = {
    type: actualType,
    hinting: true,
    inflate: actualType === 'woff' ? pako.inflate : undefined,
  };

  const font = Font.create(actualBuffer, options);

  const eotBuffer = font.write({
    type: 'eot',
    toBuffer: true,
  });

  if (eotBuffer instanceof ArrayBuffer) {
    return eotBuffer;
  }

  // Ensure we return an ArrayBuffer
  return eotBuffer.buffer.slice(eotBuffer.byteOffset, eotBuffer.byteOffset + eotBuffer.byteLength);
}
