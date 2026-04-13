// src/hb-subset.js
// Thin wrapper around hb-subset.wasm for variable font instancing.
// Converts a variable TTF to a static TTF by pinning the wght axis to 400 (Regular).

let hbExports = null;
let initPromise = null;

/**
 * Resolve the hb-subset.wasm URL.
 */
function resolveWasmUrl(wasmUrl) {
  if (wasmUrl) return wasmUrl;
  try {
    // eslint-disable-next-line no-undef
    if (typeof __HB_SUBSET_WASM_INLINE__ === 'string') {
      // eslint-disable-next-line no-undef
      return __HB_SUBSET_WASM_INLINE__;
    }
  } catch {
    // not defined
  }
  if (typeof document !== 'undefined' && document.currentScript && document.currentScript.src) {
    const scriptDir = document.currentScript.src.replace(/\/[^/]*$/, '/');
    return scriptDir + 'hb-subset.wasm';
  }
  return undefined;
}

async function ensureHbReady(wasmUrl) {
  if (hbExports) return;
  if (!initPromise) {
    initPromise = (async () => {
      const url = resolveWasmUrl(wasmUrl);
      let wasmSource;
      if (typeof url === 'string' && url.startsWith('data:')) {
        // base64 data URI — decode to ArrayBuffer
        const base64 = url.split(',')[1];
        const binary = atob(base64);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
        wasmSource = bytes.buffer;
      } else if (url) {
        const resp = await fetch(url);
        wasmSource = await resp.arrayBuffer();
      } else {
        throw new Error('hb-subset.wasm not found: provide wasmUrl or inline via __HB_SUBSET_WASM_INLINE__');
      }
      const { instance } = await WebAssembly.instantiate(wasmSource);
      hbExports = instance.exports;
    })().catch((err) => {
      initPromise = null;
      throw err;
    });
  }
  await initPromise;
}

// HarfBuzz axis tag: encode 4 ASCII chars as uint32 big-endian
function hbTag(s) {
  return (s.charCodeAt(0) << 24) | (s.charCodeAt(1) << 16) | (s.charCodeAt(2) << 8) | s.charCodeAt(3);
}

/**
 * Convert a variable font to a static font by pinning the wght axis to 400 (Regular).
 * Uses hb-subset.wasm with hb_subset_input_pin_axis_location.
 * All glyphs are preserved (no subsetting).
 *
 * @param {ArrayBuffer} ttfBuffer - Variable TTF font data
 * @param {object} [opts]
 * @param {string} [opts.hbSubsetWasmUrl] - URL to hb-subset.wasm
 * @returns {Promise<ArrayBuffer>} Static TTF font data
 */
export async function instantiateVariable(ttfBuffer, opts = {}) {
  await ensureHbReady(opts.hbSubsetWasmUrl);
  const hb = hbExports;

  const fontBytes = new Uint8Array(ttfBuffer);
  // Must re-read HEAPU8 after each allocation — malloc/harfbuzz internals may grow memory,
  // invalidating any previously cached Uint8Array view of the WASM heap.
  const heapu8 = () => new Uint8Array(hb.memory.buffer);

  // Upload font to WASM heap
  const fontPtr = hb.malloc(fontBytes.byteLength);
  heapu8().set(fontBytes, fontPtr);

  const blob = hb.hb_blob_create(fontPtr, fontBytes.byteLength, 2 /* HB_MEMORY_MODE_WRITABLE */, 0, 0);
  const face = hb.hb_face_create(blob, 0);
  hb.hb_blob_destroy(blob);

  const input = hb.hb_subset_input_create_or_fail();
  if (!input) {
    hb.hb_face_destroy(face);
    hb.free(fontPtr);
    throw new Error('hb_subset_input_create_or_fail returned null');
  }

  // Keep all glyphs — we only want to instance, not subset
  hb.hb_subset_input_keep_everything(input);

  // Pin all axes to their default values first (handles wdth, ital, opsz, etc.)
  hb.hb_subset_input_pin_all_axes_to_default(input, face);

  // Override wght to 400 (Regular) — some fonts have default wght != 400
  // (e.g. Noto Sans SC defaults to 100/Thin). Always force Regular weight.
  hb.hb_subset_input_pin_axis_location(input, face, hbTag('wght'), 400.0);

  // Execute subset (arg order: face, input)
  const subsetFace = hb.hb_subset_or_fail(face, input);
  hb.hb_subset_input_destroy(input);

  if (!subsetFace) {
    hb.hb_face_destroy(face);
    hb.free(fontPtr);
    throw new Error('hb_subset_or_fail returned null');
  }

  // Extract result
  const resultBlob = hb.hb_face_reference_blob(subsetFace);
  const dataOffset = hb.hb_blob_get_data(resultBlob, 0);
  const dataLength = hb.hb_blob_get_length(resultBlob);

  // Copy out before freeing (memory may be reused)
  const result = heapu8().slice(dataOffset, dataOffset + dataLength).buffer;

  // Cleanup
  hb.hb_blob_destroy(resultBlob);
  hb.hb_face_destroy(subsetFace);
  hb.hb_face_destroy(face);
  hb.free(fontPtr);

  return result;
}
