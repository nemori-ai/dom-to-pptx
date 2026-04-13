// src/font-utils.js
import { Font, woff2 } from 'fonteditor-core';
import pako from 'pako';
import { instantiateVariable } from './hb-subset.js';

let woff2InitPromise = null;

/**
 * Resolves the woff2.wasm URL.
 */
function resolveWasmUrl(wasmUrl) {
  if (wasmUrl) return wasmUrl;
  try {
    // eslint-disable-next-line no-undef
    if (typeof __WOFF2_WASM_INLINE__ === 'string') {
      // eslint-disable-next-line no-undef
      return __WOFF2_WASM_INLINE__;
    }
  } catch {
    // __WOFF2_WASM_INLINE__ not defined
  }
  if (typeof document !== 'undefined' && document.currentScript && document.currentScript.src) {
    const scriptDir = document.currentScript.src.replace(/\/[^/]*$/, '/');
    return scriptDir + 'woff2.wasm';
  }
  return undefined;
}

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
 * Read name records from the TTF name table in a single pass.
 * Matches fontello/ttf2eot: platformID=3, encodingID=1, languageID=0x0409.
 * Returns raw UTF-16BE Uint8Array for nameID 1 (family), 2 (subfamily), 4 (full), 5 (version).
 * @param {ArrayBuffer} ttf
 * @returns {{ familyName: Uint8Array, subfamilyName: Uint8Array, fullName: Uint8Array, versionString: Uint8Array }}
 */
function readTtfNameRecords(ttf) {
  const view = new DataView(ttf);
  const numTables = view.getUint16(4);
  const result = {
    familyName: new Uint8Array(0),
    subfamilyName: new Uint8Array(0),
    fullName: new Uint8Array(0),
    versionString: new Uint8Array(0),
  };

  for (let i = 0; i < numTables; i++) {
    const off = 12 + i * 16;
    const tag =
      String.fromCharCode(view.getUint8(off)) +
      String.fromCharCode(view.getUint8(off + 1)) +
      String.fromCharCode(view.getUint8(off + 2)) +
      String.fromCharCode(view.getUint8(off + 3));
    if (tag !== 'name') continue;

    const tableOff = view.getUint32(off + 8);
    const count = view.getUint16(tableOff + 2);
    const stringOff = view.getUint16(tableOff + 4);

    for (let j = 0; j < count; j++) {
      const recOff = tableOff + 6 + j * 12;
      const platID = view.getUint16(recOff);
      const encID = view.getUint16(recOff + 2);
      const langID = view.getUint16(recOff + 4);
      const nameID = view.getUint16(recOff + 6);
      const length = view.getUint16(recOff + 8);
      const strOff = view.getUint16(recOff + 10);

      if (platID !== 3 || encID !== 1 || langID !== 0x0409) continue;

      const bytes = new Uint8Array(ttf, tableOff + stringOff + strOff, length);
      switch (nameID) {
        case 1: result.familyName = bytes; break;
        case 2: result.subfamilyName = bytes; break;
        case 4: result.fullName = bytes; break;
        case 5: result.versionString = bytes; break;
      }
    }
    break;
  }
  return result;
}

/**
 * Build an EOT (Embedded OpenType) wrapper around raw TTF data.
 * Uses the fontello/ttf2eot structure (v2.1, 82-byte prefix) which is proven
 * compatible with PowerPoint for both Latin and CJK fonts.
 * @param {ArrayBuffer} ttfBuffer - Raw TrueType font data
 * @returns {ArrayBuffer} EOT file
 */
function buildEot(ttfBuffer) {
  const ttf = ttfBuffer;
  const ttfView = new DataView(ttf);
  const fontDataSize = ttf.byteLength;

  // 82-byte prefix: 70 bytes of fixed fields + 12 bytes of empty name slots
  // (FamilyNameSize=0 + Padding2=0 + StyleNameSize=0 + Padding3=0 + VersionNameSize=0 + Padding4=0)
  // The real name data is appended after the prefix, matching fontello/ttf2eot structure.
  const EOT_PREFIX_SIZE = 82;
  const prefix = new ArrayBuffer(EOT_PREFIX_SIZE);
  const pv = new DataView(prefix);
  const pb = new Uint8Array(prefix);

  // FontDataSize
  pv.setUint32(4, fontDataSize, true);
  // Version 0x00020001 (v2.1)
  pv.setUint32(8, 0x00020001, true);
  // Charset = 1 (Unicode)
  pb[26] = 1;
  // Magic "LP"
  pv.setUint16(34, 0x504c, true);

  // Parse TTF tables: OS/2, head (with early exit like fontello/ttf2eot).
  // Name table is parsed separately by readTtfNameRecords, so only OS/2+head are tracked here.
  const numTables = ttfView.getUint16(4);
  const names = readTtfNameRecords(ttf);
  let haveOS2 = false, haveHead = false;

  for (let i = 0; i < numTables; i++) {
    const off = 12 + i * 16;
    const tag =
      String.fromCharCode(ttfView.getUint8(off)) +
      String.fromCharCode(ttfView.getUint8(off + 1)) +
      String.fromCharCode(ttfView.getUint8(off + 2)) +
      String.fromCharCode(ttfView.getUint8(off + 3));

    if (tag === 'OS/2') {
      haveOS2 = true;
      const tOff = ttfView.getUint32(off + 8);
      const os2Ver = ttfView.getUint16(tOff);
      // PANOSE (10 bytes at OS/2 offset 32)
      pb.set(new Uint8Array(ttf, tOff + 32, 10), 16);
      // Italic (fsSelection bit 0)
      pb[27] = ttfView.getUint16(tOff + 62) & 1;
      // Weight
      pv.setUint32(28, ttfView.getUint16(tOff + 4), true);
      // UnicodeRange (4 × uint32, big-endian → little-endian)
      for (let j = 0; j < 4; j++) {
        pv.setUint32(36 + j * 4, ttfView.getUint32(tOff + 42 + j * 4), true);
      }
      // CodePageRange (OS/2 version >= 1)
      if (os2Ver >= 1) {
        for (let j = 0; j < 2; j++) {
          pv.setUint32(52 + j * 4, ttfView.getUint32(tOff + 78 + j * 4), true);
        }
      }
    } else if (tag === 'head') {
      haveHead = true;
      const tOff = ttfView.getUint32(off + 8);
      pv.setUint32(60, ttfView.getUint32(tOff + 8), true); // CheckSumAdjustment
    }
    if (haveOS2 && haveHead) break;
  }

  if (!haveOS2 || !haveHead) {
    throw new Error('buildEot: required OS/2 or head table not found');
  }

  // Build name buffers in strbuf format: [uint16LE size][utf16le data][uint16LE 0x0000]
  // Swaps UTF-16BE (from TTF name table) to UTF-16LE.
  function strbuf(utf16be) {
    const len = utf16be.length;
    const buf = new Uint8Array(2 + len + 2);
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    dv.setUint16(0, len, true);
    for (let i = 0; i < len; i += 2) {
      buf[2 + i] = utf16be[i + 1];
      buf[2 + i + 1] = utf16be[i];
    }
    return buf;
  }

  const familyBuf = strbuf(names.familyName);
  const subfamilyBuf = strbuf(names.subfamilyName);
  const versionBuf = strbuf(names.versionString);
  const fullBuf = strbuf(names.fullName);

  // Assemble: prefix(82) + 4 name bufs + null terminator(2) + font data
  const nullTerm = new Uint8Array(2); // [0x00, 0x00]
  const totalSize = EOT_PREFIX_SIZE + familyBuf.length + subfamilyBuf.length +
    versionBuf.length + fullBuf.length + 2 + fontDataSize;

  const eot = new ArrayBuffer(totalSize);
  const bytes = new Uint8Array(eot);
  let pos = 0;

  bytes.set(pb, pos); pos += EOT_PREFIX_SIZE;
  bytes.set(familyBuf, pos); pos += familyBuf.length;
  bytes.set(subfamilyBuf, pos); pos += subfamilyBuf.length;
  bytes.set(versionBuf, pos); pos += versionBuf.length;
  bytes.set(fullBuf, pos); pos += fullBuf.length;
  bytes.set(nullTerm, pos); pos += 2;
  bytes.set(new Uint8Array(ttf), pos);

  // Write total EOT size at offset 0
  new DataView(eot).setUint32(0, totalSize, true);

  return eot;
}

/**
 * Detect whether a TTF buffer is a variable font by scanning for the 'fvar' table.
 * @param {ArrayBuffer} buf - Raw TTF data
 * @returns {boolean}
 */
function isVariableFont(buf) {
  try {
    const view = new DataView(buf);
    const numTables = view.getUint16(4);
    for (let i = 0; i < numTables; i++) {
      const off = 12 + i * 16;
      const tag =
        String.fromCharCode(view.getUint8(off)) +
        String.fromCharCode(view.getUint8(off + 1)) +
        String.fromCharCode(view.getUint8(off + 2)) +
        String.fromCharCode(view.getUint8(off + 3));
      if (tag === 'fvar') return true;
    }
  } catch {
    // If buffer is too small or malformed, not a variable font
  }
  return false;
}

/**
 * Converts various font formats to EOT for PowerPoint .fntdata embedding.
 * Builds the EOT header manually to match PowerPoint's expected format.
 * @param {string} type - 'ttf', 'woff', 'woff2', or 'otf'
 * @param {ArrayBuffer} fontBuffer - The raw font data
 * @param {object} [opts] - Options
 * @param {string} [opts.woff2WasmUrl] - URL to the woff2.wasm file
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

  // Convert non-TTF formats to TTF using fonteditor-core
  if (actualType !== 'ttf') {
    const options = {
      type: actualType,
      hinting: true,
      inflate: actualType === 'woff' ? pako.inflate : undefined,
    };
    const font = Font.create(actualBuffer, options);
    const ttfOut = font.write({ type: 'ttf', toBuffer: true });
    actualBuffer =
      ttfOut instanceof ArrayBuffer
        ? ttfOut
        : ttfOut.buffer.slice(ttfOut.byteOffset, ttfOut.byteOffset + ttfOut.byteLength);
  }

  // Instance variable fonts to static using harfbuzz subset WASM.
  // Pins wght axis to 400 (Regular) and removes all variation tables,
  // producing a proper static TTF with correct glyph outlines and metrics.
  if (isVariableFont(actualBuffer)) {
    actualBuffer = await instantiateVariable(actualBuffer, {
      hbSubsetWasmUrl: opts.hbSubsetWasmUrl,
    });
  }

  // Ensure we have a proper ArrayBuffer
  if (!(actualBuffer instanceof ArrayBuffer)) {
    actualBuffer = actualBuffer.buffer.slice(
      actualBuffer.byteOffset,
      actualBuffer.byteOffset + actualBuffer.byteLength
    );
  }

  return buildEot(actualBuffer);
}
