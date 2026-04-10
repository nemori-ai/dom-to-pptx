// src/font-utils.js
import { Font, woff2 } from 'fonteditor-core';
import pako from 'pako';

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
 * Read the font's full name from the TTF name table (nameID=4, platformID=3).
 */
function readTtfFullName(ttf) {
  const view = new DataView(ttf);
  const numTables = view.getUint16(4);
  for (let i = 0; i < numTables; i++) {
    const off = 12 + i * 16;
    const tag =
      String.fromCharCode(view.getUint8(off)) +
      String.fromCharCode(view.getUint8(off + 1)) +
      String.fromCharCode(view.getUint8(off + 2)) +
      String.fromCharCode(view.getUint8(off + 3));
    if (tag === 'name') {
      const tableOff = view.getUint32(off + 8);
      const count = view.getUint16(tableOff + 2);
      const stringOff = view.getUint16(tableOff + 4);

      const readNameRecord = (targetNameID) => {
        for (let j = 0; j < count; j++) {
          const recOff = tableOff + 6 + j * 12;
          const platID = view.getUint16(recOff);
          const nameID = view.getUint16(recOff + 6);
          const length = view.getUint16(recOff + 8);
          const strOff = view.getUint16(recOff + 10);
          if (nameID === targetNameID && platID === 3) {
            const bytes = new Uint8Array(ttf, tableOff + stringOff + strOff, length);
            let name = '';
            for (let k = 0; k < bytes.length; k += 2) {
              name += String.fromCharCode((bytes[k] << 8) | bytes[k + 1]);
            }
            return name;
          }
        }
        return null;
      };

      // Prefer nameID=16 (Typographic Family) for variable fonts where
      // nameID=4 contains weight-specific names like "Fira Code Light".
      // Fall back to nameID=4 (Full Name) then nameID=1 (Family).
      return readNameRecord(16) || readNameRecord(4) || readNameRecord(1) || '';
    }
  }
  return '';
}

/**
 * Build an EOT (Embedded OpenType) wrapper around raw TTF data.
 * Follows the EOT v2 spec (version 0x00020002) matching PowerPoint's format.
 * @param {ArrayBuffer} ttfBuffer - Raw TrueType font data
 * @returns {ArrayBuffer} EOT file
 */
function buildEot(ttfBuffer) {
  const ttf = ttfBuffer;
  const ttfView = new DataView(ttf);
  const fontDataSize = ttf.byteLength;

  // Read OS/2 table for panose, weight, fsType, charset, italic, unicode/codepage ranges
  let panose = new Uint8Array(10);
  let weight = 400;
  let fsType = 0;
  let italic = 0;
  let charset = 0;
  let unicodeRange = new Uint8Array(16); // ulUnicodeRange1-4 (4 x uint32)
  let codePageRange = new Uint8Array(8); // ulCodePageRange1-2 (2 x uint32)
  let checkSumAdjustment = 0;

  const numTables = ttfView.getUint16(4);
  for (let i = 0; i < numTables; i++) {
    const off = 12 + i * 16;
    const tag =
      String.fromCharCode(ttfView.getUint8(off)) +
      String.fromCharCode(ttfView.getUint8(off + 1)) +
      String.fromCharCode(ttfView.getUint8(off + 2)) +
      String.fromCharCode(ttfView.getUint8(off + 3));
    if (tag === 'OS/2') {
      const tOff = ttfView.getUint32(off + 8);
      const tLen = ttfView.getUint32(off + 12);
      panose = new Uint8Array(ttf, tOff + 32, 10);
      weight = ttfView.getUint16(tOff + 4); // usWeightClass
      fsType = ttfView.getUint16(tOff + 8);
      // fsSelection at offset 62
      const fsSelection = ttfView.getUint16(tOff + 62);
      italic = fsSelection & 1 ? 1 : 0;
      // ulUnicodeRange1-4 at OS/2 offset 42 (16 bytes) — requires OS/2 version >= 1
      if (tLen >= 58) unicodeRange = new Uint8Array(ttf, tOff + 42, 16);
      // ulCodePageRange1-2 at OS/2 offset 78 (8 bytes) — requires OS/2 version >= 1
      if (tLen >= 86) codePageRange = new Uint8Array(ttf, tOff + 78, 8);
    }
    if (tag === 'head') {
      const tOff = ttfView.getUint32(off + 8);
      checkSumAdjustment = ttfView.getUint32(tOff + 8);
    }
  }

  // Read full name from name table
  const fullName = readTtfFullName(ttf);

  // Encode fullName as UTF-16LE with null terminator
  const fullNameUtf16 = new Uint8Array((fullName.length + 1) * 2);
  for (let i = 0; i < fullName.length; i++) {
    const code = fullName.charCodeAt(i);
    fullNameUtf16[i * 2] = code & 0xff;
    fullNameUtf16[i * 2 + 1] = (code >> 8) & 0xff;
  }
  // last 2 bytes are already 0 (null terminator)
  const fullNameSize = fullNameUtf16.length;

  // EOT header structure (version 0x00020002):
  // Fixed header: 36 bytes
  // UnicodeRange: 16 bytes
  // CodePageRange: 8 bytes
  // CheckSumAdjustment: 4 bytes
  // Reserved: 4 bytes
  // Padding1: 2 bytes
  // FamilyNameSize(2) + FamilyName(0) + Padding2(2)
  // StyleNameSize(2) + StyleName(0) + Padding3(2)
  // VersionNameSize(2) + VersionName(0) + Padding4(2)
  // FullNameSize(2) + FullName(n)
  // Then version 2 fields:
  // Padding5(2) + RootStringSize(2) + RootString(0)

  // Empty family/style/version names, only fullName populated (matching PowerPoint behavior)
  const headerSize =
    36 + // fixed header
    16 + // unicodeRange
    8 + // codePageRange
    4 + // checkSumAdjustment
    4 + // reserved
    2 + // padding1
    2 + 0 + 2 + // family
    2 + 0 + 2 + // style
    2 + 0 + 2 + // version
    2 + fullNameSize + // fullName
    2 + // padding5 (v2)
    2 + 0; // rootStringSize + rootString (v2, empty)

  const eotSize = headerSize + fontDataSize;
  const eot = new ArrayBuffer(eotSize);
  const view = new DataView(eot);
  const bytes = new Uint8Array(eot);
  let pos = 0;

  // Fixed header
  view.setUint32(0, eotSize, true); // EOTSize
  view.setUint32(4, fontDataSize, true); // FontDataSize
  view.setUint32(8, 0x00020002, true); // Version (match PowerPoint)
  view.setUint32(12, 0x00000000, true); // Flags: 0 (full font, not subset)
  pos = 16;
  bytes.set(panose, pos); // FontPANOSE[10]
  pos = 26;
  view.setUint8(pos, charset); // Charset
  view.setUint8(pos + 1, italic); // Italic
  view.setUint32(pos + 2, weight, true); // Weight
  view.setUint16(pos + 6, fsType, true); // fsType
  view.setUint16(pos + 8, 0x504c, true); // MagicNumber "LP"
  pos = 36;

  // UnicodeRange (16 bytes) - from OS/2 ulUnicodeRange1-4 (big-endian in TTF → little-endian in EOT)
  for (let i = 0; i < 4; i++) {
    const beVal = ((unicodeRange[i * 4] << 24) | (unicodeRange[i * 4 + 1] << 16) | (unicodeRange[i * 4 + 2] << 8) | unicodeRange[i * 4 + 3]) >>> 0;
    view.setUint32(pos + i * 4, beVal, true);
  }
  pos += 16;
  // CodePageRange (8 bytes) - from OS/2 ulCodePageRange1-2 (big-endian in TTF → little-endian in EOT)
  for (let i = 0; i < 2; i++) {
    const beVal = ((codePageRange[i * 4] << 24) | (codePageRange[i * 4 + 1] << 16) | (codePageRange[i * 4 + 2] << 8) | codePageRange[i * 4 + 3]) >>> 0;
    view.setUint32(pos + i * 4, beVal, true);
  }
  pos += 8;
  // CheckSumAdjustment (4 bytes) - from head table (already read as big-endian uint32)
  view.setUint32(pos, checkSumAdjustment, true);
  pos += 4;
  // Reserved (4 bytes)
  pos += 4;
  // Padding1 (2 bytes)
  pos += 2;

  // FamilyNameSize (0) + Padding
  view.setUint16(pos, 0, true);
  pos += 2;
  view.setUint16(pos, 0, true);
  pos += 2;

  // StyleNameSize (0) + Padding
  view.setUint16(pos, 0, true);
  pos += 2;
  view.setUint16(pos, 0, true);
  pos += 2;

  // VersionNameSize (0) + Padding
  view.setUint16(pos, 0, true);
  pos += 2;
  view.setUint16(pos, 0, true);
  pos += 2;

  // FullNameSize + FullName
  view.setUint16(pos, fullNameSize, true);
  pos += 2;
  bytes.set(fullNameUtf16, pos);
  pos += fullNameSize;

  // Version 2 fields
  // Padding5
  view.setUint16(pos, 0, true);
  pos += 2;
  // RootStringSize (0 = no URL restriction)
  view.setUint16(pos, 0, true);
  pos += 2;

  // Font data (raw TTF)
  bytes.set(new Uint8Array(ttf), pos);

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
 * Strip variable-font-only tables from a TTF buffer, producing a valid static TTF.
 * Preserves all standard tables (GPOS, GSUB, GDEF, gasp, prep, etc.) intact.
 * Only removes: fvar, gvar, STAT, HVAR, MVAR, avar, CVAR.
 * @param {ArrayBuffer} buf - Raw variable TTF data
 * @returns {ArrayBuffer} Static TTF data
 */
function stripVariableTables(buf) {
  // Guard: only process simple TrueType fonts (not TTC collections or CFF-based OTF)
  const sfVersion = new DataView(buf).getUint32(0);
  if (sfVersion !== 0x00010000) return buf;

  const VARIABLE_TAGS = new Set(['fvar', 'gvar', 'STAT', 'HVAR', 'MVAR', 'avar', 'CVAR']);
  const src = new DataView(buf);
  const srcBytes = new Uint8Array(buf);
  const numTables = src.getUint16(4);

  // Collect tables to keep
  const keep = [];
  for (let i = 0; i < numTables; i++) {
    const off = 12 + i * 16;
    const tag =
      String.fromCharCode(src.getUint8(off)) +
      String.fromCharCode(src.getUint8(off + 1)) +
      String.fromCharCode(src.getUint8(off + 2)) +
      String.fromCharCode(src.getUint8(off + 3));
    if (!VARIABLE_TAGS.has(tag)) {
      const checksum = src.getUint32(off + 4);
      const offset = src.getUint32(off + 8);
      const length = src.getUint32(off + 12);
      keep.push({ tag, checksum, offset, length });
    }
  }

  // Compute new TrueType header values
  const n = keep.length;
  let entrySelector = 0, searchRange = 1;
  while (searchRange * 2 <= n) { searchRange *= 2; entrySelector++; }
  searchRange *= 16;
  const rangeShift = n * 16 - searchRange;

  // Calculate total output size (4-byte aligned per table)
  const headerSize = 12 + n * 16;
  let totalSize = headerSize;
  for (const t of keep) totalSize += (t.length + 3) & ~3;

  const out = new ArrayBuffer(totalSize);
  const outView = new DataView(out);
  const outBytes = new Uint8Array(out);

  // TrueType header
  outView.setUint32(0, 0x00010000); // sfVersion
  outView.setUint16(4, n);
  outView.setUint16(6, searchRange);
  outView.setUint16(8, entrySelector);
  outView.setUint16(10, rangeShift);

  // Write table directory entries and copy table data
  let dataOffset = headerSize;
  for (let i = 0; i < n; i++) {
    const t = keep[i];
    const dirOff = 12 + i * 16;
    for (let j = 0; j < 4; j++) outView.setUint8(dirOff + j, t.tag.charCodeAt(j));
    outView.setUint32(dirOff + 4, t.checksum);
    outView.setUint32(dirOff + 8, dataOffset);
    outView.setUint32(dirOff + 12, t.length);
    outBytes.set(new Uint8Array(buf, t.offset, t.length), dataOffset);
    dataOffset += (t.length + 3) & ~3;
  }

  // Recalculate head.checkSumAdjustment
  for (let i = 0; i < n; i++) {
    if (keep[i].tag === 'head') {
      const headOff = outView.getUint32(12 + i * 16 + 8);
      outView.setUint32(headOff + 8, 0); // zero out first
      // totalSize is always 4-byte aligned, so we can read uint32 directly
      let sum = 0;
      for (let w = 0; w < totalSize / 4; w++) {
        sum = (sum + outView.getUint32(w * 4)) >>> 0;
      }
      outView.setUint32(headOff + 8, (0xB1B0AFBA - sum) >>> 0);
      break;
    }
  }

  return out;
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

  // Strip variable font tables (fvar/gvar/STAT/etc.) from TTF binary.
  // These tables cause PowerPoint to reject the embedded font.
  // We strip at the binary level to preserve all standard tables (GPOS, GSUB, etc.)
  // that fonteditor-core's conversion would discard.
  if (isVariableFont(actualBuffer)) {
    actualBuffer = stripVariableTables(actualBuffer);
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
