// src/dual-font-processor.js
// Post-processes PPTX XML to split dual-font JSON typeface attributes
// into proper <a:latin> and <a:ea> elements for East Asian font support.

import JSZip from 'jszip';

/**
 * Decode HTML entities and parse JSON font config.
 * @param {string} str - e.g. '{&quot;latin&quot;:&quot;Arial&quot;,&quot;ea&quot;:&quot;Microsoft YaHei&quot;}'
 * @returns {{ latin: string, ea: string } | null}
 */
function parseFontJson(str) {
  try {
    const decoded = str
      .replace(/&quot;/g, '"')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>');
    const config = JSON.parse(decoded);
    if (config.latin && config.ea) return config;
  } catch (e) {
    // Not valid JSON — normal for regular font names
  }
  return null;
}

/**
 * Process a single XML string, replacing dual-font JSON typeface values
 * with proper <a:latin> and <a:ea> elements via regex.
 *
 * Uses 4 sequential patterns to handle the different XML structures
 * PptxGenJS may generate:
 *   Pattern 1:  <a:latin typeface="{json}"/> → replace with latin font
 *   Pattern 1b: <a:ea typeface="{json}"/>    → replace with ea font
 *   Pattern 2:  <a:rPr ... typeface="{json}" .../> (self-closing)
 *   Pattern 3:  <a:rPr ... typeface="{json}" ...>  (non-self-closing)
 *   Pattern 4:  typeface="{json}" anywhere (generic fallback)
 *
 * @param {string} content - Raw XML string
 * @returns {{ xml: string, modified: boolean }}
 */
export function processDualFontsInXml(content) {
  // Quick check: skip files that don't contain dual-font JSON
  if (!content.includes('{"latin"') && !content.includes('{&quot;latin')) {
    return { xml: content, modified: false };
  }

  let modified = false;
  let result = content;

  // Pattern 1: <a:latin typeface="{...}" .../> — replace with latin font only
  result = result.replace(
    /<a:latin\s+typeface="(\{"latin":"[^"]+","ea":"[^"]+"\})"([^/]*)\s*\/>/g,
    (match, jsonStr, extraAttrs) => {
      const config = parseFontJson(jsonStr);
      if (config) {
        modified = true;
        return `<a:latin typeface="${config.latin}"${extraAttrs}/>`;
      }
      return match;
    }
  );

  // Pattern 1b: <a:ea typeface="{...}" .../> — replace with ea font
  result = result.replace(
    /<a:ea\s+typeface="(\{"latin":"[^"]+","ea":"[^"]+"\})"([^/]*)\s*\/>/g,
    (match, jsonStr, extraAttrs) => {
      const config = parseFontJson(jsonStr);
      if (config) {
        modified = true;
        return `<a:ea typeface="${config.ea}"${extraAttrs}/>`;
      }
      return match;
    }
  );

  // Pattern 2: <a:rPr ... typeface="{...}" .../> (self-closing) → expand with children
  result = result.replace(
    /<a:rPr([^>]*)\s+typeface=["'](\{[^"']+\})["']([^>]*)\s*\/>/g,
    (match, before, jsonStr, after) => {
      const config = parseFontJson(jsonStr);
      if (config) {
        modified = true;
        return `<a:rPr${before}${after}><a:latin typeface="${config.latin}"/><a:ea typeface="${config.ea}"/></a:rPr>`;
      }
      return match;
    }
  );

  // Pattern 3: <a:rPr ... typeface="{...}" ...> (non-self-closing) → inject children
  result = result.replace(
    /<a:rPr([^>]*)\s+typeface=["'](\{[^"']+\})["']([^>]*)>/g,
    (match, before, jsonStr, after) => {
      const config = parseFontJson(jsonStr);
      if (config) {
        modified = true;
        return `<a:rPr${before}${after}><a:latin typeface="${config.latin}"/><a:ea typeface="${config.ea}"/>`;
      }
      return match;
    }
  );

  // Pattern 4: typeface="{...}" anywhere (generic fallback) — use latin font
  result = result.replace(
    /typeface=["'](\{"latin":"[^"]+","ea":"[^"]+"\})["']/g,
    (match, jsonStr) => {
      const config = parseFontJson(jsonStr);
      if (config) {
        modified = true;
        return `typeface="${config.latin}"`;
      }
      return match;
    }
  );

  return { xml: result, modified };
}

/**
 * Post-process PPTX blob to handle dual-font (Latin + EA) fontFace JSON strings.
 * Replaces typeface='{"latin":"Arial","ea":"Microsoft YaHei"}' with proper XML:
 *   <a:latin typeface="Arial"/><a:ea typeface="Microsoft YaHei"/>
 *
 * @param {Blob} blob - The PPTX blob to process
 * @param {Object} [options] - Options
 * @param {string[]} [options.embeddedFontNames] - Font names that are embedded (keep their pitchFamily/charset)
 * @returns {Promise<Blob>} - The processed PPTX blob
 */
export async function postProcessDualFonts(blob, options = {}) {
  const zip = await JSZip.loadAsync(blob);
  let modified = false;

  const xmlFiles = Object.keys(zip.files).filter(
    (name) => name.startsWith('ppt/') && name.endsWith('.xml')
  );

  for (const fileName of xmlFiles) {
    let content = await zip.file(fileName).async('string');
    let fileModified = false;
    const result = processDualFontsInXml(content);

    if (result.modified) {
      content = result.xml;
      fileModified = true;
    }

    // Fix PptxGenJS hardcoded algn="bl" on outerShdw/innerShdw.
    // CSS box-shadow is center-aligned; OOXML algn="ctr" matches that model.
    if (content.includes('algn="bl"') && content.includes('Shdw')) {
      content = content.replace(/(<a:(?:outer|inner)Shdw\b[^>]*?)algn="bl"/g, '$1algn="ctr"');
      fileModified = true;
    }

    // Strip pitchFamily and charset from font elements (a:latin, a:ea, a:cs)
    // EXCEPT for embedded fonts which need these attributes for PowerPoint to match them.
    // PptxGenJS hardcodes these values which can cause PowerPoint (especially Mac)
    // to reject or mismap non-embedded fonts. Letting PowerPoint resolve by typeface alone is safer.
    {
      const embeddedNames = new Set(options.embeddedFontNames || []);
      const stripped = content.replace(
        /(<a:(?:latin|ea|cs)\s+typeface="([^"]*)")\s+pitchFamily="[^"]*"\s*charset="[^"]*"/g,
        (match, prefix, fontName) => {
          if (embeddedNames.has(fontName)) return match; // keep for embedded fonts
          return prefix;
        }
      );
      if (stripped !== content) {
        content = stripped;
        fileModified = true;
      }
    }

    if (fileModified) {
      zip.file(fileName, content);
      modified = true;
    }
  }

  if (modified) {
    return await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    });
  }

  return blob;
}
