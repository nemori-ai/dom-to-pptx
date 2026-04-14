// src/font-slot-processor.js
// Post-processes PPTX XML to expand multi-font JSON typeface attributes
// into proper OOXML font slot elements: <a:latin>, <a:ea>, <a:cs>, <a:sym>.

import JSZip from 'jszip';

/**
 * Decode HTML entities and parse JSON font config.
 * Accepts both legacy 2-key format {"latin":"...","ea":"..."} and
 * full 4-key format {"latin":"...","ea":"...","cs":"...","sym":"..."}.
 *
 * @param {string} str - JSON string, possibly with HTML entities
 * @returns {{ latin: string, ea?: string, cs?: string, sym?: string } | null}
 */
function parseFontJson(str) {
  try {
    const decoded = str
      .replace(/&quot;/g, '"')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>');
    const config = JSON.parse(decoded);
    if (config.latin) return config;
  } catch (e) {
    // Not valid JSON — normal for regular font names
  }
  return null;
}

/**
 * Build font child elements string from a parsed font config.
 * Generates <a:latin>, <a:ea>, <a:cs>, and optionally <a:sym>.
 *
 * @param {object} config - Parsed font config from parseFontJson
 * @returns {string} XML elements string
 */
function buildFontElements(config) {
  let xml = `<a:latin typeface="${config.latin}"/>`;
  xml += `<a:ea typeface="${config.ea || config.latin}"/>`;
  xml += `<a:cs typeface="${config.cs || config.latin}"/>`;
  if (config.sym && config.sym !== config.latin) {
    xml += `<a:sym typeface="${config.sym}"/>`;
  }
  return xml;
}

/**
 * Process a single XML string, replacing multi-font JSON typeface values
 * with proper <a:latin>, <a:ea>, <a:cs>, and <a:sym> elements via regex.
 *
 * Uses sequential patterns to handle the different XML structures
 * PptxGenJS may generate:
 *   Pattern 1:  <a:latin typeface="{json}"/>  → replace with latin font
 *   Pattern 1b: <a:ea typeface="{json}"/>     → replace with ea font
 *   Pattern 1c: <a:cs typeface="{json}"/>     → replace with cs font
 *   Pattern 2:  <a:rPr ... typeface="{json}" .../> (self-closing)
 *   Pattern 3:  <a:rPr ... typeface="{json}" ...>  (non-self-closing)
 *   Pattern 4:  typeface="{json}" anywhere (generic fallback)
 *
 * @param {string} content - Raw XML string
 * @returns {{ xml: string, modified: boolean }}
 */
export function expandFontSlotsInXml(content) {
  // Quick check: skip files that don't contain font config JSON
  if (!content.includes('{"') && !content.includes('{&quot;')) {
    return { xml: content, modified: false };
  }

  let modified = false;
  let result = content;

  // Pattern 1: <a:latin typeface="{...}" .../> — replace with latin font only
  result = result.replace(
    /<a:latin\s+typeface="(\{[^}]+\})"([^/]*)\s*\/>/g,
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
    /<a:ea\s+typeface="(\{[^}]+\})"([^/]*)\s*\/>/g,
    (match, jsonStr, extraAttrs) => {
      const config = parseFontJson(jsonStr);
      if (config) {
        modified = true;
        return `<a:ea typeface="${config.ea || config.latin}"${extraAttrs}/>`;
      }
      return match;
    }
  );

  // Pattern 1c: <a:cs typeface="{...}" .../> — replace with cs font
  result = result.replace(
    /<a:cs\s+typeface="(\{[^}]+\})"([^/]*)\s*\/>/g,
    (match, jsonStr, extraAttrs) => {
      const config = parseFontJson(jsonStr);
      if (config) {
        modified = true;
        return `<a:cs typeface="${config.cs || config.latin}"${extraAttrs}/>`;
      }
      return match;
    }
  );

  // Pattern 2: <a:rPr ... typeface="{...}" .../> (self-closing) → expand with children
  result = result.replace(
    /<a:rPr([^>]*)\s+typeface=["'](\{[^}]+\})["']([^>]*)\s*\/>/g,
    (match, before, jsonStr, after) => {
      const config = parseFontJson(jsonStr);
      if (config) {
        modified = true;
        return `<a:rPr${before}${after}>${buildFontElements(config)}</a:rPr>`;
      }
      return match;
    }
  );

  // Pattern 3: <a:rPr ... typeface="{...}" ...> (non-self-closing)
  // PptxGenJS generates font slot children (<a:latin>/<a:ea>/<a:cs>) alongside
  // the typeface attr. Patterns 1/1b/1c already fixed those children, so here
  // we only strip the residual typeface attr from the opening tag to avoid
  // injecting duplicate font slot elements.
  result = result.replace(
    /<a:rPr([^>]*)\s+typeface=["'](\{[^}]+\})["']([^>]*)>/g,
    (match, before, jsonStr, after) => {
      const config = parseFontJson(jsonStr);
      if (config) {
        modified = true;
        return `<a:rPr${before}${after}>`;
      }
      return match;
    }
  );

  // Pattern 4: typeface="{...}" anywhere (generic fallback) — use latin font
  result = result.replace(
    /typeface=["'](\{[^}]+\})["']/g,
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
 * Post-process PPTX blob to handle multi-font fontFace JSON strings.
 * Replaces typeface='{"latin":"Arial","ea":"Microsoft YaHei","cs":"Arial","sym":"Cambria Math"}'
 * with proper XML elements.
 *
 * @param {Blob} blob - The PPTX blob to process
 * @param {Object} [options] - Options
 * @param {string[]} [options.embeddedFontNames] - Font names that are embedded (keep their pitchFamily/charset)
 * @returns {Promise<Blob>} - The processed PPTX blob
 */
export async function postProcessFontSlots(blob, options = {}) {
  const zip = await JSZip.loadAsync(blob);
  let modified = false;

  const xmlFiles = Object.keys(zip.files).filter(
    (name) => name.startsWith('ppt/') && name.endsWith('.xml')
  );

  for (const fileName of xmlFiles) {
    let content = await zip.file(fileName).async('string');
    let fileModified = false;
    const result = expandFontSlotsInXml(content);

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

    // Strip pitchFamily and charset from font elements (a:latin, a:ea, a:cs, a:sym)
    // EXCEPT for embedded fonts which need these attributes for PowerPoint to match them.
    // PptxGenJS hardcodes these values which can cause PowerPoint (especially Mac)
    // to reject or mismap non-embedded fonts. Letting PowerPoint resolve by typeface alone is safer.
    {
      const embeddedNames = new Set(options.embeddedFontNames || []);
      const stripped = content.replace(
        /(<a:(?:latin|ea|cs|sym)\s+typeface="([^"]*)")\s+pitchFamily="[^"]*"\s*charset="[^"]*"/g,
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
