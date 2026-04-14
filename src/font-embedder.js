// src/font-embedder.js
import { fontToEot } from './font-utils.js';
import { expandFontSlotsInXml } from './font-slot-processor.js';

/**
 * Extract font metadata from an EOT buffer for OOXML font embedding.
 * EOT header has panose at offset 16 (10 bytes), weight at 28, etc.
 * Returns { panose, pitchFamily, charset } or null if not found.
 */
function extractFontMeta(eotBuffer) {
  try {
    const data = new DataView(eotBuffer);
    // Verify EOT magic "LP" at offset 34
    if (data.getUint16(34, true) !== 0x504c) return null;

    // panose: 10 bytes at EOT offset 16
    const panoseBytes = new Uint8Array(eotBuffer, 16, 10);
    const panose = Array.from(panoseBytes)
      .map((b) => b.toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase();

    // Detect fixed-pitch from panose byte 4 (proportion: 9 = monospaced)
    const isFixedPitch = panoseBytes[3] === 9;
    const pitch = isFixedPitch ? 1 : 2;
    // Use FF_MODERN (3) for monospaced fonts, FF_SWISS (2) for proportional
    const family = isFixedPitch ? 3 : 2;
    const pitchFamily = (family << 4) | pitch;

    return { panose, pitchFamily: String(pitchFamily), charset: '0' };
  } catch {
    // Ignore parse errors
  }
  return null;
}

export class PPTXEmbedFonts {
  constructor() {
    this.zip = null;
    this.nextRId = 1;
    this.fonts = [];
  }

  async loadZip(zip) {
    this.zip = zip;
    await this._scanExistingRIds();
  }

  async _scanExistingRIds() {
    const relsFile = this.zip.file('ppt/_rels/presentation.xml.rels');
    if (!relsFile) return;

    const xmlStr = await relsFile.async('string');
    const rIdPattern = /Id="rId(\d+)"/g;
    let match;
    let maxId = 0;

    while ((match = rIdPattern.exec(xmlStr)) !== null) {
      const id = parseInt(match[1], 10);
      if (id > maxId) maxId = id;
    }

    this.nextRId = maxId + 1;
  }

  async addFont(fontFace, fontBuffer, type, opts = {}) {
    // Deduplicate: only embed one file per font family name
    if (this.fonts.some((f) => f.name === fontFace)) return;

    const eotData = await fontToEot(type, fontBuffer, opts);
    const rid = this.nextRId++;
    const meta = extractFontMeta(eotData);
    this.fonts.push({ name: fontFace, data: eotData, rid, meta });
  }

  async updateFiles() {
    // 1. Expand font slot JSON first so we can see the real typeface names in slides
    await this.expandFontSlotJson();

    // 2. Prune fonts that aren't actually referenced in the expanded XML.
    //    detectFontsForPPTX may normalize font names (e.g. "PingFang SC" → "Microsoft YaHei"),
    //    so the DOM-scanned name may not appear in the final XML. Embedding unreferenced fonts
    //    wastes bytes and causes panose/pitchFamily mismatches.
    await this._pruneUnreferencedFonts();

    // 3. Write the (pruned) font data into the PPTX structure
    await this.updateContentTypesXML();
    await this.updatePresentationXML();
    await this.updateRelsPresentationXML();
    this.updateFontFiles();

    // 4. Inject panose/pitchFamily/charset into slide font refs
    await this.updateSlidesFontRefs();
  }

  /**
   * Remove fonts from this.fonts that aren't referenced by any typeface attribute
   * in the expanded slide XML. This handles the case where detectFontsForPPTX
   * normalized a font name (e.g. "Noto Sans CJK SC" → "Microsoft YaHei") —
   * the original name no longer appears in the XML, so embedding it is wasted.
   */
  async _pruneUnreferencedFonts() {
    if (this.fonts.length === 0) return;

    // Collect all typeface values from expanded slide XML
    const usedTypefaces = new Set();
    const slideFiles = Object.keys(this.zip.files).filter(
      (f) => f.startsWith('ppt/slides/slide') && f.endsWith('.xml')
    );
    for (const path of slideFiles) {
      const file = this.zip.file(path);
      if (!file) continue;
      const xml = await file.async('string');
      // Match typeface="..." on font slot elements (a:latin, a:ea, a:cs, a:sym)
      const re = /<a:(?:latin|ea|cs|sym)\s+typeface="([^"]+)"/g;
      let m;
      while ((m = re.exec(xml)) !== null) {
        usedTypefaces.add(m[1]);
      }
    }

    const before = this.fonts.length;
    this.fonts = this.fonts.filter((f) => usedTypefaces.has(f.name));
    const pruned = before - this.fonts.length;
    if (pruned > 0) {
      console.log(`Pruned ${pruned} unreferenced embedded font(s)`);
    }
  }

  /**
   * Pre-process slide XML to expand font slot JSON typeface values
   * (e.g. {"latin":"Arial","ea":"Microsoft YaHei","cs":"Arial","sym":"Cambria Math"})
   * into proper <a:latin>, <a:ea>, <a:cs>, <a:sym> elements.
   * Must run before updateSlidesFontRefs() so that font name matching works
   * on real names, not JSON strings.
   */
  async expandFontSlotJson() {
    const slideFiles = Object.keys(this.zip.files).filter(
      (f) => f.startsWith('ppt/slides/slide') && f.endsWith('.xml')
    );
    for (const path of slideFiles) {
      const file = this.zip.file(path);
      if (!file) continue;
      const xmlStr = await file.async('string');
      const result = expandFontSlotsInXml(xmlStr);
      if (result.modified) {
        this.zip.file(path, result.xml);
      }
    }
  }

  /**
   * Get the names of fonts that were actually embedded (after pruning).
   * Use this for the postProcessFontSlots embeddedFontNames option.
   */
  getEmbeddedFontNames() {
    return this.fonts.map((f) => f.name);
  }

  async generateBlob() {
    if (!this.zip) throw new Error('Zip not loaded');
    return this.zip.generateAsync({
      type: 'blob',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });
  }

  // --- XML Manipulation Methods ---

  async updateContentTypesXML() {
    const file = this.zip.file('[Content_Types].xml');
    if (!file) throw new Error('[Content_Types].xml not found');

    const xmlStr = await file.async('string');
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlStr, 'text/xml');

    const types = doc.getElementsByTagName('Types')[0];
    const defaults = Array.from(doc.getElementsByTagName('Default'));

    const hasFntData = defaults.some((el) => el.getAttribute('Extension') === 'fntdata');

    if (!hasFntData) {
      const el = doc.createElementNS(types.namespaceURI, 'Default');
      el.setAttribute('Extension', 'fntdata');
      el.setAttribute('ContentType', 'application/x-fontdata');
      types.insertBefore(el, types.firstChild);
    }

    this.zip.file('[Content_Types].xml', new XMLSerializer().serializeToString(doc));
  }

  async updatePresentationXML() {
    const file = this.zip.file('ppt/presentation.xml');
    if (!file) throw new Error('ppt/presentation.xml not found');

    const xmlStr = await file.async('string');
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlStr, 'text/xml');
    const presentation = doc.getElementsByTagName('p:presentation')[0];

    const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
    const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

    // Enable embedding flags
    presentation.setAttribute('saveSubsetFonts', 'true');
    presentation.setAttribute('embedTrueTypeFonts', 'true');

    // Find or create embeddedFontLst
    let embeddedFontLst = presentation.getElementsByTagName('p:embeddedFontLst')[0];

    if (!embeddedFontLst) {
      embeddedFontLst = doc.createElementNS(NS_P, 'p:embeddedFontLst');

      // Insert before defaultTextStyle or at end
      const defaultTextStyle = presentation.getElementsByTagName('p:defaultTextStyle')[0];
      if (defaultTextStyle) {
        presentation.insertBefore(embeddedFontLst, defaultTextStyle);
      } else {
        presentation.appendChild(embeddedFontLst);
      }
    }

    // Add font references
    this.fonts.forEach((font) => {
      // Check if already exists
      const existing = Array.from(embeddedFontLst.getElementsByTagName('p:font')).find(
        (node) => node.getAttribute('typeface') === font.name
      );

      if (!existing) {
        const embedFont = doc.createElementNS(NS_P, 'p:embeddedFont');

        const fontNode = doc.createElementNS(NS_P, 'p:font');
        fontNode.setAttribute('typeface', font.name);
        if (font.meta) {
          fontNode.setAttribute('panose', font.meta.panose);
          fontNode.setAttribute('pitchFamily', font.meta.pitchFamily);
          fontNode.setAttribute('charset', font.meta.charset);
        }
        embedFont.appendChild(fontNode);

        const regular = doc.createElementNS(NS_P, 'p:regular');
        regular.setAttributeNS(NS_R, 'r:id', `rId${font.rid}`);
        embedFont.appendChild(regular);

        embeddedFontLst.appendChild(embedFont);
      }
    });

    this.zip.file('ppt/presentation.xml', new XMLSerializer().serializeToString(doc));
  }

  async updateRelsPresentationXML() {
    const file = this.zip.file('ppt/_rels/presentation.xml.rels');
    if (!file) throw new Error('presentation.xml.rels not found');

    const xmlStr = await file.async('string');
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlStr, 'text/xml');
    const relationships = doc.getElementsByTagName('Relationships')[0];
    const NS_RELS = relationships.namespaceURI;

    this.fonts.forEach((font) => {
      const rel = doc.createElementNS(NS_RELS, 'Relationship');
      rel.setAttribute('Id', `rId${font.rid}`);
      rel.setAttribute('Target', `fonts/${font.rid}.fntdata`);
      rel.setAttribute(
        'Type',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font'
      );
      relationships.appendChild(rel);
    });

    this.zip.file('ppt/_rels/presentation.xml.rels', new XMLSerializer().serializeToString(doc));
  }

  updateFontFiles() {
    this.fonts.forEach((font) => {
      this.zip.file(`ppt/fonts/${font.rid}.fntdata`, font.data);
    });
  }

  /**
   * Post-process slide XML to add panose/pitchFamily/charset to font refs
   * (a:latin, a:ea, a:cs) that match embedded font names.
   * PowerPoint requires these attributes on text run font refs to use embedded fonts.
   *
   * Note: postProcessFontSlots() runs after generateBlob() and strips pitchFamily/charset
   * from non-embedded fonts. Embedded fonts survive because: (1) embeddedFontNames opt-out
   * prevents stripping, and (2) the panose attribute we add here makes the stripping regex
   * not match (it expects pitchFamily immediately after typeface, but panose is in between).
   */
  async updateSlidesFontRefs() {
    // Build lookup: fontName -> meta
    const metaMap = {};
    for (const font of this.fonts) {
      if (font.meta) metaMap[font.name] = font.meta;
    }
    if (Object.keys(metaMap).length === 0) return;

    const slideFiles = Object.keys(this.zip.files).filter(
      (f) => f.startsWith('ppt/slides/slide') && f.endsWith('.xml')
    );

    for (const path of slideFiles) {
      const file = this.zip.file(path);
      if (!file) continue;

      let xmlStr = await file.async('string');
      let modified = false;

      for (const [fontName, meta] of Object.entries(metaMap)) {
        // Match <a:latin|ea|cs|sym typeface="FontName" .../> that don't already have
        // panose attribute. Allow other attrs between typeface and />
        const escaped = fontName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const re = new RegExp(
          `(<a:(?:latin|ea|cs|sym)\\s+typeface="${escaped}")([^>]*?)(/?>)`,
          'g'
        );
        xmlStr = xmlStr.replace(re, (match, prefix, middle, suffix) => {
          if (match.includes('panose=')) return match;
          modified = true;
          // Replace any existing pitchFamily/charset with correct values, then add panose
          let attrs = middle.replace(/\s+pitchFamily="[^"]*"/g, '').replace(/\s+charset="[^"]*"/g, '');
          return `${prefix} panose="${meta.panose}" pitchFamily="${meta.pitchFamily}" charset="${meta.charset}"${attrs}${suffix}`;
        });
      }

      if (modified) {
        this.zip.file(path, xmlStr);
      }
    }
  }
}
