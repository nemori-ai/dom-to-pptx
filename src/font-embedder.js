// src/font-embedder.js
import { fontToEot } from './font-utils.js';

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
    const eotData = await fontToEot(type, fontBuffer, opts);
    const rid = this.nextRId++;
    this.fonts.push({ name: fontFace, data: eotData, rid });
  }

  async updateFiles() {
    await this.updateContentTypesXML();
    await this.updatePresentationXML();
    await this.updateRelsPresentationXML();
    this.updateFontFiles();
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
      const el = doc.createElement('Default');
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

    // Enable embedding flags
    presentation.setAttribute('saveSubsetFonts', 'true');
    presentation.setAttribute('embedTrueTypeFonts', 'true');

    // Find or create embeddedFontLst
    let embeddedFontLst = presentation.getElementsByTagName('p:embeddedFontLst')[0];

    if (!embeddedFontLst) {
      embeddedFontLst = doc.createElement('p:embeddedFontLst');

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
        const embedFont = doc.createElement('p:embeddedFont');

        const fontNode = doc.createElement('p:font');
        fontNode.setAttribute('typeface', font.name);
        embedFont.appendChild(fontNode);

        const regular = doc.createElement('p:regular');
        regular.setAttribute('r:id', `rId${font.rid}`);
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

    this.fonts.forEach((font) => {
      const rel = doc.createElement('Relationship');
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
}
