/**
 * Minimal TrueType / OpenType font binary subsetter.
 *
 * Copies the tables needed to render a subset of glyph IDs, rewrites the
 * glyf/loca/hmtx tables, patches the font header, and emits a fully
 * self-contained .ttf stream.  Tables that are not known to TrueType (CFF,
 * CFF2, etc.) are not handled – the caller is expected to feed this helper
 * with a TrueType font binary.
 */

/* Compound glyph flag bits – copied verbatim from the .NET port. */
const ARG1AND2AREWORDS = 0x0001;
const MORECOMPONENTS = 0x0020;
const WEHAVESCALE = 0x0008;
const WEHAVEANXANDYSCALE = 0x0040;
const WEHAVETWOBYTWO = 0x0080;
const WEHAVEINSTRUCTIONS = 0x0100;

interface TableRecord {
  Tag: string;
  Checksum: number;
  Offset: number;
  Length: number;
}

interface HorizontalMetric {
  AdvanceWidth: number;
  LeftSideBearing: number;
}

interface GlyfBuildResult {
  Bytes: Uint8Array;
  LocaOffsets: number[];
}

interface CharMapping {
  Unicode: number;
  GlyphId: number;
}

interface OutputRecord {
  Tag: string;
  Checksum: number;
  Offset: number;
  Length: number;
}

interface NameRecord {
  PlatformId: number;
  EncodingId: number;
  LanguageId: number;
  NameId: number;
  Bytes: Uint8Array;
}

export class TrueTypeSubsetFontProgram {
  private _fontBytes: Uint8Array;
  private _tables: Map<string, TableRecord>;
  private _unitsPerEm: number;
  private _ascent: number;
  private _descent: number;
  private _lineGap: number;
  private _xMin = 0;
  private _yMin = 0;
  private _xMax = 0;
  private _yMax = 0;
  private _numberOfHMetrics: number;
  private _numGlyphs: number;
  private _indexToLocFormat: number;
  private _glyphOffsets: number[];
  private _metrics: HorizontalMetric[];

  constructor(fontBytes: Uint8Array) {
    if (!fontBytes || fontBytes.length < 12) throw new Error("TrueTypeSubsetFontProgram: invalid font bytes");
    this._fontBytes = fontBytes;
    this._tables = parseTableDirectory(fontBytes);
    const head = this.tableBytes("head");
    this._unitsPerEm = readUShort(head, 18);
    this._xMin = readShort(head, 36);
    this._yMin = readShort(head, 38);
    this._xMax = readShort(head, 40);
    this._yMax = readShort(head, 42);
    this._indexToLocFormat = readShort(head, 50);
    const hhea = this.tableBytes("hhea");
    this._ascent = readShort(hhea, 4);
    this._descent = readShort(hhea, 6);
    this._lineGap = readShort(hhea, 8);
    this._numberOfHMetrics = readUShort(hhea, 34);
    const maxp = this.tableBytes("maxp");
    this._numGlyphs = readUShort(maxp, 4);
    this._glyphOffsets = this.parseGlyphOffsets();
    this._metrics = this.parseHorizontalMetrics();
  }

  BuildSubset(
    initialGlyphIds: number[],
    codeToGlyphId: Map<number, number>,
    glyphIdToUnicode: Map<number, number>,
    subsetFontName: string,
  ): Uint8Array {
    if (!initialGlyphIds || initialGlyphIds.length === 0) throw new Error("TrueType subset requires at least one glyph");
    const closure = this.buildGlyphClosure(initialGlyphIds);
    const oldToNew = new Map<number, number>();
    for (let i = 0; i < closure.length; i++) oldToNew.set(closure[i], i);

    const glyfResult = this.buildGlyfTable(closure, oldToNew);
    const loca = this.buildLocaTable(glyfResult.LocaOffsets);
    const hmtx = this.buildHmtxTable(closure);
    const maxp = this.patchMaxpTable(closure.length);
    const hhea = this.patchHheaTable(closure.length);
    const head = this.patchHeadTable();
    const cmap = this.buildCmapTable(codeToGlyphId, glyphIdToUnicode, oldToNew);

    const outputTables = new Map<string, Uint8Array>();
    const gasp = this._tables.get("gasp");
    if (gasp) outputTables.set("gasp", sliceBytes(this._fontBytes, gasp.Offset, gasp.Length));
    outputTables.set("OS/2", this.buildOs2Table());
    outputTables.set("cmap", cmap);
    outputTables.set("glyf", glyfResult.Bytes);
    outputTables.set("head", head);
    outputTables.set("hhea", hhea);
    outputTables.set("hmtx", hmtx);
    outputTables.set("loca", loca);
    outputTables.set("maxp", maxp);
    outputTables.set("name", this.buildNameTable(subsetFontName));
    outputTables.set("post", this.buildPostTable());
    return this.buildFontFile(outputTables);
  }

  // ----- closures / rewriting ----------------------------------------------

  private buildGlyphClosure(initial: number[]): number[] {
    const seen = new Set<number>();
    const queue: number[] = [0, ...initial];
    while (queue.length > 0) {
      const glyphId = queue.shift()!;
      if (glyphId >= this._numGlyphs || seen.has(glyphId)) continue;
      seen.add(glyphId);
      const bytes = this.getGlyphBytes(glyphId);
      if (bytes.length < 10) continue;
      const contourCount = readShort(bytes, 0);
      if (contourCount >= 0) continue; // simple glyph
      let offset = 10;
      while (offset + 4 <= bytes.length) {
        const flags = readUShort(bytes, offset);
        const componentGlyphId = readUShort(bytes, offset + 2);
        queue.push(componentGlyphId);
        offset += 4;
        offset += (flags & ARG1AND2AREWORDS) !== 0 ? 4 : 2;
        if ((flags & WEHAVESCALE) !== 0) offset += 2;
        else if ((flags & WEHAVEANXANDYSCALE) !== 0) offset += 4;
        else if ((flags & WEHAVETWOBYTWO) !== 0) offset += 8;
        if ((flags & MORECOMPONENTS) === 0) break;
      }
    }
    const ordered = [...seen].sort((a, b) => a - b);
    if (ordered.length === 0 || ordered[0] !== 0) ordered.unshift(0);
    return ordered;
  }

  private buildGlyfTable(closure: number[], oldToNew: Map<number, number>): GlyfBuildResult {
    const locaOffsets: number[] = [];
    const out: number[] = [];
    for (const original of closure) {
      locaOffsets.push(out.length);
      const bytes = this.getGlyphBytes(original);
      if (bytes.length > 0) {
        const rewritten = this.rewriteCompoundGlyph(bytes, oldToNew);
        for (let i = 0; i < rewritten.length; i++) out.push(rewritten[i]);
        while (out.length % 4 !== 0) out.push(0);
      }
    }
    locaOffsets.push(out.length);
    return { Bytes: Uint8Array.from(out), LocaOffsets: locaOffsets };
  }

  private rewriteCompoundGlyph(glyphBytes: Uint8Array, oldToNew: Map<number, number>): Uint8Array {
    const contourCount = readShort(glyphBytes, 0);
    if (contourCount >= 0) return stripInstructions(glyphBytes, contourCount);
    const out: number[] = [];
    for (let i = 0; i < 10; i++) out.push(glyphBytes[i]);
    let offset = 10;
    while (offset + 4 <= glyphBytes.length) {
      const flags = readUShort(glyphBytes, offset);
      const cleanFlags = flags & ~WEHAVEINSTRUCTIONS;
      const oldGlyphId = readUShort(glyphBytes, offset + 2);
      const newGlyphId = oldToNew.get(oldGlyphId);
      if (newGlyphId === undefined) throw new Error(`Compound glyph references unmapped component ${oldGlyphId}`);
      out.push((cleanFlags >> 8) & 0xff, cleanFlags & 0xff);
      out.push((newGlyphId >> 8) & 0xff, newGlyphId & 0xff);
      offset += 4;
      const argumentBytes = (flags & ARG1AND2AREWORDS) !== 0 ? 4 : 2;
      for (let i = 0; i < argumentBytes; i++) out.push(glyphBytes[offset + i]);
      offset += argumentBytes;
      if ((flags & WEHAVESCALE) !== 0) {
        out.push(glyphBytes[offset], glyphBytes[offset + 1]);
        offset += 2;
      } else if ((flags & WEHAVEANXANDYSCALE) !== 0) {
        for (let i = 0; i < 4; i++) out.push(glyphBytes[offset + i]);
        offset += 4;
      } else if ((flags & WEHAVETWOBYTWO) !== 0) {
        for (let i = 0; i < 8; i++) out.push(glyphBytes[offset + i]);
        offset += 8;
      }
      if ((flags & MORECOMPONENTS) === 0) break;
    }
    return Uint8Array.from(out);
  }

  private buildLocaTable(offsets: number[]): Uint8Array {
    const out = new Uint8Array(offsets.length * 4);
    for (let i = 0; i < offsets.length; i++) writeUInt(out, i * 4, offsets[i]);
    return out;
  }

  private buildHmtxTable(closure: number[]): Uint8Array {
    const out = new Uint8Array(closure.length * 4);
    for (let i = 0; i < closure.length; i++) {
      const m = this.metricForGlyph(closure[i]);
      writeUShort(out, i * 4, m.AdvanceWidth);
      writeShort(out, i * 4 + 2, m.LeftSideBearing);
    }
    return out;
  }

  private patchMaxpTable(numGlyphs: number): Uint8Array {
    const bytes = this.tableBytes("maxp").slice();
    writeUShort(bytes, 4, numGlyphs);
    return bytes;
  }

  private patchHheaTable(numHMetrics: number): Uint8Array {
    const bytes = this.tableBytes("hhea").slice();
    writeUShort(bytes, 34, numHMetrics);
    return bytes;
  }

  private patchHeadTable(): Uint8Array {
    const bytes = this.tableBytes("head").slice();
    writeUInt(bytes, 8, 0); // checksumAdjustment placeholder
    writeShort(bytes, 50, 1); // indexToLocFormat = long
    return bytes;
  }

  private buildOs2Table(): Uint8Array {
    const original = this.tableBytes("OS/2");
    const copy = original.slice();
    if (copy.length >= 62) writeUShort(copy, 62, 0); // usDefaultChar
    return copy;
  }

  private buildPostTable(): Uint8Array {
    const original = this.tableBytes("post");
    const out = new Uint8Array(32);
    writeUInt(out, 0, 0x00030000); // version 3.0 (no glyph names)
    // Copy timestamps and metrics from the original.
    for (let i = 0; i < 28; i++) out[4 + i] = original[4 + i];
    return out;
  }

  private buildCmapTable(codeToGlyphId: Map<number, number>, glyphIdToUnicode: Map<number, number>, oldToNew: Map<number, number>): Uint8Array {
    const mappings: CharMapping[] = [];
    const seen = new Set<number>();
    for (const [code, glyphId] of codeToGlyphId) {
      const unicode = glyphIdToUnicode.get(glyphId) ?? code;
      const newGlyph = oldToNew.get(glyphId);
      if (newGlyph !== undefined && seen.add(unicode)) mappings.push({ Unicode: unicode, GlyphId: newGlyph });
    }
    mappings.sort((a, b) => a.Unicode - b.Unicode);
    const format4 = this.buildFormat4Cmap(mappings);
    const out = new Uint8Array(12 + format4.length);
    writeUShort(out, 0, 0); // version
    writeUShort(out, 2, 1); // numSubtables = 1
    writeUInt(out, 4, 12); // offset to first (only) subtable
    out.set(format4, 12);
    return out;
  }

  private buildFormat4Cmap(mappings: CharMapping[]): Uint8Array {
    const segmentCount = mappings.length + 1;
    const segCountX2 = segmentCount * 2;
    let searchRange = 2;
    let entrySelector = 0;
    while (searchRange * 2 <= segCountX2) {
      searchRange *= 2;
      entrySelector++;
    }
    const rangeShift = segCountX2 - searchRange;
    const out: number[] = [];
    putUShort(out, 4); // format
    putUShort(out, 0); // reserved
    const lengthOffset = out.length;
    putUShort(out, 0); // length placeholder
    putUShort(out, 0); // language
    putUShort(out, segCountX2);
    putUShort(out, searchRange);
    putUShort(out, entrySelector);
    putUShort(out, rangeShift);
    // endCodes
    for (const m of mappings) putUShort(out, m.Unicode);
    putUShort(out, 0xffff);
    // reserved pad
    putUShort(out, 0);
    // startCodes (same as Unicodes in our simple case)
    for (const m of mappings) putUShort(out, m.Unicode);
    putUShort(out, 0xffff);
    // idDelta
    for (const m of mappings) {
      let delta = m.GlyphId - m.Unicode;
      if (delta < 0) delta += 65536;
      putUShort(out, delta & 0xffff);
    }
    putUShort(out, 1); // idRangeOffset sentinel (no per-segment lookup)
    for (let i = 0; i < segmentCount; i++) putUShort(out, 0);
    writeUShort(out as unknown as Uint8Array, lengthOffset, out.length);
    return Uint8Array.from(out);
  }

  private buildNameTable(subsetFontName: string): Uint8Array {
    const records: NameRecord[] = [
      makeName(3, 1, 1033, 1, "Calibri"),
      makeName(3, 1, 1033, 2, "Regular"),
      makeName(3, 1, 1033, 4, subsetFontName),
      makeName(3, 1, 1033, 6, subsetFontName),
    ];
    let offset = 6 + records.length * 12;
    const out: number[] = [];
    putUShort(out, 0);
    putUShort(out, records.length);
    putUShort(out, 6 + records.length * 12);
    for (const r of records) {
      putUShort(out, r.PlatformId);
      putUShort(out, r.EncodingId);
      putUShort(out, r.LanguageId);
      putUShort(out, r.NameId);
      putUShort(out, r.Bytes.length);
      putUShort(out, offset);
      offset += r.Bytes.length;
    }
    for (const r of records) for (let i = 0; i < r.Bytes.length; i++) out.push(r.Bytes[i]);
    return Uint8Array.from(out);
  }

  // ----- font file assembly -------------------------------------------------

  private buildFontFile(tables: Map<string, Uint8Array>): Uint8Array {
    const tags = [...tables.keys()].sort();
    const numTables = tags.length;
    let searchRange = 1;
    let entrySelector = 0;
    while (searchRange * 2 <= numTables) {
      searchRange *= 2;
      entrySelector++;
    }
    searchRange = searchRange * 16;
    const rangeShift = numTables * 16 - searchRange;

    const headerSize = 12;
    const directorySize = numTables * 16;
    let out: number[] = [...new Uint8Array(headerSize + directorySize)];
    putUInt(out, 0x00010000);
    putUShort(out, numTables);
    putUShort(out, searchRange);
    putUShort(out, entrySelector);
    putUShort(out, rangeShift);
    for (let i = 0; i < numTables; i++) {
      for (let j = 0; j < 16; j++) out.push(0);
    }

    const records: OutputRecord[] = [];
    for (const tag of tags) {
      while (out.length % 4 !== 0) out.push(0);
      const bytes = tables.get(tag)!;
      const recOffset = out.length;
      for (let i = 0; i < bytes.length; i++) out.push(bytes[i]);
      while (out.length % 4 !== 0) out.push(0);
      records.push({ Tag: tag, Checksum: calculateChecksum(bytes), Offset: recOffset, Length: bytes.length });
    }

    // Write table directory
    for (let i = 0; i < records.length; i++) {
      const rec = records[i];
      const dirOffset = headerSize + i * 16;
      writeTag(out as unknown as Uint8Array, dirOffset, rec.Tag);
      writeUInt(out as unknown as Uint8Array, dirOffset + 4, rec.Checksum);
      writeUInt(out as unknown as Uint8Array, dirOffset + 8, rec.Offset);
      writeUInt(out as unknown as Uint8Array, dirOffset + 12, rec.Length);
    }

    // Fix head checksumAdjustment
    const font = Uint8Array.from(out);
    const head = records.find((r) => r.Tag === "head")!;
    writeUInt(font, head.Offset + 8, 0);
    const checksum = calculateChecksum(font);
    writeUInt(font, head.Offset + 8, (0xb1b0afba - checksum) >>> 0);
    return font;
  }

  // ----- helpers ------------------------------------------------------------

  private parseGlyphOffsets(): number[] {
    const loca = this.tableBytes("loca");
    const offsets = new Array<number>(this._numGlyphs + 1);
    if (this._indexToLocFormat === 0) {
      for (let i = 0; i <= this._numGlyphs; i++) offsets[i] = readUShort(loca, i * 2) * 2;
    } else {
      for (let i = 0; i <= this._numGlyphs; i++) offsets[i] = readUInt(loca, i * 4);
    }
    return offsets;
  }

  private parseHorizontalMetrics(): HorizontalMetric[] {
    const bytes = this.tableBytes("hmtx");
    const metrics: HorizontalMetric[] = [];
    let lastAdvance = 0;
    let lastLsb = 0;
    for (let i = 0; i < this._numGlyphs; i++) {
      if (i < this._numberOfHMetrics) {
        const offset = i * 4;
        lastAdvance = readUShort(bytes, offset);
        lastLsb = readShort(bytes, offset + 2);
      } else {
        const offset = this._numberOfHMetrics * 4 + (i - this._numberOfHMetrics) * 2;
        lastLsb = readShort(bytes, offset);
      }
      metrics.push({ AdvanceWidth: lastAdvance, LeftSideBearing: lastLsb });
    }
    return metrics;
  }

  private metricForGlyph(glyphId: number): HorizontalMetric {
    if (glyphId >= this._metrics.length) return this._metrics[this._metrics.length - 1];
    return this._metrics[glyphId];
  }

  private getGlyphBytes(glyphId: number): Uint8Array {
    if (glyphId >= this._numGlyphs) return new Uint8Array(0);
    const start = this._glyphOffsets[glyphId];
    const end = this._glyphOffsets[glyphId + 1];
    if (end <= start) return new Uint8Array(0);
    const glyf = this.tableBytes("glyf");
    return sliceBytes(glyf, start, end - start);
  }

  private tableBytes(tag: string): Uint8Array {
    const rec = this._tables.get(tag);
    if (!rec) throw new Error(`TrueType subset missing required table ${tag}`);
    return sliceBytes(this._fontBytes, rec.Offset, rec.Length);
  }
}

/* ------------------------------- free functions ------------------------------ */

function parseTableDirectory(fontBytes: Uint8Array): Map<string, TableRecord> {
  const numTables = readUShort(fontBytes, 4);
  const tables = new Map<string, TableRecord>();
  for (let i = 0; i < numTables; i++) {
    const offset = 12 + i * 16;
    let tag = "";
    for (let k = 0; k < 4; k++) tag += String.fromCharCode(fontBytes[offset + k]);
    tables.set(tag, {
      Tag: tag,
      Checksum: readUInt(fontBytes, offset + 4),
      Offset: readUInt(fontBytes, offset + 8),
      Length: readUInt(fontBytes, offset + 12),
    });
  }
  return tables;
}

function sliceBytes(src: Uint8Array, offset: number, length: number): Uint8Array {
  return src.slice(offset, offset + length);
}

function makeName(
  platformId: number,
  encodingId: number,
  languageId: number,
  nameId: number,
  text: string,
): NameRecord {
  const bytes: number[] = [];
  for (let i = 0; i < text.length; i++) {
    const cp = text.charCodeAt(i);
    bytes.push((cp >> 8) & 0xff);
    bytes.push(cp & 0xff);
  }
  return { PlatformId: platformId, EncodingId: encodingId, LanguageId: languageId, NameId: nameId, Bytes: Uint8Array.from(bytes) };
}

function calculateChecksum(bytes: Uint8Array): number {
  let sum = 0;
  const paddedLength = (bytes.length + 3) & ~3;
  for (let i = 0; i < paddedLength; i += 4) {
    let value = 0;
    if (i < bytes.length) value |= bytes[i] << 24;
    if (i + 1 < bytes.length) value |= bytes[i + 1] << 16;
    if (i + 2 < bytes.length) value |= bytes[i + 2] << 8;
    if (i + 3 < bytes.length) value |= bytes[i + 3];
    sum += value;
  }
  return sum >>> 0;
}

function stripInstructions(glyphBytes: Uint8Array, contourCount: number): Uint8Array {
  if (contourCount < 0) return glyphBytes;
  const contourArrayBytes = contourCount * 2;
  const instructionLengthOffset = 10 + contourArrayBytes;
  if (instructionLengthOffset + 2 > glyphBytes.length) return glyphBytes;
  const instructionLength = readUShort(glyphBytes, instructionLengthOffset);
  const flagsOffset = instructionLengthOffset + 2 + instructionLength;
  if (flagsOffset > glyphBytes.length) return glyphBytes;
  const out = new Uint8Array(glyphBytes.length - instructionLength - 2);
  out.set(glyphBytes.subarray(0, instructionLengthOffset), 0);
  writeUShort(out, instructionLengthOffset, 0);
  out.set(glyphBytes.subarray(flagsOffset), instructionLengthOffset + 2);
  return out;
}

function readUShort(bytes: Uint8Array, offset: number): number {
  return (bytes[offset] << 8) | bytes[offset + 1];
}
function readShort(bytes: Uint8Array, offset: number): number {
  const v = readUShort(bytes, offset);
  return v >= 0x8000 ? v - 0x10000 : v;
}
function readUInt(bytes: Uint8Array, offset: number): number {
  return ((bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3]) >>> 0;
}

function writeUShort(bytes: Uint8Array, offset: number, value: number): void {
  bytes[offset] = (value >>> 8) & 0xff;
  bytes[offset + 1] = value & 0xff;
}
function writeShort(bytes: Uint8Array, offset: number, value: number): void {
  writeUShort(bytes, offset, value & 0xffff);
}
function writeUInt(bytes: Uint8Array, offset: number, value: number): void {
  bytes[offset] = (value >>> 24) & 0xff;
  bytes[offset + 1] = (value >>> 16) & 0xff;
  bytes[offset + 2] = (value >>> 8) & 0xff;
  bytes[offset + 3] = value & 0xff;
}
function writeTag(bytes: Uint8Array, offset: number, tag: string): void {
  bytes[offset] = tag.charCodeAt(0);
  bytes[offset + 1] = tag.charCodeAt(1);
  bytes[offset + 2] = tag.charCodeAt(2);
  bytes[offset + 3] = tag.charCodeAt(3);
}

function putUShort(arr: number[], value: number): void {
  arr.push((value >>> 8) & 0xff);
  arr.push(value & 0xff);
}
function putUInt(arr: number[], value: number): void {
  arr.push((value >>> 24) & 0xff, (value >>> 16) & 0xff, (value >>> 8) & 0xff, value & 0xff);
}
