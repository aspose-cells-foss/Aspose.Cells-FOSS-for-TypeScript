/**
 * Parse and rebuild a minimal PDF, preserving existing content while letting
 * callers insert new objects (Type3 fonts, subset TrueType fonts, replacement
 * images, etc.).  Only the parts we need are modelled – the full PDF grammar
 * is far larger.
 */
export class PdfObjectTable {
  Objects = new Map<number, Uint8Array>();
  PageObjectIds: number[] = [];
  HeaderBytes: Uint8Array = new Uint8Array(0);
  RootReference = "";
  InfoReference = "";

  get MaxObjectId(): number {
    let max = 0;
    for (const k of this.Objects.keys()) if (k > max) max = k;
    return max;
  }

  static Parse(pdf: Uint8Array): PdfObjectTable {
    const text = latin1(pdf);
    const startXrefIndex = text.lastIndexOf("startxref");
    if (startXrefIndex < 0) throw new Error("PDF startxref not found");
    const xrefOffset = parseStartXref(text, startXrefIndex);
    const trailerIndex = text.indexOf("trailer", xrefOffset);
    if (trailerIndex < 0) throw new Error("PDF trailer not found");

    const table = new PdfObjectTable();
    const offsets = parseXrefOffsets(text, xrefOffset, trailerIndex);
    const sorted = [...offsets.entries()]
      .filter(([k]) => k > 0)
      .sort((a, b) => a[1] - b[1]);

    const firstOffset = sorted.length > 0 ? sorted[0][1] : 0;
    table.HeaderBytes = pdf.slice(0, firstOffset);

    for (let i = 0; i < sorted.length; i++) {
      const [objectId, start] = sorted[i];
      const end = i + 1 < sorted.length ? sorted[i + 1][1] : xrefOffset;
      const length = end - start;
      if (length <= 0) continue;
      const bytes = pdf.slice(start, start + length);
      table.Objects.set(objectId, bytes);
      const objectText = latin1(bytes);
      if (objectText.includes("/Type /Page") && !objectText.includes("/Type /Pages")) {
        table.PageObjectIds.push(objectId);
      }
    }
    const trailer = text.substring(trailerIndex, startXrefIndex);
    table.RootReference = parseReference(trailer, "/Root");
    table.InfoReference = parseReference(trailer, "/Info");
    return table;
  }

  /** Serialise back to a complete PDF byte array. */
  Rebuild(): Uint8Array {
    const ids = [...this.Objects.keys()].sort((a, b) => a - b);
    const offsetMap = new Map<number, number>();
    // Pass 1: compute offsets.
    let cursor = this.HeaderBytes.length;
    for (const id of ids) {
      offsetMap.set(id, cursor);
      cursor += this.Objects.get(id)!.length;
      if (this.Objects.get(id)!.length === 0 || this.Objects.get(id)![this.Objects.get(id)!.length - 1] !== 0x0a) cursor += 1;
    }
    const xrefStart = cursor;
    const maxId = ids.length > 0 ? ids[ids.length - 1] : 0;
    const xref: string[] = ["xref", `0 ${maxId + 1}`, "0000000000 65535 f "];
    for (let id = 1; id <= maxId; id++) {
      const offset = offsetMap.get(id);
      if (offset === undefined) {
        xref.push("0000000000 00000 f ");
      } else {
        xref.push(`${offset.toString().padStart(10, "0")} 00000 n `);
      }
    }
    xref.push("trailer");
    xref.push(`<< /Size ${maxId + 1}`);
    if (this.RootReference) xref.push(`/Root ${this.RootReference}`);
    if (this.InfoReference) xref.push(`/Info ${this.InfoReference}`);
    xref.push(">>");
    xref.push("startxref");
    xref.push(xrefStart.toString());
    xref.push("%%EOF");
    const trailerBytes = ascii(xref.join("\n"));
    const out = new Uint8Array(cursor + trailerBytes.length);
    out.set(this.HeaderBytes, 0);
    let write = this.HeaderBytes.length;
    for (const id of ids) {
      const bytes = this.Objects.get(id)!;
      out.set(bytes, write);
      write += bytes.length;
      if (bytes.length === 0 || bytes[bytes.length - 1] !== 0x0a) {
        out[write++] = 0x0a;
      }
    }
    out.set(trailerBytes, cursor);
    return out;
  }
}

function parseXrefOffsets(text: string, xrefOffset: number, trailerIndex: number): Map<number, number> {
  const offsets = new Map<number, number>();
  const body = text.substring(xrefOffset, trailerIndex);
  const lines = body.split(/\r?\n/);
  let first = lines.shift();
  if (!first || first.trim() !== "xref") throw new Error("Unsupported PDF xref table");
  for (const line of lines) {
    const trimmed = line.trim();
    if (trimmed.length === 0) continue;
    const parts = trimmed.split(/\s+/);
    if (parts.length !== 2) break;
    const start = parseInt(parts[0], 10);
    const count = parseInt(parts[1], 10);
    for (let i = 0; i < count; i++) {
      const entryLine = lines.shift();
      if (!entryLine || entryLine.length < 17) continue;
      if (entryLine[17] === "n") {
        offsets.set(start + i, parseInt(entryLine.substring(0, 10), 10));
      }
    }
  }
  return offsets;
}

function parseStartXref(text: string, startXrefIndex: number): number {
  let cursor = startXrefIndex + "startxref".length;
  while (cursor < text.length && /\s/.test(text[cursor])) cursor++;
  let end = cursor;
  while (end < text.length && /\d/.test(text[end])) end++;
  return parseInt(text.substring(cursor, end), 10);
}

function parseReference(trailer: string, key: string): string {
  const index = trailer.indexOf(key);
  if (index < 0) return "";
  let i = index + key.length;
  while (i < trailer.length && /\s/.test(trailer[i])) i++;
  let j = i;
  while (j < trailer.length && trailer[j] !== "\n" && trailer[j] !== "\r" && trailer[j] !== ">") j++;
  return trailer.substring(i, j).trim();
}

function latin1(bytes: Uint8Array): string {
  let s = "";
  for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
  return s;
}

function ascii(text: string): Uint8Array {
  const out = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) out[i] = text.charCodeAt(i) & 0xff;
  return out;
}
