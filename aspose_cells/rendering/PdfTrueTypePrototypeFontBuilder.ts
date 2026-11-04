import { PdfTrueTypePrototypeUsage } from "./PdfTrueTypePrototypeUsage";

/**
 * Emits the four PDF objects required for a /TrueType font resource:
 * font file stream, /FontDescriptor, /ToUnicode CMap, and the /Font dictionary.
 */
export class PdfTrueTypePrototypeFontBuilder {
  BuildObjects(
    usages: PdfTrueTypePrototypeUsage[],
    nextObjectId: number,
  ): { objects: Map<number, Uint8Array>; nextObjectId: number } {
    const objects = new Map<number, Uint8Array>();
    let id = nextObjectId;
    for (let i = 0; i < usages.length; i++) {
      const usage = usages[i];
      usage.ResourceName = `TTF${i + 1}`;
      usage.SubsetFontName = buildSubsetFontName(i + 1, usage.BaseFontName);
      usage.FontFileObjectId = id++;
      usage.FontDescriptorObjectId = id++;
      usage.ToUnicodeObjectId = id++;
      usage.FontObjectId = id++;

      const subsetBytes = usage.BuildSubsetFontBytes();
      objects.set(usage.FontFileObjectId, buildFontFileStream(usage.FontFileObjectId, subsetBytes));
      objects.set(usage.FontDescriptorObjectId, buildFontDescriptor(usage));
      objects.set(usage.ToUnicodeObjectId, buildStreamObject(usage.ToUnicodeObjectId, buildToUnicodeContent(usage)));
      objects.set(usage.FontObjectId, buildFontObject(usage));
    }
    return { objects, nextObjectId: id };
  }
}

function buildFontObject(usage: PdfTrueTypePrototypeUsage): Uint8Array {
  const widths: string[] = [];
  for (let code = usage.FirstChar; code <= usage.LastChar; code++) {
    if (widths.length > 0) widths.push(" ");
    widths.push(fmt(usage.WidthForCode(code)));
  }
  return ascii(
    `${usage.FontObjectId} 0 obj\n<< /Type /Font\n/Subtype /TrueType\n/Name /${usage.ResourceName}\n/BaseFont /${usage.SubsetFontName}\n/Encoding /WinAnsiEncoding\n/FontDescriptor ${usage.FontDescriptorObjectId} 0 R\n/FirstChar ${usage.FirstChar}\n/LastChar ${usage.LastChar}\n/Widths [${widths.join("")}]\n/ToUnicode ${usage.ToUnicodeObjectId} 0 R\n>>\nendobj\n`,
  );
}

function buildFontDescriptor(usage: PdfTrueTypePrototypeUsage): Uint8Array {
  return ascii(
    `${usage.FontDescriptorObjectId} 0 obj\n<< /Type /FontDescriptor\n/FontName /${usage.SubsetFontName}\n/Flags 32\n/ItalicAngle 0\n/Ascent 810\n/Descent -232\n/CapHeight 810\n/StemV 44\n/FontBBox [0 0 1000 1000]\n/FontFile2 ${usage.FontFileObjectId} 0 R\n>>\nendobj\n`,
  );
}

function buildToUnicodeContent(usage: PdfTrueTypePrototypeUsage): Uint8Array {
  const sb: string[] = [];
  sb.push(
    "/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /",
  );
  sb.push(usage.ResourceName);
  sb.push(" def\n/CMapType 2 def\n1 begincodespacerange\n<00> <FF>\nendcodespacerange\n");
  const codes = usage.OrderedCodes();
  sb.push(`${codes.length} beginbfchar\n`);
  for (const code of codes) {
    sb.push(`<${code.toString(16).toUpperCase().padStart(2, "0")}> <00${code.toString(16).toUpperCase().padStart(2, "0")}>\n`);
  }
  sb.push("endbfchar\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n");
  return ascii(sb.join(""));
}

function buildStreamObject(objectId: number, streamBytes: Uint8Array): Uint8Array {
  const header = ascii(`${objectId} 0 obj\n<< /Length ${streamBytes.length} >>\nstream\n`);
  const footer = ascii("\nendstream\nendobj\n");
  const out = new Uint8Array(header.length + streamBytes.length + footer.length);
  out.set(header, 0);
  out.set(streamBytes, header.length);
  out.set(footer, header.length + streamBytes.length);
  return out;
}

function buildFontFileStream(objectId: number, streamBytes: Uint8Array): Uint8Array {
  const header = ascii(
    `${objectId} 0 obj\n<< /Length ${streamBytes.length}\n/Length1 ${streamBytes.length} >>\nstream\n`,
  );
  const footer = ascii("\nendstream\nendobj\n");
  const out = new Uint8Array(header.length + streamBytes.length + footer.length);
  out.set(header, 0);
  out.set(streamBytes, header.length);
  out.set(footer, header.length + streamBytes.length);
  return out;
}

function buildSubsetFontName(usageIndex: number, baseFontName: string): string {
  let value = usageIndex;
  const chars: string[] = ["A", "A", "A", "A", "A", "A"];
  for (let i = 5; i >= 0; i--) {
    chars[i] = String.fromCharCode(0x41 + (value % 26));
    value = Math.floor(value / 26);
  }
  return chars.join("") + "+" + baseFontName;
}

function ascii(text: string): Uint8Array {
  const out = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) out[i] = text.charCodeAt(i) & 0xff;
  return out;
}

function fmt(value: number): string {
  const rounded = Math.round(value * 1000) / 1000;
  if (rounded === 0) return "0";
  return rounded.toString();
}
