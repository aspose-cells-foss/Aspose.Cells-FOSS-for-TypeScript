import { PdfObjectTable } from "./PdfObjectTable";
import type { PdfTextFontBackend } from "./PdfTextFontBackend";
import type { PdfTextRunRecord } from "./PdfTextDocumentSession";

/**
 * Inserts a custom font backend's objects (Type3 or subset TrueType) into an
 * existing PDF, pre-computing encoded text content streams.  Mirrors the .NET
 * `PdfTextObjectWriter.Rewrite` entry point.
 */
export class PdfTextObjectWriter {
  private _backend: PdfTextFontBackend;

  constructor(backend: PdfTextFontBackend) {
    this._backend = backend;
  }

  Rewrite(originalPdf: Uint8Array, runs: PdfTextRunRecord[]): Uint8Array {
    if (!originalPdf || runs.length === 0) return originalPdf;
    this._backend.Initialize(
      runs.map((r) => ({ text: r.text, typeface: (r as unknown as { typeface?: unknown }).typeface, usePositiveYFontMatrix: false })),
    );

    const table = PdfObjectTable.Parse(originalPdf);
    let nextObjectId = table.MaxObjectId + 1;
    const fontObjects = this._backend.BuildObjects(nextObjectId);
    for (const [k, v] of fontObjects.objects) table.Objects.set(k, v);

    const fontResourceText = this._backend.BuildFontResourceText();
    for (let pageIndex = 0; pageIndex < table.PageObjectIds.length; pageIndex++) {
      const pageRuns = runs.filter((r) => getPageNumber(r) === pageIndex + 1);
      if (pageRuns.length === 0) continue;
      const pageObjectId = table.PageObjectIds[pageIndex];
      const pageObjectText = latin1(table.Objects.get(pageObjectId)!);
      const pageHeight = parsePageHeight(pageObjectText);
      const prologueObjectId = fontObjects.nextObjectId++;
      const contentObjectId = fontObjects.nextObjectId++;
      table.Objects.set(prologueObjectId, buildFlatStream(prologueObjectId, ascii("q\n")));
      table.Objects.set(contentObjectId, buildFlatStream(contentObjectId, buildPageTextContent(pageRuns, this._backend, pageHeight)));

      let updatedPageText = addFontResources(pageObjectText, fontResourceText);
      updatedPageText = appendContentReferences(updatedPageText, prologueObjectId, contentObjectId);
      table.Objects.set(pageObjectId, ascii(updatedPageText));
    }
    return table.Rebuild();
  }
}

function getPageNumber(run: PdfTextRunRecord): number {
  return (run as unknown as { pageNumber?: number }).pageNumber ?? 1;
}

function addFontResources(pageText: string, fontResources: string): string {
  if (pageText.includes("/Font <<")) return pageText;
  const marker = "/Resources <<";
  const index = pageText.indexOf(marker);
  if (index < 0) return pageText;
  return pageText.substring(0, index + marker.length) + fontResources + pageText.substring(index + marker.length);
}

function appendContentReferences(pageText: string, prologueId: number, contentId: number): string {
  const marker = "/Contents ";
  const index = pageText.indexOf(marker);
  if (index < 0) return pageText;
  let start = index + marker.length;
  if (pageText[start] === "[") {
    const end = pageText.indexOf("]", start);
    if (end < 0) return pageText;
    return pageText.substring(0, end) + ` ${prologueId} 0 R ${contentId} 0 R` + pageText.substring(end);
  }
  let end = pageText.indexOf("\n", start);
  if (end < 0) end = pageText.length;
  const existing = pageText.substring(start, end).trim();
  const replacement = `[${prologueId} 0 R ${existing} ${contentId} 0 R]`;
  return pageText.substring(0, start) + replacement + pageText.substring(end);
}

function buildPageTextContent(
  runs: PdfTextRunRecord[],
  backend: PdfTextFontBackend,
  _pageHeight: number,
): Uint8Array {
  const out: string[] = ["Q\n"];
  let start = 0;
  while (start < runs.length) {
    let end = start;
    while (end + 1 < runs.length && canBatch(runs[end], runs[end + 1])) end++;
    appendRunBatch(out, runs, start, end, backend);
    start = end + 1;
  }
  return ascii(out.join(""));
}

function appendRunBatch(out: string[], runs: PdfTextRunRecord[], start: number, end: number, backend: PdfTextFontBackend): void {
  const firstRun = runs[start];
  out.push("q\nBT\n");
  if ((firstRun as unknown as { invisible?: boolean }).invisible) out.push("3 Tr\n");
  out.push(colorCommand(firstRun));
  out.push("\n");
  let activeResource = "";
  let activeFontSize = -1;
  for (let i = start; i <= end; i++) {
    const run = runs[i];
    const encoded = backend.EncodeRun({
      text: run.text,
      typeface: (run as unknown as { typeface?: unknown }).typeface,
      usePositiveYFontMatrix: false,
    });
    for (const segment of encoded) {
      if (!segment.EncodedBytes || segment.EncodedBytes.length === 0) continue;
      if (segment.ResourceName !== activeResource || Math.abs(activeFontSize - run.fontSize) > 0.001) {
        out.push(`/${segment.ResourceName} ${fmt(run.fontSize)} Tf\n`);
        activeResource = segment.ResourceName;
        activeFontSize = run.fontSize;
      }
      out.push(`1 0 0 1 ${fmt(run.xPt)} ${fmt(run.yPt)} Tm\n<`);
      for (const b of segment.EncodedBytes) out.push(b.toString(16).toUpperCase().padStart(2, "0"));
      out.push("> Tj\n");
    }
  }
  out.push("ET\nQ\n");
}

function canBatch(a: PdfTextRunRecord, b: PdfTextRunRecord): boolean {
  if (!a || !b) return false;
  if ((a as unknown as { pageNumber?: number }).pageNumber !== (b as unknown as { pageNumber?: number }).pageNumber) return false;
  if (a.color !== b.color) return false;
  return true;
}

function colorCommand(run: PdfTextRunRecord): string {
  // Color is stored as a string ("#rrggbb" or "#rrggbbaa") in PdfTextRunRecord.
  // Use black as fallback if missing.
  const color = run.color || "#000000";
  const m = color.match(/^#?([0-9a-fA-F]{6})/);
  if (!m) return "0 0 0 rg";
  const hex = m[1];
  const r = parseInt(hex.substring(0, 2), 16) / 255;
  const g = parseInt(hex.substring(2, 4), 16) / 255;
  const bl = parseInt(hex.substring(4, 6), 16) / 255;
  return `${fmt(r)} ${fmt(g)} ${fmt(bl)} rg`;
}

function parsePageHeight(pageText: string): number {
  const marker = "/MediaBox [";
  const index = pageText.indexOf(marker);
  if (index < 0) return 842;
  const end = pageText.indexOf("]", index);
  if (end < 0) return 842;
  const parts = pageText.substring(index + marker.length, end).trim().split(/\s+/);
  if (parts.length !== 4) return 842;
  const y0 = parseFloat(parts[1]);
  const y1 = parseFloat(parts[3]);
  return Math.abs(y1 - y0);
}

function buildFlatStream(objectId: number, bytes: Uint8Array): Uint8Array {
  const header = ascii(`${objectId} 0 obj\n<< /Length ${bytes.length} >>\nstream\n`);
  const footer = ascii("\nendstream\nendobj\n");
  const out = new Uint8Array(header.length + bytes.length + footer.length);
  out.set(header, 0);
  out.set(bytes, header.length);
  out.set(footer, header.length + bytes.length);
  return out;
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

function fmt(value: number): string {
  const rounded = Math.round(value * 1000) / 1000;
  if (rounded === 0) return "0";
  return rounded.toString();
}
