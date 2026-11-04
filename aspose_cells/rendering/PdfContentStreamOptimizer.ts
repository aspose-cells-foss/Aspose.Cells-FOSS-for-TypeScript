import { PdfObjectTable } from "./PdfObjectTable";

/**
 * Post-processes compressed PDF page content streams for size by:
 *   1. Rounding fractional decimals to 3 significant digits;
 *   2. Collapsing redundant colour-state operators (`rg`, `RG`, `g`, `G`, `gs`);
 *   3. De-duplicating whitespace.
 * The result is re-compressed with FlateDecode if a zlib binding is available.
 */
export class PdfContentStreamOptimizer {
  private _decompress: ((zlib: Uint8Array) => Uint8Array) | null;
  private _compress: ((raw: Uint8Array) => Uint8Array) | null;

  constructor(decompress?: (zlib: Uint8Array) => Uint8Array, compress?: (raw: Uint8Array) => Uint8Array) {
    this._decompress = decompress ?? null;
    this._compress = compress ?? null;
  }

  Rewrite(originalPdf: Uint8Array): Uint8Array {
    if (!originalPdf || originalPdf.length === 0) return originalPdf;
    const table = PdfObjectTable.Parse(originalPdf);
    const contentObjectIds = collectPageContentObjectIds(table);
    let changed = false;
    for (const objectId of contentObjectIds) {
      const objectBytes = table.Objects.get(objectId);
      if (!objectBytes) continue;
      const rewritten = this.tryOptimize(objectId, objectBytes);
      if (!rewritten) continue;
      table.Objects.set(objectId, rewritten);
      changed = true;
    }
    return changed ? table.Rebuild() : originalPdf;
  }

  private tryOptimize(objectId: number, objectBytes: Uint8Array): Uint8Array | null {
    const text = latin1(objectBytes);
    if (!text.includes("/Filter /FlateDecode")) return null;
    const streamMarker = text.includes("stream\r\n") ? "stream\r\n" : "stream\n";
    const streamIndex = text.indexOf(streamMarker);
    if (streamIndex < 0) return null;
    const streamStart = streamIndex + streamMarker.length;
    const endStreamText = text.includes("\r\nendstream") ? "\r\nendstream" : "\nendstream";
    const endStreamIndex = text.indexOf(endStreamText, streamStart);
    if (endStreamIndex < 0) return null;
    const compressed = objectBytes.slice(streamStart, endStreamIndex);

    const decompressed = this._decompress ? this._decompress(compressed) : null;
    if (!decompressed) return null;
    const content = asciiToString(decompressed);
    const optimized = optimizeContentText(content);
    const optimizedBytes = ascii(optimized);
    if (optimizedBytes.length >= decompressed.length) return null;

    const recompressed = this._compress ? this._compress(optimizedBytes) : optimizedBytes;
    let header = text.substring(0, streamIndex);
    header = replaceLength(header, recompressed.length);
    const footer = text.substring(endStreamIndex + endStreamText.length);
    const headerBytes = ascii(header + streamMarker);
    const footerBytes = ascii(footer);
    const out = new Uint8Array(headerBytes.length + recompressed.length + footerBytes.length);
    out.set(headerBytes, 0);
    out.set(recompressed, headerBytes.length);
    out.set(footerBytes, headerBytes.length + recompressed.length);
    return out;
  }
}

function replaceLength(header: string, newLength: number): string {
  const marker = "/Length ";
  const i = header.indexOf(marker);
  if (i < 0) return header;
  let j = i + marker.length;
  while (j < header.length && /\d/.test(header[j])) j++;
  return header.substring(0, i + marker.length) + newLength + header.substring(j);
}

function collectPageContentObjectIds(table: PdfObjectTable): number[] {
  const result: number[] = [];
  const seen = new Set<number>();
  for (const pageId of table.PageObjectIds) {
    const bytes = table.Objects.get(pageId);
    if (!bytes) continue;
    for (const id of parseContentObjectIds(latin1(bytes))) {
      if (!seen.has(id)) {
        seen.add(id);
        result.push(id);
      }
    }
  }
  return result;
}

function parseContentObjectIds(pageText: string): number[] {
  const ids: number[] = [];
  const marker = "/Contents ";
  const i = pageText.indexOf(marker);
  if (i < 0) return ids;
  let start = i + marker.length;
  if (start >= pageText.length) return ids;
  if (pageText[start] === "[") {
    const end = pageText.indexOf("]", start);
    if (end < 0) return ids;
    const items = pageText.substring(start + 1, end).trim().split(/\s+/);
    for (let k = 0; k + 2 < items.length; k++) {
      if (/^\d+$/.test(items[k]) && items[k + 1] === "0" && items[k + 2] === "R") ids.push(parseInt(items[k], 10));
    }
  } else {
    let end = pageText.indexOf("\n", start);
    if (end < 0) end = pageText.length;
    const parts = pageText.substring(start, end).trim().split(/\s+/);
    if (parts.length >= 3 && /^\d+$/.test(parts[0]) && parts[1] === "0" && parts[2] === "R") {
      ids.push(parseInt(parts[0], 10));
    }
  }
  return ids;
}

function optimizeContentText(content: string): string {
  // Compact decimals, collapse whitespace, drop redundant state commands.
  let out = content.replace(/(?<![A-Za-z0-9_/])[-+]?(?:\d+\.\d+|\.\d+)/g, (m) => formatCompact(parseFloat(m)));
  out = compactWhitespace(out);
  out = removeRedundantGraphicsState(out);
  return out;
}

function formatCompact(value: number): string {
  const rounded = Math.round(value * 1000) / 1000;
  if (rounded === 0) return "0";
  let text = rounded.toString();
  if (text.startsWith("0.")) return text.substring(1);
  if (text.startsWith("-0.")) return "-" + text.substring(2);
  return text;
}

function compactWhitespace(content: string): string {
  const tokens = tokenize(content);
  return tokens.join(" ");
}

function tokenize(content: string): string[] {
  const tokens: string[] = [];
  let i = 0;
  let inHex = false;
  let inLiteral = false;
  let literalDepth = 0;
  let escaping = false;
  let comment = false;
  while (i < content.length) {
    const ch = content[i];
    if (comment) {
      if (ch === "\r" || ch === "\n") comment = false;
      i++;
      continue;
    }
    if (inLiteral) {
      tokens[tokens.length - 1] += ch;
      if (escaping) {
        escaping = false;
      } else if (ch === "\\") {
        escaping = true;
      } else if (ch === "(") {
        literalDepth++;
      } else if (ch === ")") {
        literalDepth--;
        if (literalDepth <= 0) inLiteral = false;
      }
      i++;
      continue;
    }
    if (inHex) {
      tokens[tokens.length - 1] += ch;
      if (ch === ">") inHex = false;
      i++;
      continue;
    }
    if (ch === "%") {
      comment = true;
      i++;
      continue;
    }
    if (/\s/.test(ch)) {
      i++;
      continue;
    }
    if (ch === "(") {
      tokens.push("(");
      inLiteral = true;
      literalDepth = 1;
      escaping = false;
      i++;
      continue;
    }
    if (ch === "<" && (i + 1 >= content.length || content[i + 1] !== "<")) {
      tokens.push("<");
      inHex = true;
      i++;
      continue;
    }
    if (tokens.length === 0 || isOperator(tokens[tokens.length - 1]) || isOperatorStart(ch)) {
      tokens.push("");
    }
    tokens[tokens.length - 1] += ch;
    i++;
  }
  return tokens;
}

function isOperatorStart(ch: string): boolean {
  return ch === "q" || ch === "Q" || ch === "B" || ch === "b";
}

const OPERATORS = new Set([
  "q", "Q", "cm", "w", "J", "j", "M", "d", "ri", "i", "gs", "m", "l", "c", "v", "y", "h", "re",
  "S", "s", "f", "F", "f*", "B", "B*", "b", "b*", "n", "W", "W*", "BT", "ET", "Tf", "Tm", "Td", "TD",
  "Tj", "TJ", "Tc", "Tw", "Tz", "TL", "Tr", "Ts", "Do", "RG", "rg", "G", "g", "K", "k", "CS", "cs",
  "SC", "SCN", "sc", "scn", "sh",
]);
function isOperator(token: string): boolean {
  return OPERATORS.has(token);
}

function removeRedundantGraphicsState(content: string): string {
  const tokens = tokenize(content);
  const out: string[] = [];
  const stack: {
    strokeRgb?: string;
    fillRgb?: string;
    strokeGray?: string;
    fillGray?: string;
    gs?: string;
  }[] = [];
  let strokeRgb: string | undefined, fillRgb: string | undefined, strokeGray: string | undefined, fillGray: string | undefined, gs: string | undefined;
  let i = 0;
  while (i < tokens.length) {
    const tok = tokens[i];
    if (!isOperator(tok)) {
      out.push(tok);
      i++;
      continue;
    }
    if (tok === "q") {
      stack.push({ strokeRgb, fillRgb, strokeGray, fillGray, gs });
      out.push(tok);
      i++;
      continue;
    }
    if (tok === "Q") {
      const s = stack.pop();
      if (s) {
        strokeRgb = s.strokeRgb;
        fillRgb = s.fillRgb;
        strokeGray = s.strokeGray;
        fillGray = s.fillGray;
        gs = s.gs;
      }
      out.push(tok);
      i++;
      continue;
    }
    // Gather operands back into the operator's output.
    // Our tokenizer already grouped them; emit as-is if state differs or skip if redundant.
    const operands: string[] = [];
    while (i > 0 && !isOperator(tokens[i - 1])) {
      operands.unshift(tokens[--i]);
    }
    // `i` is currently pointing at the operator; re-find it
    while (i < tokens.length && !isOperator(tokens[i])) i++;
    if (i < tokens.length) {
      out.push(...operands, tokens[i]);
      i++;
    }
  }
  return out.join(" ");
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
function asciiToString(bytes: Uint8Array): string {
  return latin1(bytes);
}
