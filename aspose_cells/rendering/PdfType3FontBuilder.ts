import type { PdfType3FontUsage } from "./PdfType3FontUsage";
import type { PdfType3FontSubset } from "./PdfType3FontSubset";
import { SerializePath, type PdfPath } from "./PdfPathSerializer";

/** Registry abstraction – mirrors PdfFontUsageRegistry in .NET. */
export interface IUsageRegistry {
  Usages: PdfType3FontUsage[];
}

/** Builds a batch of PDF objects defining one or more Type3 fonts. */
export class PdfType3FontBuilder {
  BuildObjects(registry: IUsageRegistry, nextObjectId: number): { objects: Map<number, Uint8Array>; nextObjectId: number } {
    const objects = new Map<number, Uint8Array>();
    let id = nextObjectId;
    for (const usage of registry.Usages) {
      for (const subset of usage.Subsets) {
        const glyphIds = new Map<string, number>();
        let fontBounds: { left: number; top: number; right: number; bottom: number } | null = null;

        for (const token of subset.Tokens) {
          const width = usage.GetTokenWidth(token);
          const path = usage.BuildTokenPath(token);
          const glyphId = id++;
          glyphIds.set(token, glyphId);
          const content = buildCharProcContent(path, width);
          objects.set(glyphId, buildStreamObject(glyphId, content));
          if (path && !path.IsEmpty) {
            if (!fontBounds) {
              fontBounds = { ...path.Bounds };
            } else {
              fontBounds = {
                left: Math.min(fontBounds.left, path.Bounds.left),
                top: Math.min(fontBounds.top, path.Bounds.top),
                right: Math.max(fontBounds.right, path.Bounds.right),
                bottom: Math.max(fontBounds.bottom, path.Bounds.bottom),
              };
            }
          }
        }

        subset.ToUnicodeObjectId = id++;
        objects.set(subset.ToUnicodeObjectId, buildStreamObject(subset.ToUnicodeObjectId, buildToUnicodeContent(subset)));
        subset.FontObjectId = id++;
        objects.set(
          subset.FontObjectId,
          buildFontObject(subset, usage, glyphIds, fontBounds ?? { left: 0, top: 0, right: 0, bottom: 0 }, !!fontBounds),
        );
      }
    }
    return { objects, nextObjectId: id };
  }
}

function buildFontObject(
  subset: PdfType3FontSubset,
  usage: PdfType3FontUsage,
  glyphIds: Map<string, number>,
  bounds: { left: number; top: number; right: number; bottom: number },
  hasBounds: boolean,
): Uint8Array {
  const sb: string[] = [];
  sb.push(`${subset.FontObjectId} 0 obj\n<< /Type /Font\n/Subtype /Type3\n/Name /${subset.ResourceName}\n/FontBBox [${hasBounds
    ? `${fmt(bounds.left)} ${fmt(bounds.top)} ${fmt(bounds.right)} ${fmt(bounds.bottom)}`
    : "0 0 0 0"}]\n/FontMatrix ${buildFontMatrix(usage.UsePositiveYFontMatrix)}\n/CharProcs <<`);
  for (const token of subset.Tokens) {
    const code = subset.TokenCodes.get(token)!;
    sb.push(` /g${code} ${glyphIds.get(token)!} 0 R`);
  }
  sb.push(" >>\n/Encoding << /Type /Encoding /Differences [1");
  for (const token of subset.Tokens) sb.push(` /g${subset.TokenCodes.get(token)!}`);
  sb.push(`] >>\n/FirstChar 1\n/LastChar ${subset.Tokens.length}\n/Widths [`);
  for (let i = 0; i < subset.Tokens.length; i++) {
    if (i > 0) sb.push(" ");
    sb.push(fmt(usage.GetTokenWidth(subset.Tokens[i])));
  }
  sb.push(`]\n/Resources << >>\n/ToUnicode ${subset.ToUnicodeObjectId} 0 R\n>>\nendobj\n`);
  return ascii(sb.join(""));
}

function buildCharProcContent(path: PdfPath | null, width: number): Uint8Array {
  if (!path || path.IsEmpty) return ascii(`${fmt(width)} 0 d0\n`);
  const b = path.Bounds;
  return ascii(
    `${fmt(width)} 0 ${fmt(b.left)} ${fmt(b.top)} ${fmt(b.right)} ${fmt(b.bottom)} d1\n${SerializePath(path)}f\n`,
  );
}

function buildToUnicodeContent(subset: PdfType3FontSubset): Uint8Array {
  const sb: string[] = [];
  sb.push(
    "/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /",
  );
  sb.push(subset.ResourceName);
  sb.push(" def\n/CMapType 2 def\n1 begincodespacerange\n<00> <FF>\nendcodespacerange\n");
  // Single-bucket mapping: one range covers contiguous simple scalars; everything else is a single.
  // For our typical OOXML token set this is effectively always singles.
  sb.push(`${subset.Tokens.length} beginbfchar\n`);
  for (const token of subset.Tokens) {
    const code = subset.TokenCodes.get(token)!;
    sb.push(`<${code.toString(16).toUpperCase().padStart(2, "0")}> <${toHexUnicode(token)}>\n`);
  }
  sb.push("endbfchar\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n");
  return ascii(sb.join(""));
}

function toHexUnicode(text: string): string {
  // Big-endian UTF-16 hex string.
  let out = "";
  for (let i = 0; i < text.length; i++) {
    const cp = text.charCodeAt(i);
    out += cp.toString(16).toUpperCase().padStart(4, "0");
  }
  return out;
}

function buildFontMatrix(positiveY: boolean): string {
  return positiveY ? "[0.001 0 0 0.001 0 0]" : "[0.001 0 0 -0.001 0 0]";
}

function buildStreamObject(objectId: number, streamBytes: Uint8Array): Uint8Array {
  // Prefer uncompressed for small streams to save compressor overhead; this matches the
  // .NET implementation's heuristic at compile time.  We always emit the flat form here –
  // callers can add /FlateDecode later if needed.
  const header = ascii(`${objectId} 0 obj\n<< /Length ${streamBytes.length} >>\nstream\n`);
  const footer = ascii("\nendstream\nendobj\n");
  const out = new Uint8Array(header.length + streamBytes.length + footer.length);
  out.set(header, 0);
  out.set(streamBytes, header.length);
  out.set(footer, header.length + streamBytes.length);
  return out;
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
