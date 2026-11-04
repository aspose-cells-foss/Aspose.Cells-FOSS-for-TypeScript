import { PdfFontUsageRegistry } from "./PdfFontUsageRegistry";
import { PdfType3FontBuilder } from "./PdfType3FontBuilder";
import type { PdfTextFontBackend } from "./PdfTextFontBackend";
import { PdfEncodedTextSegment } from "./PdfEncodedTextSegment";
import type { PdfType3FontUsage } from "./PdfType3FontUsage";

/**
 * Type3-font based backend.  Every glyph gets its own CharProcs stream; encoding
 * is a single-byte custom encoding per subset (up to 255 glyphs each).
 */
export class PdfType3TextFontBackend implements PdfTextFontBackend {
  private _registry: PdfFontUsageRegistry;
  private _builder: PdfType3FontBuilder;
  private _createUsage: (typeface: unknown, usePositiveY: boolean) => PdfType3FontUsage;

  constructor(
    createUsage: (typeface: unknown, usePositiveY: boolean) => PdfType3FontUsage,
    builder?: PdfType3FontBuilder,
  ) {
    this._createUsage = createUsage;
    this._registry = new PdfFontUsageRegistry(createUsage);
    this._builder = builder ?? new PdfType3FontBuilder();
  }

  Initialize(runs: { text: string; typeface: unknown; usePositiveYFontMatrix: boolean }[]): void {
    this._registry.AddRuns(runs);
    this._registry.FinalizeSubsets();
  }

  BuildObjects(nextObjectId: number): { objects: Map<number, Uint8Array>; nextObjectId: number } {
    return this._builder.BuildObjects(this._registry as unknown as { Usages: PdfType3FontUsage[] }, nextObjectId);
  }

  BuildFontResourceText(): string {
    const parts: string[] = ["/Font <<"];
    for (const usage of this._registry.Usages) {
      for (const subset of usage.Subsets) {
        parts.push(` /${subset.ResourceName} ${subset.FontObjectId} 0 R`);
      }
    }
    parts.push(" >>");
    return parts.join("");
  }

  EncodeRun(run: { text: string; typeface: unknown; usePositiveYFontMatrix: boolean }): PdfEncodedTextSegment[] {
    const segments: PdfEncodedTextSegment[] = [];
    if (!run?.text) return segments;
    const usage = this._registry.UsageFor(run.typeface, run.usePositiveYFontMatrix);
    let cursor = 0;
    let i = 0;
    while (i < run.text.length) {
      let subset: { ResourceName: string } | null = null;
      const bytes: number[] = [];
      const segmentStart = cursor;
      while (i < run.text.length) {
        const token = run.text.substring(i, readTokenLen(run.text, i));
        const resolved = usage.TryResolveToken(token);
        if (!resolved) break;
        if (subset == null) {
          subset = resolved.subset;
        } else if (subset !== resolved.subset) {
          break;
        }
        bytes.push(resolved.code);
        cursor += resolved.width1000 / 1000;
        i += token.length;
      }
      if (subset && bytes.length > 0) {
        const seg = new PdfEncodedTextSegment();
        seg.ResourceName = subset.ResourceName;
        seg.EncodedBytes = Uint8Array.from(bytes);
        seg.StartXPt = segmentStart;
        segments.push(seg);
      }
    }
    return segments;
  }

  Dispose(): void {
    this._registry.Dispose();
  }
}

function readTokenLen(text: string, i: number): number {
  const ch = text.charCodeAt(i);
  if (ch >= 0xd800 && ch <= 0xdbff && i + 1 < text.length) {
    const lo = text.charCodeAt(i + 1);
    if (lo >= 0xdc00 && lo <= 0xdfff) return 2;
  }
  return 1;
}
