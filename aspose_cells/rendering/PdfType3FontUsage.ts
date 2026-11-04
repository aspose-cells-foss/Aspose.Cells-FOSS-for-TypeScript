import { PdfType3FontSubset } from "./PdfType3FontSubset";
import type { PdfPath } from "./PdfPathSerializer";

/** Font metrics helper (SkiaSharp- or Canvas2D-backed). */
export interface FontAt1000 {
  MeasureText(text: string): number;
  BuildTextPath(text: string): PdfPath | null;
}

/** One resolved font's glyph coverage and Type3 subset metadata. */
export class PdfType3FontUsage {
  readonly Typeface: unknown;
  UsePositiveYFontMatrix: boolean;

  /** token → width at 1000 units */
  private _tokenWidths = new Map<string, number>();
  readonly Subsets: PdfType3FontSubset[] = [];

  private _fontAt1000: FontAt1000;

  constructor(typeface: unknown, usePositiveYFontMatrix: boolean, fontAt1000: FontAt1000) {
    this.Typeface = typeface;
    this.UsePositiveYFontMatrix = usePositiveYFontMatrix;
    this._fontAt1000 = fontAt1000;
  }

  BaseFontName = "SkiaFont";

  AddText(text: string): void {
    if (!text) return;
    let i = 0;
    while (i < text.length) {
      const token = readToken(text, i);
      if (!this._tokenWidths.has(token)) {
        this._tokenWidths.set(token, this._fontAt1000.MeasureText(token));
      }
      i += token.length;
    }
  }

  /** Assigns PDF resource indices and slices tokens into ≤ 255-glyph subsets. */
  FinalizeSubsets(startIndex: number): number {
    let resourceIndex = startIndex;
    if (this.Subsets.length > 0) return resourceIndex;
    const orderedTokens = [...this._tokenWidths.keys()].sort();
    let current: PdfType3FontSubset | null = null;
    for (const token of orderedTokens) {
      if (current == null || current.Tokens.length >= 255) {
        current = new PdfType3FontSubset();
        current.ResourceName = `T3F${resourceIndex++}`;
        current.UsePositiveYFontMatrix = this.UsePositiveYFontMatrix;
        this.Subsets.push(current);
      }
      const code = current.Tokens.length + 1;
      current.Tokens.push(token);
      current.TokenCodes.set(token, code);
    }
    return resourceIndex;
  }

  TryResolveToken(token: string): { subset: PdfType3FontSubset; code: number; width1000: number } | null {
    const width1000 = this._tokenWidths.get(token);
    if (width1000 === undefined) return null;
    for (const subset of this.Subsets) {
      const code = subset.TokenCodes.get(token);
      if (code !== undefined) return { subset, code, width1000 };
    }
    return null;
  }

  GetTokenWidth(token: string): number {
    return this._tokenWidths.get(token) ?? 0;
  }

  BuildTokenPath(token: string): PdfPath | null {
    return this._fontAt1000.BuildTextPath(token);
  }

  Dispose(): void {
    // Path/handle cleanup is owned by the provider.
  }
}

function readToken(text: string, i: number): string {
  const ch = text.charCodeAt(i);
  if (ch >= 0xd800 && ch <= 0xdbff && i + 1 < text.length) {
    const lo = text.charCodeAt(i + 1);
    if (lo >= 0xdc00 && lo <= 0xdfff) return text.substring(i, i + 2);
  }
  return text.substring(i, i + 1);
}
