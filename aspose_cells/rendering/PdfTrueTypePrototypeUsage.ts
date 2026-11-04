import { TrueTypeSubsetFontProgram } from "./TrueTypeSubsetFontProgram";

/** Abstracts TrueType font queries behind a backend-agnostic interface. */
export interface ITrueTypeFont {
  familyName: string;
  /** Read the full font file bytes from disk/memory. */
  getFontBytes(): Uint8Array;
  /** Map codepoint → glyph ID. */
  getGlyph(codepoint: number): number;
  /** Width (at 1000 units per em) for the codepoint. */
  measureWidth1000(codepoint: number): number;
}

/**
 * One TrueType font subsetting session.  Aggregates every ASCII codepoint (32-126)
 * actually used by a run and emits a self-contained subset font file + widths
 * table suitable for a PDF /TrueType font resource.
 */
export class PdfTrueTypePrototypeUsage {
  readonly Typeface: unknown;
  ResourceName = "";
  FontObjectId = 0;
  FontDescriptorObjectId = 0;
  FontFileObjectId = 0;
  ToUnicodeObjectId = 0;
  SubsetFontName = "";

  private _font: ITrueTypeFont;
  private _program: TrueTypeSubsetFontProgram;
  private _codes = new Set<number>();
  private _widths1000 = new Map<number, number>();
  private _codeToGlyph = new Map<number, number>();
  private _glyphToUnicode = new Map<number, number>();

  constructor(typeface: unknown, font: ITrueTypeFont) {
    this.Typeface = typeface;
    this._font = font;
    const bytes = font.getFontBytes();
    this._program = new TrueTypeSubsetFontProgram(bytes);
  }

  get BaseFontName(): string {
    const family = this._font.familyName || "SkiaFont";
    let out = "";
    for (let i = 0; i < family.length; i++) {
      const ch = family.charCodeAt(i);
      if ((ch >= 0x30 && ch <= 0x39) || (ch >= 0x41 && ch <= 0x5a) || (ch >= 0x61 && ch <= 0x7a)) {
        out += String.fromCharCode(ch);
      }
    }
    return out || "SkiaFont";
  }

  get FirstChar(): number {
    let first = 255;
    for (const c of this._codes) if (c < first) first = c;
    return this._codes.size === 0 ? 32 : first;
  }
  get LastChar(): number {
    let last = 32;
    for (const c of this._codes) if (c > last) last = c;
    return last;
  }

  AddText(text: string): void {
    if (!text) return;
    for (let i = 0; i < text.length; i++) {
      const cp = text.charCodeAt(i);
      if (cp > 255) continue;
      const code = cp;
      if (!this._codes.has(code)) {
        this._codes.add(code);
        const glyphId = this._font.getGlyph(cp);
        this._codeToGlyph.set(code, glyphId);
        if (!this._glyphToUnicode.has(glyphId)) this._glyphToUnicode.set(glyphId, cp);
        this._widths1000.set(code, this._font.measureWidth1000(cp));
      }
    }
  }

  EncodeText(text: string): Uint8Array {
    const bytes: number[] = [];
    for (let i = 0; i < text.length; i++) {
      const cp = text.charCodeAt(i);
      if (cp > 255) continue;
      bytes.push(cp);
    }
    return Uint8Array.from(bytes);
  }

  WidthForCode(code: number): number {
    return this._widths1000.get(code) ?? 0;
  }

  OrderedCodes(): number[] {
    return [...this._codes].sort((a, b) => a - b);
  }

  OrderedGlyphIds(): number[] {
    return [...this._codeToGlyph.values()].sort((a, b) => a - b);
  }

  BuildSubsetFontBytes(): Uint8Array {
    return this._program.BuildSubset(
      this.OrderedGlyphIds(),
      this._codeToGlyph,
      this._glyphToUnicode,
      this.SubsetFontName,
    );
  }

  Dispose(): void {
    this._codes.clear();
    this._widths1000.clear();
    this._codeToGlyph.clear();
    this._glyphToUnicode.clear();
  }
}
