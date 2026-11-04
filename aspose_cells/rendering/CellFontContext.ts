import { FontRegistry, type FontValue, type TypefaceInfo } from "./FontRegistry";

/** Single typeface-homogeneous slice of a text string with its measured width in points. */
export class GlyphRun {
  constructor(
    public Typeface: TypefaceInfo,
    public Text: string,
    public WidthPt: number,
  ) {}
}

export interface ITextMeasurer {
  /** Width of `text` at `sizePt` points, in points. */
  MeasureText(text: string, sizePt: number, typeface: TypefaceInfo): number;
  /** Ascent + descent of `typeface` at `sizePt` points. */
  MeasureMetrics(sizePt: number, typeface: TypefaceInfo): { ascent: number; descent: number; leading: number };
}

/**
 * Wraps a resolved font configuration and provides per-typeface measurements
 * in points (72 DPI PDF canvas).  Layout uses these widths directly against
 * the column-width values produced by {@link FontRegistry}.
 */
export class CellFontContext {
  SizePt: number;

  private _registry: FontRegistry;
  private _primary: TypefaceInfo;
  private _measurer: ITextMeasurer;

  constructor(registry: FontRegistry, font: FontValue | null | undefined, measurer: ITextMeasurer) {
    this._registry = registry;
    this._measurer = measurer;
    this._primary = registry.Resolve(font);
    this.SizePt = font?.Size ?? 11;
    if (this.SizePt <= 0) this.SizePt = 11;
  }

  /** Approximate line height in points (Excel ≈ em-size × 1.2). */
  get LineHeightPt(): number {
    return this.SizePt * 1.2;
  }

  /** Full line spacing from font metrics, +1 pt cell padding. */
  get WrappedLineHeightPt(): number {
    const m = this._measurer.MeasureMetrics(this.SizePt, this._primary);
    return Math.max(this.LineHeightPt, -m.ascent + m.descent + m.leading + 1);
  }

  /** Glyph height (ascent + descent only). */
  get TextHeightPt(): number {
    const m = this._measurer.MeasureMetrics(this.SizePt, this._primary);
    return -m.ascent + m.descent;
  }

  private measure(text: string, typeface: TypefaceInfo): number {
    return this._measurer.MeasureText(text, this.SizePt, typeface);
  }

  /** Split into runs, each backed by a single typeface. */
  SplitRuns(text: string): GlyphRun[] {
    const runs: GlyphRun[] = [];
    if (!text) return runs;
    let currentFace: TypefaceInfo | null = null;
    let currentBuilder = "";
    let i = 0;
    while (i < text.length) {
      let codepoint: number;
      let consumed: number;
      const ch = text.charCodeAt(i);
      if (ch >= 0xd800 && ch <= 0xdbff && i + 1 < text.length) {
        const lo = text.charCodeAt(i + 1);
        if (lo >= 0xdc00 && lo <= 0xdfff) {
          codepoint = 0x10000 + (ch - 0xd800) * 0x400 + (lo - 0xdc00);
          consumed = 2;
        } else {
          codepoint = ch;
          consumed = 1;
        }
      } else {
        codepoint = ch;
        consumed = 1;
      }
      const face = this._registry.ResolveForCodepoint(this._primary, codepoint);
      if (currentFace == null) {
        currentFace = face;
      } else if (currentFace !== face) {
        runs.push(new GlyphRun(currentFace, currentBuilder, this.measure(currentBuilder, currentFace)));
        currentBuilder = "";
        currentFace = face;
      }
      currentBuilder += text.substring(i, i + consumed);
      i += consumed;
    }
    if (currentBuilder.length > 0 && currentFace != null) {
      runs.push(new GlyphRun(currentFace, currentBuilder, this.measure(currentBuilder, currentFace)));
    }
    return runs;
  }

  /** Sum of SplitRuns widths. */
  Measure(text: string): number {
    return this.SplitRuns(text).reduce((a, r) => a + r.WidthPt, 0);
  }

  /**
   * Width measured at 1000pt units then scaled down – matches Type3 font widths.
   * This is what the PDF writer uses for column-width → glyph-advance mapping.
   */
  SplitPdfRuns(text: string): GlyphRun[] {
    const runs = this.SplitRuns(text);
    for (const r of runs) {
      let widthAt1000 = 0;
      let i = 0;
      while (i < r.Text.length) {
        let consumed = 1;
        const ch = r.Text.charCodeAt(i);
        if (ch >= 0xd800 && ch <= 0xdbff && i + 1 < r.Text.length) {
          const lo = r.Text.charCodeAt(i + 1);
          if (lo >= 0xdc00 && lo <= 0xdfff) consumed = 2;
        }
        widthAt1000 += this._measurer.MeasureText(r.Text.substring(i, i + consumed), 1000, r.Typeface);
        i += consumed;
      }
      r.WidthPt = (widthAt1000 * this.SizePt) / 1000;
    }
    return runs;
  }

  MeasurePdfText(text: string): number {
    return this.SplitPdfRuns(text).reduce((a, r) => a + r.WidthPt, 0);
  }

  /** Greedy line breaking on whitespace, hard-break single tokens that overflow. */
  WrapLines(text: string, maxWidthPt: number, pdfMetrics = false): string[] {
    const lines: string[] = [];
    if (!text) {
      lines.push("");
      return lines;
    }
    const paragraphs = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
    const measure = (s: string) => (pdfMetrics ? this.MeasurePdfText(s) : this.Measure(s));
    for (const para of paragraphs) {
      if (maxWidthPt <= 0 || para.length === 0) {
        lines.push(para);
        continue;
      }
      const tokens = tokenize(para);
      let current = "";
      for (const tok of tokens) {
        const candidate = current.length === 0 ? tok : current + tok;
        if (measure(candidate.replace(/\s+$/, "")) <= maxWidthPt) {
          current = candidate;
          continue;
        }
        if (current.length > 0) {
          lines.push(current.replace(/\s+$/, ""));
          current = "";
        }
        if (measure(tok.replace(/\s+$/, "")) > maxWidthPt) {
          for (const ch of tok) {
            const cand2 = current + ch;
            if (/\s/.test(ch)) {
              current = cand2;
              continue;
            }
            if (current.length > 0 && measure(cand2) > maxWidthPt) {
              lines.push(current);
              current = "";
            }
            current += ch;
          }
        } else {
          current = tok;
        }
      }
      lines.push(current.replace(/\s+$/, ""));
    }
    return lines;
  }
}

function tokenize(paragraph: string): string[] {
  const tokens: string[] = [];
  let builder = "";
  for (const ch of paragraph) {
    builder += ch;
    if (ch === " " || ch === "\t") {
      tokens.push(builder);
      builder = "";
    }
  }
  if (builder.length > 0) tokens.push(builder);
  return tokens;
}
