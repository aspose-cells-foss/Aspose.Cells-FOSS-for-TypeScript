/**
 * Lightweight TypeScript port of Aspose.Cells_FOSS.Rendering.FontRegistry.
 *
 * Caches typeface references and provides a codepoint-coverage fallback
 * chain so every glyph in the workbook resolves.  The actual glyph lookup
 * is abstracted behind {@link ITypefaceProvider} so this class works with
 * any backend (pdfkit, @napi-rs/canvas, freetype-wasm, …).
 */

export interface FontValue {
  Name?: string;
  Size?: number;
  Bold?: boolean;
  Italic?: boolean;
}

export interface TypefaceInfo {
  familyName: string;
  /** True if the face contains a glyph for `codepoint`. */
  containsGlyph(codepoint: number): boolean;
  /** Width (px at 96 DPI) of codepoint '0' at `sizePt`. */
  measureDigitWidthPxAt(sizePt: number): number;
}

export interface ITypefaceProvider {
  /** Resolve (name, bold, italic) → TypefaceInfo or default if unavailable. */
  resolve(name: string, bold: boolean, italic: boolean): TypefaceInfo;
  /** Walk platform fallback chain for a glyph. */
  resolveForCodepoint(primary: TypefaceInfo, codepoint: number): TypefaceInfo;
}

/** Calibrated digit widths at 11 pt / 96 DPI (Excel-matched). */
const ExcelDigitWidthPxAt11: Record<string, number> = {
  Simsun: 7.34,
  "宋体": 7.34,
  Nsimsun: 7.34,
  "新宋体": 7.34,
};

const DefaultDigitWidthPx = 7.5;

export class FontRegistry {
  DefaultFontName?: string;

  private _provider: ITypefaceProvider;

  constructor(provider: ITypefaceProvider) {
    this._provider = provider;
  }

  Resolve(font: FontValue | null | undefined): TypefaceInfo {
    const name = font?.Name?.length ? font.Name : "Calibri";
    const bold = !!font?.Bold;
    const italic = !!font?.Italic;
    return this._provider.resolve(name, bold, italic);
  }

  ResolveForCodepoint(primary: TypefaceInfo, codepoint: number): TypefaceInfo {
    if (primary.containsGlyph(codepoint)) return primary;
    return this._provider.resolveForCodepoint(primary, codepoint);
  }

  /** Max digit width in pixels at 96 DPI (used to convert Excel character widths). */
  MeasureMaxDigitWidth(normalFont: FontValue | null | undefined): number {
    const sizePt = normalFont?.Size ?? 11;

    // DengXian override – covers both OOXML name and resolved family.
    if (
      normalFont?.Name &&
      (/^dengxian$/i.test(normalFont.Name) || normalFont.Name === "等线")
    ) {
      return 7.75 * (sizePt / 11);
    }

    const typeface = this.Resolve(normalFont);

    if (normalFont?.Name) {
      const override = ExcelDigitWidthPxAt11[normalFont.Name.toLowerCase()];
      if (override !== undefined) return override * (sizePt / 11);
    }
    const overrideFamily = ExcelDigitWidthPxAt11[typeface.familyName.toLowerCase()];
    if (overrideFamily !== undefined) return overrideFamily * (sizePt / 11);

    const widthPx = typeface.measureDigitWidthPxAt(sizePt);
    if (widthPx <= 0) return DefaultDigitWidthPx;
    if (/^tahoma$/i.test(typeface.familyName)) return widthPx;
    return Math.ceil(widthPx * 2) / 2;
  }
}
