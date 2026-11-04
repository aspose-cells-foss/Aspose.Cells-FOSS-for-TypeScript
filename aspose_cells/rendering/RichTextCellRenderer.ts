import type { RichTextSlice } from "./RichTextSlice";
import type { CellFontContext } from "./CellFontContext";
import type { FontRegistry } from "./FontRegistry";
import { RgbaColor } from "./ColorValue";

/**
 * Builds {@link RichTextSlice}s from a cell's rich-text runs, resolving each
 * run against the workbook's default style and precomputing widths.  Drawing
 * the actual glyphs is left to callers (which need a Canvas 2D / SkiaSharp
 * surface – something this port intentionally avoids).
 */
export interface RichTextRunModel {
  StartIndex: number;
  Length: number;
  Font?: {
    Name?: string;
    Size?: number;
    Bold?: boolean;
    Italic?: boolean;
    Underline?: number;
    StrikeThrough?: boolean;
    Color?: { R?: number; G?: number; B?: number; A?: number; ThemeIndex?: number; Tint?: number; Indexed?: number };
  };
}

export class RichTextCellRenderer {
  private _fonts: FontRegistry;
  private _makeContext: (font: { Name?: string; Size?: number; Bold?: boolean; Italic?: boolean }) => CellFontContext;
  private _resolveColor: (color?: { R?: number; G?: number; B?: number; A?: number; ThemeIndex?: number; Tint?: number; Indexed?: number }) => RgbaColor;

  constructor(
    fonts: FontRegistry,
    makeContext: (font: { Name?: string; Size?: number; Bold?: boolean; Italic?: boolean }) => CellFontContext,
    resolveColor: (color?: { R?: number; G?: number; B?: number; A?: number; ThemeIndex?: number; Tint?: number; Indexed?: number }) => RgbaColor,
  ) {
    this._fonts = fonts;
    this._makeContext = makeContext;
    this._resolveColor = resolveColor;
  }

  /** Build all font-homogeneous slices from a cell's raw text + rich-text runs. */
  BuildSlices(
    text: string,
    defaultFont: { Name?: string; Size?: number; Bold?: boolean; Italic?: boolean },
    runs: RichTextRunModel[],
  ): RichTextSlice[] {
    const slices: RichTextSlice[] = [];
    let currentIndex = 0;
    for (const run of runs) {
      if (run.StartIndex > currentIndex) {
        this.addSlice(slices, text.substring(currentIndex, run.StartIndex), defaultFont);
      }
      const runFont = mergeFont(defaultFont, run.Font);
      this.addSlice(slices, text.substring(run.StartIndex, run.StartIndex + run.Length), runFont);
      currentIndex = run.StartIndex + run.Length;
    }
    if (currentIndex < text.length) this.addSlice(slices, text.substring(currentIndex), defaultFont);
    return slices;
  }

  private addSlice(
    slices: RichTextSlice[],
    text: string,
    font: { Name?: string; Size?: number; Bold?: boolean; Italic?: boolean; Color?: { R?: number; G?: number; B?: number; A?: number; ThemeIndex?: number; Tint?: number; Indexed?: number } },
  ): void {
    if (!text) return;
    const context = this._makeContext(font);
    const color = this._resolveColor(font.Color);
    slices.push({
      Text: text,
      FontName: font.Name ?? "",
      FontSize: font.Size ?? 11,
      Bold: !!font.Bold,
      Italic: !!font.Italic,
      FontContext: context,
      WidthPt: context.Measure(text),
      AscentPt: 0,
      DescentPt: 0,
      Color: color,
    });
  }
}

function mergeFont(
  base: { Name?: string; Size?: number; Bold?: boolean; Italic?: boolean },
  rich?: { Name?: string; Size?: number; Bold?: boolean; Italic?: boolean; Color?: { R?: number; G?: number; B?: number; A?: number; ThemeIndex?: number; Tint?: number; Indexed?: number } },
): { Name?: string; Size?: number; Bold?: boolean; Italic?: boolean; Color?: { R?: number; G?: number; B?: number; A?: number; ThemeIndex?: number; Tint?: number; Indexed?: number } } {
  if (!rich) return base;
  return {
    Name: rich.Name ?? base.Name,
    Size: rich.Size ?? base.Size,
    Bold: rich.Bold ?? base.Bold,
    Italic: rich.Italic ?? base.Italic,
    Color: rich.Color,
  };
}
