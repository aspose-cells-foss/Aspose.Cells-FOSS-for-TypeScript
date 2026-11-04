import { RgbaColor } from "./ColorValue";

/** One per cell, collected by {@link ConditionalFormatEvaluator}. */
export interface DataBarEffect {
  Fraction: number;
  Color: RgbaColor;
  ShowBorder?: boolean;
  Gradient?: boolean;
}

export interface IconEffect {
  Color: RgbaColor;
  BorderColor?: RgbaColor;
  HasBorder?: boolean;
  AngleDeg?: number;
  IsArrow?: boolean;
}

export class CellEffect {
  HasFill = false;
  Fill: RgbaColor | null = null;
  HasFontColor = false;
  FontColor: RgbaColor | null = null;
  HasFontBold = false;
  HasFontItalic = false;
  HasFontUnderline = false;
  Bar: DataBarEffect | null = null;
  IconGlyph: IconEffect | null = null;
  SuppressText = false;
}

interface FormatCondition {
  Type?: string;
  Priority?: number;
  Operator?: string;
  Formula1?: string;
  Formula2?: string;
  Rank?: number;
  Percent?: boolean;
  Top?: boolean;
  StandardDeviation?: number;
  Above?: boolean;
  TimePeriod?: string;
  ColorScaleCount?: number;
  MinColor?: RgbaColor;
  MidColor?: RgbaColor;
  MaxColor?: RgbaColor;
  BarColor?: RgbaColor;
  DataBarAutoMinimum?: boolean;
  DataBarMinLength?: number;
  DataBarMaxLength?: number;
  DataBarGradient?: boolean;
  ShowBorder?: boolean;
  IconSetType?: string;
  ShowIconOnly?: boolean;
  ReverseIcons?: boolean;
  Style?: {
    Pattern?: string;
    ForegroundColor?: RgbaColor;
    BackgroundColor?: RgbaColor;
    Font?: { Bold?: boolean; Italic?: boolean; Underline?: boolean; Color?: RgbaColor };
  };
}

interface ConditionalFormatting {
  Conditions: FormatCondition[];
  Areas: { StartRow: number; EndRow: number; StartColumn: number; EndColumn: number }[];
}

export class ConditionalFormatEvaluator {
  private _effects = new Map<string, CellEffect>();

  constructor(formattings: ConditionalFormatting[] | undefined | null, _evaluationDate: Date) {
    if (!formattings || formattings.length === 0) return;
    // Higher priority (lower number) wins.
    const ordered = [...formattings].sort(
      (a, b) => (a.Conditions[0]?.Priority ?? 0) - (b.Conditions[0]?.Priority ?? 0),
    );
    for (const formatting of ordered) {
      for (const condition of formatting.Conditions) {
        for (const area of formatting.Areas) {
          this.applyDifferentialArea(condition, area);
        }
      }
    }
  }

  private effectFor(row: number, col: number): CellEffect {
    const k = key(row, col);
    const existing = this._effects.get(k);
    if (existing) return existing;
    const fresh = new CellEffect();
    this._effects.set(k, fresh);
    return fresh;
  }

  private applyDifferentialArea(
    condition: FormatCondition,
    area: { StartRow: number; EndRow: number; StartColumn: number; EndColumn: number },
  ): void {
    // Without workbook data we can't run the matcher, but we still record a
    // no-op CellEffect slot so callers know the CF area exists.  Actual
    // matching is the caller's responsibility when the workbook model is
    // wired in.
    void condition;
    for (let r = area.StartRow; r <= area.EndRow; r++) {
      for (let c = area.StartColumn; c <= area.EndColumn; c++) {
        this.effectFor(r, c);
      }
    }
  }

  TryGet(row: number, col: number): CellEffect | null {
    return this._effects.get(key(row, col)) ?? null;
  }
}

function key(row: number, col: number): string {
  return `${row}::${col}`;
}
