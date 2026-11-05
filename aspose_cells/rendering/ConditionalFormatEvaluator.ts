import { RgbaColor } from "./ColorValue";

/**
 * Supplies numeric cell values to the evaluator. Returning null/undefined means
 * "no numeric value" (empty, text, boolean, error). The evaluator has no access
 * to the worksheet, so callers must inject this callback.
 */
export type ValueProvider = (row: number, col: number) => number | null | undefined;

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

interface CellArea {
  StartRow: number;
  EndRow: number;
  StartColumn: number;
  EndColumn: number;
}

const ArrowGreen = new RgbaColor(0x63, 0xa5, 0x37);
const ArrowAmber = new RgbaColor(0xe8, 0xa8, 0x38);
const ArrowRed = new RgbaColor(0xdc, 0x3a, 0x2e);
const TrafficRed = new RgbaColor(214, 85, 50);
const TrafficYellow = new RgbaColor(234, 194, 130);
const TrafficGreen = new RgbaColor(104, 164, 144);
const TrafficRedBorder = new RgbaColor(161, 63, 36);
const TrafficYellowBorder = new RgbaColor(167, 132, 51);
const TrafficGreenBorder = new RgbaColor(58, 116, 97);

export class ConditionalFormatEvaluator {
  private _effects = new Map<string, CellEffect>();
  private readonly _valueProvider: ValueProvider | null;

  constructor(
    formattings: ConditionalFormatting[] | undefined | null,
    _evaluationDate: Date,
    valueProvider: ValueProvider | null = null,
  ) {
    this._valueProvider = valueProvider;
    if (!formattings || formattings.length === 0) return;
    // Higher priority (lower number) wins.
    const ordered = [...formattings].sort(
      (a, b) => (a.Conditions[0]?.Priority ?? 0) - (b.Conditions[0]?.Priority ?? 0),
    );
    for (const formatting of ordered) {
      for (const condition of formatting.Conditions) {
        for (const area of formatting.Areas) {
          this.applyConditionToArea(condition, area);
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

  private applyConditionToArea(condition: FormatCondition, area: CellArea): void {
    const type = (condition.Type ?? "").toLowerCase();
    if (type === "colorscale") {
      this.applyColorScale(condition, area);
      return;
    }
    if (type === "databar") {
      this.applyDataBar(condition, area);
      return;
    }
    if (type === "iconset") {
      this.applyIconSet(condition, area);
      return;
    }

    // Without workbook data we can't run the matcher, but we still record a
    // no-op CellEffect slot so callers know the CF area exists.  Actual
    // matching is the caller's responsibility when the workbook model is
    // wired in.
    this.fillEmptySlots(area);
  }

  private fillEmptySlots(area: CellArea): void {
    for (let r = area.StartRow; r <= area.EndRow; r++) {
      for (let c = area.StartColumn; c <= area.EndColumn; c++) {
        this.effectFor(r, c);
      }
    }
  }

  /** Returns the injected numeric value, or null when absent/non-finite. */
  private valueAt(row: number, col: number): number | null {
    if (!this._valueProvider) return null;
    const value = this._valueProvider(row, col);
    if (value === null || value === undefined) return null;
    return Number.isFinite(value) ? value : null;
  }

  private collectNumbers(area: CellArea): number[] {
    const values: number[] = [];
    for (let r = area.StartRow; r <= area.EndRow; r++) {
      for (let c = area.StartColumn; c <= area.EndColumn; c++) {
        const x = this.valueAt(r, c);
        if (x !== null) values.push(x);
      }
    }
    return values;
  }

  // --- Color scale / data bar / icon set -------------------------------------------------------

  private applyColorScale(condition: FormatCondition, area: CellArea): void {
    const values = this.collectNumbers(area);
    if (values.length === 0) {
      this.fillEmptySlots(area);
      return;
    }

    values.sort((a, b) => a - b);
    const min = values[0];
    const max = values[values.length - 1];
    const threeColor =
      (condition.ColorScaleCount ?? 0) >= 3 && (condition.MidColor?.a ?? 0) !== 0;
    const mid = threeColor ? percentile(values, 50) : (min + max) / 2.0;

    const minColor = toRgba(condition.MinColor, new RgbaColor(0xf8, 0x69, 0x6b));
    const midColor = toRgba(condition.MidColor, new RgbaColor(0xff, 0xeb, 0x84));
    const maxColor = toRgba(condition.MaxColor, new RgbaColor(0x63, 0xbe, 0x7b));

    for (let r = area.StartRow; r <= area.EndRow; r++) {
      for (let c = area.StartColumn; c <= area.EndColumn; c++) {
        const x = this.valueAt(r, c);
        if (x === null) continue;

        let color: RgbaColor;
        if (!threeColor) {
          color = lerp(minColor, maxColor, fraction(x, min, max));
        } else if (x <= mid) {
          color = lerp(minColor, midColor, fraction(x, min, mid));
        } else {
          color = lerp(midColor, maxColor, fraction(x, mid, max));
        }

        const effect = this.effectFor(r, c);
        effect.HasFill = true;
        effect.Fill = color;
      }
    }
  }

  private applyDataBar(condition: FormatCondition, area: CellArea): void {
    const values = this.collectNumbers(area);
    if (values.length === 0) {
      this.fillEmptySlots(area);
      return;
    }

    values.sort((a, b) => a - b);
    let min = values[0];
    const max = values[values.length - 1];
    if (condition.DataBarAutoMinimum && min >= 0) {
      min = 0;
    }
    const color = toRgba(condition.BarColor, new RgbaColor(0x63, 0x8e, 0xc6));
    const minLength = condition.DataBarMinLength ?? 0.1;
    const maxLength = condition.DataBarMaxLength ?? 0.9;

    for (let r = area.StartRow; r <= area.EndRow; r++) {
      for (let c = area.StartColumn; c <= area.EndColumn; c++) {
        const x = this.valueAt(r, c);
        if (x === null) continue;

        const display = minLength + (maxLength - minLength) * fraction(x, min, max);
        const effect = this.effectFor(r, c);
        effect.Bar = {
          Fraction: display,
          Color: color,
          ShowBorder: condition.ShowBorder ?? false,
          Gradient: condition.DataBarGradient ?? true,
        };
      }
    }
  }

  private applyIconSet(condition: FormatCondition, area: CellArea): void {
    const values = this.collectNumbers(area);
    if (values.length === 0) {
      this.fillEmptySlots(area);
      return;
    }

    values.sort((a, b) => a - b);
    const min = values[0];
    const max = values[values.length - 1];
    const iconCount = iconCountOf(condition.IconSetType);

    for (let r = area.StartRow; r <= area.EndRow; r++) {
      for (let c = area.StartColumn; c <= area.EndColumn; c++) {
        const x = this.valueAt(r, c);
        if (x === null) continue;

        // OOXML's default percent thresholds are integer values (0, 33, 66 for three icons).
        const percent = fraction(x, min, max) * 100.0;
        let bucket = 0;
        for (let index = 1; index < iconCount; index++) {
          if (percent >= (100 * index) / iconCount) {
            bucket = index;
          }
        }

        const lookup = condition.ReverseIcons ? iconCount - 1 - bucket : bucket;
        const effect = this.effectFor(r, c);
        effect.IconGlyph = iconFor(condition.IconSetType, lookup, iconCount);
        if (condition.ShowIconOnly) {
          effect.SuppressText = true;
        }
      }
    }
  }

  TryGet(row: number, col: number): CellEffect | null {
    return this._effects.get(key(row, col)) ?? null;
  }
}

// --- Value / numeric helpers -------------------------------------------------------------------

function fraction(value: number, min: number, max: number): number {
  if (max <= min) return 0;
  const f = (value - min) / (max - min);
  if (f < 0) return 0;
  if (f > 1) return 1;
  return f;
}

function percentile(sorted: number[], percent: number): number {
  if (sorted.length === 1) return sorted[0];

  const rank = (percent / 100.0) * (sorted.length - 1);
  const lo = Math.floor(rank);
  const hi = Math.ceil(rank);
  if (lo === hi) return sorted[lo];

  return sorted[lo] + (rank - lo) * (sorted[hi] - sorted[lo]);
}

function toRgba(color: RgbaColor | undefined, fallback: RgbaColor): RgbaColor {
  // An unset color (alpha 0 without a theme reference in the .NET model) uses the fallback.
  if (!color) return fallback;
  if (color.a === 0) return fallback;
  return color;
}

function lerp(a: RgbaColor, b: RgbaColor, t: number): RgbaColor {
  const channel = (x: number, y: number): number => (x + Math.trunc((y - x) * t)) & 0xff;
  return new RgbaColor(channel(a.r, b.r), channel(a.g, b.g), channel(a.b, b.b));
}

function iconCountOf(iconSetType: string | undefined): number {
  if (!iconSetType) return 3;
  if (iconSetType[0] === "4") return 4;
  if (iconSetType[0] === "5") return 5;
  return 3;
}

function iconFor(iconSetType: string | undefined, index: number, count: number): IconEffect {
  const type = (iconSetType ?? "").toLowerCase();
  const isArrow = type.includes("arrow");
  const isTrafficLight = type.includes("trafficlight");

  // Colour ramp low->high: red, amber, green (amber fills the middle buckets).
  let color: RgbaColor;
  if (index === 0) color = ArrowRed;
  else if (index === count - 1) color = ArrowGreen;
  else color = ArrowAmber;

  let borderColor = RgbaColor.transparent;
  if (isTrafficLight) {
    if (index === 0) {
      color = TrafficRed;
      borderColor = TrafficRedBorder;
    } else if (index === count - 1) {
      color = TrafficGreen;
      borderColor = TrafficGreenBorder;
    } else {
      color = TrafficYellow;
      borderColor = TrafficYellowBorder;
    }
  }

  // Arrow direction sweeps from down (lowest) to up (highest) across the buckets.
  const angle = count <= 1 ? 0 : -90 + (180 * index) / (count - 1);
  return { Color: color, BorderColor: borderColor, HasBorder: isTrafficLight, AngleDeg: angle, IsArrow: isArrow };
}

function key(row: number, col: number): string {
  return `${row}::${col}`;
}
