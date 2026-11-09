import { BorderStyle } from "./Core/BorderStyle";
import { CellRecord } from "./Core/CellRecord";
import { CellValueKind } from "./Core/CellValueKind";
import { ColorValue } from "./Core/ColorValue";
import { FormatConditionModel } from "./Core/FormatConditionModel";
import { NumberFormatValue } from "./Core/NumberFormatValue";
import { StyleValue } from "./Core/StyleValue";

/** Minimal structural stand-in for System.Xml.Linq.XDocument, which only needs to be stored. */
export interface XDocument {
  readonly documentElement: unknown;
}

export class StylesheetSaveContext {
  private readonly _styleIndices: ReadonlyMap<string, number>;
  private readonly _differentialStyleIndices: ReadonlyMap<string, number>;
  public readonly Document: XDocument;
  public readonly DifferentialFormatCount: number;
  public readonly HasStyles: boolean;

  public constructor(
    document: XDocument,
    styleIndices: ReadonlyMap<string, number>,
    differentialStyleIndices: ReadonlyMap<string, number>,
    differentialFormatCount: number,
    hasStyles: boolean,
  ) {
    this.Document = document;
    this._styleIndices = styleIndices;
    this._differentialStyleIndices = differentialStyleIndices;
    this.DifferentialFormatCount = differentialFormatCount;
    this.HasStyles = hasStyles;
  }

  public GetStyleIndex(record: CellRecord): number;
  public GetStyleIndex(style: StyleValue | null): number;
  public GetStyleIndex(recordOrStyle: CellRecord | StyleValue | null): number {
    if (recordOrStyle == null) {
      return 0;
    }

    const style = StylesheetSaveContext.isCellRecord(recordOrStyle)
      ? getStyleForSerialization(recordOrStyle)
      : recordOrStyle;
    const index = this._styleIndices.get(getStyleKey(style));
    return index === undefined ? 0 : index;
  }

  public GetDifferentialStyleIndex(condition: FormatConditionModel): number | null {
    if (stylesEqual(condition.Style, StyleValue.Default)) {
      return null;
    }

    const index = this._differentialStyleIndices.get(getStyleKey(condition.Style));
    return index === undefined ? null : index;
  }

  private static isCellRecord(value: CellRecord | StyleValue): value is CellRecord {
    return (value as CellRecord).Kind !== undefined && (value as StyleValue).Font === undefined;
  }
}

function getStyleForSerialization(record: CellRecord): StyleValue {
  const style = record.Style == null ? StyleValue.Default.clone() : record.Style;
  if (record.Kind !== CellValueKind.DateTime || isDateNumberFormat(style.NumberFormat)) {
    return style;
  }

  const serializedStyle = style.clone();
  const numberFormat = new NumberFormatValue();
  numberFormat.Number = 14;
  numberFormat.Custom = null;
  serializedStyle.NumberFormat = numberFormat;
  return serializedStyle;
}

function isDateNumberFormat(numberFormat: NumberFormatValue): boolean {
  if (BuiltInDateFormats.has(numberFormat.Number)) {
    return true;
  }

  const customFormat = numberFormat.Custom;
  return customFormat != null && customFormat.length > 0 && looksLikeDateFormat(customFormat);
}

const BuiltInDateFormats: ReadonlySet<number> = new Set<number>([
  14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47,
]);

function looksLikeDateFormat(formatCode: string): boolean {
  if (formatCode == null || formatCode.trim().length === 0) {
    return false;
  }

  const lower = formatCode.toLowerCase();
  if (lower.includes("[$-f400]") || lower.includes("[$-f800]")) {
    return true;
  }

  let builder = "";
  let inQuote = false;
  let inBracket = false;

  for (let index = 0; index < formatCode.length; index++) {
    const character = formatCode[index]!;
    if (character === '"') {
      inQuote = !inQuote;
      continue;
    }

    if (inQuote) {
      continue;
    }

    if (character === "[") {
      inBracket = true;
      continue;
    }

    if (character === "]" && inBracket) {
      inBracket = false;
      continue;
    }

    if (inBracket) {
      builder += character.toLowerCase();
      continue;
    }

    if (character === "\\" || character === "_") {
      index++;
      continue;
    }

    if (character === "*") {
      continue;
    }

    builder += character.toLowerCase();
  }

  const normalized = builder;
  return (
    normalized.includes("yy") ||
    normalized.includes("aaa") ||
    (normalized.includes("m") && formatCode.includes("\u6708")) ||
    (normalized.includes("d") && formatCode.includes("\u65e5")) ||
    normalized.includes("yyyy") ||
    normalized.includes("dd") ||
    normalized.includes("ddd") ||
    normalized.includes("mm") ||
    normalized.includes("mmm") ||
    normalized.includes("m/") ||
    normalized.includes("/m") ||
    normalized.includes("h") ||
    normalized.includes("am/pm") ||
    normalized.includes("a/p") ||
    normalized.includes("ss")
  );
}

function getStyleKey(style: StyleValue): string {
  return (
    getFontKey(style) +
    "|" +
    (style.Pattern | 0) +
    "|" +
    getColorKey(style.ForegroundColor) +
    "|" +
    getColorKey(style.BackgroundColor) +
    "|" +
    getBordersKey(style) +
    "|" +
    style.NumberFormat.Number +
    "|" +
    (style.NumberFormat.Custom ?? "") +
    "|" +
    (style.Alignment.Horizontal | 0) +
    "|" +
    (style.Alignment.Vertical | 0) +
    "|" +
    (style.Alignment.WrapText ? "1" : "0") +
    "|" +
    style.Alignment.IndentLevel.toString() +
    "|" +
    style.Alignment.TextRotation.toString() +
    "|" +
    (style.Alignment.ShrinkToFit ? "1" : "0") +
    "|" +
    style.Alignment.ReadingOrder.toString() +
    "|" +
    style.Alignment.RelativeIndent.toString() +
    "|" +
    (style.Protection.IsLocked ? "1" : "0") +
    "|" +
    (style.Protection.IsHidden ? "1" : "0")
  );
}

function getFontKey(style: StyleValue): string {
  const font = style.Font;
  return (
    (font.Name ?? "") +
    "|" +
    formatSize(font.Size) +
    "|" +
    (font.Bold ? "1" : "0") +
    "|" +
    (font.Italic ? "1" : "0") +
    "|" +
    (font.Underline | 0) +
    "|" +
    (font.StrikeThrough ? "1" : "0") +
    "|" +
    getColorKey(font.Color)
  );
}

function getBordersKey(style: StyleValue): string {
  const borders = style.Borders;
  return (
    getBorderSideKey(borders.Left.Style, borders.Left.Color) +
    "|" +
    getBorderSideKey(borders.Right.Style, borders.Right.Color) +
    "|" +
    getBorderSideKey(borders.Top.Style, borders.Top.Color) +
    "|" +
    getBorderSideKey(borders.Bottom.Style, borders.Bottom.Color) +
    "|" +
    getBorderSideKey(borders.Diagonal.Style, borders.Diagonal.Color) +
    "|" +
    (borders.DiagonalUp ? "1" : "0") +
    "|" +
    (borders.DiagonalDown ? "1" : "0")
  );
}

function getBorderSideKey(style: BorderStyle, color: ColorValue): string {
  return (style | 0) + "|" + getColorKey(color);
}

function getColorKey(color: ColorValue): string {
  if (color.ThemeIndex !== null) {
    return "T" + color.ThemeIndex.toString() + (color.Tint !== null ? "_" + color.Tint.toString() : "");
  }

  if (color.Indexed !== null) {
    return "I" + color.Indexed.toString();
  }

  return formatHex(color.A) + formatHex(color.R) + formatHex(color.G) + formatHex(color.B);
}

function stylesEqual(left: StyleValue, right: StyleValue): boolean {
  return getStyleKey(left) === getStyleKey(right);
}

/** Mirrors the invariant "0.####" number formatting used for style keys. */
function formatSize(value: number): string {
  const text = value.toFixed(4);
  const trimmed = text.replace(/0+$/, "").replace(/\.$/, "");
  return trimmed === "-0" ? "0" : trimmed;
}

function formatHex(value: number): string {
  return (value & 0xff).toString(16).toUpperCase().padStart(2, "0");
}
