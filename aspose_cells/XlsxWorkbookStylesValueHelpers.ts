import { AlignmentValue } from "./Core/AlignmentValue";
import { BorderSideValue } from "./Core/BorderSideValue";
import { BorderStyle } from "./Core/BorderStyle";
import { BordersValue } from "./Core/BordersValue";
import { ColorValue } from "./Core/ColorValue";
import { FillPatternKind } from "./Core/FillPatternKind";
import { FontValue } from "./Core/FontValue";
import { HorizontalAlignment } from "./Core/HorizontalAlignment";
import { ProtectionValue } from "./Core/ProtectionValue";
import { VerticalAlignment } from "./Core/VerticalAlignment";

/** Fill representation used by the stylesheet writer. Mirrors the .NET FillValue type. */
export class FillValue {
  public Pattern: FillPatternKind = FillPatternKind.None;
  public ForegroundColor: ColorValue = new ColorValue();
  public BackgroundColor: ColorValue = new ColorValue();
}

/** Cell format entry used by the stylesheet writer. Mirrors the .NET CellFormatValue type. */
export class CellFormatValue {
  public NumFmtId: number = 0;
  public FontId: number = 0;
  public FillId: number = 0;
  public BorderId: number = 0;
  public Alignment: AlignmentValue = new AlignmentValue();
  public Protection: ProtectionValue = new ProtectionValue();
}

/** Formats a double like the invariant "0.####" format specifier (at most four decimals, no trailing zeros). */
export function formatNumber4(value: number): string {
  if (!Number.isFinite(value)) {
    return "0";
  }

  const text = value.toFixed(4);
  const trimmed = text.indexOf(".") >= 0 ? text.replace(/0+$/, "").replace(/\.$/, "") : text;
  return trimmed === "-0" ? "0" : trimmed;
}

/** Formats a double like the invariant "R" round-trip format specifier. */
export function formatDoubleR(value: number): string {
  if (!Number.isFinite(value)) {
    return "0";
  }

  return String(value);
}

/** Formats a single ARGB component as a two-digit uppercase hexadecimal value. */
export function formatHexByte(value: number): string {
  return (value & 0xff).toString(16).toUpperCase().padStart(2, "0");
}

export function toArgbHex(color: ColorValue): string {
  return formatHexByte(color.A) + formatHexByte(color.R) + formatHexByte(color.G) + formatHexByte(color.B);
}

export function isEmptyColor(color: ColorValue): boolean {
  return (
    color.ThemeIndex === null &&
    color.Indexed === null &&
    color.A === 0 &&
    color.R === 0 &&
    color.G === 0 &&
    color.B === 0
  );
}

export function fontEquals(left: FontValue, right: FontValue): boolean {
  return (
    left.Name === right.Name &&
    left.Size === right.Size &&
    left.Bold === right.Bold &&
    left.Italic === right.Italic &&
    left.Underline === right.Underline &&
    left.StrikeThrough === right.StrikeThrough &&
    left.Color.Equals(right.Color) &&
    left.Family === right.Family &&
    left.Scheme === right.Scheme
  );
}

export function bordersEqual(left: BordersValue, right: BordersValue): boolean {
  return (
    borderSideEquals(left.Left, right.Left) &&
    borderSideEquals(left.Right, right.Right) &&
    borderSideEquals(left.Top, right.Top) &&
    borderSideEquals(left.Bottom, right.Bottom) &&
    borderSideEquals(left.Diagonal, right.Diagonal) &&
    left.DiagonalUp === right.DiagonalUp &&
    left.DiagonalDown === right.DiagonalDown
  );
}

function borderSideEquals(left: BorderSideValue, right: BorderSideValue): boolean {
  return left.Style === right.Style && left.Color.Equals(right.Color);
}

export function parseBorderStyle(value: string | null | undefined): BorderStyle {
  switch ((value ?? "").toLowerCase()) {
    case "thin":
      return BorderStyle.Thin;
    case "medium":
      return BorderStyle.Medium;
    case "thick":
      return BorderStyle.Thick;
    case "dotted":
      return BorderStyle.Dotted;
    case "dashed":
      return BorderStyle.Dashed;
    case "double":
      return BorderStyle.Double;
    case "hair":
      return BorderStyle.Hair;
    case "mediumdashed":
      return BorderStyle.MediumDashed;
    case "dashdot":
      return BorderStyle.DashDot;
    case "mediumdashdot":
      return BorderStyle.MediumDashDot;
    case "dashdotdot":
      return BorderStyle.DashDotDot;
    case "mediumdashdotdot":
      return BorderStyle.MediumDashDotDot;
    case "slantdashdot":
      return BorderStyle.SlantedDashDot;
    default:
      return BorderStyle.None;
  }
}

export function getBorderStyleName(value: BorderStyle): string {
  switch (value) {
    case BorderStyle.Thin:
      return "thin";
    case BorderStyle.Medium:
      return "medium";
    case BorderStyle.Thick:
      return "thick";
    case BorderStyle.Dotted:
      return "dotted";
    case BorderStyle.Dashed:
      return "dashed";
    case BorderStyle.Double:
      return "double";
    case BorderStyle.Hair:
      return "hair";
    case BorderStyle.MediumDashed:
      return "mediumDashed";
    case BorderStyle.DashDot:
      return "dashDot";
    case BorderStyle.MediumDashDot:
      return "mediumDashDot";
    case BorderStyle.DashDotDot:
      return "dashDotDot";
    case BorderStyle.MediumDashDotDot:
      return "mediumDashDotDot";
    case BorderStyle.SlantedDashDot:
      return "slantDashDot";
    default:
      return "";
  }
}

export function getHorizontalAlignmentName(value: HorizontalAlignment): string {
  switch (value) {
    case HorizontalAlignment.Left:
      return "left";
    case HorizontalAlignment.Center:
      return "center";
    case HorizontalAlignment.Right:
      return "right";
    case HorizontalAlignment.Fill:
      return "fill";
    case HorizontalAlignment.Justify:
      return "justify";
    case HorizontalAlignment.CenterContinuous:
      return "centerContinuous";
    case HorizontalAlignment.Distributed:
      return "distributed";
    default:
      return "";
  }
}

export function getVerticalAlignmentName(value: VerticalAlignment): string {
  switch (value) {
    case VerticalAlignment.Center:
      return "center";
    case VerticalAlignment.Top:
      return "top";
    case VerticalAlignment.Justify:
      return "justify";
    case VerticalAlignment.Distributed:
      return "distributed";
    default:
      return "";
  }
}
