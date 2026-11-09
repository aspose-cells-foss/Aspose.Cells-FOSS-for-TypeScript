import { Color } from "../../aspose_cells";
import { attribute, XmlNode } from "../../aspose_cells/Xml/XmlReader";

/** Shared value normalization helpers ported from ComparisonValueHelpers.cs. */

export const DefaultColorHex = "00000000";

/** Built-in number format ids that represent date/time values. */
export const BuiltInDateFormats = new Set<number>([
  14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47,
]);

const IntMin = -2147483648;
const IntMax = 2147483647;

/** Returns the normalized A1 reference, or null when the reference is invalid. */
export function tryNormalizeCellReference(cellReference: string | null | undefined): string | null {
  if (cellReference === null || cellReference === undefined || cellReference.trim().length === 0) {
    return null;
  }

  const trimmed = cellReference.trim().toUpperCase();
  let splitIndex = 0;
  while (
    splitIndex < trimmed.length &&
    trimmed.charCodeAt(splitIndex) >= 65 &&
    trimmed.charCodeAt(splitIndex) <= 90
  ) {
    splitIndex++;
  }

  if (splitIndex === 0 || splitIndex === trimmed.length) {
    return null;
  }

  for (let index = splitIndex; index < trimmed.length; index++) {
    const code = trimmed.charCodeAt(index);
    if (code < 48 || code > 57) {
      return null;
    }
  }

  const row = Number.parseInt(trimmed.substring(splitIndex), 10);
  if (Number.isNaN(row) || row < 1) {
    return null;
  }

  return trimmed;
}

/** Builds the dictionary key for a sheet/cell pair. */
export function buildCellKey(sheetName: string, cellReference: string): string {
  return sheetName + "!" + cellReference.toUpperCase();
}

/** Normalizes a cell value to the stable text used by the comparison. */
export function normalizeValue(value: unknown): string {
  if (value === null || value === undefined) {
    return "<null>";
  }

  if (typeof value === "string") {
    return value;
  }

  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }

  if (value instanceof Date) {
    return value.toISOString();
  }

  if (typeof value === "number") {
    return Number.isFinite(value) ? value.toString() : String(value);
  }

  return String(value);
}

/** Normalizes a formula body so that a leading equals sign is always present. */
export function normalizeFormulaText(formulaBody: string | null | undefined): string {
  if (formulaBody === null || formulaBody === undefined || formulaBody.trim().length === 0) {
    return "";
  }

  const normalized = formulaBody.trim();
  return normalized.startsWith("=") ? normalized : "=" + normalized;
}

/** Formats a double with up to four decimal digits, mirroring "0.####". */
export function normalizeDouble(value: number): string {
  const fixed = value.toFixed(4);
  return fixed.indexOf(".") >= 0 ? fixed.replace(/\.?0+$/, "") : fixed;
}

function hexByte(value: number): string {
  return (value & 0xff).toString(16).toUpperCase().padStart(2, "0");
}

/** Normalizes a library color to an eight digit ARGB hex string. */
export function normalizeColor(color: Color): string {
  if (color.Equals(Color.Empty)) {
    return DefaultColorHex;
  }

  return hexByte(color.A) + hexByte(color.R) + hexByte(color.G) + hexByte(color.B);
}

/** Normalizes an Open XML color element to an eight digit ARGB hex string. */
export function normalizeOpenXmlColor(colorElement: XmlNode | null): string {
  if (colorElement === null) {
    return DefaultColorHex;
  }

  const rgb = attribute(colorElement, "rgb");
  if (rgb === null || rgb.trim().length === 0) {
    return DefaultColorHex;
  }

  let normalized = rgb.trim().toUpperCase();
  if (normalized.length === 6) {
    normalized = "FF" + normalized;
  }

  return normalized.length === 8 ? normalized : DefaultColorHex;
}

/** Maps an Open XML horizontal alignment token to the library enum name. */
export function parseHorizontalAlignment(value: string | null): string {
  switch ((value ?? "").toLowerCase()) {
    case "left":
      return "Left";
    case "center":
      return "Center";
    case "right":
      return "Right";
    case "fill":
      return "Fill";
    case "justify":
      return "Justify";
    case "centercontinuous":
      return "CenterContinuous";
    case "distributed":
      return "Distributed";
    default:
      return "General";
  }
}

/** Maps an Open XML vertical alignment token to the library enum name. */
export function parseVerticalAlignment(value: string | null): string {
  switch ((value ?? "").toLowerCase()) {
    case "center":
      return "Center";
    case "top":
      return "Top";
    case "justify":
      return "Justify";
    case "distributed":
      return "Distributed";
    default:
      return "Bottom";
  }
}

/** Maps an Open XML border style token to the library enum name. */
export function parseBorderStyle(value: string | null): string {
  switch ((value ?? "").toLowerCase()) {
    case "thin":
      return "Thin";
    case "medium":
      return "Medium";
    case "thick":
      return "Thick";
    case "dotted":
      return "Dotted";
    case "dashed":
      return "Dashed";
    case "double":
      return "Double";
    case "hair":
      return "Hair";
    case "mediumdashed":
      return "MediumDashed";
    case "dashdot":
      return "DashDot";
    case "mediumdashdot":
      return "MediumDashDot";
    case "dashdotdot":
      return "DashDotDot";
    case "mediumdashdotdot":
      return "MediumDashDotDot";
    case "slantdashdot":
      return "SlantedDashDot";
    default:
      return "None";
  }
}

/** Maps an Open XML fill pattern token to the library enum name. */
export function parseFillPattern(value: string | null): string {
  switch ((value ?? "").toLowerCase()) {
    case "solid":
      return "Solid";
    case "mediumgray":
      return "MediumGray";
    case "darkgray":
      return "DarkGray";
    case "lightgray":
      return "LightGray";
    case "gray125":
      return "Gray125";
    case "gray0625":
      return "Gray0625";
    case "darkhorizontal":
      return "DarkHorizontal";
    case "darkvertical":
      return "DarkVertical";
    case "darkdown":
      return "DarkDown";
    case "darkup":
      return "DarkUp";
    case "darkgrid":
      return "DarkGrid";
    case "darktrellis":
      return "DarkTrellis";
    case "lighthorizontal":
      return "LightHorizontal";
    case "lightvertical":
      return "LightVertical";
    case "lightdown":
      return "LightDown";
    case "lightup":
      return "LightUp";
    case "lightgrid":
      return "LightGrid";
    case "lighttrellis":
      return "LightTrellis";
    default:
      return "None";
  }
}

/** Parses an Open XML boolean attribute value. */
export function parseBoolAttribute(value: string | null): boolean {
  if (value === null) {
    return false;
  }

  return value === "1" || value.toLowerCase() === "true";
}

/** Parses an optional Open XML boolean attribute value. */
export function parseOptionalBoolAttribute(value: string | null): boolean | null {
  return value === null ? null : parseBoolAttribute(value);
}

/** Parses an Open XML integer attribute value using strict integer syntax. */
export function parseIntAttribute(value: string | null): number | null {
  if (value === null) {
    return null;
  }

  const trimmed = value.trim();
  if (!/^[+-]?\d+$/.test(trimmed)) {
    return null;
  }

  const parsed = Number.parseInt(trimmed, 10);
  return Number.isNaN(parsed) ? null : parsed;
}

/** Parses a floating point value, returning the default when parsing fails. */
export function parseDoubleValue(value: string | null, defaultValue: number): number {
  if (value === null || value.trim().length === 0) {
    return defaultValue;
  }

  const parsed = Number(value.trim());
  return Number.isNaN(parsed) ? defaultValue : parsed;
}

export interface ParsedNumber {
  readonly value: unknown;
  readonly kind: string;
}

/** Parses a numeric cell value, preferring integer results within the Int32 range. */
export function tryParseNumber(rawValue: string | null): ParsedNumber | null {
  if (rawValue === null || rawValue.trim().length === 0) {
    return null;
  }

  const text = rawValue;
  if (text.indexOf("e") < 0 && text.indexOf("E") < 0) {
    const parsed = Number(text.trim());
    if (!Number.isNaN(parsed)) {
      if (
        text.indexOf(".") < 0 &&
        text.indexOf(",") < 0 &&
        Number.isInteger(parsed) &&
        parsed >= IntMin &&
        parsed <= IntMax
      ) {
        return { value: parsed, kind: "Number" };
      }

      return { value: parsed, kind: "Number" };
    }
  }

  const doubleValue = Number(text.trim());
  if (!Number.isNaN(doubleValue)) {
    if (Math.abs(doubleValue % 1) < Number.EPSILON && doubleValue >= IntMin && doubleValue <= IntMax) {
      return { value: doubleValue, kind: "Number" };
    }

    return { value: doubleValue, kind: "Number" };
  }

  return null;
}

/** Detects whether a number format code represents a date or time value. */
export function looksLikeDateFormat(formatCode: string | null | undefined): boolean {
  if (formatCode === null || formatCode === undefined || formatCode.trim().length === 0) {
    return false;
  }

  let normalized = "";
  let inQuote = false;
  let inBracket = false;

  for (let index = 0; index < formatCode.length; index++) {
    const character = formatCode.charAt(index);
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
      normalized += character.toLowerCase();
      continue;
    }

    if (character === "\\" || character === "_") {
      index++;
      continue;
    }

    if (character === "*") {
      continue;
    }

    normalized += character.toLowerCase();
  }

  return (
    normalized.indexOf("yy") >= 0 ||
    normalized.indexOf("dd") >= 0 ||
    normalized.indexOf("mm") >= 0 ||
    normalized.indexOf("m/") >= 0 ||
    normalized.indexOf("/m") >= 0 ||
    normalized.indexOf("h") >= 0 ||
    normalized.indexOf("ss") >= 0
  );
}

/** Returns the value at the index, or the fallback when the index is out of range. */
export function getOrDefault<T>(values: readonly T[], index: number, fallback: T): T {
  return index >= 0 && index < values.length ? values[index] : fallback;
}
