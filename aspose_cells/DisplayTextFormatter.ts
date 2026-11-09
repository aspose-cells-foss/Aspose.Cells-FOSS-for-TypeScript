import { DateSerialConverter } from "./Core/DateSerialConverter";
import { DateSystem } from "./Core/DateSystem";
import { StyleValue } from "./Core/StyleValue";
import { NumberFormat } from "./NumberFormat";

/**
 * Simple port of DisplayTextFormatter.
 *
 * FormatStringValue returns the stable raw representation of a logical cell value.
 * FormatDisplayValue applies the style number format for the supported scalar kinds.
 */
export class DisplayText {
  /** Returns the invariant raw text for a cell value. */
  public static FormatStringValue(value: unknown): string {
    if (value === null || value === undefined) {
      return "";
    }

    if (typeof value === "string") {
      return value;
    }

    if (typeof value === "boolean") {
      return value ? "TRUE" : "FALSE";
    }

    if (value instanceof Date) {
      return DisplayText.formatRawDateTime(value);
    }

    if (typeof value === "number") {
      return String(value);
    }

    return String(value);
  }

  /** Returns the style-aware display text for a cell value. */
  public static FormatDisplayValue(
    value: unknown,
    style: StyleValue,
    _culture: { Name: string } | null,
    dateSystem: DateSystem,
  ): string {
    if (value === null || value === undefined) {
      return "";
    }

    if (typeof value === "string") {
      return value;
    }

    if (typeof value === "boolean") {
      return value ? "TRUE" : "FALSE";
    }

    const formatCode = NumberFormat.ResolveFormatCode(style.NumberFormat.Number, style.NumberFormat.Custom);

    if (value instanceof Date) {
      if (!formatCode || formatCode === "General") {
        return DisplayText.formatRawDateTime(value);
      }

      return DisplayText.formatNumeric(DateSerialConverter.ToSerial(value, dateSystem), formatCode);
    }

    if (typeof value === "number") {
      if (!formatCode || formatCode === "General") {
        return DisplayText.formatGeneral(value);
      }

      return DisplayText.formatNumeric(value, formatCode);
    }

    return String(value);
  }

  private static formatGeneral(value: number): string {
    if (!Number.isFinite(value)) {
      return String(value);
    }

    // Excel's General format rounds to 15 significant digits.
    return Number(value.toPrecision(15)).toString();
  }

  private static formatRawDateTime(value: Date): string {
    const month = value.getMonth() + 1;
    const day = value.getDate();
    const year = value.getFullYear();
    const datePart = `${month}/${day}/${year}`;

    if (value.getHours() === 0 && value.getMinutes() === 0 && value.getSeconds() === 0) {
      return datePart;
    }

    return `${datePart} ${value.getHours()}:${DisplayText.pad(value.getMinutes())}`;
  }

  private static pad(value: number): string {
    return value < 10 ? `0${value}` : String(value);
  }

  private static formatNumeric(value: number, formatCode: string): string {
    const sections = DisplayText.splitSections(formatCode);
    const section = DisplayText.selectSection(sections, value);
    if (section === null) {
      return DisplayText.formatGeneral(value);
    }

    const isNegativeSection = value < 0 && sections.length > 1 && section === sections[1];
    const numeric = isNegativeSection ? Math.abs(value) : value;

    let percentCount = 0;
    for (const character of section) {
      if (character === "%") {
        percentCount++;
      }
    }

    const plain = section
      .replace(/\[[^\]]*\]/g, "")
      .replace(/"[^"]*"/g, "")
      .replace(/[\\_*]./g, "");

    const hasPlaceholder = /[0#?]/.test(plain);
    if (!hasPlaceholder) {
      const literal = section.replace(/\[[^\]]*\]/g, "").replace(/"/g, "").trim();
      return literal.length > 0 ? literal : DisplayText.formatGeneral(value);
    }

    const dotIndex = plain.indexOf(".");
    let decimals = 0;
    if (dotIndex >= 0) {
      for (let index = dotIndex + 1; index < plain.length; index++) {
        const character = plain[index];
        if (character === "0" || character === "#") {
          decimals++;
        } else {
          break;
        }
      }
    }

    const useGrouping = plain.indexOf(",") >= 0;
    let scaled = numeric * Math.pow(100, percentCount);
    let text = Math.abs(scaled).toFixed(decimals);

    if (decimals > 0 && DisplayText.hasOptionalOnlyDecimals(plain)) {
      text = text.replace(/0+$/, "").replace(/\.$/, "");
    }

    if (useGrouping) {
      text = DisplayText.addGrouping(text);
    }

    if (isNegativeSection) {
      return section.indexOf("(") >= 0 ? `(${text})` : `-${text}`;
    }

    return scaled < 0 ? `-${text}` : text;
  }

  private static hasOptionalOnlyDecimals(plain: string): boolean {
    const dotIndex = plain.indexOf(".");
    if (dotIndex < 0) {
      return false;
    }

    let hasZero = false;
    for (let index = dotIndex + 1; index < plain.length; index++) {
      const character = plain[index];
      if (character === "0") {
        hasZero = true;
        break;
      }
      if (character !== "#" && character !== " ") {
        break;
      }
    }

    return !hasZero;
  }

  private static addGrouping(text: string): string {
    const parts = text.split(".");
    const integer = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    return parts.length > 1 ? `${integer}.${parts[1]}` : integer;
  }

  private static splitSections(formatCode: string): string[] {
    const sections: string[] = [];
    let current = "";
    let inQuote = false;

    for (let index = 0; index < formatCode.length; index++) {
      const character = formatCode[index];
      if (character === '"') {
        inQuote = !inQuote;
        current += character;
        continue;
      }

      if (!inQuote && character === ";") {
        sections.push(current);
        current = "";
        continue;
      }

      current += character;
    }

    sections.push(current);
    return sections;
  }

  private static selectSection(sections: string[], value: number): string | null {
    if (sections.length === 0) {
      return null;
    }

    if (sections.length === 1) {
      return sections[0];
    }

    if (sections.length === 2) {
      return value < 0 ? sections[1] : sections[0];
    }

    if (value > 0) {
      return sections[0];
    }

    if (value < 0) {
      return sections[1];
    }

    return sections[2];
  }
}
