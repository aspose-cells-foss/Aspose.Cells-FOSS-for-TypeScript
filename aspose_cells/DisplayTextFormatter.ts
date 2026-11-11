import { DateSerialConverter } from "./Core/DateSerialConverter";
import { DateSystem } from "./Core/DateSystem";
import { StyleValue } from "./Core/StyleValue";
import { DisplayTextCulture } from "./DisplayTextCulture";
import { DisplayTextDateFormatSupport } from "./DisplayTextDateFormatSupport";
import { DisplayTextFormatterSupport } from "./DisplayTextFormatterSupport";
import { DisplayTextLocaleSupport } from "./DisplayTextLocaleSupport";
import { NumberFormat } from "./NumberFormat";

/**
 * Produces the raw and display strings for logical cell values.
 *
 * FormatStringValue returns the stable raw representation of a logical cell
 * value. FormatDisplayValue applies the style number format, locale directives,
 * workbook culture, and date system.
 *
 * Ported from Aspose.Cells_FOSS.DisplayTextFormatter. JavaScript Date values
 * are treated as wall-clock values, matching .NET DateTime semantics.
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
    culture: { Name: string } | null,
    dateSystem: DateSystem,
  ): string {
    if (value === null || value === undefined) {
      return "";
    }

    const cultureName = DisplayTextCulture.NormalizeName(culture === null ? "" : culture.Name);

    if (typeof value === "string") {
      return DisplayText.formatTextValue(value, style);
    }

    if (typeof value === "boolean") {
      return value ? "TRUE" : "FALSE";
    }

    if (value instanceof Date) {
      return DisplayText.formatDateTimeValue(value, style, cultureName, dateSystem);
    }

    if (typeof value === "number") {
      return DisplayText.formatNumericValue(value, style, cultureName);
    }

    return String(value);
  }

  private static formatTextValue(value: string, style: StyleValue): string {
    const formatCode = NumberFormat.ResolveFormatCode(style.NumberFormat.Number, style.NumberFormat.Custom);
    if (formatCode.trim().length === 0 || formatCode === "General") {
      return value;
    }

    // Excel text formats can have a dedicated fourth section; only that section
    // should shape the displayed text without mutating the underlying value.
    const sections = DisplayTextFormatterSupport.ParseSections(formatCode);
    const textSection = DisplayTextFormatterSupport.SelectTextSection(sections);
    if (textSection === null) {
      return value;
    }

    const pattern = DisplayTextFormatterSupport.StripDirectiveBrackets(textSection.Raw, false);
    const formatted = DisplayTextFormatterSupport.ExpandSectionPattern(pattern, value, true);
    if (formatted.length === 0) {
      return value;
    }

    return formatted;
  }

  private static formatNumericValue(value: number, style: StyleValue, cultureName: string): string {
    const formatCode = NumberFormat.ResolveFormatCode(style.NumberFormat.Number, style.NumberFormat.Custom);
    if (formatCode.trim().length === 0 || formatCode === "General") {
      return DisplayText.formatGeneralDouble(value);
    }

    // Numeric formats can split positive, negative, zero, and text behavior into
    // separate sections, so the display path first selects the matching section.
    const sections = DisplayTextFormatterSupport.ParseSections(formatCode);
    if (sections.length === 0) {
      return DisplayText.FormatStringValue(value);
    }

    const selection = DisplayTextFormatterSupport.SelectNumericSection(sections, value);
    const selectedSection = selection.Section;
    if (selectedSection === null || selectedSection.Raw.trim().length === 0) {
      return DisplayText.FormatStringValue(value);
    }

    // Fraction placeholders are not directly compatible with custom numeric
    // format strings, so they are handled separately before standard formatting.
    const fractionResult = DisplayTextFormatterSupport.TryFormatFraction(
      value,
      selectedSection.Raw,
      selection.UseAbsoluteValue,
    );
    if (fractionResult !== null) {
      return fractionResult;
    }

    const localized = DisplayTextLocaleSupport.ApplyLocaleDirectives(selectedSection.Raw, cultureName);
    const sanitizedSection = DisplayTextFormatterSupport.SanitizeNumericSection(localized.Section);
    if (sanitizedSection.trim().length === 0) {
      return DisplayText.FormatStringValue(value);
    }

    if (!DisplayTextFormatterSupport.ContainsNumericPlaceholder(sanitizedSection)) {
      const literal = DisplayTextFormatterSupport.ExpandSectionPattern(sanitizedSection, "", false);
      if (literal.length > 0) {
        return literal;
      }

      return DisplayText.FormatStringValue(value);
    }

    let formattedValue = value;
    if (selection.UseAbsoluteValue && value < 0) {
      formattedValue = Math.abs(value);
    }

    try {
      return DisplayTextFormatterSupport.FormatNumber(formattedValue, sanitizedSection, localized.Culture);
    } catch {
      return DisplayText.FormatStringValue(value);
    }
  }

  private static formatGeneralDouble(value: number): string {
    if (!Number.isFinite(value)) {
      return DisplayText.FormatStringValue(value);
    }

    // Excel's General format rounds to 15 significant digits, which collapses
    // binary artifacts such as 1499.9899999999998 back to 1499.99.
    return Number(value.toPrecision(15)).toString();
  }

  private static formatDateTimeValue(
    value: Date,
    style: StyleValue,
    cultureName: string,
    dateSystem: DateSystem,
  ): string {
    const custom = style.NumberFormat.Custom;
    if (
      (custom === null || custom.length === 0) &&
      (style.NumberFormat.Number === 14 || style.NumberFormat.Number === 22) &&
      cultureName.length > 0
    ) {
      const pattern =
        style.NumberFormat.Number === 22
          ? DisplayTextCulture.GetShortDatePattern(cultureName) +
            " " +
            DisplayTextCulture.GetShortTimePattern(cultureName)
          : DisplayTextCulture.GetShortDatePattern(cultureName);
      return DisplayTextDateFormatSupport.FormatDateTimeValue(value, pattern, cultureName);
    }

    const formatCode = NumberFormat.ResolveFormatCode(style.NumberFormat.Number, style.NumberFormat.Custom);
    if (formatCode.trim().length === 0 || formatCode === "General") {
      return DisplayText.formatRawDateTime(value);
    }

    // Date/time display uses the first section that still carries date tokens
    // after locale and directive cleanup.
    const sections = DisplayTextFormatterSupport.ParseSections(formatCode);
    const section = DisplayTextFormatterSupport.SelectDateTimeSection(sections);
    if (section === null || section.Raw.trim().length === 0) {
      return DisplayText.formatRawDateTime(value);
    }

    const localized = DisplayTextLocaleSupport.ApplyLocaleDirectives(section.Raw, cultureName);
    let sectionCulture = localized.Culture;
    const sectionFormat = DisplayTextFormatterSupport.StripDirectiveBrackets(localized.Section, true);

    // Elapsed-time formats are duration-style output, not calendar rendering.
    if (DisplayTextDateFormatSupport.ContainsElapsedTimeToken(sectionFormat)) {
      const serial = DateSerialConverter.ToWallClockSerial(value, dateSystem);
      return DisplayTextDateFormatSupport.FormatElapsedTimeValue(
        Math.round(serial * 86400000),
        sectionFormat,
        sectionCulture,
      );
    }

    if (sectionFormat.trim().length === 0) {
      return DisplayText.formatRawDateTime(value);
    }

    if (
      DisplayTextDateFormatSupport.ContainsTextualMonthToken(sectionFormat) &&
      !DisplayTextLocaleSupport.ContainsLocaleDirective(section.Raw) &&
      DisplayTextCulture.UsesEnglishMonthFallback(sectionCulture)
    ) {
      sectionCulture = "";
    }

    try {
      return DisplayTextDateFormatSupport.FormatDateTimeValue(value, sectionFormat, sectionCulture);
    } catch {
      return DisplayText.formatRawDateTime(value);
    }
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
}
