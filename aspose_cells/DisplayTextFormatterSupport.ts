import { DisplayFormatSectionInfo } from "./DisplayFormatSectionInfo";
import { DisplayTextCulture } from "./DisplayTextCulture";
import { DisplayTextDateFormatSupport } from "./DisplayTextDateFormatSupport";

/**
 * Helpers shared by the display-text formatter: section parsing and selection,
 * directive stripping, pattern expansion, and numeric custom-format rendering.
 *
 * Ported from Aspose.Cells_FOSS.DisplayTextFormatterSupport.
 */
export class DisplayTextFormatterSupport {
  /** Selects the numeric section that applies to the supplied value. */
  public static SelectNumericSection(
    sections: DisplayFormatSectionInfo[],
    numericValue: number,
  ): { Section: DisplayFormatSectionInfo | null; UseAbsoluteValue: boolean } {
    const numericSectionCount = Math.min(sections.length, 3);
    if (numericSectionCount === 0) {
      return { Section: null, UseAbsoluteValue: false };
    }

    let hasCondition = false;
    let fallbackSection: DisplayFormatSectionInfo | null = null;
    for (let index = 0; index < numericSectionCount; index++) {
      const section = sections[index];
      if (section.HasCondition) {
        hasCondition = true;
        if (
          DisplayTextFormatterSupport.EvaluateCondition(
            section.ConditionOperator,
            numericValue,
            section.ConditionValue,
          )
        ) {
          return {
            Section: section,
            UseAbsoluteValue: DisplayTextFormatterSupport.ShouldUseAbsoluteValue(
              section.Raw,
              numericValue,
            ),
          };
        }
      } else if (fallbackSection === null) {
        fallbackSection = section;
      }
    }

    if (hasCondition && fallbackSection !== null) {
      return {
        Section: fallbackSection,
        UseAbsoluteValue: DisplayTextFormatterSupport.ShouldUseAbsoluteValue(
          fallbackSection.Raw,
          numericValue,
        ),
      };
    }

    if (numericSectionCount === 1) {
      // A lone section formats the signed value itself.
      return { Section: sections[0], UseAbsoluteValue: false };
    }

    if (numericSectionCount === 2) {
      if (numericValue < 0) {
        // A dedicated negative section receives the magnitude. Its own literals
        // (minus sign, parentheses, currency) supply the sign.
        return { Section: sections[1], UseAbsoluteValue: true };
      }

      return {
        Section: sections[0],
        UseAbsoluteValue: DisplayTextFormatterSupport.ShouldUseAbsoluteValue(
          sections[0].Raw,
          numericValue,
        ),
      };
    }

    if (numericValue > 0) {
      return {
        Section: sections[0],
        UseAbsoluteValue: DisplayTextFormatterSupport.ShouldUseAbsoluteValue(
          sections[0].Raw,
          numericValue,
        ),
      };
    }

    if (numericValue < 0) {
      return { Section: sections[1], UseAbsoluteValue: true };
    }

    return { Section: sections[2], UseAbsoluteValue: false };
  }

  /** Selects the section that shapes text values. */
  public static SelectTextSection(
    sections: DisplayFormatSectionInfo[],
  ): DisplayFormatSectionInfo | null {
    if (sections.length >= 4) {
      return sections[3];
    }

    for (const section of sections) {
      const text = DisplayTextFormatterSupport.StripDirectiveBrackets(section.Raw, false);
      if (text.indexOf("@") >= 0) {
        return section;
      }
    }

    return null;
  }

  /** Selects the first non-empty section that carries date tokens. */
  public static SelectDateTimeSection(
    sections: DisplayFormatSectionInfo[],
  ): DisplayFormatSectionInfo | null {
    const dateSectionCount = Math.min(sections.length, 3);
    for (let index = 0; index < dateSectionCount; index++) {
      if (sections[index].Raw.trim().length > 0) {
        return sections[index];
      }
    }

    return null;
  }

  /** Splits a format code into sections and parses conditional tokens. */
  public static ParseSections(formatCode: string): DisplayFormatSectionInfo[] {
    const rawSections = DisplayTextFormatterSupport.SplitSections(formatCode);
    return rawSections.map((raw) => {
      const section = new DisplayFormatSectionInfo();
      section.Raw = raw;
      const condition = DisplayTextFormatterSupport.TryParseSectionCondition(raw);
      if (condition !== null) {
        section.HasCondition = true;
        section.ConditionOperator = condition.Operator;
        section.ConditionValue = condition.Value;
      }

      return section;
    });
  }

  /** Splits a format code on semicolons that are not inside a quoted literal. */
  public static SplitSections(formatCode: string): string[] {
    const sections: string[] = [];
    let builder = "";
    let inQuote = false;

    for (let index = 0; index < formatCode.length; index++) {
      const character = formatCode[index];
      if (character === '"') {
        builder += character;
        inQuote = !inQuote;
        continue;
      }

      if (!inQuote) {
        if (character === "\\" && index + 1 < formatCode.length) {
          builder += character;
          index++;
          builder += formatCode[index];
          continue;
        }

        if (character === ";") {
          sections.push(builder);
          builder = "";
          continue;
        }
      }

      builder += character;
    }

    sections.push(builder);
    return sections;
  }

  /** Parses a conditional token such as "[>=100]" from the section. */
  public static TryParseSectionCondition(
    section: string,
  ): { Operator: string; Value: number } | null {
    let inQuote = false;

    for (let index = 0; index < section.length; index++) {
      const character = section[index];
      if (character === '"') {
        inQuote = !inQuote;
        continue;
      }

      if (inQuote || character !== "[") {
        continue;
      }

      const endIndex = section.indexOf("]", index + 1);
      if (endIndex < 0) {
        break;
      }

      const token = section.substring(index + 1, endIndex).trim();
      const condition = DisplayTextFormatterSupport.TryParseConditionToken(token);
      if (condition !== null) {
        return condition;
      }

      index = endIndex;
    }

    return null;
  }

  /** Parses an operator/value pair from a conditional token. */
  public static TryParseConditionToken(
    token: string,
  ): { Operator: string; Value: number } | null {
    if (token.trim().length === 0) {
      return null;
    }

    const operators = [">=", "<=", "<>", ">", "<", "="];
    for (const candidate of operators) {
      if (token.startsWith(candidate)) {
        const numberPart = token.substring(candidate.length).trim();
        const parsedValue = Number(numberPart);
        if (numberPart.length > 0 && Number.isFinite(parsedValue)) {
          return { Operator: candidate, Value: parsedValue };
        }
      }
    }

    return null;
  }

  /** Evaluates a parsed conditional token against the numeric value. */
  public static EvaluateCondition(
    conditionOperator: string,
    numericValue: number,
    conditionValue: number,
  ): boolean {
    switch (conditionOperator) {
      case ">":
        return numericValue > conditionValue;
      case ">=":
        return numericValue >= conditionValue;
      case "<":
        return numericValue < conditionValue;
      case "<=":
        return numericValue <= conditionValue;
      case "=":
        return Math.abs(numericValue - conditionValue) < 1e-12;
      case "<>":
        return Math.abs(numericValue - conditionValue) >= 1e-12;
      default:
        return false;
    }
  }

  /** Indicates whether a negative value must be formatted as its magnitude. */
  public static ShouldUseAbsoluteValue(section: string, numericValue: number): boolean {
    if (numericValue >= 0) {
      return false;
    }

    const sanitizedSection = DisplayTextFormatterSupport.SanitizeNumericSection(section);
    return sanitizedSection.indexOf("-") < 0;
  }

  /** Removes directives and alignment hints from a numeric section. */
  public static SanitizeNumericSection(section: string): string {
    const withoutDirectives = DisplayTextFormatterSupport.StripDirectiveBrackets(section, false);
    let builder = "";
    let inQuote = false;

    for (let index = 0; index < withoutDirectives.length; index++) {
      const character = withoutDirectives[index];
      if (character === '"') {
        builder += character;
        inQuote = !inQuote;
        continue;
      }

      if (!inQuote) {
        if (character === "_") {
          index++;
          continue;
        }

        if (character === "*") {
          index++;
          continue;
        }

        if (character === "\\") {
          if (index + 1 < withoutDirectives.length) {
            index++;
            builder += withoutDirectives[index];
          }

          continue;
        }

        if (character === "?") {
          builder += "#";
          continue;
        }

        if (character === "[" || character === "]") {
          continue;
        }
      }

      builder += character;
    }

    return builder;
  }

  /** Removes bracketed directives, optionally preserving elapsed-time tokens. */
  public static StripDirectiveBrackets(section: string, preserveElapsedTokens: boolean): string {
    let builder = "";
    let inQuote = false;

    for (let index = 0; index < section.length; index++) {
      const character = section[index];
      if (character === '"') {
        builder += character;
        inQuote = !inQuote;
        continue;
      }

      if (!inQuote && character === "[") {
        const endIndex = section.indexOf("]", index + 1);
        if (endIndex < 0) {
          continue;
        }

        const token = section.substring(index + 1, endIndex);
        if (preserveElapsedTokens && DisplayTextDateFormatSupport.IsElapsedToken(token)) {
          builder += `[${token}]`;
        }

        index = endIndex;
        continue;
      }

      builder += character;
    }

    return builder;
  }

  /** Indicates whether the pattern contains a numeric placeholder. */
  public static ContainsNumericPlaceholder(pattern: string): boolean {
    for (let index = 0; index < pattern.length; index++) {
      const character = pattern[index];
      if (
        character === "0" ||
        character === "#" ||
        character === "?" ||
        character === "." ||
        character === "%" ||
        character === "E" ||
        character === "e" ||
        character === "/"
      ) {
        return true;
      }
    }

    return false;
  }

  /** Expands a section pattern, replacing text placeholders and literals. */
  public static ExpandSectionPattern(
    pattern: string,
    valueText: string,
    replaceTextPlaceholder: boolean,
  ): string {
    let builder = "";
    let inQuote = false;

    for (let index = 0; index < pattern.length; index++) {
      const character = pattern[index];
      if (character === '"') {
        inQuote = !inQuote;
        continue;
      }

      if (!inQuote) {
        if (character === "@" && replaceTextPlaceholder) {
          builder += valueText;
          continue;
        }

        if (character === "_") {
          index++;
          continue;
        }

        if (character === "*") {
          index++;
          continue;
        }

        if (character === "\\") {
          if (index + 1 < pattern.length) {
            index++;
            builder += pattern[index];
          }

          continue;
        }
      }

      builder += character;
    }

    return builder;
  }

  /**
   * Renders a numeric value with a .NET-style custom numeric format string.
   * Only placeholders, literals, percent scaling, and scientific notation are
   * supported, which covers the Excel formats handled by FormatNumericValue.
   */
  public static FormatNumber(value: number, pattern: string, cultureName: string): string {
    if (!Number.isFinite(value)) {
      return String(value);
    }

    const separators = DisplayTextCulture.GetNumberSeparators(cultureName);
    let percentCount = 0;
    let inQuote = false;
    for (let index = 0; index < pattern.length; index++) {
      const character = pattern[index];
      if (character === '"') {
        inQuote = !inQuote;
        continue;
      }

      if (!inQuote && character === "\\") {
        index++;
        continue;
      }

      if (!inQuote && character === "%") {
        percentCount++;
      }
    }

    const scaled = value * Math.pow(100, percentCount);
    const exponentIndex = DisplayTextFormatterSupport.findExponentIndex(pattern);
    if (exponentIndex >= 0) {
      return DisplayTextFormatterSupport.renderExponent(
        scaled,
        pattern,
        exponentIndex,
        separators.Decimal,
        separators.Group,
      );
    }

    return DisplayTextFormatterSupport.renderPlain(
      scaled,
      pattern,
      separators.Decimal,
      separators.Group,
    );
  }

  /**
   * Formats values that carry a fraction pattern such as "# ?/?".
   * Returns null when the section does not describe a fraction.
   */
  public static TryFormatFraction(
    numericValue: number,
    section: string,
    useAbsoluteValue: boolean,
  ): string | null {
    const sanitizedSection = DisplayTextFormatterSupport.SanitizeNumericSection(section);
    const slashIndex = sanitizedSection.indexOf("/");
    if (slashIndex <= 0) {
      return null;
    }

    let fixedEnd = slashIndex + 1;
    while (fixedEnd < sanitizedSection.length && /[0-9]/.test(sanitizedSection[fixedEnd])) {
      fixedEnd++;
    }

    if (fixedEnd > slashIndex + 1) {
      const fixedDenominator = Number.parseInt(
        sanitizedSection.substring(slashIndex + 1, fixedEnd),
        10,
      );
      if (!Number.isFinite(fixedDenominator) || fixedDenominator <= 0) {
        return null;
      }

      const absolute = Math.abs(numericValue);
      let whole = Math.floor(absolute);
      let numerator = Math.round((absolute - whole) * fixedDenominator);
      if (numerator >= fixedDenominator) {
        whole++;
        numerator = 0;
      }

      const sign = numericValue < 0 && !useAbsoluteValue ? "-" : "";
      if (numerator === 0) {
        return sign + String(whole);
      }

      return (
        sign +
        (whole === 0 ? "" : String(whole) + " ") +
        String(numerator) +
        "/" +
        String(fixedDenominator)
      );
    }

    let denominatorDigits = 0;
    for (let index = slashIndex + 1; index < sanitizedSection.length; index++) {
      const character = sanitizedSection[index];
      if (character === "#" || character === "0") {
        denominatorDigits++;
        continue;
      }

      if (/\s/.test(character)) {
        continue;
      }

      break;
    }

    if (denominatorDigits <= 0) {
      return null;
    }

    const absoluteValue = Math.abs(numericValue);
    let wholePart = Math.floor(absoluteValue);
    const fractionalPart = absoluteValue - wholePart;
    if (fractionalPart < 1e-12) {
      return DisplayTextFormatterSupport.formatWholeFractionResult(
        wholePart,
        useAbsoluteValue,
        numericValue,
      );
    }

    let maxDenominator = 1;
    for (let index = 0; index < denominatorDigits; index++) {
      maxDenominator *= 10;
    }

    // Match Excel's visible fraction behavior by finding the closest rational
    // value that fits within the placeholder width from the format section.
    maxDenominator -= 1;
    let bestNumerator = 0;
    let bestDenominator = 1;
    let bestError = Number.MAX_VALUE;

    for (let denominator = 1; denominator <= maxDenominator; denominator++) {
      let numerator = Math.round(fractionalPart * denominator);
      if (numerator === 0) {
        continue;
      }

      if (numerator > denominator) {
        numerator = denominator;
      }

      const candidate = numerator / denominator;
      const error = Math.abs(fractionalPart - candidate);
      if (error < bestError) {
        bestError = error;
        bestNumerator = numerator;
        bestDenominator = denominator;
      }
    }

    if (bestNumerator === 0) {
      return DisplayTextFormatterSupport.formatWholeFractionResult(
        wholePart,
        useAbsoluteValue,
        numericValue,
      );
    }

    const greatestCommonDivisor = DisplayTextFormatterSupport.greatestCommonDivisor(
      bestNumerator,
      bestDenominator,
    );
    bestNumerator /= greatestCommonDivisor;
    bestDenominator /= greatestCommonDivisor;

    if (bestNumerator === bestDenominator) {
      wholePart++;
      bestNumerator = 0;
    }

    const prefix = !useAbsoluteValue && numericValue < 0 ? "-" : "";
    if (bestNumerator === 0) {
      return prefix + String(wholePart);
    }

    if (wholePart === 0) {
      return prefix + String(bestNumerator) + "/" + String(bestDenominator);
    }

    return prefix + String(wholePart) + " " + String(bestNumerator) + "/" + String(bestDenominator);
  }

  private static formatWholeFractionResult(
    wholePart: number,
    useAbsoluteValue: boolean,
    numericValue: number,
  ): string {
    if (!useAbsoluteValue && numericValue < 0) {
      return "-" + String(wholePart);
    }

    return String(wholePart);
  }

  private static greatestCommonDivisor(left: number, right: number): number {
    let first = Math.abs(left);
    let second = Math.abs(right);
    while (second !== 0) {
      const remainder = first % second;
      first = second;
      second = remainder;
    }

    return first === 0 ? 1 : first;
  }

  private static findExponentIndex(pattern: string): number {
    let inQuote = false;
    for (let index = 0; index < pattern.length; index++) {
      const character = pattern[index];
      if (character === '"') {
        inQuote = !inQuote;
        continue;
      }

      if (!inQuote && character === "\\") {
        index++;
        continue;
      }

      if (!inQuote && (character === "E" || character === "e")) {
        return index;
      }
    }

    return -1;
  }

  private static renderExponent(
    value: number,
    pattern: string,
    exponentIndex: number,
    decimalSeparator: string,
    groupSeparator: string,
  ): string {
    const mantissaPattern = pattern.substring(0, exponentIndex);
    let exponentPattern = pattern.substring(exponentIndex + 1);
    const absolute = Math.abs(value);
    let exponent = absolute === 0 ? 0 : Math.floor(Math.log10(absolute));
    let mantissa = absolute === 0 ? 0 : absolute / Math.pow(10, exponent);

    const maxDecimals = DisplayTextFormatterSupport.countDecimalPlaceholders(mantissaPattern);
    const rounded = Number(mantissa.toFixed(maxDecimals));
    if (rounded >= 10) {
      mantissa /= 10;
      exponent++;
    } else {
      mantissa = rounded;
    }

    let signCharacter = "";
    if (exponentPattern.startsWith("+") || exponentPattern.startsWith("-")) {
      signCharacter = exponentPattern[0];
      exponentPattern = exponentPattern.substring(1);
    }

    const digitIndex = exponentPattern.search(/[0#]/);
    const prefix = digitIndex >= 0 ? exponentPattern.substring(0, digitIndex) : exponentPattern;
    const digitPattern = digitIndex >= 0 ? exponentPattern.substring(digitIndex) : "";
    const exponentDigits = Math.max(
      1,
      digitPattern.split("").filter((character) => character === "0" || character === "#").length,
    );

    let sign = "";
    if (signCharacter === "+") {
      sign = exponent >= 0 ? "+" : "-";
    } else if (signCharacter === "-") {
      sign = exponent >= 0 ? "" : "-";
    }

    const mantissaText = DisplayTextFormatterSupport.renderPlain(
      mantissa,
      mantissaPattern,
      decimalSeparator,
      groupSeparator,
    );
    const exponentText = String(Math.abs(exponent)).padStart(exponentDigits, "0");
    return (
      mantissaText + pattern[exponentIndex] + prefix + sign + exponentText
    );
  }

  private static countDecimalPlaceholders(pattern: string): number {
    let inQuote = false;
    let inDecimal = false;
    let count = 0;

    for (let index = 0; index < pattern.length; index++) {
      const character = pattern[index];
      if (character === '"') {
        inQuote = !inQuote;
        continue;
      }

      if (inQuote) {
        continue;
      }

      if (character === "\\") {
        index++;
        continue;
      }

      if (character === ".") {
        inDecimal = true;
        continue;
      }

      if (inDecimal && (character === "0" || character === "#")) {
        count++;
      }
    }

    return count;
  }

  private static renderPlain(
    value: number,
    pattern: string,
    decimalSeparator: string,
    groupSeparator: string,
  ): string {
    let builder = "";
    let inQuote = false;
    let numberWritten = false;

    for (let index = 0; index < pattern.length; index++) {
      const character = pattern[index];
      if (character === '"') {
        inQuote = !inQuote;
        continue;
      }

      if (inQuote) {
        builder += character;
        continue;
      }

      if (character === "\\") {
        if (index + 1 < pattern.length) {
          index++;
          builder += pattern[index];
        }

        continue;
      }

      if (character === "_" || character === "*") {
        index++;
        continue;
      }

      if (!numberWritten && DisplayTextFormatterSupport.isNumberPatternCharacter(character)) {
        let end = index;
        while (
          end < pattern.length &&
          DisplayTextFormatterSupport.isNumberPatternCharacter(pattern[end])
        ) {
          end++;
        }

        const run = pattern.substring(index, end);
        builder += DisplayTextFormatterSupport.buildNumberText(
          Math.abs(value),
          run,
          decimalSeparator,
          groupSeparator,
        );
        numberWritten = true;
        index = end - 1;
        continue;
      }

      builder += character;
    }

    return builder;
  }

  private static isNumberPatternCharacter(character: string): boolean {
    return character === "0" || character === "#" || character === "?" || character === "." || character === ",";
  }

  private static buildNumberText(
    value: number,
    run: string,
    decimalSeparator: string,
    groupSeparator: string,
  ): string {
    const useGrouping = run.indexOf(",") >= 0;
    const clean = run.split(",").join("");
    const dotIndex = clean.indexOf(".");
    const integerPattern = dotIndex < 0 ? clean : clean.substring(0, dotIndex);
    const decimalPattern = dotIndex < 0 ? "" : clean.substring(dotIndex + 1);
    const maxDecimals = decimalPattern.length;
    const minDecimals = decimalPattern.split("").filter((character) => character === "0").length;
    const minIntegerDigits = integerPattern.split("").filter((character) => character === "0").length;

    const text = value.toFixed(maxDecimals);
    let integerText = text;
    let decimalText = "";
    const textDotIndex = text.indexOf(".");
    if (textDotIndex >= 0) {
      integerText = text.substring(0, textDotIndex);
      decimalText = text.substring(textDotIndex + 1);
    }

    while (decimalText.length > minDecimals && decimalText.endsWith("0")) {
      decimalText = decimalText.substring(0, decimalText.length - 1);
    }

    if (integerText.length < minIntegerDigits) {
      integerText = "0".repeat(minIntegerDigits - integerText.length) + integerText;
    }

    if (minIntegerDigits === 0 && integerPattern.length === 0) {
      integerText = "";
    } else if (minIntegerDigits === 0) {
      integerText = integerText.replace(/^0+(?=\d)/, "");
      if (/^0+$/.test(integerText) && integerPattern.indexOf("0") < 0) {
        integerText = "";
      }
    }

    if (useGrouping && integerText.length > 0) {
      integerText = DisplayTextFormatterSupport.groupDigits(integerText, groupSeparator);
    }

    if (decimalText.length > 0) {
      return integerText + decimalSeparator + decimalText;
    }

    return integerText;
  }

  private static groupDigits(text: string, groupSeparator: string): string {
    let result = "";
    let counter = 0;
    for (let index = text.length - 1; index >= 0; index--) {
      result = text[index] + result;
      counter++;
      if (counter === 3 && index > 0) {
        result = groupSeparator + result;
        counter = 0;
      }
    }

    return result;
  }
}
