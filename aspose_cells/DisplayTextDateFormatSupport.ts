import { DisplayTextCulture } from "./DisplayTextCulture";

/**
 * Formats date/time values using Excel's custom date and elapsed-time tokens.
 *
 * Ported from Aspose.Cells_FOSS.DisplayTextDateFormatSupport. JavaScript Date
 * values are treated as wall-clock values: the local date and time components
 * are the ones rendered, mirroring .NET DateTime semantics without time-zone
 * conversion.
 */
export class DisplayTextDateFormatSupport {
  /** Formats a date value using the supplied Excel date format code. */
  public static FormatDateTimeValue(value: Date, formatCode: string, cultureName: string): string {
    if (
      DisplayTextDateFormatSupport.indexOfIgnoreCase(formatCode, "ss") >= 0 &&
      formatCode.indexOf(".0") < 0
    ) {
      value = new Date(Math.round(value.getTime() / 1000) * 1000);
    }

    const culture = DisplayTextCulture.NormalizeName(cultureName);
    let builder = "";
    const hasAmPm =
      DisplayTextDateFormatSupport.indexOfIgnoreCase(formatCode, "AM/PM") >= 0 ||
      DisplayTextDateFormatSupport.indexOfIgnoreCase(formatCode, "A/P") >= 0;
    let inQuote = false;

    for (let index = 0; index < formatCode.length; index++) {
      if (DisplayTextDateFormatSupport.MatchesToken(formatCode, index, "aaa")) {
        builder += DisplayTextDateFormatSupport.GetAbbreviatedDayName(culture, value.getDay());
        index += 2;
        continue;
      }

      if (DisplayTextDateFormatSupport.MatchesToken(formatCode, index, "AM/PM")) {
        builder += DisplayTextDateFormatSupport.GetAmPmDesignator(value, culture, false);
        index += 4;
        continue;
      }

      if (DisplayTextDateFormatSupport.MatchesToken(formatCode, index, "A/P")) {
        builder += DisplayTextDateFormatSupport.GetAmPmDesignator(value, culture, true);
        index += 2;
        continue;
      }

      const character = formatCode[index];
      if (character === '"') {
        inQuote = !inQuote;
        continue;
      }

      if (inQuote) {
        builder += character;
        continue;
      }

      if (character === "\\") {
        if (index + 1 < formatCode.length) {
          index++;
          builder += formatCode[index];
        }

        continue;
      }

      if (character === "_" || character === "*") {
        index++;
        continue;
      }

      if (character === ".") {
        const appended = DisplayTextDateFormatSupport.TryAppendFractionalSeconds(
          value,
          formatCode,
          culture,
          index,
        );
        if (appended !== null) {
          builder += appended.Text;
          index = appended.NextIndex;
          continue;
        }
      }

      if (character === "y" || character === "Y") {
        const count = DisplayTextDateFormatSupport.CountRepeated(formatCode, index, character);
        builder += DisplayTextDateFormatSupport.FormatYear(value.getFullYear(), count);
        index += count - 1;
        continue;
      }

      if (character === "d" || character === "D") {
        const count = DisplayTextDateFormatSupport.CountRepeated(formatCode, index, character);
        builder += DisplayTextDateFormatSupport.FormatDay(value, culture, count);
        index += count - 1;
        continue;
      }

      if (character === "h" || character === "H") {
        const count = DisplayTextDateFormatSupport.CountRepeated(formatCode, index, character);
        builder += DisplayTextDateFormatSupport.FormatHour(value.getHours(), count, hasAmPm);
        index += count - 1;
        continue;
      }

      if (character === "s" || character === "S") {
        const count = DisplayTextDateFormatSupport.CountRepeated(formatCode, index, character);
        builder += DisplayTextDateFormatSupport.FormatTwoDigit(value.getSeconds(), count);
        index += count - 1;
        continue;
      }

      if (character === "m" || character === "M") {
        const count = DisplayTextDateFormatSupport.CountRepeated(formatCode, index, character);
        if (DisplayTextDateFormatSupport.IsMinuteContext(formatCode, index, count)) {
          builder += DisplayTextDateFormatSupport.FormatTwoDigit(value.getMinutes(), count);
        } else {
          builder += DisplayTextDateFormatSupport.FormatMonth(value.getMonth() + 1, culture, count);
        }

        index += count - 1;
        continue;
      }

      builder += character;
    }

    return builder;
  }

  /** Indicates whether the token appears at the supplied offset, ignoring case. */
  public static MatchesToken(formatCode: string, startIndex: number, token: string): boolean {
    if (startIndex + token.length > formatCode.length) {
      return false;
    }

    return (
      formatCode.substring(startIndex, startIndex + token.length).toLowerCase() === token.toLowerCase()
    );
  }

  /** Counts consecutive repetitions of a token character. */
  public static CountRepeated(formatCode: string, startIndex: number, token: string): number {
    let count = 1;
    for (let index = startIndex + 1; index < formatCode.length; index++) {
      if (formatCode[index].toLowerCase() !== token.toLowerCase()) {
        break;
      }

      count++;
    }

    return count;
  }

  /** Determines whether an "m" token represents minutes rather than a month. */
  public static IsMinuteContext(formatCode: string, startIndex: number, count: number): boolean {
    const previous = DisplayTextDateFormatSupport.FindNeighborToken(formatCode, startIndex - 1, -1);
    const next = DisplayTextDateFormatSupport.FindNeighborToken(formatCode, startIndex + count, 1);
    if (previous === ":" || next === ":") {
      return true;
    }

    if (
      previous === "h" ||
      previous === "H" ||
      previous === "s" ||
      previous === "S"
    ) {
      return true;
    }

    return next === "h" || next === "H" || next === "s" || next === "S";
  }

  /** Finds the nearest non-literal, non-whitespace character in a direction. */
  public static FindNeighborToken(formatCode: string, startIndex: number, direction: number): string {
    let inQuote = false;
    if (direction < 0) {
      for (let index = startIndex; index >= 0; index--) {
        const character = formatCode[index];
        if (character === '"') {
          inQuote = !inQuote;
          continue;
        }

        if (inQuote) {
          continue;
        }

        if (character === "\\" || character === "_" || character === "*") {
          index--;
          continue;
        }

        if (/\s/.test(character)) {
          continue;
        }

        return character;
      }
    } else {
      for (let index = startIndex; index < formatCode.length; index++) {
        const character = formatCode[index];
        if (character === '"') {
          inQuote = !inQuote;
          continue;
        }

        if (inQuote) {
          continue;
        }

        if (character === "\\" || character === "_" || character === "*") {
          index++;
          continue;
        }

        if (/\s/.test(character)) {
          continue;
        }

        return character;
      }
    }

    return "";
  }

  /** Indicates whether a bracketed token is an elapsed-time token. */
  public static IsElapsedToken(token: string): boolean {
    const normalized = token.trim().toLowerCase();
    return (
      normalized === "h" ||
      normalized === "hh" ||
      normalized === "m" ||
      normalized === "mm" ||
      normalized === "s" ||
      normalized === "ss"
    );
  }

  /** Indicates whether the format contains an elapsed-time token such as "[h]". */
  public static ContainsElapsedTimeToken(formatCode: string): boolean {
    for (let index = 0; index < formatCode.length; index++) {
      if (formatCode[index] !== "[") {
        continue;
      }

      const endIndex = formatCode.indexOf("]", index + 1);
      if (endIndex < 0) {
        continue;
      }

      const token = formatCode.substring(index + 1, endIndex);
      if (DisplayTextDateFormatSupport.IsElapsedToken(token)) {
        return true;
      }
    }

    return false;
  }

  /** Indicates whether the format contains a textual month token (mmm or longer). */
  public static ContainsTextualMonthToken(formatCode: string): boolean {
    let inQuote = false;
    for (let index = 0; index < formatCode.length; index++) {
      const character = formatCode[index];
      if (character === '"') {
        inQuote = !inQuote;
        continue;
      }

      if (inQuote) {
        continue;
      }

      if (character === "\\" || character === "_" || character === "*") {
        index++;
        continue;
      }

      if (character === "m" || character === "M") {
        const count = DisplayTextDateFormatSupport.CountRepeated(formatCode, index, character);
        if (
          !DisplayTextDateFormatSupport.IsMinuteContext(formatCode, index, count) &&
          count >= 3
        ) {
          return true;
        }

        index += count - 1;
      }
    }

    return false;
  }

  /**
   * Formats an elapsed duration. The duration is supplied as milliseconds so
   * callers can round the Excel serial the same way TimeSpan.FromDays does.
   */
  public static FormatElapsedTimeValue(
    totalMilliseconds: number,
    formatCode: string,
    cultureName: string,
  ): string {
    if (
      DisplayTextDateFormatSupport.indexOfIgnoreCase(formatCode, "ss") >= 0 &&
      formatCode.indexOf(".0") < 0
    ) {
      totalMilliseconds = Math.round(totalMilliseconds / 1000) * 1000;
    }

    const culture = DisplayTextCulture.NormalizeName(cultureName);
    const separators = DisplayTextCulture.GetNumberSeparators(culture);
    const totalHours = Math.floor(totalMilliseconds / 3600000);
    const totalMinutes = Math.floor(totalMilliseconds / 60000);
    const totalSeconds = Math.floor(totalMilliseconds / 1000);
    const hoursComponent = DisplayTextDateFormatSupport.Modulo(Math.floor(totalMilliseconds / 3600000), 24);
    const minutesComponent = DisplayTextDateFormatSupport.Modulo(Math.floor(totalMilliseconds / 60000), 60);
    const secondsComponent = DisplayTextDateFormatSupport.Modulo(Math.floor(totalMilliseconds / 1000), 60);
    const millisecondsComponent = DisplayTextDateFormatSupport.Modulo(totalMilliseconds, 1000);

    let builder = "";
    let inQuote = false;

    for (let index = 0; index < formatCode.length; index++) {
      const character = formatCode[index];
      if (character === '"') {
        inQuote = !inQuote;
        continue;
      }

      if (inQuote) {
        builder += character;
        continue;
      }

      if (character === "[") {
        const endIndex = formatCode.indexOf("]", index + 1);
        if (endIndex > index) {
          const token = formatCode.substring(index + 1, endIndex).toLowerCase();
          if (token === "h") {
            builder += String(totalHours);
          } else if (token === "hh") {
            builder += DisplayTextDateFormatSupport.PadStart(String(totalHours), 2);
          } else if (token === "m") {
            builder += String(totalMinutes);
          } else if (token === "mm") {
            builder += DisplayTextDateFormatSupport.PadStart(String(totalMinutes), 2);
          } else if (token === "s") {
            builder += String(totalSeconds);
          } else if (token === "ss") {
            builder += DisplayTextDateFormatSupport.PadStart(String(totalSeconds), 2);
          } else {
            builder += `[${token}]`;
          }

          index = endIndex;
          continue;
        }
      }

      if (character === "h" || character === "H") {
        const count = DisplayTextDateFormatSupport.CountRepeated(formatCode, index, character);
        builder += DisplayTextDateFormatSupport.FormatTwoDigit(hoursComponent, count);
        index += count - 1;
        continue;
      }

      if (character === "m" || character === "M") {
        const count = DisplayTextDateFormatSupport.CountRepeated(formatCode, index, character);
        builder += DisplayTextDateFormatSupport.FormatTwoDigit(minutesComponent, count);
        index += count - 1;
        continue;
      }

      if (character === "s" || character === "S") {
        const count = DisplayTextDateFormatSupport.CountRepeated(formatCode, index, character);
        builder += DisplayTextDateFormatSupport.FormatTwoDigit(secondsComponent, count);
        index += count - 1;
        continue;
      }

      if (character === "." && index + 1 < formatCode.length && formatCode[index + 1] === "0") {
        const zeroCount = DisplayTextDateFormatSupport.CountRepeated(formatCode, index + 1, "0");
        builder += separators.Decimal;
        builder += DisplayTextDateFormatSupport.FormatFractionDigits(millisecondsComponent, zeroCount);
        index += zeroCount;
        continue;
      }

      if (character === "\\") {
        if (index + 1 < formatCode.length) {
          index++;
          builder += formatCode[index];
        }

        continue;
      }

      if (character === "_" || character === "*") {
        index++;
        continue;
      }

      builder += character;
    }

    return builder;
  }

  private static indexOfIgnoreCase(value: string, search: string): number {
    return value.toLowerCase().indexOf(search.toLowerCase());
  }

  private static Modulo(value: number, divisor: number): number {
    const remainder = value % divisor;
    return remainder < 0 ? remainder + divisor : remainder;
  }

  private static PadStart(value: string, width: number): string {
    return value.length >= width ? value : "0".repeat(width - value.length) + value;
  }

  private static FormatYear(year: number, count: number): string {
    if (count <= 1) {
      return String(year % 100);
    }

    if (count === 2) {
      return DisplayTextDateFormatSupport.PadStart(String(year % 100), 2);
    }

    return DisplayTextDateFormatSupport.PadStart(String(year), count);
  }

  private static FormatDay(value: Date, culture: string, count: number): string {
    if (count === 1) {
      return String(value.getDate());
    }

    if (count === 2) {
      return DisplayTextDateFormatSupport.PadStart(String(value.getDate()), 2);
    }

    if (count === 3) {
      return DisplayTextDateFormatSupport.GetAbbreviatedDayName(culture, value.getDay());
    }

    return DisplayTextDateFormatSupport.GetDayName(culture, value.getDay());
  }

  private static FormatMonth(month: number, culture: string, count: number): string {
    if (count === 1) {
      return String(month);
    }

    if (count === 2) {
      return DisplayTextDateFormatSupport.PadStart(String(month), 2);
    }

    if (count === 3) {
      return DisplayTextCulture.GetAbbreviatedMonthNames(culture)[month - 1];
    }

    const monthName = DisplayTextCulture.GetMonthNames(culture)[month - 1];
    if (count === 4) {
      return monthName;
    }

    if (monthName.length === 0) {
      return String(month);
    }

    return monthName[0];
  }

  private static FormatHour(hour: number, count: number, hasAmPm: boolean): string {
    if (hasAmPm) {
      hour %= 12;
      if (hour === 0) {
        hour = 12;
      }
    }

    if (count === 1) {
      return String(hour);
    }

    return DisplayTextDateFormatSupport.PadStart(String(hour), 2);
  }

  private static FormatTwoDigit(value: number, count: number): string {
    if (count === 1) {
      return String(value);
    }

    return DisplayTextDateFormatSupport.PadStart(String(value), 2);
  }

  private static TryAppendFractionalSeconds(
    value: Date,
    formatCode: string,
    culture: string,
    index: number,
  ): { Text: string; NextIndex: number } | null {
    if (index + 1 >= formatCode.length || formatCode[index + 1] !== "0") {
      return null;
    }

    const previous = DisplayTextDateFormatSupport.FindNeighborToken(formatCode, index - 1, -1);
    if (previous !== "s" && previous !== "S") {
      return null;
    }

    const zeroCount = DisplayTextDateFormatSupport.CountRepeated(formatCode, index + 1, "0");
    const separators = DisplayTextCulture.GetNumberSeparators(culture);
    const text =
      separators.Decimal +
      DisplayTextDateFormatSupport.FormatFractionDigits(value.getMilliseconds(), zeroCount);
    return { Text: text, NextIndex: index + zeroCount };
  }

  private static FormatFractionDigits(milliseconds: number, zeroCount: number): string {
    const digits = DisplayTextDateFormatSupport.PadStart(String(milliseconds), 3);
    if (zeroCount <= 0) {
      return "";
    }

    if (zeroCount === 1) {
      return digits[0];
    }

    if (zeroCount === 2) {
      return digits[0] + digits[1];
    }

    let result = digits;
    for (let index = 3; index < zeroCount; index++) {
      result += "0";
    }

    return result;
  }

  private static GetDayName(culture: string, dayOfWeek: number): string {
    return DisplayTextCulture.GetDayNames(culture)[dayOfWeek];
  }

  private static GetAbbreviatedDayName(culture: string, dayOfWeek: number): string {
    return DisplayTextCulture.GetAbbreviatedDayNames(culture)[dayOfWeek];
  }

  private static GetAmPmDesignator(value: Date, culture: string, abbreviated: boolean): string {
    const designators = DisplayTextCulture.GetAmPmDesignators(culture);
    let designator = value.getHours() < 12 ? designators.AM : designators.PM;
    if (designator.length === 0) {
      designator = value.getHours() < 12 ? "AM" : "PM";
    }

    if (!abbreviated) {
      return designator;
    }

    return designator.substring(0, 1);
  }
}
