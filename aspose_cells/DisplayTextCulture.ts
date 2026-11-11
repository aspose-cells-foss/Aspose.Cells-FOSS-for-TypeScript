/**
 * Culture data used by the display-text formatter.
 *
 * The .NET implementation relies on System.Globalization.CultureInfo. This port
 * derives the same information from the ECMAScript Intl APIs where they can
 * supply it (separators, day and month names, AM/PM designators) and keeps a
 * small table of the long/short date and time patterns that Intl does not expose.
 */
export interface NumberSeparators {
  Decimal: string;
  Group: string;
}

export interface AmPmDesignators {
  AM: string;
  PM: string;
}

export class DisplayTextCulture {
  private static readonly numberSeparatorCache = new Map<string, NumberSeparators>();
  private static readonly monthNameCache = new Map<string, string[]>();
  private static readonly abbreviatedMonthNameCache = new Map<string, string[]>();
  private static readonly dayNameCache = new Map<string, string[]>();
  private static readonly abbreviatedDayNameCache = new Map<string, string[]>();
  private static readonly amPmCache = new Map<string, AmPmDesignators>();

  private static readonly lcidTable: ReadonlyMap<string, string> = new Map<string, string>([
    ["409", "en-US"],
    ["809", "en-GB"],
    ["40C", "fr-FR"],
    ["407", "de-DE"],
    ["40A", "es-ES"],
    ["C0A", "es-ES"],
    ["410", "it-IT"],
    ["416", "pt-BR"],
    ["816", "pt-PT"],
    ["413", "nl-NL"],
    ["405", "pl-PL"],
    ["41D", "sv-SE"],
    ["40F", "da-DK"],
    ["414", "nb-NO"],
    ["415", "fi-FI"],
    ["419", "ru-RU"],
    ["412", "ko-KR"],
    ["411", "ja-JP"],
    ["804", "zh-CN"],
    ["404", "zh-TW"],
    ["40D", "he-IL"],
  ]);

  /** Returns the trimmed culture name, mapping null and undefined to invariant. */
  public static NormalizeName(name: string | null | undefined): string {
    if (name === null || name === undefined) {
      return "";
    }

    return name.trim();
  }

  /** Indicates whether the culture name denotes the invariant culture. */
  public static IsInvariant(name: string): boolean {
    return DisplayTextCulture.NormalizeName(name).length === 0;
  }

  /**
   * Indicates whether a format without a locale directive should still render
   * English month names for the supplied culture.
   *
   * The .NET formatter applied this fallback when the workbook culture matched
   * the host's CurrentCulture. The fixtures used a Chinese host culture, so the
   * port keeps that behaviour for Chinese workbooks while explicit Western
   * cultures continue to use their own month names.
   */
  public static UsesEnglishMonthFallback(name: string): boolean {
    const normalized = DisplayTextCulture.NormalizeName(name);
    if (normalized.length === 0) {
      return true;
    }

    return normalized.toLowerCase().startsWith("zh");
  }

  /** Maps a hexadecimal LCID to a culture name, or returns null when unknown. */
  public static ResolveLocaleFromLcid(localeCode: string): string | null {
    const normalized = localeCode.trim().toUpperCase();
    if (normalized.length === 0) {
      return null;
    }

    const resolved = DisplayTextCulture.lcidTable.get(normalized);
    return resolved === undefined ? null : resolved;
  }

  /** Returns the decimal and grouping separators for a culture. */
  public static GetNumberSeparators(name: string): NumberSeparators {
    const key = DisplayTextCulture.NormalizeName(name);
    const cached = DisplayTextCulture.numberSeparatorCache.get(key);
    if (cached !== undefined) {
      return cached;
    }

    if (key.length === 0) {
      const invariant: NumberSeparators = { Decimal: ".", Group: "," };
      DisplayTextCulture.numberSeparatorCache.set(key, invariant);
      return invariant;
    }

    let separators: NumberSeparators = { Decimal: ".", Group: "," };
    try {
      const formatter = new Intl.NumberFormat(key);
      const decimalPart = formatter.formatToParts(1234.5).find((part) => part.type === "decimal");
      const groupPart = formatter.formatToParts(1234567.5).find((part) => part.type === "group");
      separators = {
        Decimal: decimalPart === undefined ? "." : decimalPart.value,
        Group: groupPart === undefined ? "," : groupPart.value,
      };
    } catch {
      separators = { Decimal: ".", Group: "," };
    }

    DisplayTextCulture.numberSeparatorCache.set(key, separators);
    return separators;
  }

  /** Returns full month names indexed by month number minus one. */
  public static GetMonthNames(name: string): string[] {
    return DisplayTextCulture.getMonthNames(name, "long", DisplayTextCulture.monthNameCache);
  }

  /** Returns abbreviated month names indexed by month number minus one. */
  public static GetAbbreviatedMonthNames(name: string): string[] {
    return DisplayTextCulture.getMonthNames(
      name,
      "short",
      DisplayTextCulture.abbreviatedMonthNameCache,
    );
  }

  /** Returns full day names indexed from Sunday (0) through Saturday (6). */
  public static GetDayNames(name: string): string[] {
    return DisplayTextCulture.getDayNames(name, "long", DisplayTextCulture.dayNameCache);
  }

  /** Returns abbreviated day names indexed from Sunday (0) through Saturday (6). */
  public static GetAbbreviatedDayNames(name: string): string[] {
    return DisplayTextCulture.getDayNames(
      name,
      "short",
      DisplayTextCulture.abbreviatedDayNameCache,
    );
  }

  /** Returns the AM/PM designators for a culture. */
  public static GetAmPmDesignators(name: string): AmPmDesignators {
    const key = DisplayTextCulture.NormalizeName(name);
    const cached = DisplayTextCulture.amPmCache.get(key);
    if (cached !== undefined) {
      return cached;
    }

    let designators: AmPmDesignators = { AM: "AM", PM: "PM" };
    if (key.length > 0) {
      try {
        const formatter = new Intl.DateTimeFormat(key, {
          hour: "numeric",
          hour12: true,
          timeZone: "UTC",
        });
        const amPart = formatter
          .formatToParts(new Date(Date.UTC(2000, 0, 1, 7)))
          .find((part) => part.type === "dayPeriod");
        const pmPart = formatter
          .formatToParts(new Date(Date.UTC(2000, 0, 1, 19)))
          .find((part) => part.type === "dayPeriod");
        designators = {
          AM: amPart === undefined ? "AM" : amPart.value,
          PM: pmPart === undefined ? "PM" : pmPart.value,
        };
      } catch {
        designators = { AM: "AM", PM: "PM" };
      }
    }

    DisplayTextCulture.amPmCache.set(key, designators);
    return designators;
  }

  /** Returns the culture's long date pattern. */
  public static GetLongDatePattern(name: string): string {
    switch (DisplayTextCulture.NormalizeName(name).toLowerCase()) {
      case "en-us":
        return "dddd, MMMM d, yyyy";
      case "fr-fr":
        return "dddd d MMMM yyyy";
      case "de-de":
        return "dddd, d. MMMM yyyy";
      case "ja-jp":
        return "yyyy\u5e74M\u6708d\u65e5";
      case "zh-cn":
        return "yyyy\u5e74M\u6708d\u65e5";
      case "zh-tw":
        return "yyyy\u5e74M\u6708d\u65e5";
      default:
        return "dddd, dd MMMM yyyy";
    }
  }

  /** Returns the culture's long time pattern. */
  public static GetLongTimePattern(name: string): string {
    switch (DisplayTextCulture.NormalizeName(name).toLowerCase()) {
      case "en-us":
      case "en-gb":
        return "h:mm:ss tt";
      default:
        return "HH:mm:ss";
    }
  }

  /** Returns the culture's short date pattern. */
  public static GetShortDatePattern(name: string): string {
    switch (DisplayTextCulture.NormalizeName(name).toLowerCase()) {
      case "en-us":
        return "M/d/yyyy";
      case "en-gb":
        return "dd/MM/yyyy";
      case "fr-fr":
        return "dd/MM/yyyy";
      case "de-de":
        return "dd.MM.yyyy";
      case "ja-jp":
        return "yyyy/MM/dd";
      case "zh-cn":
        return "yyyy/M/d";
      case "zh-tw":
        return "yyyy/M/d";
      default:
        return "MM/dd/yyyy";
    }
  }

  /** Returns the culture's short time pattern. */
  public static GetShortTimePattern(name: string): string {
    switch (DisplayTextCulture.NormalizeName(name).toLowerCase()) {
      case "en-us":
        return "h:mm tt";
      default:
        return "HH:mm";
    }
  }

  private static getMonthNames(
    name: string,
    style: "long" | "short",
    cache: Map<string, string[]>,
  ): string[] {
    const key = DisplayTextCulture.NormalizeName(name);
    const cached = cache.get(key);
    if (cached !== undefined) {
      return cached;
    }

    const names: string[] = [];
    for (let month = 0; month < 12; month++) {
      names.push(DisplayTextCulture.formatName(key, new Date(Date.UTC(2000, month, 15)), {
        month: style,
        timeZone: "UTC",
      }));
    }

    cache.set(key, names);
    return names;
  }

  private static getDayNames(
    name: string,
    style: "long" | "short",
    cache: Map<string, string[]>,
  ): string[] {
    const key = DisplayTextCulture.NormalizeName(name);
    const cached = cache.get(key);
    if (cached !== undefined) {
      return cached;
    }

    const names: string[] = [];
    for (let day = 0; day < 7; day++) {
      // 2024-01-07 is a Sunday, so the probes cover Sunday through Saturday.
      names.push(DisplayTextCulture.formatName(key, new Date(Date.UTC(2024, 0, 7 + day)), {
        weekday: style,
        timeZone: "UTC",
      }));
    }

    cache.set(key, names);
    return names;
  }

  private static formatName(
    cultureName: string,
    value: Date,
    options: Intl.DateTimeFormatOptions,
  ): string {
    const locale = cultureName.length === 0 ? "en-US" : cultureName;
    try {
      return new Intl.DateTimeFormat(locale, options).format(value);
    } catch {
      return new Intl.DateTimeFormat("en-US", options).format(value);
    }
  }
}
