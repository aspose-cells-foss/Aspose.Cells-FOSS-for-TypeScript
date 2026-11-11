import { DisplayTextCulture } from "./DisplayTextCulture";

/** Result of applying the locale directives embedded in a format section. */
export interface LocalizedFormatSection {
  /** Section text with locale directives removed or replaced. */
  Section: string;

  /** Culture resolved from the directive, or the fallback culture. */
  Culture: string;
}

/**
 * Handles Excel locale directives such as "[$-409]" and "[$-F800]".
 *
 * Ported from Aspose.Cells_FOSS.DisplayTextLocaleSupport.
 */
export class DisplayTextLocaleSupport {
  /** Indicates whether the section contains a "[$-...]" locale directive. */
  public static ContainsLocaleDirective(section: string): boolean {
    if (section === null || section === undefined || section.length === 0) {
      return false;
    }

    let inQuote = false;
    for (let index = 0; index < section.length; index++) {
      const character = section[index];
      if (character === '"') {
        inQuote = !inQuote;
        continue;
      }

      if (!inQuote && character === "[") {
        const endIndex = section.indexOf("]", index + 1);
        if (endIndex > index) {
          const token = section.substring(index + 1, endIndex);
          if (token.trim().length > 0 && token[0] === "$") {
            return true;
          }

          index = endIndex;
        }
      }
    }

    return false;
  }

  /**
   * Replaces locale directives in the section and resolves the culture that the
   * directive selects. The fallback culture is used when no directive is found.
   */
  public static ApplyLocaleDirectives(
    section: string,
    fallbackCulture: string,
  ): LocalizedFormatSection {
    let culture = DisplayTextCulture.NormalizeName(fallbackCulture);
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
        if (endIndex > index) {
          const token = section.substring(index + 1, endIndex);
          const resolved = DisplayTextLocaleSupport.TryResolveLocaleDirective(token, culture);
          if (resolved !== null) {
            builder += resolved.Replacement;
            culture = resolved.Culture;
            // Excel's system long-date/time directives replace the stored
            // fallback pattern rather than prefixing it.
            if (
              token.toLowerCase() === "$-f800" ||
              token.toLowerCase() === "$-f400"
            ) {
              return { Section: builder, Culture: culture };
            }

            index = endIndex;
            continue;
          }
        }
      }

      builder += character;
    }

    return { Section: builder, Culture: culture };
  }

  private static TryResolveLocaleDirective(
    token: string,
    fallbackCulture: string,
  ): { Replacement: string; Culture: string } | null {
    if (token.trim().length === 0 || token[0] !== "$") {
      return null;
    }

    const dashIndex = token.lastIndexOf("-");
    if (dashIndex < 0) {
      return { Replacement: DisplayTextLocaleSupport.QuoteLiteral(token.substring(1)), Culture: fallbackCulture };
    }

    if (dashIndex === 0 || dashIndex >= token.length - 1) {
      return null;
    }

    const symbol = token.substring(1, dashIndex);
    const localeCode = token.substring(dashIndex + 1);
    const resolvedCulture = DisplayTextCulture.ResolveLocaleFromLcid(localeCode);
    let culture = resolvedCulture === null ? fallbackCulture : resolvedCulture;
    let replacement = "";

    if (localeCode.toLowerCase() === "f800") {
      replacement = DisplayTextCulture.GetLongDatePattern(fallbackCulture);
      culture = fallbackCulture;
      return { Replacement: replacement, Culture: culture };
    }

    if (localeCode.toLowerCase() === "f400") {
      replacement = DisplayTextCulture.NormalizeName(fallbackCulture).toLowerCase().startsWith("zh")
        ? "H:mm:ss"
        : DisplayTextCulture.GetLongTimePattern(fallbackCulture);
      culture = fallbackCulture;
      return { Replacement: replacement, Culture: culture };
    }

    if (symbol.length > 0) {
      replacement = DisplayTextLocaleSupport.QuoteLiteral(symbol);
    }

    return { Replacement: replacement, Culture: culture };
  }

  private static QuoteLiteral(value: string): string {
    let builder = '"';
    for (let index = 0; index < value.length; index++) {
      if (value[index] === '"') {
        builder += '"';
      }

      builder += value[index];
    }

    builder += '"';
    return builder;
  }
}
