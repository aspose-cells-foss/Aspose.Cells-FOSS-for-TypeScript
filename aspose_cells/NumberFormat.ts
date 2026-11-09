/**
 * Provides number format operations.
 * Ported from Aspose.Cells_FOSS.NumberFormat.
 */
export class NumberFormat {
  private static readonly BuiltInFormats: ReadonlyMap<number, string> = NumberFormat.createBuiltInFormats();

  /** Gets the built-in format code for a format identifier. */
  public static GetBuiltInFormat(formatId: number): string {
    const formatCode = NumberFormat.BuiltInFormats.get(formatId);
    return formatCode === undefined ? "General" : formatCode;
  }

  /** Returns whether the specified format code is a built-in format. */
  public static IsBuiltInFormat(formatCode: string | null): boolean {
    return NumberFormat.GetBuiltInFormatId(formatCode) !== null;
  }

  /** Gets the built-in format identifier for a format code. */
  public static GetBuiltInFormatId(formatCode: string | null): number | null {
    if (formatCode === null || formatCode === undefined || formatCode.trim().length === 0) {
      return 0;
    }

    const trimmed = formatCode.trim();
    for (const [id, code] of NumberFormat.BuiltInFormats) {
      if (code === trimmed) {
        return id;
      }
    }

    return null;
  }

  /** Resolves the effective format code for a style. */
  public static ResolveFormatCode(number: number, custom: string | null): string {
    if (custom !== null && custom !== undefined && custom.length > 0) {
      return custom;
    }

    return NumberFormat.GetBuiltInFormat(number);
  }

  private static createBuiltInFormats(): ReadonlyMap<number, string> {
    const formats = new Map<number, string>();
    formats.set(0, "General");
    formats.set(1, "0");
    formats.set(2, "0.00");
    formats.set(3, "#,##0");
    formats.set(4, "#,##0.00");
    formats.set(5, "$#,##0_);($#,##0)");
    formats.set(6, "$#,##0_);[Red]($#,##0)");
    formats.set(7, "$#,##0.00_);($#,##0.00)");
    formats.set(8, "$#,##0.00_);[Red]($#,##0.00)");
    formats.set(9, "0%");
    formats.set(10, "0.00%");
    formats.set(11, "0.00E+00");
    formats.set(12, "# ?/?");
    formats.set(13, "# ??/??");
    formats.set(14, "mm-dd-yy");
    formats.set(15, "d-mmm-yy");
    formats.set(16, "d-mmm");
    formats.set(17, "mmm-yy");
    formats.set(18, "h:mm AM/PM");
    formats.set(19, "h:mm:ss AM/PM");
    formats.set(20, "h:mm");
    formats.set(21, "h:mm:ss");
    formats.set(22, "m/d/yy h:mm");
    formats.set(37, "#,##0_);(#,##0)");
    formats.set(38, "#,##0_);[Red](#,##0)");
    formats.set(39, "#,##0.00_);(#,##0.00)");
    formats.set(40, "#,##0.00_);[Red](#,##0.00)");
    formats.set(41, "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)");
    formats.set(42, "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)");
    formats.set(43, "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)");
    formats.set(44, "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)");
    formats.set(45, "mm:ss");
    formats.set(46, "[h]:mm:ss");
    formats.set(47, "mm:ss.0");
    formats.set(48, "##0.0E+0");
    formats.set(49, "@");
    formats.set(50, "General");

    for (let id = 51; id <= 82; id++) {
      const red = (id - 51) % 2 === 1;
      formats.set(id, red ? "0_);[Red](0)" : "0_);(0)");
    }

    return formats;
  }
}
