import { RgbaColor } from "./ColorValue";

interface TableStyleReference {
  Name: string;
}

interface ListObject {
  StartRow: number;
  EndRow: number;
  StartColumn: number;
  EndColumn: number;
  ShowHeaderRow?: boolean;
  ShowTotals?: boolean;
  ShowRowStripes?: boolean;
  TableStyleName?: string;
}

/**
 * Applies Excel built-in TableStyleMedium2 to a cell, header, totals row, or
 * stripe row.  Custom styles declared via `RawTableStylesXml` are parsed but
 * this port only wires up the standard style – extension points are left as
 * TODO markers.
 */
export class TableStyleResolver {
  private static readonly HeaderFill = new RgbaColor(78, 110, 255);     // #FF4E6EFF
  private static readonly DataFill = new RgbaColor(255, 255, 255, 255);  // #FFFFFFFF
  private static readonly StripeFill = new RgbaColor(245, 230, 192, 255);
  private static readonly BorderFill = new RgbaColor(205, 192, 159, 255);
  private static readonly LightText = new RgbaColor(255, 255, 255, 255);

  Resolve(
    listObjects: ListObject[] | undefined | null,
    row: number,
    column: number,
    baseStyle: { Pattern?: "solid" | "none"; ForegroundColor?: RgbaColor; BackgroundColor?: RgbaColor; Font?: { Bold?: boolean; Color?: RgbaColor }; Borders?: { Left?: { Style?: number; Color?: RgbaColor }; Right?: { Style?: number; Color?: RgbaColor }; Top?: { Style?: number; Color?: RgbaColor }; Bottom?: { Style?: number; Color?: RgbaColor } } },
  ): { Pattern?: "solid" | "none"; ForegroundColor?: RgbaColor; BackgroundColor?: RgbaColor; Font?: { Bold?: boolean; Color?: RgbaColor }; Borders?: { Left?: { Style?: number; Color?: RgbaColor }; Right?: { Style?: number; Color?: RgbaColor }; Top?: { Style?: number; Color?: RgbaColor }; Bottom?: { Style?: number; Color?: RgbaColor } } } {
    const defaultStyle = baseStyle ?? {};
    if (!listObjects || listObjects.length === 0) return defaultStyle;
    for (const table of listObjects) {
      if (row >= table.StartRow && row <= table.EndRow && column >= table.StartColumn && column <= table.EndColumn) {
        return this.applyTableStyle(table, row, defaultStyle);
      }
    }
    return defaultStyle;
  }

  private applyTableStyle(
    table: ListObject,
    row: number,
    base: { Pattern?: "solid" | "none"; ForegroundColor?: RgbaColor; BackgroundColor?: RgbaColor; Font?: { Bold?: boolean; Color?: RgbaColor }; Borders?: { Left?: { Style?: number; Color?: RgbaColor }; Right?: { Style?: number; Color?: RgbaColor }; Top?: { Style?: number; Color?: RgbaColor }; Bottom?: { Style?: number; Color?: RgbaColor } } },
  ): { Pattern?: "solid" | "none"; ForegroundColor?: RgbaColor; BackgroundColor?: RgbaColor; Font?: { Bold?: boolean; Color?: RgbaColor }; Borders?: { Left?: { Style?: number; Color?: RgbaColor }; Right?: { Style?: number; Color?: RgbaColor }; Top?: { Style?: number; Color?: RgbaColor }; Bottom?: { Style?: number; Color?: RgbaColor } } } {
    const result = { ...base, Borders: deepCloneBorders(base.Borders) };
    if (table.ShowHeaderRow && row === table.StartRow) {
      result.Pattern = "solid";
      result.ForegroundColor = TableStyleResolver.HeaderFill;
      result.BackgroundColor = TableStyleResolver.HeaderFill;
      result.Font = { ...result.Font, Bold: true, Color: TableStyleResolver.LightText };
      applyBorders(result, TableStyleResolver.LightText);
      return result;
    }
    if (table.ShowTotals && row === table.EndRow) {
      result.Pattern = "solid";
      result.ForegroundColor = TableStyleResolver.StripeFill;
      result.BackgroundColor = TableStyleResolver.StripeFill;
      result.Font = { ...result.Font, Bold: true };
      applyBorders(result, TableStyleResolver.BorderFill);
      return result;
    }
    result.Pattern = "solid";
    result.ForegroundColor = TableStyleResolver.DataFill;
    result.BackgroundColor = TableStyleResolver.DataFill;
    applyBorders(result, TableStyleResolver.BorderFill);
    if (table.ShowRowStripes) {
      const dataStartRow = table.StartRow + (table.ShowHeaderRow ? 1 : 0);
      const dataEndRow = table.EndRow - (table.ShowTotals ? 1 : 0);
      if (row >= dataStartRow && row <= dataEndRow) {
        const stripeIndex = row - dataStartRow;
        if (stripeIndex % 2 === 0) {
          result.ForegroundColor = TableStyleResolver.StripeFill;
          result.BackgroundColor = TableStyleResolver.StripeFill;
        }
      }
    }
    return result;
  }
}

function applyBorders(
  style: { Borders?: { Left?: { Style?: number; Color?: RgbaColor }; Right?: { Style?: number; Color?: RgbaColor }; Top?: { Style?: number; Color?: RgbaColor }; Bottom?: { Style?: number; Color?: RgbaColor } } },
  color: RgbaColor,
): void {
  style.Borders = style.Borders ?? {};
  for (const key of ["Left", "Right", "Top", "Bottom"] as const) {
    style.Borders[key] = { Style: 1, Color: color };
  }
}

function deepCloneBorders(
  borders?: { Left?: { Style?: number; Color?: RgbaColor }; Right?: { Style?: number; Color?: RgbaColor }; Top?: { Style?: number; Color?: RgbaColor }; Bottom?: { Style?: number; Color?: RgbaColor } },
): { Left?: { Style?: number; Color?: RgbaColor }; Right?: { Style?: number; Color?: RgbaColor }; Top?: { Style?: number; Color?: RgbaColor }; Bottom?: { Style?: number; Color?: RgbaColor } } {
  if (!borders) return {};
  return {
    Left: borders.Left ? { Style: borders.Left.Style, Color: borders.Left.Color } : undefined,
    Right: borders.Right ? { Style: borders.Right.Style, Color: borders.Right.Color } : undefined,
    Top: borders.Top ? { Style: borders.Top.Style, Color: borders.Top.Color } : undefined,
    Bottom: borders.Bottom ? { Style: borders.Bottom.Style, Color: borders.Bottom.Color } : undefined,
  };
}
