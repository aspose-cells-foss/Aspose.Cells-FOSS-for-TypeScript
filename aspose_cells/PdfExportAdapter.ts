import { LegacyWorkbook } from "./workbook";
import type { LegacyWorksheet } from "./worksheet";
import { cellRef } from "./util";
import { WorkbookPdfExporter } from "./rendering/WorkbookPdfExporter";
import { RenderColor } from "./rendering/RenderColor";
import {
  ColorValue as RenderColorValue,
  RgbaColor,
} from "./rendering/ColorValue";
import { PageMargins } from "./rendering/types";
import type { PageSetupData } from "./rendering/types";
import {
  CellValue,
  PdfSaveOptions as LegacyPdfSaveOptions,
  SaveFormat as LegacySaveFormat,
} from "./types";
import type {
  Alignment,
  Border,
  BorderLine,
  Fill,
  Font as LegacyFont,
  ImageInfo,
  Protection,
  ShapeInfo,
  Style as LegacyStyle,
} from "./types";
import type { SaveOptions } from "./SaveOptions";
import type { PdfSaveOptions } from "./PdfSaveOptions";
import type { WorkbookModel } from "./Core/WorkbookModel";
import type { WorksheetModel } from "./Core/WorksheetModel";
import type { StyleValue } from "./Core/StyleValue";
import type { ColorValue } from "./Core/ColorValue";
import type { CellRecord } from "./Core/CellRecord";
import { FillPatternKind } from "./Core/FillPatternKind";
import { BorderStyle } from "./Core/BorderStyle";
import { HorizontalAlignment } from "./Core/HorizontalAlignment";
import { VerticalAlignment } from "./Core/VerticalAlignment";
import { FontUnderlineType } from "./FontUnderlineType";
import { SheetVisibility } from "./Core/SheetVisibility";

/** EMU per pixel, used only to fill the ImageInfo width/height hints. */
const EMU_PER_PIXEL = 9525;

/**
 * Builds the legacy object graph consumed by the PDF renderer from the Core workbook model.
 *
 * The rendering pipeline (WorkbookPdfExporter and its helpers) reads LegacyWorkbook /
 * LegacyWorksheet / LegacyCell / Style values. This adapter mirrors the .NET PDF bridge,
 * which reads the Core model directly, so that the public Workbook API can reuse the
 * existing renderer without changing it.
 */
export class PdfExportAdapter {
  /** Built-in number format ids defined by ECMA-376. */
  private static readonly BuiltInNumberFormats: Record<number, string> = {
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "mm-dd-yy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yy h:mm",
    37: "#,##0 ;(#,##0)",
    38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)",
    40: "#,##0.00;[Red](#,##0.00)",
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    48: "##0.0E+0",
    49: "@",
  };

  /** Converts the Core workbook into the legacy graph expected by the PDF renderer. */
  static toLegacyWorkbook(model: WorkbookModel): LegacyWorkbook {
    const workbook = new LegacyWorkbook(false);
    const colors = RenderColor.fromWorkbook(model.RawThemeXml ?? undefined);

    for (const sheetModel of model.Worksheets) {
      if (sheetModel.Visibility !== SheetVisibility.Visible) {
        continue;
      }

      const worksheet = workbook.worksheets.addWorksheet(sheetModel.Name);
      PdfExportAdapter.applyPageSetup(worksheet, sheetModel);
      PdfExportAdapter.applySheetFormat(worksheet, sheetModel);
      PdfExportAdapter.applyMergeRegions(worksheet, sheetModel);
      PdfExportAdapter.applyCells(worksheet, sheetModel, colors);
      PdfExportAdapter.applyPictures(worksheet, sheetModel);
      PdfExportAdapter.applyShapes(worksheet, sheetModel);
      PdfExportAdapter.applyCharts(worksheet, sheetModel);
    }

    return workbook;
  }

  /** Maps the public PdfSaveOptions fields onto the legacy renderer options. */
  static toLegacySaveOptions(options: SaveOptions): LegacyPdfSaveOptions {
    const legacy = new LegacyPdfSaveOptions();
    legacy.SaveFormat = LegacySaveFormat.PDF;

    const source = options as Partial<PdfSaveOptions>;
    legacy.OnePagePerSheet = source.OnePagePerSheet === true;
    legacy.AllColumnsInOnePagePerSheet = source.AllColumnsInOnePagePerSheet === true;

    if (typeof source.DefaultFont === "string" && source.DefaultFont.length > 0) {
      legacy.DefaultFont = source.DefaultFont;
    }

    if (source.ConditionalFormattingEvaluationDate instanceof Date) {
      legacy.ConditionalFormattingEvaluationDate = source.ConditionalFormattingEvaluationDate;
    }

    return legacy;
  }

  /**
   * Starts an asynchronous PDF export to a file path.
   *
   * Save is synchronous, so the returned promise is not awaited. The open write stream keeps
   * the Node event loop alive until the file is complete, which is what makes the fire-and-forget
   * call safe for callers that exit naturally after Save returns.
   */
  static exportToFile(model: WorkbookModel, outputPath: string, options: SaveOptions): void {
    const workbook = PdfExportAdapter.toLegacyWorkbook(model);
    const legacyOptions = PdfExportAdapter.toLegacySaveOptions(options);
    WorkbookPdfExporter.export(workbook, outputPath, legacyOptions).catch(() => {
      // The synchronous public API has no channel for reporting asynchronous stream failures.
    });
  }

  /** Renders the PDF into memory and forwards the bytes to a caller-provided stream. */
  static exportToStream(
    model: WorkbookModel,
    target: { write(buffer: Buffer): void },
    options: SaveOptions,
  ): void {
    const workbook = PdfExportAdapter.toLegacyWorkbook(model);
    const legacyOptions = PdfExportAdapter.toLegacySaveOptions(options);
    WorkbookPdfExporter.exportToBuffer(workbook, legacyOptions)
      .then((buffer) => target.write(buffer))
      .catch(() => {
        // The synchronous public API has no channel for reporting asynchronous stream failures.
      });
  }

  private static applyPageSetup(worksheet: LegacyWorksheet, model: WorksheetModel): void {
    const setup = model.PageSetup;
    const margins = new PageMargins();
    margins.left = setup.Margins.Left;
    margins.right = setup.Margins.Right;
    margins.top = setup.Margins.Top;
    margins.bottom = setup.Margins.Bottom;
    margins.header = setup.Margins.Header;
    margins.footer = setup.Margins.Footer;

    const data: PageSetupData = {
      paperSize: setup.PaperSize,
      orientation: setup.Orientation as unknown as PageSetupData["orientation"],
      margins,
      printArea: setup.PrintArea ?? undefined,
      scale: setup.Scale ?? undefined,
      fitToPage: setup.FitToPage,
      fitToWidth: setup.FitToWidth ?? undefined,
      fitToHeight: setup.FitToHeight ?? undefined,
      horizontalPageBreaks: [...setup.HorizontalPageBreaks],
      verticalPageBreaks: [...setup.VerticalPageBreaks],
      pageOrderOverThenDown: setup.PageOrderOverThenDown,
    };

    // The renderer resolves page setup from this optional slot when present.
    (worksheet as unknown as { __pdfPageSetup?: PageSetupData }).__pdfPageSetup = data;
  }

  private static applySheetFormat(worksheet: LegacyWorksheet, model: WorksheetModel): void {
    if (model.DefaultColumnWidth !== null && model.DefaultColumnWidth > 0) {
      worksheet.defaultColumnWidth = model.DefaultColumnWidth;
    }

    if (model.DefaultRowHeight !== null && model.DefaultRowHeight > 0) {
      worksheet.defaultRowHeight = model.DefaultRowHeight;
    }

    for (const column of model.Columns) {
      const first = Math.max(0, column.MinColumnIndex);
      const last = Math.max(first, column.MaxColumnIndex);

      if (column.Hidden) {
        for (let index = first; index <= last; index++) {
          worksheet.setColumnWidth(index, 0);
        }
        continue;
      }

      if (column.Width !== null && column.Width > 0) {
        for (let index = first; index <= last; index++) {
          worksheet.setColumnWidth(index, column.Width);
        }
      }
    }

    for (const [rowIndex, row] of model.Rows) {
      if (row.Hidden) {
        worksheet.setRowHidden(rowIndex, true);
        continue;
      }

      if (row.Height !== null && row.Height > 0) {
        worksheet.setRowHeight(rowIndex, row.Height);
      }
    }
  }

  private static applyMergeRegions(worksheet: LegacyWorksheet, model: WorksheetModel): void {
    for (const merge of model.MergeRegions) {
      const first = cellRef(merge.FirstRow, merge.FirstColumn);
      const lastRow = merge.FirstRow + Math.max(1, merge.TotalRows) - 1;
      const lastColumn = merge.FirstColumn + Math.max(1, merge.TotalColumns) - 1;
      worksheet.mergeCells(`${first}:${cellRef(lastRow, lastColumn)}`);
    }
  }

  private static applyCells(
    worksheet: LegacyWorksheet,
    model: WorksheetModel,
    colors: RenderColor,
  ): void {
    const hiddenRows = new Set<number>();
    for (const [rowIndex, row] of model.Rows) {
      if (row.Hidden) {
        hiddenRows.add(rowIndex);
      }
    }

    const hiddenColumns = new Set<number>();
    for (const column of model.Columns) {
      if (!column.Hidden) {
        continue;
      }

      const first = Math.max(0, column.MinColumnIndex);
      const last = Math.max(first, column.MaxColumnIndex);
      for (let index = first; index <= last; index++) {
        hiddenColumns.add(index);
      }
    }

    for (const [address, record] of model.Cells) {
      if (hiddenRows.has(address.RowIndex) || hiddenColumns.has(address.ColumnIndex)) {
        continue;
      }

      const cell = worksheet.cells.getOrCreate(address.RowIndex, address.ColumnIndex);
      cell.value = PdfExportAdapter.toLegacyCellValue(record.Value);
      if (record.Formula !== null && record.Formula.length > 0) {
        cell.formula = record.Formula;
      }
      cell.style = PdfExportAdapter.toLegacyStyle(record.Style, colors);
      cell.setStyleIndex(0);
    }
  }

  private static applyPictures(worksheet: LegacyWorksheet, model: WorksheetModel): void {
    for (let index = 0; index < model.Pictures.length; index++) {
      const picture = model.Pictures[index];
      if (picture.ImageData.length === 0) {
        continue;
      }

      const name = picture.Name.length > 0 ? picture.Name : `Picture ${index + 1}`;
      const image: ImageInfo = {
        name,
        data: Buffer.from(picture.ImageData),
        width: picture.ExtentCx > 0 ? picture.ExtentCx / EMU_PER_PIXEL : 0,
        height: picture.ExtentCy > 0 ? picture.ExtentCy / EMU_PER_PIXEL : 0,
        ext: (picture.ImageExtension || "png").toLowerCase(),
      };
      worksheet.addImage(image);

      worksheet.addShape({
        name,
        type: "picture",
        fromRow: picture.UpperLeftRow,
        fromRowOff: picture.UpperLeftRowOffset,
        fromCol: picture.UpperLeftColumn,
        fromColOff: picture.UpperLeftColumnOffset,
        toRow: picture.LowerRightRow,
        toRowOff: picture.LowerRightRowOffset,
        toCol: picture.LowerRightColumn,
        toColOff: picture.LowerRightColumnOffset,
        imageName: name,
      } as ShapeInfo);
    }
  }

  private static applyShapes(worksheet: LegacyWorksheet, model: WorksheetModel): void {
    for (let index = 0; index < model.Shapes.length; index++) {
      const shape = model.Shapes[index];
      worksheet.addShape({
        name: shape.Name.length > 0 ? shape.Name : `Shape ${index + 1}`,
        type: shape.GeometryType,
        fromRow: shape.UpperLeftRow,
        fromRowOff: shape.UpperLeftRowOffset,
        fromCol: shape.UpperLeftColumn,
        fromColOff: shape.UpperLeftColumnOffset,
        toRow: shape.LowerRightRow,
        toRowOff: shape.LowerRightRowOffset,
        toCol: shape.LowerRightColumn,
        toColOff: shape.LowerRightColumnOffset,
      });
    }
  }

  private static applyCharts(worksheet: LegacyWorksheet, model: WorksheetModel): void {
    for (let index = 0; index < model.Charts.length; index++) {
      const chart = model.Charts[index];
      worksheet.addShape({
        name: chart.Name.length > 0 ? chart.Name : `Chart ${index + 1}`,
        type: "chart",
        fromRow: chart.UpperLeftRow,
        fromRowOff: chart.UpperLeftRowOffset,
        fromCol: chart.UpperLeftColumn,
        fromColOff: chart.UpperLeftColumnOffset,
        toRow: chart.LowerRightRow,
        toRowOff: chart.LowerRightRowOffset,
        toCol: chart.LowerRightColumn,
        toColOff: chart.LowerRightColumnOffset,
      });
    }
  }

  private static toLegacyCellValue(value: CellRecord["Value"]): CellValue {
    if (value === null || value === undefined) {
      return null;
    }

    if (
      typeof value === "string" ||
      typeof value === "number" ||
      typeof value === "boolean" ||
      value instanceof Date
    ) {
      return value;
    }

    return String(value);
  }

  private static toLegacyStyle(style: StyleValue, colors: RenderColor): LegacyStyle {
    const legacy: LegacyStyle = {};

    const font: LegacyFont = {};
    if (style.Font.Name) {
      font.name = style.Font.Name;
    }
    if (style.Font.Size > 0) {
      font.size = style.Font.Size;
    }
    if (style.Font.Bold) {
      font.bold = true;
    }
    if (style.Font.Italic) {
      font.italic = true;
    }
    if (style.Font.Underline !== FontUnderlineType.None) {
      font.underline = true;
    }
    if (style.Font.StrikeThrough) {
      font.strikeout = true;
    }
    const fontColor = PdfExportAdapter.toHexColor(style.Font.Color, colors, RgbaColor.black);
    if (fontColor !== undefined) {
      font.color = fontColor;
    }
    legacy.font = font;

    if (style.Pattern !== FillPatternKind.None) {
      const fill: Fill = { patternType: PdfExportAdapter.toFillPattern(style.Pattern) };
      const foreground = PdfExportAdapter.toHexColor(
        style.ForegroundColor,
        colors,
        RgbaColor.transparent,
      );
      if (foreground !== undefined) {
        fill.fgColor = foreground;
      }
      const background = PdfExportAdapter.toHexColor(
        style.BackgroundColor,
        colors,
        RgbaColor.transparent,
      );
      if (background !== undefined) {
        fill.bgColor = background;
      }
      legacy.fill = fill;
    }

    const border: Border = {};
    const left = PdfExportAdapter.toBorderLine(style.Borders.Left, colors);
    const right = PdfExportAdapter.toBorderLine(style.Borders.Right, colors);
    const top = PdfExportAdapter.toBorderLine(style.Borders.Top, colors);
    const bottom = PdfExportAdapter.toBorderLine(style.Borders.Bottom, colors);
    if (left !== undefined) border.left = left;
    if (right !== undefined) border.right = right;
    if (top !== undefined) border.top = top;
    if (bottom !== undefined) border.bottom = bottom;
    if (left !== undefined || right !== undefined || top !== undefined || bottom !== undefined) {
      legacy.border = border;
    }

    const alignment: Alignment = {};
    const horizontal = PdfExportAdapter.toHorizontalAlignment(style.Alignment.Horizontal);
    if (horizontal !== undefined) {
      alignment.horizontal = horizontal;
    }
    alignment.vertical = PdfExportAdapter.toVerticalAlignment(style.Alignment.Vertical);
    if (style.Alignment.WrapText) {
      alignment.wrapText = true;
    }
    legacy.alignment = alignment;

    const numberFormat = PdfExportAdapter.toNumberFormat(style);
    if (numberFormat !== undefined) {
      legacy.numberFormat = numberFormat;
    }

    const protection: Protection = {};
    if (!style.Protection.IsLocked) {
      protection.locked = false;
    }
    if (style.Protection.IsHidden) {
      protection.hidden = true;
    }
    if (protection.locked !== undefined || protection.hidden !== undefined) {
      legacy.protection = protection;
    }

    return legacy;
  }

  private static toBorderLine(
    side: { Style: BorderStyle; Color: ColorValue },
    colors: RenderColor,
  ): BorderLine | undefined {
    if (side.Style === BorderStyle.None) {
      return undefined;
    }

    const line: BorderLine = {
      style: PdfExportAdapter.toBorderStyleName(side.Style) as BorderLine["style"],
    };
    const color = PdfExportAdapter.toHexColor(side.Color, colors, RgbaColor.black);
    if (color !== undefined) {
      line.color = color;
    }
    return line;
  }

  private static toHexColor(
    color: ColorValue,
    colors: RenderColor,
    fallback: RgbaColor,
  ): string | undefined {
    const value = new RenderColorValue();
    value.A = color.A;
    value.R = color.R;
    value.G = color.G;
    value.B = color.B;

    if (color.ThemeIndex !== null) {
      value.ThemeIndex = color.ThemeIndex;
    }
    if (color.Tint !== null) {
      value.Tint = color.Tint;
    }
    if (color.Indexed !== null) {
      value.Indexed = color.Indexed;
    }

    const resolved = colors.resolve(value, fallback);
    if (resolved.a === 0) {
      return undefined;
    }

    const hex = (channel: number): string =>
      Math.max(0, Math.min(255, channel)).toString(16).padStart(2, "0");
    return `#${hex(resolved.r)}${hex(resolved.g)}${hex(resolved.b)}`;
  }

  private static toFillPattern(pattern: FillPatternKind): Fill["patternType"] {
    switch (pattern) {
      case FillPatternKind.Solid:
        return "solid";
      case FillPatternKind.MediumGray:
      case FillPatternKind.DarkHorizontal:
      case FillPatternKind.DarkVertical:
      case FillPatternKind.DarkDown:
      case FillPatternKind.DarkUp:
      case FillPatternKind.DarkGrid:
      case FillPatternKind.DarkTrellis:
        return "darkGray";
      case FillPatternKind.DarkGray:
        return "darkGray";
      case FillPatternKind.LightGray:
      case FillPatternKind.LightHorizontal:
      case FillPatternKind.LightVertical:
      case FillPatternKind.LightDown:
      case FillPatternKind.LightUp:
      case FillPatternKind.LightGrid:
      case FillPatternKind.LightTrellis:
        return "lightGray";
      case FillPatternKind.Gray125:
      case FillPatternKind.Gray0625:
        return "gray125";
      default:
        return "mediumGray";
    }
  }

  private static toBorderStyleName(style: BorderStyle): string {
    switch (style) {
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
        return "dashed";
      case BorderStyle.DashDot:
        return "dashDot";
      case BorderStyle.MediumDashDot:
        return "dashDot";
      case BorderStyle.DashDotDot:
        return "dashDotDot";
      case BorderStyle.MediumDashDotDot:
        return "dashDotDot";
      case BorderStyle.SlantedDashDot:
        return "slantDashDot";
      default:
        return "thin";
    }
  }

  private static toHorizontalAlignment(
    alignment: HorizontalAlignment,
  ): Alignment["horizontal"] {
    switch (alignment) {
      case HorizontalAlignment.Left:
        return "left";
      case HorizontalAlignment.Center:
        return "center";
      case HorizontalAlignment.Right:
        return "right";
      case HorizontalAlignment.Fill:
        return "fill";
      case HorizontalAlignment.Justify:
      case HorizontalAlignment.Distributed:
        return "justify";
      case HorizontalAlignment.CenterContinuous:
        return "centerContinuous";
      default:
        return "general";
    }
  }

  private static toVerticalAlignment(
    alignment: VerticalAlignment,
  ): NonNullable<Alignment["vertical"]> {
    switch (alignment) {
      case VerticalAlignment.Center:
        return "center";
      case VerticalAlignment.Top:
        return "top";
      case VerticalAlignment.Justify:
      case VerticalAlignment.Distributed:
        return "justify";
      default:
        return "bottom";
    }
  }

  private static toNumberFormat(style: StyleValue): string | undefined {
    if (style.NumberFormat.Custom !== null && style.NumberFormat.Custom.length > 0) {
      return style.NumberFormat.Custom;
    }

    return PdfExportAdapter.BuiltInNumberFormats[style.NumberFormat.Number];
  }
}
