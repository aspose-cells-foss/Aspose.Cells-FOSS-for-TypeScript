import {
  BorderStyleType,
  CellArea,
  Color,
  ConditionalFormattingCollection,
  FillPattern,
  FontUnderlineType,
  FormatConditionType,
  HorizontalAlignmentType,
  OperatorType,
  PageOrientationType,
  PaperSizeType,
  Style,
  Validation,
  ValidationAlertType,
  ValidationCollection,
  ValidationType,
  VerticalAlignmentType,
  VisibilityType,
  Workbook,
  Worksheet,
} from "../../aspose_cells";
import {
  AssertEx,
  MemoryStream,
  TemporaryDirectory,
  ZipPackageHelper,
  resolveRepositoryRoot,
  getOrCreateOutputDirectory,
} from "./TestInfrastructure";

export { AutoFilterScenarioFactory } from "./AutoFilterScenarioFactory";

/** Returns the current UTC date with the time-of-day zeroed. */
function todayUtcDate(): Date {
  const now = new Date();
  return new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));
}

/** Scenario factory for mixed cell value workbooks. */
export class WorkbookScenarioFactory {
  public static CreateMixedCellWorkbook(useDate1904: boolean = false): Workbook {
    const workbook = new Workbook();
    workbook.Settings.Date1904 = useDate1904;

    const sheet = workbook.Worksheets[0];
    sheet.Name = "Data";
    sheet.Cells["A1"].PutValue("Hello");
    sheet.Cells["B1"].PutValue(123);
    sheet.Cells["C1"].PutValue(true);
    sheet.Cells["D1"].PutValue(12.5);
    sheet.Cells["E1"].PutValue(6.02214076e23);
    sheet.Cells["F1"].PutValue(new Date(Date.UTC(2024, 4, 6, 7, 8, 9)));
    sheet.Cells["G1"].PutValue(20);
    sheet.Cells["G1"].Formula = "=B1*2";
    return workbook;
  }
}

/** Scenario factory for worksheet-level settings. */
export class WorksheetScenarioFactory {
  public static CreateWorksheetSettingsWorkbook(): Workbook {
    const workbook = new Workbook();
    const sheet = workbook.Worksheets[0];
    sheet.Name = "Layout";
    sheet.VisibilityType = VisibilityType.Hidden;
    sheet.TabColor = Color.FromArgb(255, 34, 68, 102);
    sheet.ShowGridlines = false;
    sheet.ShowRowColumnHeaders = false;
    sheet.ShowZeros = false;
    sheet.RightToLeft = true;
    sheet.Zoom = 85;
    sheet.Protect();
    sheet.Protection.Objects = true;
    sheet.Protection.Scenarios = true;
    sheet.Protection.FormatCells = true;
    sheet.Protection.InsertRows = true;
    sheet.Protection.AutoFilter = true;
    sheet.Protection.SelectLockedCells = true;
    sheet.Protection.SelectUnlockedCells = true;

    sheet.Cells["A1"].PutValue("Merged");
    sheet.Cells["C4"].PutValue(99);
    sheet.Cells.Rows[1].Height = 22.5;
    sheet.Cells.Rows[3].IsHidden = true;
    sheet.Cells.Columns[0].Width = 18.25;
    sheet.Cells.Columns[2].IsHidden = true;
    sheet.Cells.Merge(0, 0, 2, 2);

    const visibleSheetIndex = workbook.Worksheets.Add();
    const visibleSheet = workbook.Worksheets[visibleSheetIndex];
    visibleSheet.Name = "Visible";
    visibleSheet.Cells["A1"].PutValue("Visible");
    workbook.Worksheets.ActiveSheetName = "Visible";

    return workbook;
  }

  public static AssertWorksheetSettings(workbook: Workbook): void {
    const sheet = workbook.Worksheets["Layout"];
    AssertEx.notNull(sheet, "Expected the worksheet settings scenario to contain the 'Layout' sheet.");
    AssertEx.equal(VisibilityType.Hidden, sheet.VisibilityType);
    AssertEx.equal(Color.FromArgb(255, 34, 68, 102), sheet.TabColor);
    AssertEx.false(sheet.ShowGridlines);
    AssertEx.false(sheet.ShowRowColumnHeaders);
    AssertEx.false(sheet.ShowZeros);
    AssertEx.true(sheet.RightToLeft);
    AssertEx.equal(85, sheet.Zoom);
    AssertEx.true(sheet.Protection.IsProtected);
    AssertEx.true(sheet.Protection.Objects);
    AssertEx.true(sheet.Protection.Scenarios);
    AssertEx.true(sheet.Protection.FormatCells);
    AssertEx.true(sheet.Protection.InsertRows);
    AssertEx.true(sheet.Protection.AutoFilter);
    AssertEx.true(sheet.Protection.SelectLockedCells);
    AssertEx.true(sheet.Protection.SelectUnlockedCells);
    AssertEx.equal("Merged", sheet.Cells["A1"].StringValue);
    AssertEx.equal(99, Number(sheet.Cells["C4"].Value));
    AssertEx.equal(22.5, sheet.Cells.Rows[1].Height ?? 0);
    AssertEx.true(sheet.Cells.Rows[3].IsHidden);
    AssertEx.equal(18.25, sheet.Cells.Columns[0].Width ?? 0);
    AssertEx.true(sheet.Cells.Columns[2].IsHidden);
    AssertEx.equal(1, sheet.Cells.MergedCells.length);
    AssertEx.equal(0, sheet.Cells.MergedCells[0].StartRow);
    AssertEx.equal(0, sheet.Cells.MergedCells[0].StartColumn);
    AssertEx.equal(2, sheet.Cells.MergedCells[0].EndRow - sheet.Cells.MergedCells[0].StartRow + 1);
    AssertEx.equal(2, sheet.Cells.MergedCells[0].EndColumn - sheet.Cells.MergedCells[0].StartColumn + 1);
  }

  public static AssertWorksheetSettingsScenarioHasVisibleSheet(workbook: Workbook): void {
    AssertEx.true(workbook.Worksheets.Count >= 2);
    const visibleSheet = workbook.Worksheets["Visible"];
    AssertEx.notNull(visibleSheet, "Expected the worksheet settings scenario to contain a visible sheet.");
    AssertEx.equal(VisibilityType.Visible, visibleSheet.VisibilityType);
    AssertEx.equal("Visible", visibleSheet.Cells["A1"].StringValue);
    AssertEx.equal("Visible", workbook.Worksheets.ActiveSheetName);
  }
}

/** Scenario factory for page setup settings. */
export class PageSetupScenarioFactory {
  public static CreatePageSetupWorkbook(): Workbook {
    const workbook = new Workbook();
    const sheet = workbook.Worksheets[0];
    sheet.Name = "Print Sheet";
    sheet.Cells["A1"].PutValue("Title");
    sheet.Cells["C10"].PutValue(42);

    const pageSetup = sheet.PageSetup;
    pageSetup.LeftMarginInch = 0.25;
    pageSetup.RightMarginInch = 0.4;
    pageSetup.TopMarginInch = 0.5;
    pageSetup.BottomMarginInch = 0.6;
    pageSetup.HeaderMarginInch = 0.2;
    pageSetup.FooterMarginInch = 0.22;
    pageSetup.Orientation = PageOrientationType.Landscape;
    pageSetup.PaperSize = PaperSizeType.PaperA4;
    pageSetup.FirstPageNumber = 3;
    pageSetup.Scale = 95;
    pageSetup.FitToPagesWide = 1;
    pageSetup.FitToPagesTall = 2;
    pageSetup.PrintArea = "$A$1:$C$10";
    pageSetup.PrintTitleRows = "$1:$2";
    pageSetup.PrintTitleColumns = "$A:$B";
    pageSetup.LeftHeader = "Left Header";
    pageSetup.CenterHeader = "Center Header";
    pageSetup.RightHeader = "Right Header";
    pageSetup.LeftFooter = "Left Footer";
    pageSetup.CenterFooter = "Center Footer";
    pageSetup.RightFooter = "Right Footer";
    pageSetup.PrintGridlines = true;
    pageSetup.PrintHeadings = true;
    pageSetup.CenterHorizontally = true;
    pageSetup.CenterVertically = true;
    pageSetup.AddHorizontalPageBreak(4);
    pageSetup.AddHorizontalPageBreak(7);
    pageSetup.AddVerticalPageBreak(2);

    return workbook;
  }

  public static AssertPageSetup(workbook: Workbook): void {
    const pageSetup = workbook.Worksheets[0].PageSetup;
    AssertEx.equal(0.25, pageSetup.LeftMarginInch);
    AssertEx.equal(0.4, pageSetup.RightMarginInch);
    AssertEx.equal(0.5, pageSetup.TopMarginInch);
    AssertEx.equal(0.6, pageSetup.BottomMarginInch);
    AssertEx.equal(0.2, pageSetup.HeaderMarginInch);
    AssertEx.equal(0.22, pageSetup.FooterMarginInch);
    AssertEx.equal(PageOrientationType.Landscape, pageSetup.Orientation);
    AssertEx.equal(PaperSizeType.PaperA4, pageSetup.PaperSize);
    AssertEx.equal(3, pageSetup.FirstPageNumber ?? 0);
    AssertEx.equal(95, pageSetup.Scale ?? 0);
    AssertEx.equal(1, pageSetup.FitToPagesWide ?? 0);
    AssertEx.equal(2, pageSetup.FitToPagesTall ?? 0);
    AssertEx.equal("$A$1:$C$10", pageSetup.PrintArea);
    AssertEx.equal("$1:$2", pageSetup.PrintTitleRows);
    AssertEx.equal("$A:$B", pageSetup.PrintTitleColumns);
    AssertEx.equal("Left Header", pageSetup.LeftHeader);
    AssertEx.equal("Center Header", pageSetup.CenterHeader);
    AssertEx.equal("Right Header", pageSetup.RightHeader);
    AssertEx.equal("Left Footer", pageSetup.LeftFooter);
    AssertEx.equal("Center Footer", pageSetup.CenterFooter);
    AssertEx.equal("Right Footer", pageSetup.RightFooter);
    AssertEx.true(pageSetup.PrintGridlines);
    AssertEx.true(pageSetup.PrintHeadings);
    AssertEx.true(pageSetup.CenterHorizontally);
    AssertEx.true(pageSetup.CenterVertically);
    AssertEx.equal(2, pageSetup.HorizontalPageBreaks.length);
    AssertEx.equal(4, pageSetup.HorizontalPageBreaks[0]);
    AssertEx.equal(7, pageSetup.HorizontalPageBreaks[1]);
    AssertEx.equal(1, pageSetup.VerticalPageBreaks.length);
    AssertEx.equal(2, pageSetup.VerticalPageBreaks[0]);
    AssertEx.equal("Title", workbook.Worksheets[0].Cells["A1"].StringValue);
    AssertEx.equal("42", workbook.Worksheets[0].Cells["C10"].StringValue);
  }
}

/** Scenario factory for worksheet hyperlinks. */
export class HyperlinkScenarioFactory {
  public static CreateHyperlinkWorkbook(): Workbook {
    const workbook = new Workbook();
    const sheet = workbook.Worksheets[0];
    sheet.Name = "Links";
    workbook.Worksheets.Add("Target Sheet");
    workbook.Worksheets[1].Cells.get(2, 2).PutValue("Target");

    sheet.Cells["A1"].PutValue("Docs");
    const external = sheet.Hyperlinks[sheet.Hyperlinks.Add("A1", 1, 1, "https://example.com/docs?q=1")];
    external.TextToDisplay = "Docs";
    external.ScreenTip = "External docs";

    sheet.Cells["B2"].PutValue("Jump");
    const internalLink = sheet.Hyperlinks[sheet.Hyperlinks.Add("B2", 1, 1, "'Target Sheet'!C3")];
    internalLink.TextToDisplay = "Jump";
    internalLink.ScreenTip = "Jump to target";

    const rangeLink = sheet.Hyperlinks[sheet.Hyperlinks.Add("C4", 2, 2, "mailto:test@example.com")];
    rangeLink.TextToDisplay = "Mail";
    rangeLink.ScreenTip = "Send mail";

    return workbook;
  }

  public static AssertHyperlinks(workbook: Workbook): void {
    const sheet = workbook.Worksheets[0];
    AssertEx.equal(3, sheet.Hyperlinks.Count);

    const external = sheet.Hyperlinks[0];
    AssertEx.equal("A1", external.Area);
    AssertEx.equal("https://example.com/docs?q=1", external.Address);
    AssertEx.equal("External docs", external.ScreenTip);
    AssertEx.equal("Docs", external.TextToDisplay);

    const internalLink = sheet.Hyperlinks[1];
    AssertEx.equal("B2", internalLink.Area);
    AssertEx.equal("'Target Sheet'!C3", internalLink.Address);
    AssertEx.equal("Jump to target", internalLink.ScreenTip);
    AssertEx.equal("Jump", internalLink.TextToDisplay);

    const rangeLink = sheet.Hyperlinks[2];
    AssertEx.equal("C4:D5", rangeLink.Area);
    AssertEx.equal("mailto:test@example.com", rangeLink.Address);
    AssertEx.equal("Send mail", rangeLink.ScreenTip);
    AssertEx.equal("Mail", rangeLink.TextToDisplay);
  }
}

/** Scenario factory for cell style features. */
export class StyleScenarioFactory {
  private static readonly FillPatternShowcase: FillPattern[] = [
    FillPattern.Solid,
    FillPattern.MediumGray,
    FillPattern.DarkGray,
    FillPattern.LightGray,
    FillPattern.Gray125,
    FillPattern.Gray0625,
    FillPattern.DarkHorizontal,
    FillPattern.DarkVertical,
    FillPattern.DarkDown,
    FillPattern.DarkUp,
    FillPattern.DarkGrid,
    FillPattern.DarkTrellis,
    FillPattern.LightHorizontal,
    FillPattern.LightVertical,
    FillPattern.LightDown,
    FillPattern.LightUp,
    FillPattern.LightGrid,
    FillPattern.LightTrellis,
  ];

  private static readonly BorderStyleShowcase: BorderStyleType[] = [
    BorderStyleType.Thin,
    BorderStyleType.Medium,
    BorderStyleType.Thick,
    BorderStyleType.Dotted,
    BorderStyleType.Dashed,
    BorderStyleType.Double,
    BorderStyleType.Hair,
    BorderStyleType.MediumDashed,
    BorderStyleType.DashDot,
    BorderStyleType.MediumDashDot,
    BorderStyleType.DashDotDot,
    BorderStyleType.MediumDashDotDot,
    BorderStyleType.SlantedDashDot,
  ];

  private static readonly HorizontalAlignmentShowcase: HorizontalAlignmentType[] = [
    HorizontalAlignmentType.General,
    HorizontalAlignmentType.Left,
    HorizontalAlignmentType.Center,
    HorizontalAlignmentType.Right,
    HorizontalAlignmentType.Fill,
    HorizontalAlignmentType.Justify,
    HorizontalAlignmentType.CenterContinuous,
    HorizontalAlignmentType.Distributed,
  ];

  private static readonly VerticalAlignmentShowcase: VerticalAlignmentType[] = [
    VerticalAlignmentType.Bottom,
    VerticalAlignmentType.Center,
    VerticalAlignmentType.Top,
    VerticalAlignmentType.Justify,
    VerticalAlignmentType.Distributed,
  ];

  private static readonly NumberFormatShowcase: string[] = [
    "0",
    "0.00",
    "#,##0.00",
    "0%",
    "0.00E+00",
    "# ?/?",
    "[$-409]#,##0.00",
  ];

  private static readonly RotationShowcase: number[] = [
    0,
    45,
    90,
    135,
    180,
    255,
  ];

  private static readonly ReadingOrderShowcase: number[] = [
    0,
    1,
    2,
  ];

  private static readonly FontSizeShowcase: number[] = [
    8,
    10,
    12,
    14,
    18,
    24,
  ];

  private static readonly IndentLevelShowcase: number[] = [
    0,
    1,
    2,
    3,
    4,
  ];

  private static readonly ShowcaseStartColumn = 5;
  private static readonly FillPatternRow = 0;
  private static readonly BorderStyleRow = 1;
  private static readonly HorizontalAlignmentRow = 2;
  private static readonly VerticalAlignmentRow = 3;
  private static readonly FontRow = 4;
  private static readonly NumberFormatRow = 5;
  private static readonly BorderSidesRow = 6;
  private static readonly RotationRow = 7;
  private static readonly ReadingOrderRow = 8;
  private static readonly ProtectionRow = 9;
  private static readonly FontSizeRow = 10;
  private static readonly IndentRow = 11;
  private static readonly WrapShrinkRow = 12;

  public static CreateStyledWorkbook(): Workbook {
    const workbook = new Workbook();
    const sheet = workbook.Worksheets[0];
    sheet.Name = "Styled";

    const primaryCell = sheet.Cells["A1"];
    primaryCell.PutValue(1234.567);
    const primaryStyle = primaryCell.GetStyle();
    StyleScenarioFactory.ApplyPrimaryStyle(primaryStyle);
    primaryCell.SetStyle(primaryStyle);

    const customCell = sheet.Cells["B2"];
    const customStyle = customCell.GetStyle();
    StyleScenarioFactory.ApplyCustomNumberStyle(customStyle);
    customCell.SetStyle(customStyle);

    const dateCell = sheet.Cells["C3"];
    dateCell.PutValue(new Date(2024, 4, 6, 7, 8, 9));
    const dateStyle = dateCell.GetStyle();
    StyleScenarioFactory.ApplyDateStyle(dateStyle);
    dateCell.SetStyle(dateStyle);

    const verticalTextCell = sheet.Cells["D4"];
    verticalTextCell.PutValue("Vertical");
    const verticalTextStyle = verticalTextCell.GetStyle();
    StyleScenarioFactory.ApplyVerticalTextStyle(verticalTextStyle);
    verticalTextCell.SetStyle(verticalTextStyle);

    StyleScenarioFactory.AddShowcaseLabel(sheet, StyleScenarioFactory.FillPatternRow, "Fill Patterns");
    StyleScenarioFactory.AddShowcaseLabel(sheet, StyleScenarioFactory.BorderStyleRow, "Border Styles");
    StyleScenarioFactory.AddShowcaseLabel(sheet, StyleScenarioFactory.HorizontalAlignmentRow, "Horizontal Align");
    StyleScenarioFactory.AddShowcaseLabel(sheet, StyleScenarioFactory.VerticalAlignmentRow, "Vertical Align");
    StyleScenarioFactory.AddShowcaseLabel(sheet, StyleScenarioFactory.FontRow, "Font Settings");
    StyleScenarioFactory.AddShowcaseLabel(sheet, StyleScenarioFactory.NumberFormatRow, "Number Formats");
    StyleScenarioFactory.AddShowcaseLabel(sheet, StyleScenarioFactory.BorderSidesRow, "Border Sides");
    StyleScenarioFactory.AddShowcaseLabel(sheet, StyleScenarioFactory.RotationRow, "Text Rotation");
    StyleScenarioFactory.AddShowcaseLabel(sheet, StyleScenarioFactory.ReadingOrderRow, "Reading Order");
    StyleScenarioFactory.AddShowcaseLabel(sheet, StyleScenarioFactory.ProtectionRow, "Protection");
    StyleScenarioFactory.AddShowcaseLabel(sheet, StyleScenarioFactory.FontSizeRow, "Font Sizes");
    StyleScenarioFactory.AddShowcaseLabel(sheet, StyleScenarioFactory.IndentRow, "Indent Levels");
    StyleScenarioFactory.AddShowcaseLabel(sheet, StyleScenarioFactory.WrapShrinkRow, "Wrap And Shrink");

    StyleScenarioFactory.AddFillPatternShowcase(sheet);
    StyleScenarioFactory.AddBorderStyleShowcase(sheet);
    StyleScenarioFactory.AddHorizontalAlignmentShowcase(sheet);
    StyleScenarioFactory.AddVerticalAlignmentShowcase(sheet);
    StyleScenarioFactory.AddFontShowcase(sheet);
    StyleScenarioFactory.AddNumberFormatShowcase(sheet);
    StyleScenarioFactory.AddBorderSideShowcase(sheet);
    StyleScenarioFactory.AddRotationShowcase(sheet);
    StyleScenarioFactory.AddReadingOrderShowcase(sheet);
    StyleScenarioFactory.AddProtectionShowcase(sheet);
    StyleScenarioFactory.AddFontSizeShowcase(sheet);
    StyleScenarioFactory.AddIndentShowcase(sheet);
    StyleScenarioFactory.AddWrapShrinkShowcase(sheet);

    return workbook;
  }

  public static ApplyPrimaryStyle(style: Style): void {
    style.Font.Name = "Arial";
    style.Font.Size = 14.5;
    style.Font.IsBold = true;
    style.Font.IsItalic = true;
    style.Font.Underline = FontUnderlineType.Single;
    style.Font.IsStrikeout = true;
    style.Font.Color = Color.FromArgb(255, 17, 34, 51);
    style.Pattern = FillPattern.LightGrid;
    style.ForegroundColor = Color.FromArgb(255, 210, 220, 30);
    style.BackgroundColor = Color.FromArgb(255, 12, 45, 78);
    style.Borders.Left.LineStyle = BorderStyleType.Dotted;
    style.Borders.Left.Color = Color.FromArgb(255, 200, 0, 0);
    style.Borders.Right.LineStyle = BorderStyleType.MediumDashDot;
    style.Borders.Right.Color = Color.FromArgb(255, 240, 120, 0);
    style.Borders.Top.LineStyle = BorderStyleType.Double;
    style.Borders.Top.Color = Color.FromArgb(255, 0, 0, 255);
    style.Borders.Bottom.LineStyle = BorderStyleType.DashDotDot;
    style.Borders.Bottom.Color = Color.FromArgb(255, 0, 120, 0);
    style.Borders.Diagonal.LineStyle = BorderStyleType.SlantedDashDot;
    style.Borders.Diagonal.Color = Color.FromArgb(255, 128, 0, 128);
    style.Borders.DiagonalUp = true;
    style.Borders.DiagonalDown = true;
    style.NumberFormat = "#,##0.00";
    style.HorizontalAlignment = HorizontalAlignmentType.Distributed;
    style.VerticalAlignment = VerticalAlignmentType.Distributed;
    style.WrapText = true;
    style.IndentLevel = 2;
    style.TextRotation = 45;
    style.ShrinkToFit = true;
    style.ReadingOrder = 2;
    style.RelativeIndent = 1;
    style.IsLocked = false;
    style.IsHidden = true;
  }

  public static ApplyCustomNumberStyle(style: Style): void {
    style.Font.IsBold = true;
    style.Borders.Top.LineStyle = BorderStyleType.MediumDashed;
    style.Borders.Top.Color = Color.FromArgb(255, 0, 0, 255);
    style.HorizontalAlignment = HorizontalAlignmentType.CenterContinuous;
    style.VerticalAlignment = VerticalAlignmentType.Top;
    style.NumberFormat = "0.0000";
  }

  public static ApplyDateStyle(style: Style): void {
    style.Font.Name = "Times New Roman";
    style.Font.Size = 12;
    style.Font.Color = Color.FromArgb(255, 0, 96, 160);
    style.Pattern = FillPattern.Solid;
    style.ForegroundColor = Color.FromArgb(255, 221, 235, 247);
    style.NumberFormat = "m/d/yyyy h:mm";
    style.HorizontalAlignment = HorizontalAlignmentType.Center;
    style.VerticalAlignment = VerticalAlignmentType.Center;
  }

  public static ApplyVerticalTextStyle(style: Style): void {
    style.Font.Name = "Consolas";
    style.Font.Size = 10;
    style.Font.IsItalic = true;
    style.Pattern = FillPattern.DarkHorizontal;
    style.ForegroundColor = Color.FromArgb(255, 217, 217, 217);
    style.BackgroundColor = Color.FromArgb(255, 255, 255, 255);
    style.Borders.Left.LineStyle = BorderStyleType.Thin;
    style.Borders.Left.Color = Color.FromArgb(255, 64, 64, 64);
    style.HorizontalAlignment = HorizontalAlignmentType.Justify;
    style.VerticalAlignment = VerticalAlignmentType.Justify;
    style.WrapText = true;
    style.TextRotation = 255;
    style.ReadingOrder = 1;
    style.RelativeIndent = 2;
    style.IsLocked = true;
    style.IsHidden = false;
  }

  public static AssertPrimaryStyle(style: Style): void {
    AssertEx.equal("Arial", style.Font.Name);
    AssertEx.equal(14.5, style.Font.Size);
    AssertEx.true(style.Font.IsBold);
    AssertEx.true(style.Font.IsItalic);
    AssertEx.equal(FontUnderlineType.Single, style.Font.Underline);
    AssertEx.true(style.Font.IsStrikeout);
    AssertEx.equal(Color.FromArgb(255, 17, 34, 51), style.Font.Color);
    AssertEx.equal(FillPattern.LightGrid, style.Pattern);
    AssertEx.equal(Color.FromArgb(255, 210, 220, 30), style.ForegroundColor);
    AssertEx.equal(Color.FromArgb(255, 12, 45, 78), style.BackgroundColor);
    AssertEx.equal(BorderStyleType.Dotted, style.Borders.Left.LineStyle);
    AssertEx.equal(Color.FromArgb(255, 200, 0, 0), style.Borders.Left.Color);
    AssertEx.equal(BorderStyleType.MediumDashDot, style.Borders.Right.LineStyle);
    AssertEx.equal(Color.FromArgb(255, 240, 120, 0), style.Borders.Right.Color);
    AssertEx.equal(BorderStyleType.Double, style.Borders.Top.LineStyle);
    AssertEx.equal(Color.FromArgb(255, 0, 0, 255), style.Borders.Top.Color);
    AssertEx.equal(BorderStyleType.DashDotDot, style.Borders.Bottom.LineStyle);
    AssertEx.equal(Color.FromArgb(255, 0, 120, 0), style.Borders.Bottom.Color);
    AssertEx.equal(BorderStyleType.SlantedDashDot, style.Borders.Diagonal.LineStyle);
    AssertEx.equal(Color.FromArgb(255, 128, 0, 128), style.Borders.Diagonal.Color);
    AssertEx.true(style.Borders.DiagonalUp);
    AssertEx.true(style.Borders.DiagonalDown);
    AssertEx.equal(4, style.Number);
    AssertEx.null(style.Custom);
    AssertEx.equal("#,##0.00", style.NumberFormat);
    AssertEx.equal(HorizontalAlignmentType.Distributed, style.HorizontalAlignment);
    AssertEx.equal(VerticalAlignmentType.Distributed, style.VerticalAlignment);
    AssertEx.true(style.WrapText);
    AssertEx.equal(2, style.IndentLevel);
    AssertEx.equal(45, style.TextRotation);
    AssertEx.true(style.ShrinkToFit);
    AssertEx.equal(2, style.ReadingOrder);
    AssertEx.equal(1, style.RelativeIndent);
    AssertEx.false(style.IsLocked);
    AssertEx.true(style.IsHidden);
  }

  public static AssertCustomNumberStyle(style: Style): void {
    AssertEx.true(style.Font.IsBold);
    AssertEx.equal(BorderStyleType.MediumDashed, style.Borders.Top.LineStyle);
    AssertEx.equal(Color.FromArgb(255, 0, 0, 255), style.Borders.Top.Color);
    AssertEx.equal(HorizontalAlignmentType.CenterContinuous, style.HorizontalAlignment);
    AssertEx.equal(VerticalAlignmentType.Top, style.VerticalAlignment);
    AssertEx.equal("0.0000", style.Custom);
    AssertEx.equal("0.0000", style.NumberFormat);
  }

  public static AssertDateStyle(style: Style): void {
    AssertEx.equal("Times New Roman", style.Font.Name);
    AssertEx.equal(12, style.Font.Size);
    AssertEx.equal(Color.FromArgb(255, 0, 96, 160), style.Font.Color);
    AssertEx.equal(FillPattern.Solid, style.Pattern);
    AssertEx.equal(Color.FromArgb(255, 221, 235, 247), style.ForegroundColor);
    AssertEx.equal("m/d/yyyy h:mm", style.NumberFormat);
    AssertEx.equal(HorizontalAlignmentType.Center, style.HorizontalAlignment);
    AssertEx.equal(VerticalAlignmentType.Center, style.VerticalAlignment);
  }

  public static AssertVerticalTextStyle(style: Style): void {
    AssertEx.equal("Consolas", style.Font.Name);
    AssertEx.equal(10, style.Font.Size);
    AssertEx.true(style.Font.IsItalic);
    AssertEx.equal(FillPattern.DarkHorizontal, style.Pattern);
    AssertEx.equal(Color.FromArgb(255, 217, 217, 217), style.ForegroundColor);
    AssertEx.equal(Color.FromArgb(255, 255, 255, 255), style.BackgroundColor);
    AssertEx.equal(BorderStyleType.Thin, style.Borders.Left.LineStyle);
    AssertEx.equal(Color.FromArgb(255, 64, 64, 64), style.Borders.Left.Color);
    AssertEx.equal(HorizontalAlignmentType.Justify, style.HorizontalAlignment);
    AssertEx.equal(VerticalAlignmentType.Justify, style.VerticalAlignment);
    AssertEx.true(style.WrapText);
    AssertEx.equal(255, style.TextRotation);
    AssertEx.equal(1, style.ReadingOrder);
    AssertEx.equal(2, style.RelativeIndent);
    AssertEx.true(style.IsLocked);
    AssertEx.false(style.IsHidden);
  }

  public static AssertStyleShowcase(workbook: Workbook): void {
    const sheet = workbook.Worksheets["Styled"];
    StyleScenarioFactory.AssertPrimaryStyle(sheet.Cells["A1"].GetStyle());
    StyleScenarioFactory.AssertCustomNumberStyle(sheet.Cells["B2"].GetStyle());
    StyleScenarioFactory.AssertDateStyle(sheet.Cells["C3"].GetStyle());
    StyleScenarioFactory.AssertVerticalTextStyle(sheet.Cells["D4"].GetStyle());

    for (let index = 0; index < StyleScenarioFactory.FillPatternShowcase.length; index++) {
      const style = sheet.Cells.get(StyleScenarioFactory.FillPatternRow, StyleScenarioFactory.ShowcaseStartColumn + index).GetStyle();
      AssertEx.equal(StyleScenarioFactory.FillPatternShowcase[index], style.Pattern);
      AssertEx.equal(StyleScenarioFactory.GetShowcaseColor(index), style.ForegroundColor);
      AssertEx.equal(StyleScenarioFactory.GetShowcaseAccentColor(index), style.BackgroundColor);
    }

    for (let index = 0; index < StyleScenarioFactory.BorderStyleShowcase.length; index++) {
      const style = sheet.Cells.get(StyleScenarioFactory.BorderStyleRow, StyleScenarioFactory.ShowcaseStartColumn + index).GetStyle();
      AssertEx.equal(StyleScenarioFactory.BorderStyleShowcase[index], style.Borders.Top.LineStyle);
      AssertEx.equal(StyleScenarioFactory.GetShowcaseColor(index), style.Borders.Top.Color);
      AssertEx.equal(BorderStyleType.Thin, style.Borders.Bottom.LineStyle);
      AssertEx.equal(StyleScenarioFactory.GetShowcaseAccentColor(index), style.Borders.Bottom.Color);
    }

    for (let index = 0; index < StyleScenarioFactory.HorizontalAlignmentShowcase.length; index++) {
      const style = sheet.Cells.get(StyleScenarioFactory.HorizontalAlignmentRow, StyleScenarioFactory.ShowcaseStartColumn + index).GetStyle();
      AssertEx.equal(StyleScenarioFactory.HorizontalAlignmentShowcase[index], style.HorizontalAlignment);
      AssertEx.true(style.WrapText);
    }

    for (let index = 0; index < StyleScenarioFactory.VerticalAlignmentShowcase.length; index++) {
      const style = sheet.Cells.get(StyleScenarioFactory.VerticalAlignmentRow, StyleScenarioFactory.ShowcaseStartColumn + index).GetStyle();
      AssertEx.equal(StyleScenarioFactory.VerticalAlignmentShowcase[index], style.VerticalAlignment);
      AssertEx.true(style.ShrinkToFit);
    }

    StyleScenarioFactory.AssertFontShowcase(sheet);
    StyleScenarioFactory.AssertNumberFormatShowcase(sheet);
    StyleScenarioFactory.AssertBorderSideShowcase(sheet);
    StyleScenarioFactory.AssertRotationShowcase(sheet);
    StyleScenarioFactory.AssertReadingOrderShowcase(sheet);
    StyleScenarioFactory.AssertProtectionShowcase(sheet);
    StyleScenarioFactory.AssertFontSizeShowcase(sheet);
    StyleScenarioFactory.AssertIndentShowcase(sheet);
    StyleScenarioFactory.AssertWrapShrinkShowcase(sheet);
  }

  private static AddFillPatternShowcase(sheet: Worksheet): void {
    for (let index = 0; index < StyleScenarioFactory.FillPatternShowcase.length; index++) {
      const cell = sheet.Cells.get(StyleScenarioFactory.FillPatternRow, StyleScenarioFactory.ShowcaseStartColumn + index);
      cell.PutValue("P" + index.toString());
      const style = cell.GetStyle();
      style.Pattern = StyleScenarioFactory.FillPatternShowcase[index];
      style.ForegroundColor = StyleScenarioFactory.GetShowcaseColor(index);
      style.BackgroundColor = StyleScenarioFactory.GetShowcaseAccentColor(index);
      cell.SetStyle(style);
    }
  }

  private static AddBorderStyleShowcase(sheet: Worksheet): void {
    for (let index = 0; index < StyleScenarioFactory.BorderStyleShowcase.length; index++) {
      const cell = sheet.Cells.get(StyleScenarioFactory.BorderStyleRow, StyleScenarioFactory.ShowcaseStartColumn + index);
      cell.PutValue("B" + index.toString());
      const style = cell.GetStyle();
      style.Borders.Top.LineStyle = StyleScenarioFactory.BorderStyleShowcase[index];
      style.Borders.Top.Color = StyleScenarioFactory.GetShowcaseColor(index);
      style.Borders.Bottom.LineStyle = BorderStyleType.Thin;
      style.Borders.Bottom.Color = StyleScenarioFactory.GetShowcaseAccentColor(index);
      cell.SetStyle(style);
    }
  }

  private static AddHorizontalAlignmentShowcase(sheet: Worksheet): void {
    for (let index = 0; index < StyleScenarioFactory.HorizontalAlignmentShowcase.length; index++) {
      const cell = sheet.Cells.get(StyleScenarioFactory.HorizontalAlignmentRow, StyleScenarioFactory.ShowcaseStartColumn + index);
      cell.PutValue("Align " + index.toString() + " sample");
      const style = cell.GetStyle();
      style.HorizontalAlignment = StyleScenarioFactory.HorizontalAlignmentShowcase[index];
      style.WrapText = true;
      style.Pattern = FillPattern.Solid;
      style.ForegroundColor = StyleScenarioFactory.GetShowcaseAccentColor(index);
      cell.SetStyle(style);
    }
  }

  private static AddVerticalAlignmentShowcase(sheet: Worksheet): void {
    for (let index = 0; index < StyleScenarioFactory.VerticalAlignmentShowcase.length; index++) {
      const cell = sheet.Cells.get(StyleScenarioFactory.VerticalAlignmentRow, StyleScenarioFactory.ShowcaseStartColumn + index);
      cell.PutValue("Vertical " + index.toString());
      const style = cell.GetStyle();
      style.VerticalAlignment = StyleScenarioFactory.VerticalAlignmentShowcase[index];
      style.ShrinkToFit = true;
      style.Pattern = FillPattern.LightTrellis;
      style.ForegroundColor = StyleScenarioFactory.GetShowcaseColor(index);
      cell.SetStyle(style);
    }
  }

  private static AddFontShowcase(sheet: Worksheet): void {
    const nameCell = sheet.Cells.get(StyleScenarioFactory.FontRow, StyleScenarioFactory.ShowcaseStartColumn);
    nameCell.PutValue("Cambria 16");
    const nameStyle = nameCell.GetStyle();
    nameStyle.Font.Name = "Cambria";
    nameStyle.Font.Size = 16;
    nameStyle.Pattern = FillPattern.Solid;
    nameStyle.ForegroundColor = Color.FromArgb(255, 242, 242, 242);
    nameCell.SetStyle(nameStyle);

    const boldCell = sheet.Cells.get(StyleScenarioFactory.FontRow, StyleScenarioFactory.ShowcaseStartColumn + 1);
    boldCell.PutValue("Bold");
    const boldStyle = boldCell.GetStyle();
    boldStyle.Font.IsBold = true;
    boldStyle.Font.Size = 12;
    boldStyle.ForegroundColor = Color.FromArgb(255, 31, 78, 121);
    boldCell.SetStyle(boldStyle);

    const italicCell = sheet.Cells.get(StyleScenarioFactory.FontRow, StyleScenarioFactory.ShowcaseStartColumn + 2);
    italicCell.PutValue("Italic");
    const italicStyle = italicCell.GetStyle();
    italicStyle.Font.IsItalic = true;
    italicStyle.Font.Size = 12;
    italicStyle.ForegroundColor = Color.FromArgb(255, 128, 100, 162);
    italicCell.SetStyle(italicStyle);

    const underlineCell = sheet.Cells.get(StyleScenarioFactory.FontRow, StyleScenarioFactory.ShowcaseStartColumn + 3);
    underlineCell.PutValue("Underline");
    const underlineStyle = underlineCell.GetStyle();
    underlineStyle.Font.Underline = FontUnderlineType.Single;
    underlineStyle.Font.Color = Color.FromArgb(255, 0, 112, 192);
    underlineCell.SetStyle(underlineStyle);

    const strikeCell = sheet.Cells.get(StyleScenarioFactory.FontRow, StyleScenarioFactory.ShowcaseStartColumn + 4);
    strikeCell.PutValue("Strike");
    const strikeStyle = strikeCell.GetStyle();
    strikeStyle.Font.IsStrikeout = true;
    strikeStyle.Font.Color = Color.FromArgb(255, 192, 0, 0);
    strikeCell.SetStyle(strikeStyle);

    const colorCell = sheet.Cells.get(StyleScenarioFactory.FontRow, StyleScenarioFactory.ShowcaseStartColumn + 5);
    colorCell.PutValue("Color");
    const colorStyle = colorCell.GetStyle();
    colorStyle.Font.IsBold = true;
    colorStyle.Font.Color = Color.FromArgb(255, 0, 176, 80);
    colorStyle.Pattern = FillPattern.Solid;
    colorStyle.ForegroundColor = Color.FromArgb(255, 226, 239, 218);
    colorCell.SetStyle(colorStyle);
  }

  private static AddNumberFormatShowcase(sheet: Worksheet): void {
    for (let index = 0; index < StyleScenarioFactory.NumberFormatShowcase.length; index++) {
      const cell = sheet.Cells.get(StyleScenarioFactory.NumberFormatRow, StyleScenarioFactory.ShowcaseStartColumn + index);
      if (index === 3) {
        cell.PutValue(0.375);
      } else if (index === 5) {
        cell.PutValue(12.75);
      } else {
        cell.PutValue(1234.567 + index);
      }

      const style = cell.GetStyle();
      style.NumberFormat = StyleScenarioFactory.NumberFormatShowcase[index];
      style.HorizontalAlignment = HorizontalAlignmentType.Right;
      style.Pattern = FillPattern.Solid;
      style.ForegroundColor = StyleScenarioFactory.GetShowcaseAccentColor(index + 10);
      cell.SetStyle(style);
    }
  }

  private static AddBorderSideShowcase(sheet: Worksheet): void {
    const leftCell = sheet.Cells.get(StyleScenarioFactory.BorderSidesRow, StyleScenarioFactory.ShowcaseStartColumn);
    leftCell.PutValue("Left");
    const leftStyle = leftCell.GetStyle();
    leftStyle.Borders.Left.LineStyle = BorderStyleType.Thick;
    leftStyle.Borders.Left.Color = Color.FromArgb(255, 192, 0, 0);
    leftStyle.Pattern = FillPattern.Solid;
    leftStyle.ForegroundColor = Color.FromArgb(255, 255, 242, 204);
    leftCell.SetStyle(leftStyle);

    const rightCell = sheet.Cells.get(StyleScenarioFactory.BorderSidesRow, StyleScenarioFactory.ShowcaseStartColumn + 1);
    rightCell.PutValue("Right");
    const rightStyle = rightCell.GetStyle();
    rightStyle.Borders.Right.LineStyle = BorderStyleType.Double;
    rightStyle.Borders.Right.Color = Color.FromArgb(255, 0, 112, 192);
    rightStyle.Pattern = FillPattern.Solid;
    rightStyle.ForegroundColor = Color.FromArgb(255, 221, 235, 247);
    rightCell.SetStyle(rightStyle);

    const topCell = sheet.Cells.get(StyleScenarioFactory.BorderSidesRow, StyleScenarioFactory.ShowcaseStartColumn + 2);
    topCell.PutValue("Top");
    const topStyle = topCell.GetStyle();
    topStyle.Borders.Top.LineStyle = BorderStyleType.DashDot;
    topStyle.Borders.Top.Color = Color.FromArgb(255, 112, 48, 160);
    topStyle.Pattern = FillPattern.Solid;
    topStyle.ForegroundColor = Color.FromArgb(255, 234, 209, 220);
    topCell.SetStyle(topStyle);

    const bottomCell = sheet.Cells.get(StyleScenarioFactory.BorderSidesRow, StyleScenarioFactory.ShowcaseStartColumn + 3);
    bottomCell.PutValue("Bottom");
    const bottomStyle = bottomCell.GetStyle();
    bottomStyle.Borders.Bottom.LineStyle = BorderStyleType.MediumDashed;
    bottomStyle.Borders.Bottom.Color = Color.FromArgb(255, 0, 176, 80);
    bottomStyle.Pattern = FillPattern.Solid;
    bottomStyle.ForegroundColor = Color.FromArgb(255, 226, 239, 218);
    bottomCell.SetStyle(bottomStyle);

    const diagonalUpCell = sheet.Cells.get(StyleScenarioFactory.BorderSidesRow, StyleScenarioFactory.ShowcaseStartColumn + 4);
    diagonalUpCell.PutValue("Diag Up");
    const diagonalUpStyle = diagonalUpCell.GetStyle();
    diagonalUpStyle.Borders.Diagonal.LineStyle = BorderStyleType.SlantedDashDot;
    diagonalUpStyle.Borders.Diagonal.Color = Color.FromArgb(255, 255, 0, 0);
    diagonalUpStyle.Borders.DiagonalUp = true;
    diagonalUpStyle.Pattern = FillPattern.Solid;
    diagonalUpStyle.ForegroundColor = Color.FromArgb(255, 252, 228, 214);
    diagonalUpCell.SetStyle(diagonalUpStyle);

    const diagonalDownCell = sheet.Cells.get(StyleScenarioFactory.BorderSidesRow, StyleScenarioFactory.ShowcaseStartColumn + 5);
    diagonalDownCell.PutValue("Diag Down");
    const diagonalDownStyle = diagonalDownCell.GetStyle();
    diagonalDownStyle.Borders.Diagonal.LineStyle = BorderStyleType.SlantedDashDot;
    diagonalDownStyle.Borders.Diagonal.Color = Color.FromArgb(255, 0, 176, 240);
    diagonalDownStyle.Borders.DiagonalDown = true;
    diagonalDownStyle.Pattern = FillPattern.Solid;
    diagonalDownStyle.ForegroundColor = Color.FromArgb(255, 222, 234, 246);
    diagonalDownCell.SetStyle(diagonalDownStyle);
  }

  private static AddRotationShowcase(sheet: Worksheet): void {
    for (let index = 0; index < StyleScenarioFactory.RotationShowcase.length; index++) {
      const cell = sheet.Cells.get(StyleScenarioFactory.RotationRow, StyleScenarioFactory.ShowcaseStartColumn + index);
      cell.PutValue("Rot " + StyleScenarioFactory.RotationShowcase[index].toString());
      const style = cell.GetStyle();
      style.TextRotation = StyleScenarioFactory.RotationShowcase[index];
      style.HorizontalAlignment = HorizontalAlignmentType.Center;
      style.VerticalAlignment = VerticalAlignmentType.Center;
      style.WrapText = true;
      style.Pattern = FillPattern.Solid;
      style.ForegroundColor = StyleScenarioFactory.GetShowcaseColor(index + 20);
      cell.SetStyle(style);
    }
  }

  private static AddReadingOrderShowcase(sheet: Worksheet): void {
    for (let index = 0; index < StyleScenarioFactory.ReadingOrderShowcase.length; index++) {
      const cell = sheet.Cells.get(StyleScenarioFactory.ReadingOrderRow, StyleScenarioFactory.ShowcaseStartColumn + index);
      cell.PutValue("Order " + StyleScenarioFactory.ReadingOrderShowcase[index].toString());
      const style = cell.GetStyle();
      style.ReadingOrder = StyleScenarioFactory.ReadingOrderShowcase[index];
      style.IndentLevel = index;
      if (StyleScenarioFactory.ReadingOrderShowcase[index] > 0) {
        style.RelativeIndent = index + 1;
      }
      style.HorizontalAlignment = HorizontalAlignmentType.Distributed;
      style.Pattern = FillPattern.Solid;
      style.ForegroundColor = StyleScenarioFactory.GetShowcaseAccentColor(index + 20);
      cell.SetStyle(style);
    }
  }

  private static AddProtectionShowcase(sheet: Worksheet): void {
    const lockedCell = sheet.Cells.get(StyleScenarioFactory.ProtectionRow, StyleScenarioFactory.ShowcaseStartColumn);
    lockedCell.PutValue("Locked");
    const lockedStyle = lockedCell.GetStyle();
    lockedStyle.IsLocked = true;
    lockedStyle.IsHidden = false;
    lockedStyle.Pattern = FillPattern.Solid;
    lockedStyle.ForegroundColor = Color.FromArgb(255, 226, 239, 218);
    lockedCell.SetStyle(lockedStyle);

    const unlockedCell = sheet.Cells.get(StyleScenarioFactory.ProtectionRow, StyleScenarioFactory.ShowcaseStartColumn + 1);
    unlockedCell.PutValue("Unlocked");
    const unlockedStyle = unlockedCell.GetStyle();
    unlockedStyle.IsLocked = false;
    unlockedStyle.IsHidden = false;
    unlockedStyle.Pattern = FillPattern.Solid;
    unlockedStyle.ForegroundColor = Color.FromArgb(255, 255, 242, 204);
    unlockedCell.SetStyle(unlockedStyle);

    const hiddenCell = sheet.Cells.get(StyleScenarioFactory.ProtectionRow, StyleScenarioFactory.ShowcaseStartColumn + 2);
    hiddenCell.PutValue("Hidden");
    const hiddenStyle = hiddenCell.GetStyle();
    hiddenStyle.IsLocked = true;
    hiddenStyle.IsHidden = true;
    hiddenStyle.Pattern = FillPattern.Solid;
    hiddenStyle.ForegroundColor = Color.FromArgb(255, 242, 220, 219);
    hiddenCell.SetStyle(hiddenStyle);
  }

  private static AddFontSizeShowcase(sheet: Worksheet): void {
    for (let index = 0; index < StyleScenarioFactory.FontSizeShowcase.length; index++) {
      const cell = sheet.Cells.get(StyleScenarioFactory.FontSizeRow, StyleScenarioFactory.ShowcaseStartColumn + index);
      cell.PutValue("Size " + StyleScenarioFactory.FontSizeShowcase[index].toFixed(0));
      const style = cell.GetStyle();
      style.Font.Name = "Calibri";
      style.Font.Size = StyleScenarioFactory.FontSizeShowcase[index];
      style.Pattern = FillPattern.Solid;
      style.ForegroundColor = StyleScenarioFactory.GetShowcaseAccentColor(index + 30);
      cell.SetStyle(style);
    }
  }

  private static AddIndentShowcase(sheet: Worksheet): void {
    for (let index = 0; index < StyleScenarioFactory.IndentLevelShowcase.length; index++) {
      const cell = sheet.Cells.get(StyleScenarioFactory.IndentRow, StyleScenarioFactory.ShowcaseStartColumn + index);
      cell.PutValue("Indent " + StyleScenarioFactory.IndentLevelShowcase[index].toString());
      const style = cell.GetStyle();
      style.HorizontalAlignment = HorizontalAlignmentType.Left;
      style.IndentLevel = StyleScenarioFactory.IndentLevelShowcase[index];
      style.RelativeIndent = StyleScenarioFactory.IndentLevelShowcase[index];
      style.Pattern = FillPattern.Solid;
      style.ForegroundColor = StyleScenarioFactory.GetShowcaseColor(index + 30);
      cell.SetStyle(style);
    }
  }

  private static AddWrapShrinkShowcase(sheet: Worksheet): void {
    const wrapCell = sheet.Cells.get(StyleScenarioFactory.WrapShrinkRow, StyleScenarioFactory.ShowcaseStartColumn);
    wrapCell.PutValue("Wrap text sample");
    const wrapStyle = wrapCell.GetStyle();
    wrapStyle.WrapText = true;
    wrapStyle.HorizontalAlignment = HorizontalAlignmentType.Justify;
    wrapStyle.Pattern = FillPattern.Solid;
    wrapStyle.ForegroundColor = Color.FromArgb(255, 252, 228, 214);
    wrapCell.SetStyle(wrapStyle);

    const shrinkCell = sheet.Cells.get(StyleScenarioFactory.WrapShrinkRow, StyleScenarioFactory.ShowcaseStartColumn + 1);
    shrinkCell.PutValue("Shrink to fit sample");
    const shrinkStyle = shrinkCell.GetStyle();
    shrinkStyle.ShrinkToFit = true;
    shrinkStyle.HorizontalAlignment = HorizontalAlignmentType.Center;
    shrinkStyle.Pattern = FillPattern.Solid;
    shrinkStyle.ForegroundColor = Color.FromArgb(255, 221, 235, 247);
    shrinkCell.SetStyle(shrinkStyle);

    const wrapShrinkCell = sheet.Cells.get(StyleScenarioFactory.WrapShrinkRow, StyleScenarioFactory.ShowcaseStartColumn + 2);
    wrapShrinkCell.PutValue("Wrap and shrink");
    const wrapShrinkStyle = wrapShrinkCell.GetStyle();
    wrapShrinkStyle.WrapText = true;
    wrapShrinkStyle.ShrinkToFit = true;
    wrapShrinkStyle.HorizontalAlignment = HorizontalAlignmentType.Distributed;
    wrapShrinkStyle.VerticalAlignment = VerticalAlignmentType.Center;
    wrapShrinkStyle.Pattern = FillPattern.Solid;
    wrapShrinkStyle.ForegroundColor = Color.FromArgb(255, 226, 239, 218);
    wrapShrinkCell.SetStyle(wrapShrinkStyle);

    const justifyCell = sheet.Cells.get(StyleScenarioFactory.WrapShrinkRow, StyleScenarioFactory.ShowcaseStartColumn + 3);
    justifyCell.PutValue("Distributed indent");
    const justifyStyle = justifyCell.GetStyle();
    justifyStyle.WrapText = true;
    justifyStyle.HorizontalAlignment = HorizontalAlignmentType.Distributed;
    justifyStyle.IndentLevel = 3;
    justifyStyle.RelativeIndent = 2;
    justifyStyle.Pattern = FillPattern.Solid;
    justifyStyle.ForegroundColor = Color.FromArgb(255, 242, 242, 242);
    justifyCell.SetStyle(justifyStyle);
  }

  private static AssertFontShowcase(sheet: Worksheet): void {
    const nameStyle = sheet.Cells.get(StyleScenarioFactory.FontRow, StyleScenarioFactory.ShowcaseStartColumn).GetStyle();
    AssertEx.equal("Cambria", nameStyle.Font.Name);
    AssertEx.equal(16, nameStyle.Font.Size);
    AssertEx.equal(FillPattern.Solid, nameStyle.Pattern);

    const boldStyle = sheet.Cells.get(StyleScenarioFactory.FontRow, StyleScenarioFactory.ShowcaseStartColumn + 1).GetStyle();
    AssertEx.true(boldStyle.Font.IsBold);
    AssertEx.equal(12, boldStyle.Font.Size);
    AssertEx.equal(Color.FromArgb(255, 31, 78, 121), boldStyle.ForegroundColor);

    const italicStyle = sheet.Cells.get(StyleScenarioFactory.FontRow, StyleScenarioFactory.ShowcaseStartColumn + 2).GetStyle();
    AssertEx.true(italicStyle.Font.IsItalic);
    AssertEx.equal(12, italicStyle.Font.Size);
    AssertEx.equal(Color.FromArgb(255, 128, 100, 162), italicStyle.ForegroundColor);

    const underlineStyle = sheet.Cells.get(StyleScenarioFactory.FontRow, StyleScenarioFactory.ShowcaseStartColumn + 3).GetStyle();
    AssertEx.equal(FontUnderlineType.Single, underlineStyle.Font.Underline);
    AssertEx.equal(Color.FromArgb(255, 0, 112, 192), underlineStyle.Font.Color);

    const strikeStyle = sheet.Cells.get(StyleScenarioFactory.FontRow, StyleScenarioFactory.ShowcaseStartColumn + 4).GetStyle();
    AssertEx.true(strikeStyle.Font.IsStrikeout);
    AssertEx.equal(Color.FromArgb(255, 192, 0, 0), strikeStyle.Font.Color);

    const colorStyle = sheet.Cells.get(StyleScenarioFactory.FontRow, StyleScenarioFactory.ShowcaseStartColumn + 5).GetStyle();
    AssertEx.true(colorStyle.Font.IsBold);
    AssertEx.equal(Color.FromArgb(255, 0, 176, 80), colorStyle.Font.Color);
    AssertEx.equal(FillPattern.Solid, colorStyle.Pattern);
    AssertEx.equal(Color.FromArgb(255, 226, 239, 218), colorStyle.ForegroundColor);
  }

  private static AssertNumberFormatShowcase(sheet: Worksheet): void {
    for (let index = 0; index < StyleScenarioFactory.NumberFormatShowcase.length; index++) {
      const style = sheet.Cells.get(StyleScenarioFactory.NumberFormatRow, StyleScenarioFactory.ShowcaseStartColumn + index).GetStyle();
      AssertEx.equal(StyleScenarioFactory.NumberFormatShowcase[index], style.NumberFormat);
      AssertEx.equal(HorizontalAlignmentType.Right, style.HorizontalAlignment);
      AssertEx.equal(FillPattern.Solid, style.Pattern);
      AssertEx.equal(StyleScenarioFactory.GetShowcaseAccentColor(index + 10), style.ForegroundColor);
    }
  }

  private static AssertBorderSideShowcase(sheet: Worksheet): void {
    const leftStyle = sheet.Cells.get(StyleScenarioFactory.BorderSidesRow, StyleScenarioFactory.ShowcaseStartColumn).GetStyle();
    AssertEx.equal(BorderStyleType.Thick, leftStyle.Borders.Left.LineStyle);
    AssertEx.equal(Color.FromArgb(255, 192, 0, 0), leftStyle.Borders.Left.Color);

    const rightStyle = sheet.Cells.get(StyleScenarioFactory.BorderSidesRow, StyleScenarioFactory.ShowcaseStartColumn + 1).GetStyle();
    AssertEx.equal(BorderStyleType.Double, rightStyle.Borders.Right.LineStyle);
    AssertEx.equal(Color.FromArgb(255, 0, 112, 192), rightStyle.Borders.Right.Color);

    const topStyle = sheet.Cells.get(StyleScenarioFactory.BorderSidesRow, StyleScenarioFactory.ShowcaseStartColumn + 2).GetStyle();
    AssertEx.equal(BorderStyleType.DashDot, topStyle.Borders.Top.LineStyle);
    AssertEx.equal(Color.FromArgb(255, 112, 48, 160), topStyle.Borders.Top.Color);

    const bottomStyle = sheet.Cells.get(StyleScenarioFactory.BorderSidesRow, StyleScenarioFactory.ShowcaseStartColumn + 3).GetStyle();
    AssertEx.equal(BorderStyleType.MediumDashed, bottomStyle.Borders.Bottom.LineStyle);
    AssertEx.equal(Color.FromArgb(255, 0, 176, 80), bottomStyle.Borders.Bottom.Color);

    const diagonalUpStyle = sheet.Cells.get(StyleScenarioFactory.BorderSidesRow, StyleScenarioFactory.ShowcaseStartColumn + 4).GetStyle();
    AssertEx.equal(BorderStyleType.SlantedDashDot, diagonalUpStyle.Borders.Diagonal.LineStyle);
    AssertEx.equal(Color.FromArgb(255, 255, 0, 0), diagonalUpStyle.Borders.Diagonal.Color);
    AssertEx.true(diagonalUpStyle.Borders.DiagonalUp);
    AssertEx.false(diagonalUpStyle.Borders.DiagonalDown);

    const diagonalDownStyle = sheet.Cells.get(StyleScenarioFactory.BorderSidesRow, StyleScenarioFactory.ShowcaseStartColumn + 5).GetStyle();
    AssertEx.equal(BorderStyleType.SlantedDashDot, diagonalDownStyle.Borders.Diagonal.LineStyle);
    AssertEx.equal(Color.FromArgb(255, 0, 176, 240), diagonalDownStyle.Borders.Diagonal.Color);
    AssertEx.false(diagonalDownStyle.Borders.DiagonalUp);
    AssertEx.true(diagonalDownStyle.Borders.DiagonalDown);
  }

  private static AssertRotationShowcase(sheet: Worksheet): void {
    for (let index = 0; index < StyleScenarioFactory.RotationShowcase.length; index++) {
      const style = sheet.Cells.get(StyleScenarioFactory.RotationRow, StyleScenarioFactory.ShowcaseStartColumn + index).GetStyle();
      AssertEx.equal(StyleScenarioFactory.RotationShowcase[index], style.TextRotation);
      AssertEx.equal(HorizontalAlignmentType.Center, style.HorizontalAlignment);
      AssertEx.equal(VerticalAlignmentType.Center, style.VerticalAlignment);
      AssertEx.true(style.WrapText);
    }
  }

  private static AssertReadingOrderShowcase(sheet: Worksheet): void {
    for (let index = 0; index < StyleScenarioFactory.ReadingOrderShowcase.length; index++) {
      const style = sheet.Cells.get(StyleScenarioFactory.ReadingOrderRow, StyleScenarioFactory.ShowcaseStartColumn + index).GetStyle();
      AssertEx.equal(StyleScenarioFactory.ReadingOrderShowcase[index], style.ReadingOrder);
      AssertEx.equal(index, style.IndentLevel);
      AssertEx.equal(
        StyleScenarioFactory.ReadingOrderShowcase[index] > 0 ? index + 1 : 0,
        style.RelativeIndent,
      );
      AssertEx.equal(HorizontalAlignmentType.Distributed, style.HorizontalAlignment);
      AssertEx.equal(FillPattern.Solid, style.Pattern);
      AssertEx.equal(StyleScenarioFactory.GetShowcaseAccentColor(index + 20), style.ForegroundColor);
    }
  }

  private static AssertProtectionShowcase(sheet: Worksheet): void {
    const lockedStyle = sheet.Cells.get(StyleScenarioFactory.ProtectionRow, StyleScenarioFactory.ShowcaseStartColumn).GetStyle();
    AssertEx.true(lockedStyle.IsLocked);
    AssertEx.false(lockedStyle.IsHidden);

    const unlockedStyle = sheet.Cells.get(StyleScenarioFactory.ProtectionRow, StyleScenarioFactory.ShowcaseStartColumn + 1).GetStyle();
    AssertEx.false(unlockedStyle.IsLocked);
    AssertEx.false(unlockedStyle.IsHidden);

    const hiddenStyle = sheet.Cells.get(StyleScenarioFactory.ProtectionRow, StyleScenarioFactory.ShowcaseStartColumn + 2).GetStyle();
    AssertEx.true(hiddenStyle.IsLocked);
    AssertEx.true(hiddenStyle.IsHidden);
  }

  private static AssertFontSizeShowcase(sheet: Worksheet): void {
    for (let index = 0; index < StyleScenarioFactory.FontSizeShowcase.length; index++) {
      const style = sheet.Cells.get(StyleScenarioFactory.FontSizeRow, StyleScenarioFactory.ShowcaseStartColumn + index).GetStyle();
      AssertEx.equal("Calibri", style.Font.Name);
      AssertEx.equal(StyleScenarioFactory.FontSizeShowcase[index], style.Font.Size);
      AssertEx.equal(FillPattern.Solid, style.Pattern);
      AssertEx.equal(StyleScenarioFactory.GetShowcaseAccentColor(index + 30), style.ForegroundColor);
    }
  }

  private static AssertIndentShowcase(sheet: Worksheet): void {
    for (let index = 0; index < StyleScenarioFactory.IndentLevelShowcase.length; index++) {
      const style = sheet.Cells.get(StyleScenarioFactory.IndentRow, StyleScenarioFactory.ShowcaseStartColumn + index).GetStyle();
      AssertEx.equal(HorizontalAlignmentType.Left, style.HorizontalAlignment);
      AssertEx.equal(StyleScenarioFactory.IndentLevelShowcase[index], style.IndentLevel);
      AssertEx.equal(FillPattern.Solid, style.Pattern);
      AssertEx.equal(StyleScenarioFactory.GetShowcaseColor(index + 30), style.ForegroundColor);
    }
  }

  private static AssertWrapShrinkShowcase(sheet: Worksheet): void {
    const wrapStyle = sheet.Cells.get(StyleScenarioFactory.WrapShrinkRow, StyleScenarioFactory.ShowcaseStartColumn).GetStyle();
    AssertEx.true(wrapStyle.WrapText);
    AssertEx.equal(HorizontalAlignmentType.Justify, wrapStyle.HorizontalAlignment);

    const shrinkStyle = sheet.Cells.get(StyleScenarioFactory.WrapShrinkRow, StyleScenarioFactory.ShowcaseStartColumn + 1).GetStyle();
    AssertEx.true(shrinkStyle.ShrinkToFit);
    AssertEx.equal(HorizontalAlignmentType.Center, shrinkStyle.HorizontalAlignment);

    const wrapShrinkStyle = sheet.Cells.get(StyleScenarioFactory.WrapShrinkRow, StyleScenarioFactory.ShowcaseStartColumn + 2).GetStyle();
    AssertEx.true(wrapShrinkStyle.WrapText);
    AssertEx.true(wrapShrinkStyle.ShrinkToFit);
    AssertEx.equal(HorizontalAlignmentType.Distributed, wrapShrinkStyle.HorizontalAlignment);
    AssertEx.equal(VerticalAlignmentType.Center, wrapShrinkStyle.VerticalAlignment);

    const justifyStyle = sheet.Cells.get(StyleScenarioFactory.WrapShrinkRow, StyleScenarioFactory.ShowcaseStartColumn + 3).GetStyle();
    AssertEx.true(justifyStyle.WrapText);
    AssertEx.equal(HorizontalAlignmentType.Distributed, justifyStyle.HorizontalAlignment);
    AssertEx.equal(3, justifyStyle.IndentLevel);
    AssertEx.equal(0, justifyStyle.RelativeIndent);
  }

  private static AddShowcaseLabel(sheet: Worksheet, rowIndex: number, label: string): void {
    const cell = sheet.Cells.get(rowIndex, StyleScenarioFactory.ShowcaseStartColumn - 1);
    cell.PutValue(label);
    const style = cell.GetStyle();
    style.Font.IsBold = true;
    style.Pattern = FillPattern.Solid;
    style.ForegroundColor = Color.FromArgb(255, 217, 225, 242);
    style.Borders.Bottom.LineStyle = BorderStyleType.Thin;
    style.Borders.Bottom.Color = Color.FromArgb(255, 79, 129, 189);
    cell.SetStyle(style);
  }

  private static GetShowcaseColor(index: number): Color {
    return Color.FromArgb(
      255,
      (40 + index * 29) % 256,
      (90 + index * 37) % 256,
      (140 + index * 43) % 256,
    );
  }

  private static GetShowcaseAccentColor(index: number): Color {
    return Color.FromArgb(
      255,
      (220 + index * 17) % 256,
      (180 + index * 19) % 256,
      (120 + index * 23) % 256,
    );
  }
}

/** Scenario factory for worksheet data validations. */
export class ValidationScenarioFactory {
  public static CreateValidationWorkbook(): Workbook {
    const workbook = new Workbook();
    const sheet = workbook.Worksheets[0];
    sheet.Name = "Validation Sheet";

    sheet.Cells["A1"].PutValue("Open");
    sheet.Cells["B2"].PutValue(5);
    sheet.Cells["C3"].PutValue(7);
    sheet.Cells["E2"].PutValue(8);
    sheet.Cells["G1"].PutValue("ABCDE");

    const listIndex = sheet.Validations.Add(CellArea.CreateCellArea("A1", "A3"));
    const listValidation = sheet.Validations[listIndex];
    listValidation.Type = ValidationType.List;
    listValidation.Formula1 = "\"Open,Closed\"";
    listValidation.IgnoreBlank = true;
    listValidation.InCellDropDown = true;
    listValidation.ShowInput = true;
    listValidation.InputTitle = "Status";
    listValidation.InputMessage = "Pick a status";
    listValidation.ShowError = true;
    listValidation.ErrorTitle = "Invalid";
    listValidation.ErrorMessage = "Choose from the list";

    const decimalIndex = sheet.Validations.Add(CellArea.CreateCellArea("B2", "C3"));
    const decimalValidation = sheet.Validations[decimalIndex];
    decimalValidation.Type = ValidationType.Decimal;
    decimalValidation.Operator = OperatorType.Between;
    decimalValidation.Formula1 = "1.5";
    decimalValidation.Formula2 = "9.5";
    decimalValidation.IgnoreBlank = false;
    decimalValidation.InCellDropDown = false;
    decimalValidation.ShowError = true;
    decimalValidation.ErrorTitle = "Range";
    decimalValidation.ErrorMessage = "Enter 1.5-9.5";
    decimalValidation.AddArea(CellArea.CreateCellArea("E2", "E3"));

    const customIndex = sheet.Validations.Add(CellArea.CreateCellArea("G1", "G1"));
    const customValidation = sheet.Validations[customIndex];
    customValidation.Type = ValidationType.Custom;
    customValidation.AlertStyle = ValidationAlertType.Warning;
    customValidation.Formula1 = "LEN(G1)<=5";
    customValidation.ShowInput = true;
    customValidation.InputTitle = "Code";
    customValidation.InputMessage = "Up to 5 chars";

    return workbook;
  }

  public static AssertValidations(workbook: Workbook): void {
    const sheet = workbook.Worksheets["Validation Sheet"];
    AssertEx.equal(3, sheet.Validations.Count);

    const listValidation = ValidationScenarioFactory.GetValidationByArea(sheet.Validations, "A1:A3");
    AssertEx.equal(ValidationType.List, listValidation.Type);
    AssertEx.equal(OperatorType.None, listValidation.Operator);
    AssertEx.equal(1, listValidation.Areas.length);
    ValidationScenarioFactory.AssertArea(listValidation.Areas[0], 0, 0, 3, 1);
    AssertEx.equal("\"Open,Closed\"", listValidation.Formula1);
    AssertEx.equal("", listValidation.Formula2);
    AssertEx.true(listValidation.IgnoreBlank);
    AssertEx.true(listValidation.InCellDropDown);
    AssertEx.true(listValidation.ShowInput);
    AssertEx.true(listValidation.ShowError);
    AssertEx.equal("Status", listValidation.InputTitle);
    AssertEx.equal("Pick a status", listValidation.InputMessage);
    AssertEx.equal("Invalid", listValidation.ErrorTitle);
    AssertEx.equal("Choose from the list", listValidation.ErrorMessage);

    const decimalValidation = ValidationScenarioFactory.GetValidationByArea(sheet.Validations, "B2:C3 E2:E3");
    AssertEx.equal(ValidationType.Decimal, decimalValidation.Type);
    AssertEx.equal(OperatorType.Between, decimalValidation.Operator);
    AssertEx.equal(2, decimalValidation.Areas.length);
    ValidationScenarioFactory.AssertArea(decimalValidation.Areas[0], 1, 1, 2, 2);
    ValidationScenarioFactory.AssertArea(decimalValidation.Areas[1], 1, 4, 2, 1);
    AssertEx.equal("1.5", decimalValidation.Formula1);
    AssertEx.equal("9.5", decimalValidation.Formula2);
    AssertEx.false(decimalValidation.IgnoreBlank);
    AssertEx.false(decimalValidation.InCellDropDown);
    AssertEx.false(decimalValidation.ShowInput);
    AssertEx.true(decimalValidation.ShowError);
    AssertEx.equal("Range", decimalValidation.ErrorTitle);
    AssertEx.equal("Enter 1.5-9.5", decimalValidation.ErrorMessage);

    const customValidation = ValidationScenarioFactory.GetValidationByArea(sheet.Validations, "G1");
    AssertEx.equal(ValidationType.Custom, customValidation.Type);
    AssertEx.equal(ValidationAlertType.Warning, customValidation.AlertStyle);
    AssertEx.equal(OperatorType.None, customValidation.Operator);
    AssertEx.equal(1, customValidation.Areas.length);
    ValidationScenarioFactory.AssertArea(customValidation.Areas[0], 0, 6, 1, 1);
    AssertEx.equal("LEN(G1)<=5", customValidation.Formula1);
    AssertEx.equal("", customValidation.Formula2);
    AssertEx.true(customValidation.ShowInput);
    AssertEx.false(customValidation.ShowError);
    AssertEx.equal("Code", customValidation.InputTitle);
    AssertEx.equal("Up to 5 chars", customValidation.InputMessage);

    const a1Validation = sheet.Validations.GetValidationInCell(0, 0);
    AssertEx.notNull(a1Validation);
    AssertEx.equal(ValidationType.List, a1Validation.Type);

    const c2Validation = sheet.Validations.GetValidationInCell(1, 2);
    AssertEx.notNull(c2Validation);
    AssertEx.equal(ValidationType.Decimal, c2Validation.Type);

    const missingValidation = sheet.Validations.GetValidationInCell(9, 9);
    AssertEx.null(missingValidation);
  }

  private static GetValidationByArea(validations: ValidationCollection, expectedArea: string): Validation {
    for (let index = 0; index < validations.Count; index++) {
      const validation = validations[index];
      if (ValidationScenarioFactory.BuildAreaKey(validation) === expectedArea) {
        return validation;
      }
    }

    throw new Error("Validation with area '" + expectedArea + "' was not found.");
  }

  private static BuildAreaKey(validation: Validation): string {
    const areas: string[] = [];
    for (let index = 0; index < validation.Areas.length; index++) {
      const area = validation.Areas[index];
      const startCell = ValidationScenarioFactory.BuildCellName(area.StartRow, area.StartColumn);
      if (area.EndRow === area.StartRow && area.EndColumn === area.StartColumn) {
        areas.push(startCell);
        continue;
      }

      const endCell = ValidationScenarioFactory.BuildCellName(area.EndRow, area.EndColumn);
      areas.push(startCell + ":" + endCell);
    }

    return areas.join(" ");
  }

  private static BuildCellName(rowIndex: number, columnIndex: number): string {
    let dividend = columnIndex + 1;
    let name = "";
    while (dividend > 0) {
      const remainder = (dividend - 1) % 26;
      name = String.fromCharCode("A".charCodeAt(0) + remainder) + name;
      dividend = Math.floor((dividend - remainder - 1) / 26);
    }

    return name + (rowIndex + 1).toString();
  }

  private static AssertArea(area: CellArea, firstRow: number, firstColumn: number, totalRows: number, totalColumns: number): void {
    AssertEx.equal(firstRow, area.StartRow);
    AssertEx.equal(firstColumn, area.StartColumn);
    AssertEx.equal(firstRow + totalRows - 1, area.EndRow);
    AssertEx.equal(firstColumn + totalColumns - 1, area.EndColumn);
  }
}

/** Scenario factory for conditional formatting rules. */
export class ConditionalFormattingScenarioFactory {
  public static CreateConditionalFormattingWorkbook(): Workbook {
    const workbook = new Workbook();
    const sheet = workbook.Worksheets[0];
    sheet.Name = "Conditional Formatting";

    sheet.Cells["A1"].PutValue(1);
    sheet.Cells["A2"].PutValue(5);
    sheet.Cells["A3"].PutValue(10);
    sheet.Cells["A4"].PutValue(15);
    sheet.Cells["A5"].PutValue(20);
    sheet.Cells["C1"].PutValue(11);
    sheet.Cells["C2"].PutValue(8);
    sheet.Cells["C3"].PutValue(15);

    const primaryIndex = sheet.ConditionalFormattings.Add();
    const primary: any = sheet.ConditionalFormattings.get(primaryIndex);
    primary.AddArea(CellArea.CreateCellArea("A1", "A5"));
    const betweenIndex = primary.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "5", "15");
    const between = primary.get(betweenIndex);
    between.StopIfTrue = true;
    between.Priority = 1;
    const betweenStyle = between.Style;
    betweenStyle.Pattern = FillPattern.Solid;
    betweenStyle.ForegroundColor = Color.FromArgb(255, 255, 199, 206);
    betweenStyle.Font.IsBold = true;
    betweenStyle.Font.Color = Color.FromArgb(255, 156, 0, 6);
    between.Style = betweenStyle;

    const expressionIndex = primary.AddCondition(FormatConditionType.Expression, OperatorType.None, "MOD(A1,2)=0", "");
    const expression = primary.get(expressionIndex);
    expression.Priority = 2;
    const expressionStyle = expression.Style;
    expressionStyle.Font.IsItalic = true;
    expressionStyle.Font.Color = Color.FromArgb(255, 0, 0, 255);
    expressionStyle.Borders.Bottom.LineStyle = BorderStyleType.Thin;
    expressionStyle.Borders.Bottom.Color = Color.FromArgb(255, 0, 0, 255);
    expression.Style = expressionStyle;

    const secondaryIndex = sheet.ConditionalFormattings.Add();
    const secondary: any = sheet.ConditionalFormattings.get(secondaryIndex);
    secondary.AddArea(CellArea.CreateCellArea("C1", "C3"));
    const greaterThanIndex = secondary.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "10", "");
    const greaterThan = secondary.get(greaterThanIndex);
    greaterThan.Priority = 3;
    const greaterThanStyle = greaterThan.Style;
    greaterThanStyle.Pattern = FillPattern.Solid;
    greaterThanStyle.ForegroundColor = Color.FromArgb(255, 198, 239, 206);
    greaterThanStyle.Font.Color = Color.FromArgb(255, 0, 97, 0);
    greaterThan.Style = greaterThanStyle;

    return workbook;
  }

  public static AssertConditionalFormattings(workbook: Workbook): void {
    const sheet = workbook.Worksheets["Conditional Formatting"];
    AssertEx.equal(2, sheet.ConditionalFormattings.Count);

    const primary: any = ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "A1:A5");
    AssertEx.equal(2, primary.Count);
    AssertEx.equal(1, primary.RangeCount);
    ConditionalFormattingScenarioFactory.AssertArea(primary.GetCellArea(0), 0, 0, 5, 1);

    const between = primary.get(0);
    AssertEx.equal(FormatConditionType.CellValue, between.Type);
    AssertEx.equal(OperatorType.Between, between.Operator);
    AssertEx.equal("5", between.Formula1);
    AssertEx.equal("15", between.Formula2);
    AssertEx.equal(1, between.Priority);
    AssertEx.true(between.StopIfTrue);
    AssertEx.equal(FillPattern.Solid, between.Style.Pattern);
    AssertEx.equal(Color.FromArgb(255, 255, 199, 206), between.Style.ForegroundColor);
    AssertEx.true(between.Style.Font.IsBold);
    AssertEx.equal(Color.FromArgb(255, 156, 0, 6), between.Style.Font.Color);

    const expression = primary.get(1);
    AssertEx.equal(FormatConditionType.Expression, expression.Type);
    AssertEx.equal(OperatorType.None, expression.Operator);
    AssertEx.equal("MOD(A1,2)=0", expression.Formula1);
    AssertEx.equal("", expression.Formula2);
    AssertEx.equal(2, expression.Priority);
    AssertEx.false(expression.StopIfTrue);
    AssertEx.true(expression.Style.Font.IsItalic);
    AssertEx.equal(Color.FromArgb(255, 0, 0, 255), expression.Style.Font.Color);
    AssertEx.equal(BorderStyleType.Thin, expression.Style.Borders.Bottom.LineStyle);
    AssertEx.equal(Color.FromArgb(255, 0, 0, 255), expression.Style.Borders.Bottom.Color);

    const secondary: any = ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "C1:C3");
    AssertEx.equal(1, secondary.Count);
    AssertEx.equal(1, secondary.RangeCount);
    ConditionalFormattingScenarioFactory.AssertArea(secondary.GetCellArea(0), 0, 2, 3, 1);
    const greaterThan = secondary.get(0);
    AssertEx.equal(FormatConditionType.CellValue, greaterThan.Type);
    AssertEx.equal(OperatorType.GreaterThan, greaterThan.Operator);
    AssertEx.equal("10", greaterThan.Formula1);
    AssertEx.equal("", greaterThan.Formula2);
    AssertEx.equal(3, greaterThan.Priority);
    AssertEx.equal(FillPattern.Solid, greaterThan.Style.Pattern);
    AssertEx.equal(Color.FromArgb(255, 198, 239, 206), greaterThan.Style.ForegroundColor);
    AssertEx.equal(Color.FromArgb(255, 0, 97, 0), greaterThan.Style.Font.Color);
  }

  private static GetCollectionByArea(collections: ConditionalFormattingCollection, expectedArea: string): any {
    for (let index = 0; index < collections.Count; index++) {
      const collection: any = collections.get(index);
      if (ConditionalFormattingScenarioFactory.BuildAreaKey(collection) === expectedArea) {
        return collection;
      }
    }

    throw new Error("Conditional formatting collection with area '" + expectedArea + "' was not found.");
  }

  private static BuildAreaKey(collection: any): string {
    const areas: string[] = [];
    for (let index = 0; index < collection.RangeCount; index++) {
      const area: CellArea = collection.GetCellArea(index);
      const startCell = ConditionalFormattingScenarioFactory.BuildCellName(area.StartRow, area.StartColumn);
      if (area.EndRow === area.StartRow && area.EndColumn === area.StartColumn) {
        areas.push(startCell);
        continue;
      }

      const endCell = ConditionalFormattingScenarioFactory.BuildCellName(area.EndRow, area.EndColumn);
      areas.push(startCell + ":" + endCell);
    }

    return areas.join(" ");
  }

  private static BuildCellName(rowIndex: number, columnIndex: number): string {
    let dividend = columnIndex + 1;
    let name = "";
    while (dividend > 0) {
      const remainder = (dividend - 1) % 26;
      name = String.fromCharCode("A".charCodeAt(0) + remainder) + name;
      dividend = Math.floor((dividend - remainder - 1) / 26);
    }

    return name + (rowIndex + 1).toString();
  }

  private static AssertArea(area: CellArea, firstRow: number, firstColumn: number, totalRows: number, totalColumns: number): void {
    AssertEx.equal(firstRow, area.StartRow);
    AssertEx.equal(firstColumn, area.StartColumn);
    AssertEx.equal(firstRow + totalRows - 1, area.EndRow);
    AssertEx.equal(firstColumn + totalColumns - 1, area.EndColumn);
  }

  public static CreateAdvancedConditionalFormattingWorkbook(): Workbook {
    const workbook = new Workbook();
    const sheet = workbook.Worksheets[0];
    sheet.Name = "Advanced Conditional Formatting";

    sheet.Cells["A1"].PutValue("error one");
    sheet.Cells["A2"].PutValue("ok");
    sheet.Cells["B1"].PutValue("ok");
    sheet.Cells["B2"].PutValue("warning");
    sheet.Cells["C1"].PutValue("prefix-item");
    sheet.Cells["D1"].PutValue("item-suffix");
    sheet.Cells["E1"].PutValue(todayUtcDate());
    sheet.Cells["F1"].PutValue(1);
    sheet.Cells["F2"].PutValue(1);
    sheet.Cells["G1"].PutValue(10);
    sheet.Cells["G2"].PutValue(20);
    for (let index = 1; index <= 10; index++) {
      sheet.Cells["H" + index].PutValue(index * 10);
      sheet.Cells["I" + index].PutValue(index * 10);
      sheet.Cells["J" + index].PutValue(index * 10);
      sheet.Cells["K" + index].PutValue(index * 10);
      sheet.Cells["L" + index].PutValue(index * 10);
      sheet.Cells["M" + index].PutValue(index * 10);
      sheet.Cells["N" + index].PutValue(index * 10);
    }

    const contains: any = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
    contains.AddArea(CellArea.CreateCellArea("A1", "A10"));
    const containsRule = contains.get(contains.AddCondition(FormatConditionType.ContainsText));
    containsRule.Formula1 = "error";
    const containsStyle = containsRule.Style;
    containsStyle.Pattern = FillPattern.Solid;
    containsStyle.ForegroundColor = Color.FromArgb(255, 255, 235, 156);
    containsStyle.Font.Color = Color.FromArgb(255, 156, 0, 6);
    containsRule.Style = containsStyle;

    const notContains: any = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
    notContains.AddArea(CellArea.CreateCellArea("B1", "B10"));
    const notContainsRule = notContains.get(notContains.AddCondition(FormatConditionType.NotContainsText));
    notContainsRule.Formula1 = "warning";

    const beginsWith: any = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
    beginsWith.AddArea(CellArea.CreateCellArea("C1", "C10"));
    const beginsWithRule = beginsWith.get(beginsWith.AddCondition(FormatConditionType.BeginsWith));
    beginsWithRule.Formula1 = "prefix";

    const endsWith: any = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
    endsWith.AddArea(CellArea.CreateCellArea("D1", "D10"));
    const endsWithRule = endsWith.get(endsWith.AddCondition(FormatConditionType.EndsWith));
    endsWithRule.Formula1 = "suffix";

    const timePeriod: any = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
    timePeriod.AddArea(CellArea.CreateCellArea("E1", "E10"));
    const timePeriodRule = timePeriod.get(timePeriod.AddCondition(FormatConditionType.TimePeriod));
    timePeriodRule.TimePeriod = "today";
    const timeStyle = timePeriodRule.Style;
    timeStyle.Pattern = FillPattern.Solid;
    timeStyle.ForegroundColor = Color.FromArgb(255, 221, 235, 247);
    timePeriodRule.Style = timeStyle;

    const duplicate: any = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
    duplicate.AddArea(CellArea.CreateCellArea("F1", "F10"));
    duplicate.get(duplicate.AddCondition(FormatConditionType.DuplicateValues)).Duplicate = true;

    const unique: any = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
    unique.AddArea(CellArea.CreateCellArea("G1", "G10"));
    unique.get(unique.AddCondition(FormatConditionType.UniqueValues)).Duplicate = false;

    const top10: any = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
    top10.AddArea(CellArea.CreateCellArea("H1", "H10"));
    const top10Rule = top10.get(top10.AddCondition(FormatConditionType.Top10));
    top10Rule.Top = true;
    top10Rule.Percent = true;
    top10Rule.Rank = 10;

    const bottom10: any = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
    bottom10.AddArea(CellArea.CreateCellArea("I1", "I10"));
    const bottom10Rule = bottom10.get(bottom10.AddCondition(FormatConditionType.Bottom10));
    bottom10Rule.Top = false;
    bottom10Rule.Rank = 10;

    const aboveAverage: any = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
    aboveAverage.AddArea(CellArea.CreateCellArea("J1", "J10"));
    const aboveAverageRule = aboveAverage.get(aboveAverage.AddCondition(FormatConditionType.AboveAverage));
    aboveAverageRule.Above = true;

    const belowAverage: any = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
    belowAverage.AddArea(CellArea.CreateCellArea("K1", "K10"));
    const belowAverageRule = belowAverage.get(belowAverage.AddCondition(FormatConditionType.BelowAverage));
    belowAverageRule.Above = false;
    belowAverageRule.StandardDeviation = 1;

    const colorScale: any = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
    colorScale.AddArea(CellArea.CreateCellArea("L1", "L10"));
    const colorScaleRule = colorScale.get(colorScale.AddCondition(FormatConditionType.ColorScale));
    colorScaleRule.ColorScaleCount = 3;
    colorScaleRule.MinColor = Color.FromArgb(255, 248, 105, 107);
    colorScaleRule.MidColor = Color.FromArgb(255, 255, 235, 132);
    colorScaleRule.MaxColor = Color.FromArgb(255, 99, 190, 123);

    const dataBar: any = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
    dataBar.AddArea(CellArea.CreateCellArea("M1", "M10"));
    const dataBarRule = dataBar.get(dataBar.AddCondition(FormatConditionType.DataBar));
    dataBarRule.BarColor = Color.FromArgb(255, 99, 142, 198);
    dataBarRule.NegativeBarColor = Color.FromArgb(255, 255, 0, 0);
    dataBarRule.ShowBorder = true;
    dataBarRule.Direction = "left-to-right";

    const iconSet: any = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
    iconSet.AddArea(CellArea.CreateCellArea("N1", "N10"));
    const iconSetRule = iconSet.get(iconSet.AddCondition(FormatConditionType.IconSet));
    iconSetRule.IconSetType = "4Arrows";
    iconSetRule.ReverseIcons = true;
    iconSetRule.ShowIconOnly = true;

    return workbook;
  }

  public static AssertAdvancedConditionalFormattings(workbook: Workbook): void {
    const sheet = workbook.Worksheets["Advanced Conditional Formatting"];
    AssertEx.equal(14, sheet.ConditionalFormattings.Count);

    AssertEx.equal(
      "error",
      ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "A1:A10").get(0).Formula1,
    );
    AssertEx.equal(
      FormatConditionType.ContainsText,
      ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "A1:A10").get(0).Type,
    );
    AssertEx.equal(
      FillPattern.Solid,
      ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "A1:A10").get(0).Style.Pattern,
    );

    AssertEx.equal(
      FormatConditionType.NotContainsText,
      ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "B1:B10").get(0).Type,
    );
    AssertEx.equal(
      "warning",
      ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "B1:B10").get(0).Formula1,
    );

    AssertEx.equal(
      FormatConditionType.BeginsWith,
      ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "C1:C10").get(0).Type,
    );
    AssertEx.equal(
      "prefix",
      ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "C1:C10").get(0).Formula1,
    );

    AssertEx.equal(
      FormatConditionType.EndsWith,
      ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "D1:D10").get(0).Type,
    );
    AssertEx.equal(
      "suffix",
      ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "D1:D10").get(0).Formula1,
    );

    AssertEx.equal(
      FormatConditionType.TimePeriod,
      ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "E1:E10").get(0).Type,
    );
    AssertEx.equal(
      "today",
      ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "E1:E10").get(0).TimePeriod,
    );

    AssertEx.equal(
      FormatConditionType.DuplicateValues,
      ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "F1:F10").get(0).Type,
    );
    AssertEx.true(
      ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "F1:F10").get(0).Duplicate,
    );

    AssertEx.equal(
      FormatConditionType.UniqueValues,
      ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "G1:G10").get(0).Type,
    );
    AssertEx.false(
      ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "G1:G10").get(0).Duplicate,
    );

    const topRule = ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "H1:H10").get(0);
    AssertEx.equal(FormatConditionType.Top10, topRule.Type);
    AssertEx.true(topRule.Top);
    AssertEx.true(topRule.Percent);
    AssertEx.equal(10, topRule.Rank);

    const bottomRule = ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "I1:I10").get(0);
    AssertEx.equal(FormatConditionType.Bottom10, bottomRule.Type);
    AssertEx.false(bottomRule.Top);
    AssertEx.equal(10, bottomRule.Rank);

    AssertEx.equal(
      FormatConditionType.AboveAverage,
      ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "J1:J10").get(0).Type,
    );
    AssertEx.true(
      ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "J1:J10").get(0).Above,
    );

    const belowRule = ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "K1:K10").get(0);
    AssertEx.equal(FormatConditionType.BelowAverage, belowRule.Type);
    AssertEx.false(belowRule.Above);
    AssertEx.equal(1, belowRule.StandardDeviation);

    const colorScaleRule = ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "L1:L10").get(0);
    AssertEx.equal(FormatConditionType.ColorScale, colorScaleRule.Type);
    AssertEx.equal(3, colorScaleRule.ColorScaleCount);
    AssertEx.equal(Color.FromArgb(255, 248, 105, 107), colorScaleRule.MinColor);
    AssertEx.equal(Color.FromArgb(255, 255, 235, 132), colorScaleRule.MidColor);
    AssertEx.equal(Color.FromArgb(255, 99, 190, 123), colorScaleRule.MaxColor);

    const dataBarRule = ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "M1:M10").get(0);
    AssertEx.equal(FormatConditionType.DataBar, dataBarRule.Type);
    AssertEx.equal(Color.FromArgb(255, 99, 142, 198), dataBarRule.BarColor);

    const iconSetRule = ConditionalFormattingScenarioFactory.GetCollectionByArea(sheet.ConditionalFormattings, "N1:N10").get(0);
    AssertEx.equal(FormatConditionType.IconSet, iconSetRule.Type);
    AssertEx.equal("4Arrows", iconSetRule.IconSetType);
    AssertEx.true(iconSetRule.ReverseIcons);
    AssertEx.true(iconSetRule.ShowIconOnly);
  }
}

/** Scenario factory for workbook defined names. */
export class DefinedNameScenarioFactory {
  public static CreateDefinedNamesWorkbook(): Workbook {
    const workbook = PageSetupScenarioFactory.CreatePageSetupWorkbook();
    workbook.Worksheets.Add("Scoped Sheet");
    workbook.Worksheets[1].Cells["B2"].PutValue(5);

    const global = workbook.DefinedNames.get(workbook.DefinedNames.Add("GlobalRange", "='Print Sheet'!$A$1:$C$10"));
    global.Hidden = true;
    global.Comment = "Primary range";

    const local = workbook.DefinedNames.get(workbook.DefinedNames.Add("LocalCell", "'Scoped Sheet'!$B$2", 1));
    local.Comment = "Scoped name";

    return workbook;
  }

  public static AssertDefinedNames(workbook: Workbook): void {
    PageSetupScenarioFactory.AssertPageSetup(workbook);
    AssertEx.equal(2, workbook.DefinedNames.Count);
    AssertEx.equal("5", workbook.Worksheets[1].Cells["B2"].StringValue);

    const global = workbook.DefinedNames.get(0);
    AssertEx.equal("GlobalRange", global.Name);
    AssertEx.equal("'Print Sheet'!$A$1:$C$10", global.Formula);
    AssertEx.null(global.LocalSheetIndex);
    AssertEx.true(global.Hidden);
    AssertEx.equal("Primary range", global.Comment);

    const local = workbook.DefinedNames.get(1);
    AssertEx.equal("LocalCell", local.Name);
    AssertEx.equal("'Scoped Sheet'!$B$2", local.Formula);
    AssertEx.equal(1, local.LocalSheetIndex ?? -1);
    AssertEx.false(local.Hidden);
    AssertEx.equal("Scoped name", local.Comment);
  }
}
