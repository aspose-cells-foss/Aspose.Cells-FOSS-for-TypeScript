import {
  BorderStyleType,
  CellValueType,
  Color,
  FillPattern,
  FontUnderlineType,
  HorizontalAlignmentType,
  Style,
  StyleFlag,
  VerticalAlignmentType,
  VisibilityType,
  Workbook,
} from "../../aspose_cells";
import { AssertEx } from "../Shared/TestInfrastructure";
import {
  germanDecimalSeparator,
  germanLongDate,
  germanLongTime,
  frenchGroupedNumber,
  weekdayMonthDayYear,
  yenGroupedNumber,
} from "./CulturalFormatting";
import { StyleScenarioFactory } from "../Shared/ScenarioFactories";

export function cellSetStyleStyleFlagAppliesSelectedProperties(): void {
  const workbook = new Workbook();
  const cell = workbook.Worksheets[0].Cells["A1"];
  cell.PutValue(1);

  const baseline = cell.GetStyle();
  baseline.Font.Name = "Calibri";
  baseline.Font.IsBold = false;
  baseline.Font.Size = 11;
  baseline.Font.IsItalic = false;
  baseline.Font.Underline = FontUnderlineType.None;
  baseline.Font.IsStrikeout = false;
  baseline.Font.Color = Color.FromArgb(255, 0, 0, 0);
  baseline.NumberFormat = "0";
  baseline.HorizontalAlignment = HorizontalAlignmentType.Left;
  baseline.IsHidden = false;
  baseline.IsLocked = true;
  baseline.QuotePrefix = false;
  baseline.Borders.Left.LineStyle = BorderStyleType.None;
  baseline.Borders.Right.LineStyle = BorderStyleType.None;
  baseline.Borders.Top.LineStyle = BorderStyleType.None;
  baseline.Borders.Bottom.LineStyle = BorderStyleType.None;
  baseline.Borders.Diagonal.LineStyle = BorderStyleType.None;
  baseline.Borders.DiagonalUp = false;
  baseline.Borders.DiagonalDown = false;
  cell.SetStyle(baseline);

  const incoming = cell.GetStyle();
  incoming.Font.Name = "Arial";
  incoming.Font.IsBold = true;
  incoming.Font.Size = 14;
  incoming.Font.IsItalic = true;
  incoming.Font.Underline = FontUnderlineType.Single;
  incoming.Font.IsStrikeout = true;
  incoming.Font.Color = Color.FromArgb(255, 200, 10, 20);
  incoming.NumberFormat = "#,##0.00";
  incoming.HorizontalAlignment = HorizontalAlignmentType.Center;
  incoming.IndentLevel = 3;
  incoming.IsHidden = true;
  incoming.IsLocked = false;
  incoming.QuotePrefix = true;
  incoming.Borders.Left.LineStyle = BorderStyleType.Thin;
  incoming.Borders.Left.Color = Color.FromArgb(255, 1, 2, 3);
  incoming.Borders.Right.LineStyle = BorderStyleType.Dashed;
  incoming.Borders.Right.Color = Color.FromArgb(255, 30, 40, 50);
  incoming.Borders.Top.LineStyle = BorderStyleType.Double;
  incoming.Borders.Top.Color = Color.FromArgb(255, 60, 70, 80);
  incoming.Borders.Bottom.LineStyle = BorderStyleType.Dotted;
  incoming.Borders.Bottom.Color = Color.FromArgb(255, 90, 100, 110);
  incoming.Borders.Diagonal.LineStyle = BorderStyleType.Medium;
  incoming.Borders.Diagonal.Color = Color.FromArgb(255, 9, 8, 7);
  incoming.Borders.DiagonalUp = true;
  incoming.Borders.DiagonalDown = true;

  const flag = new StyleFlag();
  flag.FontName = true;
  flag.FontBold = true;
  flag.FontSize = true;
  flag.FontItalic = true;
  flag.FontUnderline = true;
  flag.FontStrike = true;
  flag.FontColor = true;
  flag.NumberFormat = true;
  flag.Indent = true;
  flag.Hidden = true;
  flag.Locked = true;
  flag.QuotePrefix = true;
  flag.LeftBorder = true;
  flag.RightBorder = true;
  flag.TopBorder = true;
  flag.BottomBorder = true;
  flag.DiagonalUpBorder = true;
  flag.DiagonalDownBorder = true;
  cell.SetStyle(incoming, flag);

  const applied = cell.GetStyle();
  AssertEx.equal("Arial", applied.Font.Name);
  AssertEx.true(applied.Font.IsBold);
  AssertEx.equal(14, applied.Font.Size);
  AssertEx.true(applied.Font.IsItalic);
  AssertEx.equal(FontUnderlineType.Single, applied.Font.Underline);
  AssertEx.true(applied.Font.IsStrikeout);
  AssertEx.equal(Color.FromArgb(255, 200, 10, 20), applied.Font.Color);
  AssertEx.equal("#,##0.00", applied.NumberFormat);
  AssertEx.equal(HorizontalAlignmentType.Left, applied.HorizontalAlignment);
  AssertEx.equal(3, applied.IndentLevel);
  AssertEx.true(applied.IsHidden);
  AssertEx.false(applied.IsLocked);
  AssertEx.true(applied.QuotePrefix);
  AssertEx.equal(BorderStyleType.Thin, applied.Borders.Left.LineStyle);
  AssertEx.equal(Color.FromArgb(255, 1, 2, 3), applied.Borders.Left.Color);
  AssertEx.equal(BorderStyleType.Dashed, applied.Borders.Right.LineStyle);
  AssertEx.equal(Color.FromArgb(255, 30, 40, 50), applied.Borders.Right.Color);
  AssertEx.equal(BorderStyleType.Double, applied.Borders.Top.LineStyle);
  AssertEx.equal(Color.FromArgb(255, 60, 70, 80), applied.Borders.Top.Color);
  AssertEx.equal(BorderStyleType.Dotted, applied.Borders.Bottom.LineStyle);
  AssertEx.equal(Color.FromArgb(255, 90, 100, 110), applied.Borders.Bottom.Color);
  AssertEx.true(applied.Borders.DiagonalUp);
  AssertEx.true(applied.Borders.DiagonalDown);
  AssertEx.equal(BorderStyleType.Medium, applied.Borders.Diagonal.LineStyle);
  AssertEx.equal(Color.FromArgb(255, 9, 8, 7), applied.Borders.Diagonal.Color);
}

export function cellSetStyleStyleFlagAllAppliesFullStyle(): void {
  const workbook = new Workbook();
  const cell = workbook.Worksheets[0].Cells["B2"];
  cell.PutValue(1);

  const style = cell.GetStyle();
  style.Font.Name = "Arial";
  style.Font.IsBold = true;
  style.NumberFormat = "#,##0.00";
  style.HorizontalAlignment = HorizontalAlignmentType.Center;
  style.VerticalAlignment = VerticalAlignmentType.Top;
  style.WrapText = true;
  style.IndentLevel = 2;
  style.IsHidden = true;
  style.IsLocked = false;
  style.QuotePrefix = true;
  style.Borders.Left.LineStyle = BorderStyleType.Thin;
  style.Borders.Left.Color = Color.FromArgb(255, 10, 20, 30);

  const flag = new StyleFlag();
  flag.All = true;
  cell.SetStyle(style, flag);

  const applied = cell.GetStyle();
  AssertEx.true(applied.Equals(style));
}

export function cellSetStyleStyleFlagAlignmentsAppliesIndentSettings(): void {
  const workbook = new Workbook();
  const cell = workbook.Worksheets[0].Cells["C3"];
  cell.PutValue(1);

  const baseline = cell.GetStyle();
  baseline.HorizontalAlignment = HorizontalAlignmentType.Left;
  baseline.VerticalAlignment = VerticalAlignmentType.Bottom;
  baseline.WrapText = false;
  baseline.IndentLevel = 0;
  baseline.RelativeIndent = 0;
  cell.SetStyle(baseline);

  const incoming = cell.GetStyle();
  incoming.HorizontalAlignment = HorizontalAlignmentType.Center;
  incoming.VerticalAlignment = VerticalAlignmentType.Top;
  incoming.WrapText = true;
  incoming.IndentLevel = 4;
  incoming.RelativeIndent = 2;

  const flag = new StyleFlag();
  flag.Alignments = true;
  cell.SetStyle(incoming, flag);

  const applied = cell.GetStyle();
  AssertEx.equal(HorizontalAlignmentType.Center, applied.HorizontalAlignment);
  AssertEx.equal(VerticalAlignmentType.Top, applied.VerticalAlignment);
  AssertEx.true(applied.WrapText);
  AssertEx.equal(4, applied.IndentLevel);
  AssertEx.equal(2, applied.RelativeIndent);
}

export function displayStringValueUsesWorkbookCultureAndLocaleDirectives(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  workbook.Settings.Culture = "fr-FR";

  const numeric = sheet.Cells["A3"];
  numeric.PutValue(1234.5);
  const numericStyle = numeric.GetStyle();
  numericStyle.NumberFormat = "#,##0.00";
  numeric.SetStyle(numericStyle);

  const dateCell = sheet.Cells["B3"];
  const timestamp = new Date(2024, 5, 6, 7, 8, 9);
  dateCell.PutValue(timestamp);
  const dateStyle = dateCell.GetStyle();
  dateStyle.NumberFormat = "dddd, mmmm d, yyyy";
  dateCell.SetStyle(dateStyle);

  const englishDate = sheet.Cells["C3"];
  englishDate.PutValue(timestamp);
  const englishDateStyle = englishDate.GetStyle();
  englishDateStyle.NumberFormat = "[$-409]dddd, mmmm d, yyyy";
  englishDate.SetStyle(englishDateStyle);

  const yenCell = sheet.Cells["D3"];
  yenCell.PutValue(1234.5);
  const yenStyle = yenCell.GetStyle();
  yenStyle.NumberFormat = "[$\u00A5-411]#,##0.00";
  yenCell.SetStyle(yenStyle);

  const longDateCell = sheet.Cells["E3"];
  longDateCell.PutValue(timestamp);
  const longDateStyle = longDateCell.GetStyle();
  longDateStyle.NumberFormat = "[$-F800]";
  longDateCell.SetStyle(longDateStyle);

  AssertEx.equal(frenchGroupedNumber(1234.5), numeric.DisplayStringValue);
  AssertEx.equal(weekdayMonthDayYear(timestamp, "fr-FR"), dateCell.DisplayStringValue);
  AssertEx.equal(weekdayMonthDayYear(timestamp, "en-US"), englishDate.DisplayStringValue);
  AssertEx.equal(yenGroupedNumber(1234.5), yenCell.DisplayStringValue);

  workbook.Settings.Culture = "de-DE";
  AssertEx.equal(germanLongDate(timestamp), longDateCell.DisplayStringValue);
}

export function displayStringValueUsesEnglishMonthAbbreviationsWithoutLocaleDirective(): void {
  const workbook = new Workbook();
  workbook.Settings.Culture = "zh-CN";
  const cell = workbook.Worksheets[0].Cells["A5"];
  cell.PutValue(new Date(2025, 0, 1, 0, 0, 0));
  const style = cell.GetStyle();
  style.NumberFormat = "mmm-yy";
  cell.SetStyle(style);

  AssertEx.equal("Jan-25", cell.DisplayStringValue);
}

export function displayStringValueAppliesExtendedDateTokens(): void {
  const workbook = new Workbook();
  workbook.Settings.Culture = "en-US";
  const sheet = workbook.Worksheets[0];
  const timestamp = new Date(2024, 4, 6, 7, 8, 9, 345);

  const monthInitial = sheet.Cells["A4"];
  monthInitial.PutValue(timestamp);
  const monthInitialStyle = monthInitial.GetStyle();
  monthInitialStyle.NumberFormat = "mmmmm d, yyyy";
  monthInitial.SetStyle(monthInitialStyle);

  const abbreviatedDate = sheet.Cells["B4"];
  abbreviatedDate.PutValue(timestamp);
  const abbreviatedStyle = abbreviatedDate.GetStyle();
  abbreviatedStyle.NumberFormat = "ddd, mmm d yyyy";
  abbreviatedDate.SetStyle(abbreviatedStyle);

  const fractionalSeconds = sheet.Cells["C4"];
  fractionalSeconds.PutValue(timestamp);
  const fractionalStyle = fractionalSeconds.GetStyle();
  fractionalStyle.NumberFormat = "h:mm:ss.000 AM/PM";
  fractionalSeconds.SetStyle(fractionalStyle);

  const shortFraction = sheet.Cells["D4"];
  shortFraction.PutValue(timestamp);
  const shortFractionStyle = shortFraction.GetStyle();
  shortFractionStyle.NumberFormat = "hh:mm:ss.00";
  shortFraction.SetStyle(shortFractionStyle);

  const shortAmPm = sheet.Cells["E4"];
  shortAmPm.PutValue(timestamp);
  const shortAmPmStyle = shortAmPm.GetStyle();
  shortAmPmStyle.NumberFormat = "h A/P";
  shortAmPm.SetStyle(shortAmPmStyle);

  AssertEx.equal("M 6, 2024", monthInitial.DisplayStringValue);
  AssertEx.equal("Mon, May 6 2024", abbreviatedDate.DisplayStringValue);
  AssertEx.equal("7:08:09.345 AM", fractionalSeconds.DisplayStringValue);
  AssertEx.equal("07:08:09.34", shortFraction.DisplayStringValue);
  AssertEx.equal("7 A", shortAmPm.DisplayStringValue);
}

export function displayStringValueAppliesLongTimeAndElapsedFractionWithCulture(): void {
  const workbook = new Workbook();
  workbook.Settings.Culture = "de-DE";
  const sheet = workbook.Worksheets[0];
  const timestamp = new Date(2024, 5, 6, 7, 8, 9, 345);

  const longTimeCell = sheet.Cells["F4"];
  longTimeCell.PutValue(timestamp);
  const longTimeStyle = longTimeCell.GetStyle();
  longTimeStyle.NumberFormat = "[$-F400]";
  longTimeCell.SetStyle(longTimeStyle);

  const elapsedCell = sheet.Cells["G4"];
  elapsedCell.PutValue(new Date(1899, 11, 31, 7, 8, 9, 345));
  const elapsedStyle = elapsedCell.GetStyle();
  elapsedStyle.NumberFormat = "[h]:mm:ss.000";
  elapsedCell.SetStyle(elapsedStyle);

  AssertEx.equal(germanLongTime(timestamp), longTimeCell.DisplayStringValue);
  AssertEx.equal("7:08:09" + germanDecimalSeparator() + "345", elapsedCell.DisplayStringValue);
}

export function displayStringValueElapsedTimeHonors1904DateSystem(): void {
  const workbook = new Workbook();
  workbook.Settings.Date1904 = true;
  const cell = workbook.Worksheets[0].Cells["A1"];
  cell.PutValue(new Date(1904, 0, 2, 12, 0, 0));
  const style = cell.GetStyle();
  style.NumberFormat = "[h]:mm:ss";
  cell.SetStyle(style);

  AssertEx.equal("36:00:00", cell.DisplayStringValue);
}

export function formulaPropertyNormalizesAndPreservesCachedValue(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  const cell = sheet.Cells["C3"];

  cell.PutValue(20);
  cell.Formula = "A1+B1";

  AssertEx.equal(CellValueType.IsUnknown, cell.Type);
  AssertEx.equal("=A1+B1", cell.Formula);
  AssertEx.equal("20", cell.StringValue);
  AssertEx.equal(20, cell.Value as number);
}

export function blankCellsAreBlankByDefault(): void {
  const workbook = new Workbook();
  const cell = workbook.Worksheets[0].Cells["Z99"];

  AssertEx.equal(CellValueType.IsNull, cell.Type);
  AssertEx.null(cell.Value);
  AssertEx.equal("", cell.StringValue);
}

export function worksheetNameAndCollectionGuards(): void {
  const workbook = new Workbook();
  workbook.Worksheets.Add("Data");

  AssertEx.throws(() => {
    workbook.Worksheets.Add("data");
  });
  AssertEx.throws(() => {
    workbook.Worksheets[0].Cells["1A"];
  });
  AssertEx.throws(() => {
    workbook.Worksheets[0].Cells.get(-1, 0);
  });

  workbook.Worksheets[0].Cells["AB3"].PutValue("ok");
  AssertEx.equal("ok", workbook.Worksheets[0].Cells.get(2, 27).StringValue);
}

export function styleMutationRequiresSetStyleAndReturnsClones(): void {
  const workbook = new Workbook();
  const cell = workbook.Worksheets[0].Cells["A1"];

  const style = cell.GetStyle();
  style.Font.IsBold = true;
  style.HorizontalAlignment = HorizontalAlignmentType.Right;

  const untouched = cell.GetStyle();
  AssertEx.false(untouched.Font.IsBold);
  AssertEx.equal(HorizontalAlignmentType.General, untouched.HorizontalAlignment);

  cell.SetStyle(style);
  const applied = cell.GetStyle();
  AssertEx.true(applied.Font.IsBold);
  AssertEx.equal(HorizontalAlignmentType.Right, applied.HorizontalAlignment);

  applied.Font.IsItalic = true;
  AssertEx.false(cell.GetStyle().Font.IsItalic);
}

export function styleApiCoversAllPublicSettings(): void {
  const workbook = new Workbook();
  const primaryCell = workbook.Worksheets[0].Cells["A1"];
  primaryCell.PutValue(1);

  const primaryStyle = primaryCell.GetStyle();
  StyleScenarioFactory.ApplyPrimaryStyle(primaryStyle);

  const untouched = primaryCell.GetStyle();
  AssertEx.equal("Calibri", untouched.Font.Name);
  AssertEx.equal(FillPattern.None, untouched.Pattern);
  AssertEx.true(untouched.IsLocked);
  AssertEx.false(untouched.IsHidden);
  AssertEx.equal(0, untouched.IndentLevel);
  AssertEx.equal(0, untouched.TextRotation);
  AssertEx.equal(0, untouched.ReadingOrder);
  AssertEx.false(untouched.ShrinkToFit);
  AssertEx.false(untouched.Font.IsStrikeout);
  AssertEx.false(untouched.Borders.DiagonalUp);
  AssertEx.false(untouched.Borders.DiagonalDown);
  AssertEx.equal("General", untouched.NumberFormat);

  primaryCell.SetStyle(primaryStyle);
  StyleScenarioFactory.AssertPrimaryStyle(primaryCell.GetStyle());

  const mutatedClone = primaryCell.GetStyle();
  mutatedClone.Font.Name = "Mutated";
  mutatedClone.Font.IsStrikeout = false;
  mutatedClone.Pattern = FillPattern.None;
  mutatedClone.ForegroundColor = Color.Empty;
  mutatedClone.BackgroundColor = Color.Empty;
  mutatedClone.Borders.Right.LineStyle = BorderStyleType.None;
  mutatedClone.Borders.Diagonal.LineStyle = BorderStyleType.None;
  mutatedClone.Borders.DiagonalUp = false;
  mutatedClone.Borders.DiagonalDown = false;
  mutatedClone.Number = 0;
  mutatedClone.HorizontalAlignment = HorizontalAlignmentType.Left;
  mutatedClone.VerticalAlignment = VerticalAlignmentType.Bottom;
  mutatedClone.WrapText = false;
  mutatedClone.IndentLevel = 0;
  mutatedClone.TextRotation = 0;
  mutatedClone.ShrinkToFit = false;
  mutatedClone.ReadingOrder = 0;
  mutatedClone.RelativeIndent = 0;
  mutatedClone.IsLocked = true;
  mutatedClone.IsHidden = false;

  const persistedPrimary = primaryCell.GetStyle();
  AssertEx.equal("Arial", persistedPrimary.Font.Name);
  AssertEx.true(persistedPrimary.Font.IsStrikeout);
  AssertEx.equal(FillPattern.LightGrid, persistedPrimary.Pattern);
  AssertEx.equal(BorderStyleType.MediumDashDot, persistedPrimary.Borders.Right.LineStyle);
  AssertEx.equal(BorderStyleType.SlantedDashDot, persistedPrimary.Borders.Diagonal.LineStyle);
  AssertEx.true(persistedPrimary.Borders.DiagonalUp);
  AssertEx.true(persistedPrimary.Borders.DiagonalDown);
  AssertEx.equal(4, persistedPrimary.Number);
  AssertEx.equal("#,##0.00", persistedPrimary.NumberFormat);
  AssertEx.equal(HorizontalAlignmentType.Distributed, persistedPrimary.HorizontalAlignment);
  AssertEx.equal(VerticalAlignmentType.Distributed, persistedPrimary.VerticalAlignment);
  AssertEx.true(persistedPrimary.WrapText);
  AssertEx.equal(2, persistedPrimary.IndentLevel);
  AssertEx.equal(45, persistedPrimary.TextRotation);
  AssertEx.true(persistedPrimary.ShrinkToFit);
  AssertEx.equal(2, persistedPrimary.ReadingOrder);
  AssertEx.equal(1, persistedPrimary.RelativeIndent);
  AssertEx.false(persistedPrimary.IsLocked);
  AssertEx.true(persistedPrimary.IsHidden);

  const customCell = workbook.Worksheets[0].Cells["B2"];
  const customStyle = customCell.GetStyle();
  StyleScenarioFactory.ApplyCustomNumberStyle(customStyle);
  customCell.SetStyle(customStyle);

  StyleScenarioFactory.AssertCustomNumberStyle(customCell.GetStyle());
  AssertEx.equal(CellValueType.IsNull, customCell.Type);

  const numberFormatStyle = new Style();
  numberFormatStyle.NumberFormat = "0.00%";
  AssertEx.equal(10, numberFormatStyle.Number);
  AssertEx.null(numberFormatStyle.Custom);
  numberFormatStyle.NumberFormat = "[Blue]0.000";
  AssertEx.equal(0, numberFormatStyle.Number);
  AssertEx.equal("[Blue]0.000", numberFormatStyle.Custom);

  AssertEx.throws(() => {
    numberFormatStyle.IndentLevel = -1;
  });
  AssertEx.throws(() => {
    numberFormatStyle.TextRotation = 181;
  });
  AssertEx.throws(() => {
    numberFormatStyle.ReadingOrder = 3;
  });
}

export function worksheetRowColumnAndMergeApisMutateExpectedSettings(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];

  sheet.VisibilityType = VisibilityType.Hidden;
  sheet.Cells.Rows[2].Height = 19.75;
  sheet.Cells.Rows[4].IsHidden = true;
  sheet.Cells.Columns[1].Width = 25.5;
  sheet.Cells.Columns[3].IsHidden = true;
  sheet.Cells.Merge(1, 1, 2, 3);

  AssertEx.equal(VisibilityType.Hidden, sheet.VisibilityType);
  AssertEx.equal(19.75, sheet.Cells.Rows[2].Height ?? 0);
  AssertEx.true(sheet.Cells.Rows[4].IsHidden);
  AssertEx.equal(25.5, sheet.Cells.Columns[1].Width ?? 0);
  AssertEx.true(sheet.Cells.Columns[3].IsHidden);
  AssertEx.equal(1, sheet.Cells.MergedCells.length);
  AssertEx.equal(1, sheet.Cells.MergedCells[0].StartRow);
  AssertEx.equal(1, sheet.Cells.MergedCells[0].StartColumn);
  AssertEx.equal(2, sheet.Cells.MergedCells[0].EndRow - sheet.Cells.MergedCells[0].StartRow + 1);
  AssertEx.equal(3, sheet.Cells.MergedCells[0].EndColumn - sheet.Cells.MergedCells[0].StartColumn + 1);
  AssertEx.throws(() => {
    sheet.Cells.Merge(2, 2, 2, 2);
  });
}

export function worksheetViewApisMutateExpectedSettings(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];

  sheet.TabColor = Color.FromArgb(255, 34, 68, 102);
  sheet.ShowGridlines = false;
  sheet.ShowRowColumnHeaders = false;
  sheet.ShowZeros = false;
  sheet.RightToLeft = true;
  sheet.Zoom = 85;

  AssertEx.equal(Color.FromArgb(255, 34, 68, 102), sheet.TabColor);
  AssertEx.false(sheet.ShowGridlines);
  AssertEx.false(sheet.ShowRowColumnHeaders);
  AssertEx.false(sheet.ShowZeros);
  AssertEx.true(sheet.RightToLeft);
  AssertEx.equal(85, sheet.Zoom);

  sheet.TabColor = Color.Empty;
  AssertEx.equal(Color.Empty, sheet.TabColor);
  AssertEx.throws(() => {
    sheet.Zoom = 9;
  });
  AssertEx.throws(() => {
    sheet.Zoom = 401;
  });
}
