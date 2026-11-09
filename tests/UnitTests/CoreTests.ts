import {
  CellValueType,
  BorderStyleType,
  Color,
  FillPattern,
  Font,
  FontUnderlineType,
  HorizontalAlignmentType,
  Style,
  VerticalAlignmentType,
  Workbook,
} from "../../aspose_cells";
import { AssertEx } from "../Shared/TestInfrastructure";
import { invariantMonthDayYearTime } from "./CulturalFormatting";

export function a1IndexersRoundtrip(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  sheet.Cells.get(2, 27).PutValue("AB3");
  sheet.Cells.get(0, 0).PutValue(42);

  AssertEx.equal("AB3", sheet.Cells["AB3"].StringValue);
  AssertEx.equal("42", sheet.Cells.get(0, 0).StringValue);
  AssertEx.equal(CellValueType.IsString, sheet.Cells["AB3"].Type);
  AssertEx.equal(CellValueType.IsNumeric, sheet.Cells["A1"].Type);
}

export function putValueOverloadsAssignExpectedTypes(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  const timestamp = new Date(Date.UTC(2024, 5, 6, 7, 8, 9));

  sheet.Cells["A1"].PutValue("alpha");
  sheet.Cells["B1"].PutValue(123);
  sheet.Cells["C1"].PutValue(12.5);
  sheet.Cells["D1"].PutValue(6.02214076e23);
  sheet.Cells["E1"].PutValue(true);
  sheet.Cells["F1"].PutValue(timestamp);

  AssertEx.equal(CellValueType.IsString, sheet.Cells["A1"].Type);
  AssertEx.equal(CellValueType.IsNumeric, sheet.Cells["B1"].Type);
  AssertEx.equal(CellValueType.IsNumeric, sheet.Cells["C1"].Type);
  AssertEx.equal(CellValueType.IsNumeric, sheet.Cells["D1"].Type);
  AssertEx.equal(CellValueType.IsBool, sheet.Cells["E1"].Type);
  AssertEx.equal(CellValueType.IsDateTime, sheet.Cells["F1"].Type);

  AssertEx.equal("alpha", sheet.Cells["A1"].Value as string);
  AssertEx.equal(123, sheet.Cells["B1"].Value as number);
  AssertEx.equal(12.5, sheet.Cells["C1"].Value as number);
  AssertEx.true(
    Math.abs((sheet.Cells["D1"].Value as number) - 6.02214076e23) < 1e10,
    "Double PutValue should retain magnitude.",
  );
  AssertEx.equal(true, sheet.Cells["E1"].Value as boolean);
  AssertEx.equal(timestamp, sheet.Cells["F1"].Value as Date);
}

export function stringValueFormatsSupportedScalarTypes(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  const timestamp = new Date(Date.UTC(2024, 5, 6, 7, 8, 9));

  sheet.Cells["A1"].PutValue(true);
  sheet.Cells["B1"].PutValue(123);
  sheet.Cells["C1"].PutValue(12.5);
  sheet.Cells["D1"].PutValue(timestamp);

  AssertEx.equal("TRUE", sheet.Cells["A1"].StringValue);
  AssertEx.equal("TRUE", sheet.Cells["A1"].DisplayStringValue);
  AssertEx.equal("123", sheet.Cells["B1"].StringValue);
  AssertEx.equal("123", sheet.Cells["B1"].DisplayStringValue);
  AssertEx.equal("12.5", sheet.Cells["C1"].StringValue);
  AssertEx.equal("12.5", sheet.Cells["C1"].DisplayStringValue);
  const expectedDateText = invariantMonthDayYearTime(timestamp, true);
  AssertEx.equal(expectedDateText, sheet.Cells["D1"].StringValue);
  AssertEx.equal(expectedDateText, sheet.Cells["D1"].DisplayStringValue);
}

export function cellsStylePropertyRoundtrip(): void {
  const workbook = new Workbook();
  const cells = workbook.Worksheets[0].Cells;

  const style = cells.Style;
  style.Font.Name = "Tahoma";
  style.Font.IsBold = true;
  style.Number = 14;
  cells.Style = style;

  const persisted = cells.Style;
  AssertEx.equal("Tahoma", persisted.Font.Name);
  AssertEx.true(persisted.Font.IsBold);
  AssertEx.equal(14, persisted.Number);
}

export function putValueStringAndObjectOverloadsFollowAsposeShape(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];

  sheet.Cells["A1"].PutValue("123", true);
  sheet.Cells["B1"].PutValue("456", false, false);
  sheet.Cells["C1"].PutValue(true);

  AssertEx.equal(CellValueType.IsNumeric, sheet.Cells["A1"].Type);
  AssertEx.equal(123, sheet.Cells["A1"].Value as number);

  AssertEx.equal(CellValueType.IsString, sheet.Cells["B1"].Type);
  AssertEx.equal("456", sheet.Cells["B1"].Value as string);

  AssertEx.equal(CellValueType.IsBool, sheet.Cells["C1"].Type);
  AssertEx.equal(true, sheet.Cells["C1"].Value as boolean);
}

export function displayStringValueAppliesNumericFormatsAndStringValueStaysRaw(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];

  const grouped = sheet.Cells["A1"];
  grouped.PutValue(1234.567);
  const groupedStyle = grouped.GetStyle();
  groupedStyle.NumberFormat = "#,##0.00";
  grouped.SetStyle(groupedStyle);

  const percent = sheet.Cells["B1"];
  percent.PutValue(0.125);
  const percentStyle = percent.GetStyle();
  percentStyle.NumberFormat = "0.00%";
  percent.SetStyle(percentStyle);

  const scientific = sheet.Cells["C1"];
  scientific.PutValue(1234);
  const scientificStyle = scientific.GetStyle();
  scientificStyle.NumberFormat = "0.00E+00";
  scientific.SetStyle(scientificStyle);

  const fraction = sheet.Cells["D1"];
  fraction.PutValue(1.25);
  const fractionStyle = fraction.GetStyle();
  fractionStyle.NumberFormat = "# ?/?";
  fraction.SetStyle(fractionStyle);

  const negative = sheet.Cells["E1"];
  negative.PutValue(-12.3);
  const negativeStyle = negative.GetStyle();
  negativeStyle.NumberFormat = "#,##0.00_);(#,##0.00)";
  negative.SetStyle(negativeStyle);

  const color = sheet.Cells["F1"];
  color.PutValue(1.25);
  const colorStyle = color.GetStyle();
  colorStyle.NumberFormat = "[Blue]0.000";
  color.SetStyle(colorStyle);

  const conditionalHigh = sheet.Cells["G1"];
  conditionalHigh.PutValue(125);
  const conditionalStyle = conditionalHigh.GetStyle();
  conditionalStyle.NumberFormat = "[>100]0.0;\"small\"";
  conditionalHigh.SetStyle(conditionalStyle);

  const conditionalLow = sheet.Cells["H1"];
  conditionalLow.PutValue(10);
  conditionalLow.SetStyle(conditionalStyle);

  AssertEx.equal("1234.567", grouped.StringValue);
  AssertEx.equal("1,234.57", grouped.DisplayStringValue);
  AssertEx.equal("0.125", percent.StringValue);
  AssertEx.equal("12.50%", percent.DisplayStringValue);
  AssertEx.equal("1234", scientific.StringValue);
  AssertEx.equal("1.23E+03", scientific.DisplayStringValue);
  AssertEx.equal("1.25", fraction.StringValue);
  AssertEx.equal("1 1/4", fraction.DisplayStringValue);
  AssertEx.equal("-12.3", negative.StringValue);
  AssertEx.equal("(12.30)", negative.DisplayStringValue);
  AssertEx.equal("1.25", color.StringValue);
  AssertEx.equal("1.250", color.DisplayStringValue);
  AssertEx.equal("125", conditionalHigh.StringValue);
  AssertEx.equal("125.0", conditionalHigh.DisplayStringValue);
  AssertEx.equal("10", conditionalLow.StringValue);
  AssertEx.equal("small", conditionalLow.DisplayStringValue);
}

export function displayStringValueAppliesDateAndTextFormats(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  const timestamp = new Date(2024, 5, 6, 7, 8, 9);

  const builtInDate = sheet.Cells["A2"];
  builtInDate.PutValue(timestamp);
  const builtInDateStyle = builtInDate.GetStyle();
  builtInDateStyle.Number = 14;
  builtInDate.SetStyle(builtInDateStyle);

  const customDate = sheet.Cells["B2"];
  customDate.PutValue(timestamp);
  const customDateStyle = customDate.GetStyle();
  customDateStyle.NumberFormat = "d-mmm-yy h:mm AM/PM";
  customDate.SetStyle(customDateStyle);

  const elapsed = sheet.Cells["C2"];
  elapsed.PutValue(timestamp);
  const elapsedStyle = elapsed.GetStyle();
  elapsedStyle.Number = 46;
  elapsed.SetStyle(elapsedStyle);

  const textCell = sheet.Cells["D2"];
  textCell.PutValue("ABC");
  const textStyle = textCell.GetStyle();
  textStyle.NumberFormat = "0;0;0;\"Item \"@";
  textCell.SetStyle(textStyle);

  AssertEx.equal("5/6/2024 7:08", builtInDate.StringValue);
  AssertEx.equal("05-06-24", builtInDate.DisplayStringValue);
  AssertEx.equal("5/6/2024 7:08", customDate.StringValue);
  AssertEx.equal("6-May-24 7:08 AM", customDate.DisplayStringValue);
  AssertEx.equal("5/6/2024 7:08", elapsed.StringValue);
  AssertEx.equal("1090039:08:09", elapsed.DisplayStringValue);
  AssertEx.equal("ABC", textCell.StringValue);
  AssertEx.equal("Item ABC", textCell.DisplayStringValue);
}

export function fontEqualsAndToStringFollowAsposeShape(): void {
  const left = new Font();
  left.Name = "Arial";
  left.Size = 12;
  left.IsBold = true;
  left.IsItalic = false;
  left.Underline = FontUnderlineType.Single;
  left.IsStrikeout = true;
  left.Color = Color.FromArgb(255, 12, 34, 56);

  const right = new Font();
  right.Name = "Arial";
  right.Size = 12;
  right.IsBold = true;
  right.IsItalic = false;
  right.Underline = FontUnderlineType.Single;
  right.IsStrikeout = true;
  right.Color = Color.FromArgb(255, 12, 34, 56);

  AssertEx.true(left.Equals(right));

  const text = left.ToString();
  AssertEx.true(text.indexOf("Arial") >= 0);
  AssertEx.true(text.indexOf("Underline=Single") >= 0);
}

export function styleEqualsHashCodeAreConsistent(): void {
  const left = new Style();
  left.Font.Name = "Arial";
  left.Font.IsBold = true;
  left.Borders.Left.LineStyle = BorderStyleType.Thin;
  left.Borders.Left.Color = Color.FromArgb(255, 1, 2, 3);
  left.Number = 14;
  left.HorizontalAlignment = HorizontalAlignmentType.Center;
  left.VerticalAlignment = VerticalAlignmentType.Top;
  left.WrapText = true;
  left.IsLocked = false;
  left.IsHidden = true;

  const right = new Style();
  right.Font.Name = "Arial";
  right.Font.IsBold = true;
  right.Borders.Left.LineStyle = BorderStyleType.Thin;
  right.Borders.Left.Color = Color.FromArgb(255, 1, 2, 3);
  right.Number = 14;
  right.HorizontalAlignment = HorizontalAlignmentType.Center;
  right.VerticalAlignment = VerticalAlignmentType.Top;
  right.WrapText = true;
  right.IsLocked = false;
  right.IsHidden = true;

  AssertEx.true(left.Equals(right));
  AssertEx.equal(left.GetHashCode(), right.GetHashCode());
}

export function styleCopyCopiesAllPublicSettings(): void {
  const source = new Style();
  source.Font.Name = "Consolas";
  source.Font.IsBold = true;
  source.Borders.Left.LineStyle = BorderStyleType.Medium;
  source.Borders.Left.Color = Color.FromArgb(255, 9, 8, 7);
  source.Pattern = FillPattern.Solid;
  source.ForegroundColor = Color.FromArgb(255, 10, 20, 30);
  source.BackgroundColor = Color.FromArgb(255, 40, 50, 60);
  source.Number = 14;
  source.Custom = "0.00";
  source.HorizontalAlignment = HorizontalAlignmentType.Center;
  source.VerticalAlignment = VerticalAlignmentType.Top;
  source.WrapText = true;
  source.IndentLevel = 2;
  source.TextRotation = 45;
  source.ShrinkToFit = true;
  source.ReadingOrder = 2;
  source.RelativeIndent = 1;
  source.IsLocked = false;
  source.IsHidden = true;

  const target = new Style();
  target.Copy(source);

  AssertEx.true(target.Equals(source));

  source.Font.Name = "Mutated";
  AssertEx.equal("Consolas", target.Font.Name);
}

export function cellGetStyleBoolOverloadMatchesGetStyle(): void {
  const workbook = new Workbook();
  const cell = workbook.Worksheets[0].Cells["A1"];
  cell.PutValue(1);

  const style = cell.GetStyle();
  style.Font.Name = "Calibri";
  style.Font.IsBold = true;
  cell.SetStyle(style);

  const fromDefault = cell.GetStyle();
  const fromBoolFalse = cell.GetStyle(false);
  const fromBoolTrue = cell.GetStyle(true);

  AssertEx.true(fromDefault.Equals(fromBoolFalse));
  AssertEx.true(fromDefault.Equals(fromBoolTrue));
}

export function cellSetStyleBoolOverloadMatchesSetStyle(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  const left = sheet.Cells["A1"];
  const right = sheet.Cells["B1"];
  left.PutValue(1);
  right.PutValue(1);

  const style = left.GetStyle();
  style.Font.Name = "Calibri";
  style.Font.IsBold = true;
  style.NumberFormat = "#,##0.00";

  left.SetStyle(style);
  right.SetStyle(style, true);

  AssertEx.true(left.GetStyle().Equals(right.GetStyle()));
}
