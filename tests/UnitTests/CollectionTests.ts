import {
  CellArea,
  Color,
  FillPattern,
  FormatConditionType,
  ImageType,
  OperatorType,
  PageOrientationType,
  PaperSizeType,
  TableStyleType,
  TotalsCalculation,
  ValidationType,
  Workbook,
} from "../../aspose_cells";
import { AssertEx } from "../Shared/TestInfrastructure";
import { AutoFilterScenarioFactory } from "../Shared/AutoFilterScenarioFactory";
import { PictureScenarioFactory } from "../Shared/PictureScenarioFactory";
import { WorkbookMetadataScenarioFactory } from "../Shared/WorkbookMetadataScenarioFactory";

export function worksheetProtectionApisMutateExpectedSettings(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];

  sheet.Protect();
  sheet.Protection.Objects = true;
  sheet.Protection.FormatCells = true;
  sheet.Protection.InsertRows = true;
  sheet.Protection.SelectUnlockedCells = true;

  AssertEx.true(sheet.Protection.IsProtected);
  AssertEx.true(sheet.Protection.Objects);
  AssertEx.true(sheet.Protection.FormatCells);
  AssertEx.true(sheet.Protection.InsertRows);
  AssertEx.true(sheet.Protection.SelectUnlockedCells);

  sheet.Unprotect();
  AssertEx.false(sheet.Protection.IsProtected);
  AssertEx.false(sheet.Protection.Objects);
  AssertEx.false(sheet.Protection.FormatCells);
  AssertEx.false(sheet.Protection.InsertRows);
  AssertEx.false(sheet.Protection.SelectUnlockedCells);

  sheet.Protection.AutoFilter = true;
  AssertEx.true(sheet.Protection.IsProtected);
  AssertEx.true(sheet.Protection.AutoFilter);
}

export function autoFilterApisMutateExpectedSettings(): void {
  const workbook = AutoFilterScenarioFactory.CreateAutoFilterWorkbook();
  AutoFilterScenarioFactory.AssertAutoFilter(workbook);

  const sheet = workbook.Worksheets[0];
  AssertEx.throws(() => {
    sheet.AutoFilter.FilterColumns.Add(-1);
  });
  AssertEx.throws(() => {
    sheet.AutoFilter.FilterColumns.Add(0);
  });
  AssertEx.throws(() => {
    sheet.AutoFilter.SortState.SortConditions.Add("1A");
  });

  sheet.AutoFilter.FilterColumns.RemoveAt(4);
  AssertEx.equal(4, sheet.AutoFilter.FilterColumns.Count);
  sheet.AutoFilter.Clear();
  AssertEx.equal("", sheet.AutoFilter.Range);
  AssertEx.equal(0, sheet.AutoFilter.FilterColumns.Count);
  AssertEx.equal(0, sheet.AutoFilter.SortState.SortConditions.Count);
}

export function definedNameApisMutateExpectedSettings(): void {
  const workbook = new Workbook();
  workbook.Worksheets[0].Name = "Data";
  workbook.Worksheets.Add("Scoped");

  const total = workbook.DefinedNames.get(workbook.DefinedNames.Add("Total", "=SUM(Data!$A$1:$A$2)"));
  total.Hidden = true;
  total.Comment = "Workbook scope";

  const scoped = workbook.DefinedNames.get(workbook.DefinedNames.Add("Input", "'Scoped'!$B$2", 1));
  scoped.Comment = "Local scope";

  AssertEx.equal(2, workbook.DefinedNames.Count);
  AssertEx.equal("Total", total.Name);
  AssertEx.equal("SUM(Data!$A$1:$A$2)", total.Formula);
  AssertEx.null(total.LocalSheetIndex);
  AssertEx.true(total.Hidden);
  AssertEx.equal("Workbook scope", total.Comment);

  AssertEx.equal("Input", scoped.Name);
  AssertEx.equal("'Scoped'!$B$2", scoped.Formula);
  AssertEx.equal(1, scoped.LocalSheetIndex ?? -1);
  AssertEx.equal("Local scope", scoped.Comment);

  AssertEx.throws(() => {
    workbook.DefinedNames.Add("Total", "1");
  });
  AssertEx.throws(() => {
    workbook.DefinedNames.Add("_xlnm.Print_Area", "A1");
  });
  AssertEx.throws(() => {
    workbook.DefinedNames.Add("Broken", "1", 5);
  });

  scoped.Name = "Total";
  AssertEx.throws(() => {
    scoped.LocalSheetIndex = null;
  });

  workbook.DefinedNames.RemoveAt(1);
  AssertEx.equal(1, workbook.DefinedNames.Count);
  AssertEx.throws(() => {
    workbook.DefinedNames.RemoveAt(5);
  });
}

export function hyperlinkApisMutateExpectedSettings(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];

  sheet.Cells["A1"].PutValue("Docs");
  const external = sheet.Hyperlinks[sheet.Hyperlinks.Add("A1", 1, 1, "https://example.com/docs")];
  external.TextToDisplay = "Docs";
  external.ScreenTip = "External docs";

  const internalLink = sheet.Hyperlinks[sheet.Hyperlinks.Add("B2", 1, 1, "Sheet1!C3")];
  internalLink.TextToDisplay = "Jump";

  const rangeLink = sheet.Hyperlinks[sheet.Hyperlinks.Add("C4", 2, 2, "mailto:test@example.com")];
  rangeLink.ScreenTip = "Send mail";

  const defaultTextLink = sheet.Hyperlinks[sheet.Hyperlinks.Add("E7", 1, 1, "https://default-text.example.com")];

  AssertEx.equal(4, sheet.Hyperlinks.Count);
  AssertEx.equal("A1", external.Area);
  AssertEx.equal("https://example.com/docs", external.Address);
  AssertEx.equal("External docs", external.ScreenTip);
  AssertEx.equal("Docs", external.TextToDisplay);

  AssertEx.equal("B2", internalLink.Area);
  AssertEx.equal("Sheet1!C3", internalLink.Address);
  AssertEx.equal("Jump", internalLink.TextToDisplay);

  AssertEx.equal("C4:D5", rangeLink.Area);
  AssertEx.equal("mailto:test@example.com", rangeLink.Address);
  AssertEx.equal("Send mail", rangeLink.ScreenTip);
  AssertEx.equal("https://default-text.example.com", defaultTextLink.TextToDisplay);

  AssertEx.throws(() => {
    sheet.Hyperlinks.Add("A1", 1, 1, "https://overlap.example.com");
  });
  AssertEx.throws(() => {
    sheet.Hyperlinks.Add("Z1", 0, 1, "https://invalid.example.com");
  });
  AssertEx.throws(() => {
    sheet.Hyperlinks.Add("A2", 1, 1, "");
  });
  AssertEx.throws(() => {
    sheet.Hyperlinks[-1];
  });

  sheet.Hyperlinks.RemoveAt(1);
  AssertEx.equal(3, sheet.Hyperlinks.Count);
  AssertEx.equal("C4:D5", sheet.Hyperlinks[1].Area);
  AssertEx.throws(() => {
    sheet.Hyperlinks.RemoveAt(99);
  });

  const replacementIndex = sheet.Hyperlinks.Add("F1", 1, 1, "https://replace.example.com");
  const replacement = sheet.Hyperlinks[replacementIndex];
  replacement.TextToDisplay = "Replacement";
  sheet.Hyperlinks.RemoveAt(replacementIndex);
  sheet.Hyperlinks[0] = replacement;
  AssertEx.equal("F1", sheet.Hyperlinks[0].Area);
  AssertEx.equal("https://replace.example.com", sheet.Hyperlinks[0].Address);
  AssertEx.equal("Replacement", sheet.Hyperlinks[0].TextToDisplay);
  AssertEx.throws(() => {
    sheet.Hyperlinks[0] = null;
  });

  const currentCapacity = sheet.Hyperlinks.Capacity;
  sheet.Hyperlinks.Capacity = currentCapacity + 4;
  AssertEx.true(sheet.Hyperlinks.Capacity >= currentCapacity + 4);
  AssertEx.throws(() => {
    sheet.Hyperlinks.Capacity = 1;
  });

  sheet.Hyperlinks.Clear();
  AssertEx.equal(0, sheet.Hyperlinks.Count);
}

export function validationApisMutateExpectedSettings(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];

  const primary = sheet.Validations.get(sheet.Validations.Add(CellArea.CreateCellArea("A1", "B2")));
  primary.Type = ValidationType.WholeNumber;
  primary.Operator = OperatorType.Between;
  primary.Formula1 = "=1";
  primary.Formula2 = "=10";
  primary.ShowError = true;
  primary.ErrorTitle = "Whole Number";
  primary.ErrorMessage = "Enter 1-10";
  primary.AddArea(CellArea.CreateCellArea("D4", "D5"));

  AssertEx.equal(1, sheet.Validations.Count);
  AssertEx.equal(2, primary.Areas.length);
  AssertEx.equal("1", primary.Formula1);
  AssertEx.equal("10", primary.Formula2);
  AssertEx.equal(ValidationType.WholeNumber, sheet.Validations.GetValidationInCell(0, 0)!.Type);
  AssertEx.equal(ValidationType.WholeNumber, sheet.Validations.GetValidationInCell(4, 3)!.Type);
  AssertEx.throws(() => {
    sheet.Validations.Add(CellArea.CreateCellArea("B2", "C3"));
  });

  const secondIndex = sheet.Validations.Add(CellArea.CreateCellArea("F1", "F1"));
  const second = sheet.Validations.get(secondIndex);
  second.Type = ValidationType.List;
  second.Formula1 = "\"Y,N\"";
  AssertEx.equal(2, sheet.Validations.Count);

  sheet.Validations.RemoveACell(0, 0);
  AssertEx.null(sheet.Validations.GetValidationInCell(0, 0));
  AssertEx.notNull(sheet.Validations.GetValidationInCell(0, 1));

  sheet.Validations.RemoveArea(CellArea.CreateCellArea("F1", "F1"));
  AssertEx.equal(1, sheet.Validations.Count);
  AssertEx.throws(() => {
    sheet.Validations.RemoveACell(-1, 0);
  });
  AssertEx.throws(() => {
    sheet.Validations.get(-1);
  });
}

export function conditionalFormattingApisMutateExpectedSettings(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];

  const index = sheet.ConditionalFormattings.Add();
  const collection = sheet.ConditionalFormattings.get(index);
  collection.AddArea(CellArea.CreateCellArea("A1", "A3"));
  const conditionIndex = collection.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "=1", "=9");
  const condition = collection.get(conditionIndex);
  condition.StopIfTrue = true;
  condition.Priority = 1;
  const style = condition.Style;
  style.Pattern = FillPattern.Solid;
  style.ForegroundColor = Color.FromArgb(255, 255, 0, 0);
  condition.Style = style;

  AssertEx.equal(1, sheet.ConditionalFormattings.Count);
  AssertEx.equal(1, collection.RangeCount);
  AssertEx.equal(1, collection.Count);
  AssertEx.equal("1", condition.Formula1);
  AssertEx.equal("9", condition.Formula2);
  AssertEx.true(condition.StopIfTrue);
  AssertEx.equal(FillPattern.Solid, condition.Style.Pattern);

  collection.AddCondition(FormatConditionType.Expression);
  AssertEx.equal(2, collection.Count);
  collection.RemoveCondition(1);
  AssertEx.equal(1, collection.Count);

  collection.AddArea(CellArea.CreateCellArea("C1", "C2"));
  AssertEx.equal(2, collection.RangeCount);
  collection.RemoveAreaByBounds(0, 0, 1, 1);
  AssertEx.equal(2, collection.RangeCount);
  AssertEx.equal(0, collection.GetCellArea(0).StartRow);
  AssertEx.equal(2, collection.GetCellArea(0).StartColumn);

  sheet.ConditionalFormattings.RemoveArea(0, 2, 2, 1);
  AssertEx.equal(1, sheet.ConditionalFormattings.Count);
  sheet.ConditionalFormattings.RemoveAt(0);
  AssertEx.equal(0, sheet.ConditionalFormattings.Count);
  AssertEx.throws(() => {
    sheet.ConditionalFormattings.get(-1);
  });
}

export function conditionalFormattingAdvancedApisMutateExpectedSettings(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];

  const contains = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
  contains.AddArea(CellArea.CreateCellArea("A1", "A10"));
  const containsRule = contains.get(contains.AddCondition(FormatConditionType.ContainsText));
  containsRule.Formula1 = "error";
  containsRule.Priority = 2;

  const timePeriod = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
  timePeriod.AddArea(CellArea.CreateCellArea("B1", "B10"));
  const timeRule = timePeriod.get(timePeriod.AddCondition(FormatConditionType.TimePeriod));
  timeRule.TimePeriod = "today";

  const top10 = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
  top10.AddArea(CellArea.CreateCellArea("C1", "C10"));
  const topRule = top10.get(top10.AddCondition(FormatConditionType.Top10));
  topRule.Percent = true;
  topRule.Rank = 10;

  const colorScale = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
  colorScale.AddArea(CellArea.CreateCellArea("D1", "D10"));
  const colorRule = colorScale.get(colorScale.AddCondition(FormatConditionType.ColorScale));
  colorRule.ColorScaleCount = 3;
  colorRule.MinColor = Color.FromArgb(255, 248, 105, 107);
  colorRule.MidColor = Color.FromArgb(255, 255, 235, 132);
  colorRule.MaxColor = Color.FromArgb(255, 99, 190, 123);

  const dataBar = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
  dataBar.AddArea(CellArea.CreateCellArea("E1", "E10"));
  const dataBarRule = dataBar.get(dataBar.AddCondition(FormatConditionType.DataBar));
  dataBarRule.BarColor = Color.FromArgb(255, 99, 142, 198);
  dataBarRule.ShowBorder = true;
  dataBarRule.Direction = "left-to-right";

  const iconSet = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
  iconSet.AddArea(CellArea.CreateCellArea("F1", "F10"));
  const iconSetRule = iconSet.get(iconSet.AddCondition(FormatConditionType.IconSet));
  iconSetRule.IconSetType = "4Arrows";
  iconSetRule.ReverseIcons = true;
  iconSetRule.ShowIconOnly = true;

  AssertEx.equal(6, sheet.ConditionalFormattings.Count);
  AssertEx.equal(FormatConditionType.ContainsText, containsRule.Type as FormatConditionType);
  AssertEx.equal("error", containsRule.Formula1);
  AssertEx.equal(2, containsRule.Priority);
  AssertEx.equal("today", timeRule.TimePeriod);
  AssertEx.true(topRule.Percent);
  AssertEx.equal(10, topRule.Rank);
  AssertEx.equal(3, colorRule.ColorScaleCount);
  AssertEx.equal(Color.FromArgb(255, 248, 105, 107), colorRule.MinColor);
  AssertEx.equal(Color.FromArgb(255, 99, 142, 198), dataBarRule.BarColor);
  AssertEx.true(dataBarRule.ShowBorder);
  AssertEx.equal("4Arrows", iconSetRule.IconSetType);
  AssertEx.true(iconSetRule.ReverseIcons);
  AssertEx.true(iconSetRule.ShowIconOnly);
}

export function pageSetupApisMutateExpectedSettings(): void {
  const workbook = new Workbook();
  const pageSetup = workbook.Worksheets[0].PageSetup;

  pageSetup.LeftMargin = 0.508;
  pageSetup.RightMargin = 0.635;
  pageSetup.Orientation = PageOrientationType.Landscape;
  pageSetup.PaperSize = PaperSizeType.PaperA4;
  pageSetup.FirstPageNumber = 2;
  pageSetup.Scale = 90;
  pageSetup.FitToPagesWide = 1;
  pageSetup.FitToPagesTall = 3;
  pageSetup.PrintArea = "$A$1:$D$20";
  pageSetup.PrintTitleRows = "$1:$2";
  pageSetup.PrintTitleColumns = "$A:$B";
  pageSetup.LeftHeader = "LH";
  pageSetup.CenterFooter = "CF";
  pageSetup.PrintGridlines = true;
  pageSetup.CenterHorizontally = true;
  pageSetup.AddHorizontalPageBreak(5);
  pageSetup.AddVerticalPageBreak(2);

  AssertEx.equal(0.508, pageSetup.LeftMargin);
  AssertEx.equal(0.635, pageSetup.RightMargin);
  AssertEx.equal(PageOrientationType.Landscape, pageSetup.Orientation);
  AssertEx.equal(PaperSizeType.PaperA4, pageSetup.PaperSize);
  AssertEx.equal(2, pageSetup.FirstPageNumber ?? 0);
  AssertEx.equal(90, pageSetup.Scale ?? 0);
  AssertEx.equal(1, pageSetup.FitToPagesWide ?? 0);
  AssertEx.equal(3, pageSetup.FitToPagesTall ?? 0);
  AssertEx.equal("$A$1:$D$20", pageSetup.PrintArea);
  AssertEx.equal("$1:$2", pageSetup.PrintTitleRows);
  AssertEx.equal("$A:$B", pageSetup.PrintTitleColumns);
  AssertEx.equal("LH", pageSetup.LeftHeader);
  AssertEx.equal("CF", pageSetup.CenterFooter);
  AssertEx.true(pageSetup.PrintGridlines);
  AssertEx.true(pageSetup.CenterHorizontally);
  AssertEx.equal(1, pageSetup.HorizontalPageBreaks.length);
  AssertEx.equal(5, pageSetup.HorizontalPageBreaks[0]);
  AssertEx.equal(1, pageSetup.VerticalPageBreaks.length);
  AssertEx.equal(2, pageSetup.VerticalPageBreaks[0]);
  AssertEx.throws(() => {
    pageSetup.Scale = 5;
  });
  AssertEx.throws(() => {
    pageSetup.LeftMargin = -0.1;
  });
}

export function workbookMetadataApisMutateExpectedSettings(): void {
  const workbook = WorkbookMetadataScenarioFactory.CreateWorkbookMetadataWorkbook();
  WorkbookMetadataScenarioFactory.AssertWorkbookMetadata(workbook);
}

export function listObjectApisMutateExpectedSettings(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];

  sheet.Cells["A1"].PutValue("Name");
  sheet.Cells["B1"].PutValue("Score");
  sheet.Cells["A2"].PutValue("Alice");
  sheet.Cells["B2"].PutValue(95);
  sheet.Cells["A3"].PutValue("Bob");
  sheet.Cells["B3"].PutValue(87);

  const index = sheet.ListObjects.Add("A1", "B3", true);
  AssertEx.equal(0, index);
  AssertEx.equal(1, sheet.ListObjects.Count);

  const table = sheet.ListObjects[index];
  AssertEx.equal("Table1", table.DisplayName);
  AssertEx.equal(0, table.StartRow);
  AssertEx.equal(0, table.StartColumn);
  AssertEx.equal(2, table.EndRow);
  AssertEx.equal(1, table.EndColumn);
  AssertEx.true(table.ShowHeaderRow);
  AssertEx.false(table.ShowTotals);
  AssertEx.equal(2, table.ListColumns.Count);
  AssertEx.equal("Name", table.ListColumns.get(0).Name);
  AssertEx.equal("Score", table.ListColumns.get(1).Name);

  table.TableStyleType = TableStyleType.TableStyleMedium9;
  AssertEx.equal(TableStyleType.TableStyleMedium9, table.TableStyleType);
  AssertEx.equal("TableStyleMedium9", table.TableStyleName);

  table.TableStyleName = "MyCustom";
  AssertEx.equal("MyCustom", table.TableStyleName);
  AssertEx.equal(TableStyleType.Custom, table.TableStyleType);

  table.TableStyleType = TableStyleType.None;
  AssertEx.equal("", table.TableStyleName);

  table.ShowTableStyleRowStripes = true;
  table.ShowTableStyleFirstColumn = true;
  table.ShowTableStyleLastColumn = false;
  table.ShowTableStyleColumnStripes = false;
  AssertEx.true(table.ShowTableStyleRowStripes);
  AssertEx.true(table.ShowTableStyleFirstColumn);
  AssertEx.false(table.ShowTableStyleLastColumn);
  AssertEx.false(table.ShowTableStyleColumnStripes);

  table.ShowTotals = true;
  const scoreCol = table.ListColumns.get(1);
  scoreCol.TotalsCalculation = TotalsCalculation.Sum;
  scoreCol.TotalsRowLabel = "";
  AssertEx.equal(TotalsCalculation.Sum, table.ListColumns.get(1).TotalsCalculation);

  const nameCol = table.ListColumns.get(0);
  nameCol.TotalsCalculation = TotalsCalculation.None;
  nameCol.TotalsRowLabel = "Total";
  AssertEx.equal(TotalsCalculation.None, table.ListColumns.get(0).TotalsCalculation);
  AssertEx.equal("Total", table.ListColumns.get(0).TotalsRowLabel);

  scoreCol.TotalsCalculation = TotalsCalculation.Average;
  AssertEx.equal(TotalsCalculation.Average, table.ListColumns.get(1).TotalsCalculation);

  table.DisplayName = "SalesTable";
  AssertEx.equal("SalesTable", table.DisplayName);
  AssertEx.equal("SalesTable", sheet.ListObjects.getByName("SalesTable").DisplayName);

  table.Comment = "Test comment";
  AssertEx.equal("Test comment", table.Comment);

  table.ShowAutoFilter();
  table.RemoveAutoFilter();

  table.Resize(0, 0, 3, 1, true);
  AssertEx.equal(3, table.EndRow);

  AssertEx.throws(() => {
    table.DisplayName = "";
  });
  AssertEx.throws(() => {
    sheet.ListObjects.Add(0, 0, 2, 1, true);
  });
  AssertEx.throws(() => {
    sheet.ListObjects[5];
  });
  AssertEx.throws(() => {
    sheet.ListObjects.RemoveAt(5);
  });
  AssertEx.throws(() => {
    sheet.ListObjects["NoSuchTable"];
  });

  sheet.Cells["D1"].PutValue("Category");
  sheet.Cells["E1"].PutValue("Units");
  sheet.Cells["D2"].PutValue("A");
  sheet.Cells["E2"].PutValue(10);
  const index2 = sheet.ListObjects.Add(0, 3, 2, 4, true);
  const table2 = sheet.ListObjects[index2];
  AssertEx.equal(2, sheet.ListObjects.Count);
  AssertEx.equal("Table2", table2.DisplayName);
  AssertEx.equal("Category", table2.ListColumns.get(0).Name);
  AssertEx.equal("Units", table2.ListColumns.get(1).Name);

  table2.ConvertToRange();
  AssertEx.equal(1, sheet.ListObjects.Count);

  sheet.ListObjects.RemoveAt(0);
  AssertEx.equal(0, sheet.ListObjects.Count);
}

export function pictureApisMutateExpectedSettings(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];

  const jpegBytes = PictureScenarioFactory.LoadJpeg();
  const pngBytes = PictureScenarioFactory.LoadPng();

  AssertEx.equal(0, sheet.Pictures.Count);

  const idx = sheet.Pictures.Add(1, 2, 5, 6, jpegBytes);
  AssertEx.equal(0, idx);
  AssertEx.equal(1, sheet.Pictures.Count);

  const pic = sheet.Pictures.get(0);
  AssertEx.equal("Picture 1", pic.Name);
  AssertEx.equal(1, pic.UpperLeftRow);
  AssertEx.equal(2, pic.UpperLeftColumn);
  AssertEx.equal(5, pic.LowerRightRow);
  AssertEx.equal(6, pic.LowerRightColumn);
  AssertEx.equal(ImageType.Jpeg, pic.ImageType);
  AssertEx.equal(jpegBytes.length, pic.Data.length);

  pic.Name = "Logo";
  AssertEx.equal("Logo", sheet.Pictures.get(0).Name);

  pic.UpperLeftRow = 0;
  pic.UpperLeftColumn = 0;
  pic.LowerRightRow = 3;
  pic.LowerRightColumn = 3;
  AssertEx.equal(0, sheet.Pictures.get(0).UpperLeftRow);
  AssertEx.equal(3, sheet.Pictures.get(0).LowerRightRow);

  const idx2 = sheet.Pictures.Add(5, 0, 8, 4, pngBytes);
  AssertEx.equal(1, idx2);
  AssertEx.equal(2, sheet.Pictures.Count);
  AssertEx.equal(ImageType.Png, sheet.Pictures.get(1).ImageType);

  sheet.Pictures.RemoveAt(0);
  AssertEx.equal(1, sheet.Pictures.Count);
  AssertEx.equal(ImageType.Png, sheet.Pictures.get(0).ImageType);
}
