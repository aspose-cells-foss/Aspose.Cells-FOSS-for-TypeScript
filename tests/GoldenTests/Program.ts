import * as path from "path";
import {
  CellValueType,
  ChartType,
  SaveFormat,
  SaveOptions,
  Workbook,
} from "../../aspose_cells";
import {
  elementChildren,
  firstElement,
  localName,
  parseXml,
  XmlNode,
} from "../../aspose_cells/Xml/XmlReader";
import {
  AssertEx,
  getOrCreateOutputDirectory,
  MemoryStream,
  resolveRepositoryRoot,
  TemporaryDirectory,
  TestCase,
  TestRunner,
  ZipPackageHelper,
} from "../Shared/TestInfrastructure";
import {
  ConditionalFormattingScenarioFactory,
  DefinedNameScenarioFactory,
  HyperlinkScenarioFactory,
  PageSetupScenarioFactory,
  StyleScenarioFactory,
  ValidationScenarioFactory,
  WorkbookScenarioFactory,
  WorksheetScenarioFactory,
} from "../Shared/ScenarioFactories";
import { AutoFilterScenarioFactory } from "../Shared/AutoFilterScenarioFactory";
import { ChartScenarioFactory } from "../Shared/ChartScenarioFactory";
import { WorkbookMetadataScenarioFactory } from "../Shared/WorkbookMetadataScenarioFactory";

const CORE_PROPERTIES_RELATIONSHIP =
  "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const EXTENDED_PROPERTIES_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";

function xlsxRoundtripMixedScalarCellsFile(): void {
  const temp = new TemporaryDirectory("golden-file");
  const filePath = temp.getPath("mixed.xlsx");
  const workbook = WorkbookScenarioFactory.CreateMixedCellWorkbook();
  workbook.Save(filePath);

  const loaded = new Workbook(filePath);
  assertMixedWorkbook(loaded, false);
}

function xlsxRoundtripMixedScalarCellsStream(): void {
  const workbook = WorkbookScenarioFactory.CreateMixedCellWorkbook();
  const stream = new MemoryStream();
  workbook.Save(stream, SaveFormat.Xlsx);
  stream.position = 0;

  const loaded = new Workbook(stream);
  assertMixedWorkbook(loaded, false);
}

function sharedStringsExportUsesSstWhenEnabled(): void {
  const temp = new TemporaryDirectory("golden-sst");
  const filePath = temp.getPath("shared-strings.xlsx");
  const workbook = new Workbook();
  workbook.Worksheets[0].Cells["A1"].PutValue("Hello");
  workbook.Worksheets[0].Cells["A2"].PutValue("Hello");
  const options = new SaveOptions();
  options.UseSharedStrings = true;
  workbook.Save(filePath, options);

  AssertEx.true(ZipPackageHelper.entryExists(filePath, "xl/sharedStrings.xml"));
  const sharedStringsXml = ZipPackageHelper.readEntryText(filePath, "xl/sharedStrings.xml");
  const worksheetXml = ZipPackageHelper.readEntryText(filePath, "xl/worksheets/sheet1.xml");

  AssertEx.contains("<sst", sharedStringsXml);
  AssertEx.contains('uniqueCount="1"', sharedStringsXml);
  AssertEx.contains('t="s"', worksheetXml);
}

function inlineStringsExportWhenSharedStringsDisabled(): void {
  const temp = new TemporaryDirectory("golden-inline");
  const filePath = temp.getPath("inline.xlsx");
  const workbook = new Workbook();
  workbook.Worksheets[0].Cells["A1"].PutValue("Inline");
  const options = new SaveOptions();
  options.UseSharedStrings = false;
  workbook.Save(filePath, options);

  AssertEx.false(ZipPackageHelper.entryExists(filePath, "xl/sharedStrings.xml"));
  const worksheetXml = ZipPackageHelper.readEntryText(filePath, "xl/worksheets/sheet1.xml");

  AssertEx.contains('t="inlineStr"', worksheetXml);
  AssertEx.contains("<t>Inline</t>", worksheetXml);
}

function formulaCellsPersistFormulaAndCachedValue(): void {
  const temp = new TemporaryDirectory("golden-formula");
  const filePath = temp.getPath("formula.xlsx");
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  sheet.Cells["A1"].PutValue(10);
  sheet.Cells["B1"].PutValue(20);
  sheet.Cells["C1"].PutValue(30);
  sheet.Cells["C1"].Formula = "=A1+B1";
  workbook.Save(filePath);

  const worksheetXml = ZipPackageHelper.readEntryText(filePath, "xl/worksheets/sheet1.xml");
  AssertEx.contains("<f>A1+B1</f>", worksheetXml);
  AssertEx.contains("<v>30</v>", worksheetXml);

  const loaded = new Workbook(filePath);
  AssertEx.equal("=A1+B1", loaded.Worksheets[0].Cells["C1"].Formula);
  AssertEx.equal("30", loaded.Worksheets[0].Cells["C1"].StringValue);
}

function mac1904DateTimeRoundtripAndMarkup(): void {
  const temp = new TemporaryDirectory("golden-1904");
  const filePath = temp.getPath("date1904.xlsx");
  const workbook = WorkbookScenarioFactory.CreateMixedCellWorkbook(true);
  workbook.Save(filePath);

  const workbookXml = ZipPackageHelper.readEntryText(filePath, "xl/workbook.xml");
  AssertEx.contains('date1904="1"', workbookXml);

  const loaded = new Workbook(filePath);
  AssertEx.true(loaded.Settings.Date1904);
  const expected = new Date(Date.UTC(2024, 4, 6, 7, 8, 9));
  const actual = loaded.Worksheets[0].Cells["F1"].Value as Date;
  AssertEx.equal(expected, actual);
}

function displayRuntimeCultureAfterRoundtrip(): void {
  const temp = new TemporaryDirectory("golden-display-culture");
  const filePath = temp.getPath("display-culture.xlsx");
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  const timestamp = new Date(2024, 4, 6, 7, 8, 9);

  sheet.Cells["A1"].PutValue(1234.5);
  const numericStyle = sheet.Cells["A1"].GetStyle();
  numericStyle.NumberFormat = "#,##0.00";
  sheet.Cells["A1"].SetStyle(numericStyle);

  sheet.Cells["B1"].PutValue(timestamp);
  const dateStyle = sheet.Cells["B1"].GetStyle();
  dateStyle.NumberFormat = "[$-409]dddd, mmmm d, yyyy";
  sheet.Cells["B1"].SetStyle(dateStyle);

  workbook.Save(filePath);

  const loaded = new Workbook(filePath);
  loaded.Settings.Culture = "fr-FR";

  AssertEx.equal(frenchGroupedNumber(1234.5), loaded.Worksheets[0].Cells["A1"].DisplayStringValue);
  AssertEx.equal(weekdayMonthDayYear(timestamp, "en-US"), loaded.Worksheets[0].Cells["B1"].DisplayStringValue);
}

function styledCellsRoundtripAndEmitStylesheet(): void {
  const temp = new TemporaryDirectory("golden-styles");
  const filePath = temp.getPath("styles.xlsx");
  const workbook = StyleScenarioFactory.CreateStyledWorkbook();
  workbook.Save(filePath);

  AssertEx.true(ZipPackageHelper.entryExists(filePath, "xl/styles.xml"));
  const stylesheetXml = ZipPackageHelper.readEntryText(filePath, "xl/styles.xml");
  const worksheetXml = ZipPackageHelper.readEntryText(filePath, "xl/worksheets/sheet1.xml");

  AssertEx.contains("Arial", stylesheetXml);
  AssertEx.contains("0.0000", stylesheetXml);
  AssertEx.contains("<strike", stylesheetXml);
  AssertEx.contains('patternType="lightGrid"', stylesheetXml);
  AssertEx.contains('fgColor rgb="FFD2DC1E"', stylesheetXml);
  AssertEx.contains('bgColor rgb="FF0C2D4E"', stylesheetXml);
  AssertEx.contains('style="dotted"', stylesheetXml);
  AssertEx.contains('style="mediumDashDot"', stylesheetXml);
  AssertEx.contains('style="double"', stylesheetXml);
  AssertEx.contains('style="dashDotDot"', stylesheetXml);
  AssertEx.contains('style="slantDashDot"', stylesheetXml);
  AssertEx.contains('diagonalUp="1"', stylesheetXml);
  AssertEx.contains('diagonalDown="1"', stylesheetXml);
  AssertEx.contains('horizontal="distributed"', stylesheetXml);
  AssertEx.contains('vertical="distributed"', stylesheetXml);
  AssertEx.contains('indent="2"', stylesheetXml);
  AssertEx.contains('textRotation="45"', stylesheetXml);
  AssertEx.contains('shrinkToFit="1"', stylesheetXml);
  AssertEx.contains('readingOrder="2"', stylesheetXml);
  AssertEx.contains('relativeIndent="1"', stylesheetXml);
  AssertEx.contains('wrapText="1"', stylesheetXml);
  AssertEx.contains('locked="0"', stylesheetXml);
  AssertEx.contains('hidden="1"', stylesheetXml);
  AssertEx.contains('numFmtId="4"', stylesheetXml);
  assertFontElementOrder(stylesheetXml);
  AssertEx.contains('s="', worksheetXml);

  const loaded = new Workbook(filePath);
  StyleScenarioFactory.AssertPrimaryStyle(loaded.Worksheets[0].Cells["A1"].GetStyle());
  StyleScenarioFactory.AssertCustomNumberStyle(loaded.Worksheets[0].Cells["B2"].GetStyle());
  AssertEx.equal(CellValueType.IsNull, loaded.Worksheets[0].Cells["B2"].Type);
  AssertEx.equal(CellValueType.IsDateTime, loaded.Worksheets[0].Cells["C3"].Type);
}

function assertFontElementOrder(stylesheetXml: string): void {
  const document = parseXml(stylesheetXml);
  const root = rootElement(document);
  const order = new Map<string, number>([
    ["b", 0],
    ["i", 1],
    ["strike", 2],
    ["u", 3],
    ["condense", 4],
    ["extend", 5],
    ["outline", 6],
    ["shadow", 7],
    ["charset", 8],
    ["family", 9],
    ["sz", 10],
    ["color", 11],
    ["vertAlign", 12],
    ["scheme", 13],
    ["name", 14],
  ]);

  const fontsElement = firstElement(root, "fonts");
  for (const font of elementChildren(fontsElement, "font")) {
    let previousOrder = -1;
    for (let index = 0; index < font.childNodes!.length; index++) {
      const child = font.childNodes![index];
      const childName = localName(child.nodeName);
      if (childName.startsWith("#")) {
        continue;
      }

      const currentOrder = order.get(childName);
      AssertEx.true(currentOrder !== undefined, "Unexpected font child element: " + childName);
      if (currentOrder === undefined) {
        continue;
      }

      AssertEx.true(currentOrder >= previousOrder, "Font child element order is invalid.");
      previousOrder = currentOrder;
    }
  }
}

function worksheetSettingsRoundtripAndEmitExpectedMarkup(): void {
  const temp = new TemporaryDirectory("golden-worksheet-settings");
  const filePath = temp.getPath("worksheet-settings.xlsx");
  const workbook = WorksheetScenarioFactory.CreateWorksheetSettingsWorkbook();
  workbook.Save(filePath);

  const worksheetXml = ZipPackageHelper.readEntryText(filePath, "xl/worksheets/sheet1.xml");
  const workbookXml = ZipPackageHelper.readEntryText(filePath, "xl/workbook.xml");
  AssertEx.contains('dimension ref="A1:C4"', worksheetXml);
  AssertEx.contains('mergeCell ref="A1:B2"', worksheetXml);
  AssertEx.contains('hidden="1"', worksheetXml);
  AssertEx.contains('state="hidden"', workbookXml);

  const loaded = new Workbook(filePath);
  WorksheetScenarioFactory.AssertWorksheetSettings(loaded);
  WorksheetScenarioFactory.AssertWorksheetSettingsScenarioHasVisibleSheet(loaded);
}

function worksheetViewAndTabColorRoundtripAndEmitExpectedMarkup(): void {
  const temp = new TemporaryDirectory("golden-worksheet-view");
  const filePath = temp.getPath("worksheet-view.xlsx");
  const workbook = WorksheetScenarioFactory.CreateWorksheetSettingsWorkbook();
  workbook.Save(filePath);

  const worksheetXml = ZipPackageHelper.readEntryText(filePath, "xl/worksheets/sheet1.xml");
  AssertEx.contains("<sheetViews>", worksheetXml);
  AssertEx.contains('showGridLines="0"', worksheetXml);
  AssertEx.contains('showRowColHeaders="0"', worksheetXml);
  AssertEx.contains('showZeros="0"', worksheetXml);
  AssertEx.contains('rightToLeft="1"', worksheetXml);
  AssertEx.contains('zoomScale="85"', worksheetXml);
  AssertEx.contains('tabColor rgb="FF224466"', worksheetXml);

  const loaded = new Workbook(filePath);
  WorksheetScenarioFactory.AssertWorksheetSettings(loaded);
  WorksheetScenarioFactory.AssertWorksheetSettingsScenarioHasVisibleSheet(loaded);
}

function worksheetProtectionRoundtripAndEmitExpectedMarkup(): void {
  const temp = new TemporaryDirectory("golden-worksheet-protection");
  const filePath = temp.getPath("worksheet-protection.xlsx");
  const workbook = WorksheetScenarioFactory.CreateWorksheetSettingsWorkbook();
  workbook.Save(filePath);

  const worksheetXml = ZipPackageHelper.readEntryText(filePath, "xl/worksheets/sheet1.xml");
  AssertEx.contains("<sheetProtection", worksheetXml);
  AssertEx.contains('sheet="1"', worksheetXml);
  AssertEx.contains('objects="1"', worksheetXml);
  AssertEx.contains('scenarios="1"', worksheetXml);
  AssertEx.contains('formatCells="1"', worksheetXml);
  AssertEx.contains('insertRows="1"', worksheetXml);
  AssertEx.contains('autoFilter="1"', worksheetXml);
  AssertEx.contains('selectLockedCells="1"', worksheetXml);
  AssertEx.contains('selectUnlockedCells="1"', worksheetXml);

  const loaded = new Workbook(filePath);
  WorksheetScenarioFactory.AssertWorksheetSettings(loaded);
  WorksheetScenarioFactory.AssertWorksheetSettingsScenarioHasVisibleSheet(loaded);
}

function autoFilterRoundtripAndEmitExpectedMarkup(): void {
  const temp = new TemporaryDirectory("golden-autofilter");
  const filePath = temp.getPath("autofilter.xlsx");
  const workbook = AutoFilterScenarioFactory.CreateAutoFilterWorkbook();
  workbook.Save(filePath);

  const worksheetXml = ZipPackageHelper.readEntryText(filePath, "xl/worksheets/sheet1.xml");
  const workbookXml = ZipPackageHelper.readEntryText(filePath, "xl/workbook.xml");
  AssertEx.contains('<autoFilter ref="A1:E6"', worksheetXml);
  AssertEx.contains('<filterColumn colId="0" hiddenButton="1">', worksheetXml);
  AssertEx.contains('<filters><filter val="Open" /><filter val="Closed" /></filters>', worksheetXml);
  AssertEx.contains('<customFilters and="1">', worksheetXml);
  AssertEx.contains('operator="greaterThanOrEqual"', worksheetXml);
  AssertEx.contains('operator="lessThanOrEqual"', worksheetXml);
  AssertEx.contains('<colorFilter dxfId="3" cellColor="1" />', worksheetXml);
  AssertEx.contains('<dynamicFilter type="thisMonth" val="1" maxVal="31" />', worksheetXml);
  AssertEx.contains('<top10 top="0" percent="1" val="10" filterVal="2.5" />', worksheetXml);
  AssertEx.contains('<sortState ref="A2:E6" caseSensitive="1" sortMethod="pinYin">', worksheetXml);
  AssertEx.contains(
    '<sortCondition ref="B2:B6" descending="1" sortBy="value" customList="High,Medium,Low" />',
    worksheetXml,
  );
  AssertEx.contains('<sortCondition ref="C2:C6" sortBy="cellColor" dxfId="4" />', worksheetXml);
  AssertEx.contains(
    '<sortCondition ref="E2:E6" sortBy="icon" iconSet="3TrafficLights1" iconId="2" />',
    worksheetXml,
  );
  AssertEx.contains('name="_xlnm._FilterDatabase"', workbookXml);
  AssertEx.contains("'Filtered'!$A$1:$E$6", workbookXml);

  const loaded = new Workbook(filePath);
  AutoFilterScenarioFactory.AssertAutoFilter(loaded);
}

function autoFilterOmitsInvalidDxfReferences(): void {
  const temp = new TemporaryDirectory("golden-autofilter-invalid-dxf");
  const filePath = temp.getPath("autofilter-invalid-dxf.xlsx");
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  sheet.Name = "InvalidDxf";
  sheet.Cells["A1"].PutValue("Status");
  sheet.Cells["A2"].PutValue("Open");
  sheet.Cells["A3"].PutValue("Closed");
  sheet.Cells["B1"].PutValue("Amount");
  sheet.Cells["B2"].PutValue(10);
  sheet.Cells["B3"].PutValue(20);
  sheet.AutoFilter.Range = "A1:B3";

  const column = sheet.AutoFilter.FilterColumns.get(sheet.AutoFilter.FilterColumns.Add(0));
  column.ColorFilter.Enabled = true;
  column.ColorFilter.DifferentialStyleId = 9;
  column.ColorFilter.CellColor = true;

  sheet.AutoFilter.SortState.Ref = "A2:B3";
  const sortCondition = sheet.AutoFilter.SortState.SortConditions.get(
    sheet.AutoFilter.SortState.SortConditions.Add("B2:B3"),
  );
  sortCondition.SortBy = "value";
  sortCondition.DifferentialStyleId = 8;

  workbook.Save(filePath);

  const worksheetXml = ZipPackageHelper.readEntryText(filePath, "xl/worksheets/sheet1.xml");
  AssertEx.contains('<autoFilter ref="A1:B3"', worksheetXml);
  AssertEx.false(worksheetXml.includes("<colorFilter"));
  AssertEx.false(worksheetXml.includes('dxfId="8"'));
  AssertEx.false(worksheetXml.includes('dxfId="9"'));

  const loaded = new Workbook(filePath);
  AssertEx.equal("A1:B3", loaded.Worksheets[0].AutoFilter.Range);
  AssertEx.equal(0, loaded.Worksheets[0].AutoFilter.FilterColumns.Count);
  AssertEx.equal(1, loaded.Worksheets[0].AutoFilter.SortState.SortConditions.Count);
  AssertEx.null(loaded.Worksheets[0].AutoFilter.SortState.SortConditions.get(0).DifferentialStyleId);
}

function definedNamesRoundtripAndEmitExpectedMarkup(): void {
  const temp = new TemporaryDirectory("golden-defined-names");
  const filePath = temp.getPath("defined-names.xlsx");
  const workbook = DefinedNameScenarioFactory.CreateDefinedNamesWorkbook();
  workbook.Save(filePath);

  const workbookXml = ZipPackageHelper.readEntryText(filePath, "xl/workbook.xml");
  AssertEx.contains("<definedNames>", workbookXml);
  AssertEx.contains('name="GlobalRange"', workbookXml);
  AssertEx.contains('hidden="1"', workbookXml);
  AssertEx.contains('comment="Primary range"', workbookXml);
  AssertEx.contains('name="LocalCell"', workbookXml);
  AssertEx.contains('localSheetId="1"', workbookXml);
  AssertEx.contains('name="_xlnm.Print_Area"', workbookXml);
  AssertEx.contains('name="_xlnm.Print_Titles"', workbookXml);

  const loaded = new Workbook(filePath);
  DefinedNameScenarioFactory.AssertDefinedNames(loaded);
}

function worksheetDimensionIncludesMergeOnlyRanges(): void {
  const temp = new TemporaryDirectory("golden-merge-dimension");
  const filePath = temp.getPath("merge-only.xlsx");
  const workbook = new Workbook();
  workbook.Worksheets[0].Name = "MergeOnly";
  workbook.Worksheets[0].Cells.Merge(4, 5, 2, 2);
  workbook.Save(filePath);

  const worksheetXml = ZipPackageHelper.readEntryText(filePath, "xl/worksheets/sheet1.xml");
  AssertEx.contains('dimension ref="F5:G6"', worksheetXml);
  AssertEx.contains('mergeCell ref="F5:G6"', worksheetXml);

  const loaded = new Workbook(filePath);
  AssertEx.equal(1, loaded.Worksheets[0].Cells.MergedCells.length);
  AssertEx.equal(4, loaded.Worksheets[0].Cells.MergedCells[0].StartRow);
  AssertEx.equal(5, loaded.Worksheets[0].Cells.MergedCells[0].StartColumn);
}

function hyperlinksRoundtripAndEmitExpectedMarkup(): void {
  const temp = new TemporaryDirectory("golden-hyperlinks");
  const filePath = temp.getPath("hyperlinks.xlsx");
  const workbook = HyperlinkScenarioFactory.CreateHyperlinkWorkbook();
  workbook.Save(filePath);

  const worksheetXml = ZipPackageHelper.readEntryText(filePath, "xl/worksheets/sheet1.xml");
  const relsXml = ZipPackageHelper.readEntryText(filePath, "xl/worksheets/_rels/sheet1.xml.rels");
  AssertEx.contains("<hyperlinks>", worksheetXml);
  AssertEx.contains('ref="A1"', worksheetXml);
  AssertEx.contains('location="\'Target Sheet\'!C3"', worksheetXml);
  AssertEx.contains('Target="https://example.com/docs?q=1"', relsXml);
  AssertEx.contains('Target="mailto:test@example.com"', relsXml);

  const loaded = new Workbook(filePath);
  HyperlinkScenarioFactory.AssertHyperlinks(loaded);
}

function dataValidationsRoundtripAndEmitExpectedMarkup(): void {
  const temp = new TemporaryDirectory("golden-validations");
  const filePath = temp.getPath("validations.xlsx");
  const workbook = ValidationScenarioFactory.CreateValidationWorkbook();
  workbook.Save(filePath);

  const worksheetXml = ZipPackageHelper.readEntryText(filePath, "xl/worksheets/sheet1.xml");
  AssertEx.contains("<dataValidations", worksheetXml);
  AssertEx.contains('type="list"', worksheetXml);
  AssertEx.contains('type="decimal"', worksheetXml);
  AssertEx.contains('type="custom"', worksheetXml);
  AssertEx.contains('sqref="A1:A3"', worksheetXml);
  AssertEx.contains('sqref="B2:C3 E2:E3"', worksheetXml);
  AssertEx.contains('showDropDown="1"', worksheetXml);
  AssertEx.contains('<formula1>"Open,Closed"</formula1>', worksheetXml);
  AssertEx.contains("<formula2>9.5</formula2>", worksheetXml);

  const loaded = new Workbook(filePath);
  ValidationScenarioFactory.AssertValidations(loaded);
}

function conditionalFormattingsRoundtripAndEmitExpectedMarkup(): void {
  const temp = new TemporaryDirectory("golden-conditional-formatting");
  const filePath = temp.getPath("conditional-formatting.xlsx");
  const workbook = ConditionalFormattingScenarioFactory.CreateConditionalFormattingWorkbook();
  workbook.Save(filePath);

  const worksheetXml = ZipPackageHelper.readEntryText(filePath, "xl/worksheets/sheet1.xml");
  const stylesXml = ZipPackageHelper.readEntryText(filePath, "xl/styles.xml");
  AssertEx.contains("<conditionalFormatting", worksheetXml);
  AssertEx.contains('type="cellIs"', worksheetXml);
  AssertEx.contains('type="expression"', worksheetXml);
  AssertEx.contains('operator="between"', worksheetXml);
  AssertEx.contains('stopIfTrue="1"', worksheetXml);
  AssertEx.contains("<formula>MOD(A1,2)=0</formula>", worksheetXml);
  AssertEx.contains("<dxfs", stylesXml);
  AssertEx.contains('count="3"', stylesXml);

  const loaded = new Workbook(filePath);
  ConditionalFormattingScenarioFactory.AssertConditionalFormattings(loaded);
}

function advancedConditionalFormattingsRoundtripAndEmitExpectedMarkup(): void {
  const temp = new TemporaryDirectory("golden-advanced-conditional-formatting");
  const filePath = temp.getPath("advanced-conditional-formatting.xlsx");
  const workbook = ConditionalFormattingScenarioFactory.CreateAdvancedConditionalFormattingWorkbook();
  workbook.Save(filePath);

  const worksheetXml = ZipPackageHelper.readEntryText(filePath, "xl/worksheets/sheet1.xml");
  AssertEx.contains('type="containsText"', worksheetXml);
  AssertEx.contains('type="notContainsText"', worksheetXml);
  AssertEx.contains('type="beginsWith"', worksheetXml);
  AssertEx.contains('type="endsWith"', worksheetXml);
  AssertEx.contains('type="timePeriod"', worksheetXml);
  AssertEx.contains('type="duplicateValues"', worksheetXml);
  AssertEx.contains('type="uniqueValues"', worksheetXml);
  AssertEx.contains('type="top10"', worksheetXml);
  AssertEx.contains('type="aboveAverage"', worksheetXml);
  AssertEx.contains('type="colorScale"', worksheetXml);
  AssertEx.contains('type="dataBar"', worksheetXml);
  AssertEx.contains('type="iconSet"', worksheetXml);
  AssertEx.contains("<colorScale>", worksheetXml);
  AssertEx.contains("<dataBar>", worksheetXml);
  AssertEx.contains('<iconSet iconSet="4Arrows" reverse="1" showValue="0">', worksheetXml);

  const loaded = new Workbook(filePath);
  ConditionalFormattingScenarioFactory.AssertAdvancedConditionalFormattings(loaded);
}

function pageSetupRoundtripAndEmitExpectedMarkup(): void {
  const temp = new TemporaryDirectory("golden-page-setup");
  const filePath = temp.getPath("page-setup.xlsx");
  const workbook = PageSetupScenarioFactory.CreatePageSetupWorkbook();
  workbook.Save(filePath);

  const worksheetXml = ZipPackageHelper.readEntryText(filePath, "xl/worksheets/sheet1.xml");
  AssertEx.contains("pageMargins", worksheetXml);
  AssertEx.contains("pageSetup", worksheetXml);
  AssertEx.contains('orientation="landscape"', worksheetXml);
  AssertEx.contains('paperSize="9"', worksheetXml);
  AssertEx.contains("<rowBreaks", worksheetXml);
  AssertEx.contains("<colBreaks", worksheetXml);

  const loaded = new Workbook(filePath);
  PageSetupScenarioFactory.AssertPageSetup(loaded);
}

function assertMixedWorkbook(workbook: Workbook, expectedDate1904: boolean): void {
  AssertEx.equal(expectedDate1904, workbook.Settings.Date1904);
  AssertEx.equal("Data", workbook.Worksheets[0].Name);
  AssertEx.equal("Hello", workbook.Worksheets[0].Cells["A1"].StringValue);
  AssertEx.equal(123, workbook.Worksheets[0].Cells["B1"].Value as number);
  AssertEx.equal(true, workbook.Worksheets[0].Cells["C1"].Value as boolean);
  AssertEx.equal(12.5, workbook.Worksheets[0].Cells["D1"].Value as number);
  AssertEx.true(
    Math.abs((workbook.Worksheets[0].Cells["E1"].Value as number) - 6.02214076e23) < 1e10,
  );
  AssertEx.equal(
    new Date(Date.UTC(2024, 4, 6, 7, 8, 9)),
    workbook.Worksheets[0].Cells["F1"].Value as Date,
  );
  AssertEx.equal("=B1*2", workbook.Worksheets[0].Cells["G1"].Formula);
  AssertEx.equal(20, workbook.Worksheets[0].Cells["G1"].Value as number);
}

function workbookMetadataRoundtripAndEmitExpectedMarkup(): void {
  const temp = new TemporaryDirectory("golden-workbook-metadata");
  const filePath = temp.getPath("workbook-metadata.xlsx");
  const workbook = WorkbookMetadataScenarioFactory.CreateWorkbookMetadataWorkbook();
  workbook.Save(filePath);

  const workbookXml = ZipPackageHelper.readEntryText(filePath, "xl/workbook.xml");
  const coreXml = ZipPackageHelper.readEntryText(filePath, "docProps/core.xml");
  const appXml = ZipPackageHelper.readEntryText(filePath, "docProps/app.xml");

  AssertEx.contains("<workbookPr", workbookXml);
  AssertEx.contains('codeName="WorkbookCode"', workbookXml);
  AssertEx.contains('showObjects="placeholders"', workbookXml);
  AssertEx.contains("<workbookProtection", workbookXml);
  AssertEx.contains('workbookPassword="ABCD"', workbookXml);
  AssertEx.contains("<bookViews>", workbookXml);
  AssertEx.contains('activeTab="1"', workbookXml);
  AssertEx.contains('showSheetTabs="0"', workbookXml);
  AssertEx.contains("<calcPr", workbookXml);
  AssertEx.contains('calcMode="manual"', workbookXml);
  AssertEx.contains('refMode="R1C1"', workbookXml);
  AssertEx.contains("Quarterly Summary", coreXml);
  AssertEx.contains("Automation", coreXml);
  AssertEx.contains("Aspose.Cells_FOSS Tests", appXml);
  AssertEx.contains("https://example.com/base/", appXml);

  const loaded = new Workbook(filePath);
  WorkbookMetadataScenarioFactory.AssertWorkbookMetadata(loaded);
}

function extendedDocumentPropertiesDoNotInjectDefaultApplicationMetadata(): void {
  const temp = new TemporaryDirectory("golden-extended-document-properties-no-default-app");
  const filePath = temp.getPath("document-properties-no-default-app.xlsx");
  const workbook = new Workbook();
  workbook.DocumentProperties.Company = "Aspose Cells FOSS";
  workbook.DocumentProperties.Manager = "Release";
  workbook.Save(filePath);

  const appXml = ZipPackageHelper.readEntryText(filePath, "docProps/app.xml");
  AssertEx.contains("<Company>Aspose Cells FOSS</Company>", appXml);
  AssertEx.contains("<Manager>Release</Manager>", appXml);
  AssertEx.false(appXml.includes("<Application>"));
  AssertEx.false(appXml.includes("<AppVersion>"));

  const loaded = new Workbook(filePath);
  AssertEx.equal("Aspose Cells FOSS", loaded.DocumentProperties.Company);
  AssertEx.equal("Release", loaded.DocumentProperties.Manager);
  AssertEx.equal("", loaded.DocumentProperties.Extended.Application);
  AssertEx.equal("", loaded.DocumentProperties.Extended.AppVersion);
}

function workbookMetadataLoadsFromRootRelationshipTargets(): void {
  const temp = new TemporaryDirectory("golden-workbook-metadata-root-relationships");
  const filePath = temp.getPath("workbook-metadata-root-targets.xlsx");
  const workbook = WorkbookMetadataScenarioFactory.CreateWorkbookMetadataWorkbook();
  workbook.Save(filePath);

  ZipPackageHelper.moveEntry(filePath, "docProps/core.xml", "metadata/core-props.xml");
  ZipPackageHelper.moveEntry(filePath, "docProps/app.xml", "metadata/app-props.xml");
  rewriteDocumentPropertiesTargets(filePath, "/metadata/core-props.xml", "/metadata/app-props.xml");

  const loaded = new Workbook(filePath);
  WorkbookMetadataScenarioFactory.AssertWorkbookMetadata(loaded);
}

function unreferencedDocumentPropertiesPartsAreIgnored(): void {
  const temp = new TemporaryDirectory("golden-workbook-metadata-orphaned-docprops");
  const filePath = temp.getPath("workbook-metadata-orphaned-docprops.xlsx");
  const workbook = WorkbookMetadataScenarioFactory.CreateWorkbookMetadataWorkbook();
  workbook.Save(filePath);

  removeDocumentPropertiesRelationships(filePath);

  const loaded = new Workbook(filePath);
  AssertEx.equal("WorkbookCode", loaded.Properties.CodeName);
  assertDocumentPropertiesAreDefault(loaded);
}

function rewriteDocumentPropertiesTargets(packagePath: string, coreTarget: string, appTarget: string): void {
  ZipPackageHelper.rewriteXmlEntry(packagePath, "[Content_Types].xml", (document) => {
    for (const element of childElements(document.documentElement, "Override")) {
      const partName = element.getAttribute("PartName");
      if (partName === "/docProps/core.xml") {
        element.setAttribute("PartName", coreTarget);
      } else if (partName === "/docProps/app.xml") {
        element.setAttribute("PartName", appTarget);
      }
    }
  });

  ZipPackageHelper.rewriteXmlEntry(packagePath, "_rels/.rels", (document) => {
    for (const relationship of childElements(document.documentElement, "Relationship")) {
      const relationshipType = relationship.getAttribute("Type");
      if (relationshipType === CORE_PROPERTIES_RELATIONSHIP) {
        relationship.setAttribute("Target", coreTarget.replace(/^\/+/, ""));
      } else if (relationshipType === EXTENDED_PROPERTIES_RELATIONSHIP) {
        relationship.setAttribute("Target", appTarget.replace(/^\/+/, ""));
      }
    }
  });
}

function removeDocumentPropertiesRelationships(packagePath: string): void {
  ZipPackageHelper.rewriteXmlEntry(packagePath, "_rels/.rels", (document) => {
    const relationships = childElements(document.documentElement, "Relationship");
    for (let index = relationships.length - 1; index >= 0; index--) {
      const relationship = relationships[index];
      const relationshipType = relationship.getAttribute("Type");
      if (
        relationshipType === CORE_PROPERTIES_RELATIONSHIP ||
        relationshipType === EXTENDED_PROPERTIES_RELATIONSHIP
      ) {
        const parent = relationship.parentNode as Node | null;
        if (parent !== null) {
          parent.removeChild(relationship);
        }
      }
    }
  });
}

function rootElement(document: XmlNode): XmlNode | null {
  for (let index = 0; index < document.childNodes!.length; index++) {
    const child = document.childNodes![index] as XmlNode & { nodeType?: number };
    if (child.nodeType === 1) {
      return child;
    }
  }

  return null;
}

function childElements(node: Node, name: string): Element[] {
  const result: Element[] = [];
  for (let index = 0; index < node.childNodes.length; index++) {
    const child = node.childNodes[index];
    if (child.nodeType === 1 && localName(child.nodeName) === name) {
      result.push(child as Element);
    }
  }

  return result;
}

function assertDocumentPropertiesAreDefault(workbook: Workbook): void {
  AssertEx.equal("", workbook.DocumentProperties.Title);
  AssertEx.equal("", workbook.DocumentProperties.Subject);
  AssertEx.equal("", workbook.DocumentProperties.Author);
  AssertEx.equal("", workbook.DocumentProperties.Keywords);
  AssertEx.equal("", workbook.DocumentProperties.Comments);
  AssertEx.equal("", workbook.DocumentProperties.Category);
  AssertEx.equal("", workbook.DocumentProperties.Company);
  AssertEx.equal("", workbook.DocumentProperties.Manager);
  AssertEx.equal("", workbook.DocumentProperties.Core.LastModifiedBy);
  AssertEx.equal("", workbook.DocumentProperties.Core.Revision);
  AssertEx.equal("", workbook.DocumentProperties.Core.ContentStatus);
  AssertEx.null(workbook.DocumentProperties.Core.Created);
  AssertEx.null(workbook.DocumentProperties.Core.Modified);
  AssertEx.equal("", workbook.DocumentProperties.Extended.Application);
  AssertEx.equal("", workbook.DocumentProperties.Extended.AppVersion);
  AssertEx.equal(0, workbook.DocumentProperties.Extended.DocSecurity);
  AssertEx.equal("", workbook.DocumentProperties.Extended.HyperlinkBase);
  AssertEx.false(workbook.DocumentProperties.Extended.ScaleCrop);
  AssertEx.false(workbook.DocumentProperties.Extended.LinksUpToDate);
  AssertEx.false(workbook.DocumentProperties.Extended.SharedDoc);
}

function shapeFileRoundtripAndExport(): void {
  const temp = new TemporaryDirectory("shape-file-roundtrip");
  const inputPath = resolveRepositoryFile("Input", "shape1.xlsx");
  const outputPath = temp.getPath("shape1-export.xlsx");

  const workbook = new Workbook(inputPath);
  workbook.Save(outputPath);

  const loaded = new Workbook(outputPath);
  AssertEx.true(loaded.Worksheets[0].Shapes.Count > 0, "Exported file should contain shapes");
}

function chartFileRoundtripAndExport(): void {
  const inputPath = resolveRepositoryFile("Input", path.join("AreaChart", "areachart1.xlsx"));
  const outputDir = getOrCreateOutputDirectory("chart-golden");
  const outputPath = path.join(outputDir, "areachart1-export.xlsx");

  const workbook = new Workbook(inputPath);
  const originalChartCount = workbook.Worksheets[0].Charts.Count;
  AssertEx.true(originalChartCount > 0, "Input file should contain at least one chart");

  workbook.Save(outputPath);

  const drawingXml = ZipPackageHelper.readEntryText(outputPath, "xl/drawings/drawing1.xml");
  AssertEx.contains("graphicFrame", drawingXml, "Drawing XML must contain graphicFrame");
  AssertEx.contains(
    "http://schemas.openxmlformats.org/drawingml/2006/chart",
    drawingXml,
    "Drawing XML must reference chart namespace URI",
  );

  AssertEx.true(
    ZipPackageHelper.entryExists(outputPath, "xl/charts/chart1.xml"),
    "Saved XLSX must contain xl/charts/chart1.xml",
  );

  const contentTypes = ZipPackageHelper.readEntryText(outputPath, "[Content_Types].xml");
  AssertEx.contains("/xl/charts/chart1.xml", contentTypes, "Content_Types.xml must have chart1.xml override");

  const loaded = new Workbook(outputPath);
  AssertEx.equal(originalChartCount, loaded.Worksheets[0].Charts.Count, "Chart count must match after round-trip");
}

function chartProgrammaticCreationRoundtrip(): void {
  const outputDir = getOrCreateOutputDirectory("chart-golden");
  const outputPath = path.join(outputDir, "column-chart.xlsx");

  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  sheet.Cells["A1"].PutValue("Q1");
  sheet.Cells["B1"].PutValue(10);
  sheet.Cells["A2"].PutValue("Q2");
  sheet.Cells["B2"].PutValue(30);
  sheet.Cells["A3"].PutValue("Q3");
  sheet.Cells["B3"].PutValue(20);
  sheet.Charts.Add(ChartType.Column, "Sheet1!$B$1:$B$3", 5, 0, 20, 10);
  workbook.Save(outputPath);

  const chartXml = ZipPackageHelper.readEntryText(outputPath, "xl/charts/chart1.xml");
  AssertEx.contains("<c:barChart", chartXml, "Chart XML must contain barChart element");
  AssertEx.contains('barDir val="col"', chartXml, "Chart XML must have col direction for column chart");
  AssertEx.contains("Sheet1!$B$1:$B$3", chartXml, "Chart XML must preserve data range");

  const loaded = new Workbook(outputPath);
  AssertEx.equal(1, loaded.Worksheets[0].Charts.Count, "Reloaded file must have 1 chart");
  AssertEx.equal(
    ChartType.Column,
    loaded.Worksheets[0].Charts.get(0).ChartType,
    "Reloaded chart must be Column type",
  );
}

function chartAllTypesExport(): void {
  const outputDir = getOrCreateOutputDirectory("chart-all-types");

  const types: ChartType[] = [
    ChartType.Column,
    ChartType.Bar,
    ChartType.Line,
    ChartType.Area,
    ChartType.Pie,
    ChartType.Doughnut,
    ChartType.Scatter,
    ChartType.Bubble,
    ChartType.Radar,
    ChartType.Stock,
    ChartType.Column3D,
    ChartType.Bar3D,
    ChartType.Line3D,
    ChartType.Area3D,
    ChartType.Pie3D,
    ChartType.Surface3D,
    ChartType.SurfaceWireframe3D,
    ChartType.Contour,
  ];

  for (const type of types) {
    const workbook = ChartScenarioFactory.CreateChartWorkbook(type);
    const outputPath = path.join(outputDir, chartTypeName(type) + ".xlsx");
    workbook.Save(outputPath);

    const reloaded = new Workbook(outputPath);
    AssertEx.equal(1, reloaded.Worksheets[0].Charts.Count, "Expected 1 chart for type: " + type);
    AssertEx.equal(
      type,
      reloaded.Worksheets[0].Charts.get(0).ChartType,
      "Wrong chart type after export/reload for: " + type,
    );
  }
}

function commentXmlPartEmittedWhenCommentsExist(): void {
  const temp = new TemporaryDirectory("comment-golden");
  const filePath = temp.getPath("comment-exists.xlsx");

  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  const comment = sheet.Comments.Add(0, 0);
  comment.Note = "Hello";
  comment.Author = "Bob";
  workbook.Save(filePath);

  AssertEx.true(
    ZipPackageHelper.entryExists(filePath, "xl/comments1.xml"),
    "xl/comments1.xml must be present when worksheet has comments",
  );

  const commentsXml = ZipPackageHelper.readEntryText(filePath, "xl/comments1.xml");
  AssertEx.contains("<comments", commentsXml, "comments XML must have root element");
  AssertEx.contains("<author>Bob</author>", commentsXml, "authors list must contain Bob");
  AssertEx.contains("Hello", commentsXml, "comment text must appear in XML");
  AssertEx.contains('ref="A1"', commentsXml, "comment ref must be A1");

  const contentTypes = ZipPackageHelper.readEntryText(filePath, "[Content_Types].xml");
  AssertEx.contains("/xl/comments1.xml", contentTypes, "[Content_Types].xml must list comments part");
}

function commentXmlPartNotEmittedWhenNoComments(): void {
  const temp = new TemporaryDirectory("comment-golden");
  const filePath = temp.getPath("no-comments.xlsx");

  const workbook = new Workbook();
  workbook.Save(filePath);

  AssertEx.false(
    ZipPackageHelper.entryExists(filePath, "xl/comments1.xml"),
    "xl/comments1.xml must NOT be present when worksheet has no comments",
  );
  AssertEx.false(
    ZipPackageHelper.entryExists(filePath, "xl/drawings/vmlDrawing1.vml"),
    "vmlDrawing1.vml must NOT be present when worksheet has no comments",
  );
}

function commentVmlDrawingEmittedWhenCommentsExist(): void {
  const temp = new TemporaryDirectory("comment-golden");
  const filePath = temp.getPath("comment-vml.xlsx");

  const workbook = new Workbook();
  const comment = workbook.Worksheets[0].Comments.Add(2, 3);
  comment.Width = 200;
  comment.Height = 100;
  workbook.Save(filePath);

  AssertEx.true(
    ZipPackageHelper.entryExists(filePath, "xl/drawings/vmlDrawing1.vml"),
    "vmlDrawing1.vml must be present when worksheet has comments",
  );

  const vmlXml = ZipPackageHelper.readEntryText(filePath, "xl/drawings/vmlDrawing1.vml");
  AssertEx.contains("v:shape", vmlXml, "VML must contain v:shape element");
  AssertEx.contains('ObjectType="Note"', vmlXml, "VML shape must have ObjectType Note");
  AssertEx.contains("<x:Row>2</x:Row>", vmlXml, "VML Row must match comment row");
  AssertEx.contains("<x:Column>3</x:Column>", vmlXml, "VML Column must match comment column");

  const contentTypes = ZipPackageHelper.readEntryText(filePath, "[Content_Types].xml");
  AssertEx.contains("vmlDrawing1.vml", contentTypes, "[Content_Types].xml must list VML part");

  const sheetRels = ZipPackageHelper.readEntryText(filePath, "xl/worksheets/_rels/sheet1.xml.rels");
  AssertEx.contains("relationships/comments", sheetRels, "Sheet rels must contain comments relationship");
  AssertEx.contains("relationships/vmlDrawing", sheetRels, "Sheet rels must contain vmlDrawing relationship");
}

function commentRoundtripPreservesNoteAndAuthor(): void {
  const temp = new TemporaryDirectory("comment-golden");
  const filePath = temp.getPath("comment-roundtrip.xlsx");

  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  sheet.Cells["A1"].PutValue("Data");

  const c1 = sheet.Comments.Add(0, 0);
  c1.Note = "First note";
  c1.Author = "Alice";

  const c2 = sheet.Comments.Add(1, 2);
  c2.Note = "Second note";
  c2.Author = "Bob";

  workbook.Save(filePath);

  const loaded = new Workbook(filePath);
  const loadedSheet = loaded.Worksheets[0];

  AssertEx.equal(2, loadedSheet.Comments.Count, "Must preserve comment count after roundtrip");

  const found0 = loadedSheet.Comments.get("A1");
  AssertEx.notNull(found0, "A1 comment must survive roundtrip");
  AssertEx.equal("First note", found0.Note);
  AssertEx.equal("Alice", found0.Author);

  const found1 = loadedSheet.Comments.get("C2");
  AssertEx.notNull(found1, "C2 comment must survive roundtrip");
  AssertEx.equal("Second note", found1.Note);
  AssertEx.equal("Bob", found1.Author);
}

function commentVisibleRoundtrip(): void {
  const temp = new TemporaryDirectory("comment-golden");
  const filePath = temp.getPath("comment-visible.xlsx");

  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];

  const hidden = sheet.Comments.Add(0, 0);
  hidden.Note = "hidden";
  hidden.IsVisible = false;

  const visible = sheet.Comments.Add(0, 1);
  visible.Note = "visible";
  visible.IsVisible = true;

  workbook.Save(filePath);

  const vmlXml = ZipPackageHelper.readEntryText(filePath, "xl/drawings/vmlDrawing1.vml");
  AssertEx.contains("visibility:visible", vmlXml, "VML must contain visibility:visible for visible comment");

  const loaded = new Workbook(filePath);
  const loadedSheet = loaded.Worksheets[0];

  AssertEx.equal(2, loadedSheet.Comments.Count, "Comment count must be preserved");
  const hiddenLoaded = loadedSheet.Comments.get("A1");
  AssertEx.notNull(hiddenLoaded, "Hidden comment must survive roundtrip");
  AssertEx.false(hiddenLoaded.IsVisible, "Hidden comment must remain hidden after roundtrip");

  const visibleLoaded = loadedSheet.Comments.get("B1");
  AssertEx.notNull(visibleLoaded, "Visible comment must survive roundtrip");
  AssertEx.true(visibleLoaded.IsVisible, "Visible comment must remain visible after roundtrip");
}

function chartTypeName(type: ChartType): string {
  return String((ChartType as unknown as Record<number, string>)[type] ?? type);
}

function resolveRepositoryFile(firstSegment: string, secondSegment: string): string {
  return path.join(resolveRepositoryRoot(), firstSegment, secondSegment);
}

function frenchGroupedNumber(value: number): string {
  return new Intl.NumberFormat("fr-FR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(value);
}

function weekdayMonthDayYear(value: Date, locale: string): string {
  const weekday = new Intl.DateTimeFormat(locale, { weekday: "long" }).format(value);
  const month = new Intl.DateTimeFormat(locale, { month: "long" }).format(value);
  return `${weekday}, ${month} ${value.getDate()}, ${value.getFullYear()}`;
}

TestRunner.runAndExit("CellData.GoldenTests", [
  new TestCase("xlsx_roundtrip_mixed_scalar_cells_file", xlsxRoundtripMixedScalarCellsFile),
  new TestCase("xlsx_roundtrip_mixed_scalar_cells_stream", xlsxRoundtripMixedScalarCellsStream),
  new TestCase("shared_strings_export_uses_sst_when_enabled", sharedStringsExportUsesSstWhenEnabled),
  new TestCase("inline_strings_export_when_shared_strings_disabled", inlineStringsExportWhenSharedStringsDisabled),
  new TestCase("formula_cells_persist_formula_and_cached_value", formulaCellsPersistFormulaAndCachedValue),
  new TestCase("mac1904_datetime_roundtrip_and_markup", mac1904DateTimeRoundtripAndMarkup),
  new TestCase("display_runtime_culture_after_roundtrip", displayRuntimeCultureAfterRoundtrip),
  new TestCase("styled_cells_roundtrip_and_emit_stylesheet", styledCellsRoundtripAndEmitStylesheet),
  new TestCase(
    "worksheet_settings_roundtrip_and_emit_expected_markup",
    worksheetSettingsRoundtripAndEmitExpectedMarkup,
  ),
  new TestCase(
    "worksheet_view_and_tab_color_roundtrip_and_emit_expected_markup",
    worksheetViewAndTabColorRoundtripAndEmitExpectedMarkup,
  ),
  new TestCase(
    "worksheet_protection_roundtrip_and_emit_expected_markup",
    worksheetProtectionRoundtripAndEmitExpectedMarkup,
  ),
  new TestCase("autofilter_roundtrip_and_emit_expected_markup", autoFilterRoundtripAndEmitExpectedMarkup),
  new TestCase("autofilter_omits_invalid_dxf_references", autoFilterOmitsInvalidDxfReferences),
  new TestCase("defined_names_roundtrip_and_emit_expected_markup", definedNamesRoundtripAndEmitExpectedMarkup),
  new TestCase("worksheet_dimension_includes_merge_only_ranges", worksheetDimensionIncludesMergeOnlyRanges),
  new TestCase("hyperlinks_roundtrip_and_emit_expected_markup", hyperlinksRoundtripAndEmitExpectedMarkup),
  new TestCase("data_validations_roundtrip_and_emit_expected_markup", dataValidationsRoundtripAndEmitExpectedMarkup),
  new TestCase(
    "conditional_formattings_roundtrip_and_emit_expected_markup",
    conditionalFormattingsRoundtripAndEmitExpectedMarkup,
  ),
  new TestCase(
    "advanced_conditional_formattings_roundtrip_and_emit_expected_markup",
    advancedConditionalFormattingsRoundtripAndEmitExpectedMarkup,
  ),
  new TestCase("page_setup_roundtrip_and_emit_expected_markup", pageSetupRoundtripAndEmitExpectedMarkup),
  new TestCase(
    "workbook_metadata_roundtrip_and_emit_expected_markup",
    workbookMetadataRoundtripAndEmitExpectedMarkup,
  ),
  new TestCase(
    "extended_document_properties_do_not_inject_default_application_metadata",
    extendedDocumentPropertiesDoNotInjectDefaultApplicationMetadata,
  ),
  new TestCase(
    "workbook_metadata_loads_from_root_relationship_targets",
    workbookMetadataLoadsFromRootRelationshipTargets,
  ),
  new TestCase(
    "unreferenced_document_properties_parts_are_ignored",
    unreferencedDocumentPropertiesPartsAreIgnored,
  ),
  new TestCase("shape_file_roundtrip_and_export", shapeFileRoundtripAndExport),
  new TestCase("chart_file_roundtrip_and_export", chartFileRoundtripAndExport),
  new TestCase("chart_programmatic_creation_roundtrip", chartProgrammaticCreationRoundtrip),
  new TestCase("chart_all_types_export", chartAllTypesExport),
  new TestCase("comment_xml_part_emitted_when_comments_exist", commentXmlPartEmittedWhenCommentsExist),
  new TestCase("comment_xml_part_not_emitted_when_no_comments", commentXmlPartNotEmittedWhenNoComments),
  new TestCase("comment_vml_drawing_emitted_when_comments_exist", commentVmlDrawingEmittedWhenCommentsExist),
  new TestCase("comment_roundtrip_preserves_note_and_author", commentRoundtripPreservesNoteAndAuthor),
  new TestCase("comment_visible_roundtrip", commentVisibleRoundtrip),
]);
