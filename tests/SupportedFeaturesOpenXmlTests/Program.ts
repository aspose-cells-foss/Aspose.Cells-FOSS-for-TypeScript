import {
  CellArea,
  CellValueType,
  Color,
  FillPattern,
  FontUnderlineType,
  FormatConditionType,
  OperatorType,
  PageOrientationType,
  PaperSizeType,
  SaveOptions,
  ValidationAlertType,
  ValidationType,
  VisibilityType,
  Workbook,
} from "../../aspose_cells";
import {
  attribute,
  documentRoot,
  elementChildren,
  firstElement,
  XmlNode,
} from "../../aspose_cells/Xml/XmlReader";
import {
  AssertEx,
  TemporaryDirectory,
  TestCase,
  TestRunner,
  ZipPackageHelper,
} from "../Shared/TestInfrastructure";
import { OpenXmlWorkbookReader, resolvePartPath } from "../Shared/OpenXmlPackage";
import { ListObjectScenarioFactory } from "../Shared/ListObjectScenarioFactory";
import { PictureScenarioFactory } from "../Shared/PictureScenarioFactory";
import { ShapeScenarioFactory } from "../Shared/ShapeScenarioFactory";
import { StyleScenarioFactory } from "../Shared/ScenarioFactories";
import { WorkbookMetadataScenarioFactory } from "../Shared/WorkbookMetadataScenarioFactory";
import { CaseInsensitiveStringMap, CellSnapshot, StyleSnapshot } from "../CompareOpenXml/ComparisonModels";
import { buildCellKey, normalizeFormulaText, normalizeValue } from "../CompareOpenXml/ComparisonValueHelpers";
import {
  createLibraryStyleSnapshot,
  readCellDataWithOpenXmlSdk,
  readCellStylesWithOpenXmlSdk,
} from "../CompareOpenXml/OpenXmlComparisonSupport";

class WorkbookSnapshot {
  public constructor(
    public readonly UsesDate1904: boolean,
    public readonly HasSharedStrings: boolean,
    public readonly Sheets: readonly SheetSnapshot[],
    public readonly Cells: CaseInsensitiveStringMap<CellSnapshot>,
    public readonly Styles: CaseInsensitiveStringMap<StyleSnapshot>,
    public readonly Worksheets: Map<string, WorksheetSnapshot>,
  ) {}
}

class SheetSnapshot {
  public constructor(
    public readonly Name: string,
    public readonly Visibility: string,
  ) {}
}

class DefinedNameSnapshot {
  public constructor(
    public readonly Name: string,
    public readonly Formula: string,
    public readonly LocalSheetIndex: number | null,
    public readonly Hidden: boolean,
    public readonly Comment: string,
  ) {}
}

class WorksheetSnapshot {
  public constructor(
    public readonly Name: string,
    public readonly Visibility: string,
    public readonly Dimension: string,
    public readonly Rows: Map<number, RowSnapshot>,
    public readonly Columns: readonly ColumnSnapshot[],
    public readonly MergeReferences: readonly string[],
    public readonly Hyperlinks: readonly HyperlinkSnapshot[],
    public readonly Validations: readonly ValidationSnapshot[],
    public readonly ConditionalFormattings: readonly ConditionalFormattingSnapshot[],
    public readonly PageSetup: PageSetupSnapshot,
  ) {}
}

class RowSnapshot {
  public constructor(
    public readonly Height: number | null,
    public readonly Hidden: boolean,
  ) {}
}

class ColumnSnapshot {
  public constructor(
    public readonly MinColumnIndex: number,
    public readonly MaxColumnIndex: number,
    public readonly Width: number | null,
    public readonly Hidden: boolean,
  ) {}
}

class HyperlinkSnapshot {
  public constructor(
    public readonly Area: string,
    public readonly Address: string,
    public readonly Location: string,
    public readonly ScreenTip: string,
    public readonly Display: string,
  ) {}
}

class ValidationSnapshot {
  public constructor(
    public readonly Sqref: string,
    public readonly Type: string,
    public readonly Operator: string,
    public readonly ErrorStyle: string,
    public readonly AllowBlank: boolean,
    public readonly ShowDropDown: boolean,
    public readonly ShowInputMessage: boolean,
    public readonly ShowErrorMessage: boolean,
    public readonly PromptTitle: string,
    public readonly Prompt: string,
    public readonly ErrorTitle: string,
    public readonly Error: string,
    public readonly Formula1: string,
    public readonly Formula2: string,
  ) {}
}

class ConditionalFormattingSnapshot {
  public constructor(
    public readonly Sqref: string,
    public readonly Rules: readonly ConditionalFormattingRuleSnapshot[],
  ) {}
}

class ConditionalFormattingRuleSnapshot {
  public constructor(
    public readonly Type: string,
    public readonly Operator: string,
    public readonly Priority: string,
    public readonly StopIfTrue: string,
    public readonly DxfId: string,
    public readonly Formulas: readonly string[],
  ) {}
}

class PageSetupSnapshot {
  public constructor(
    public readonly PrintArea: string,
    public readonly PrintTitleRows: string,
    public readonly PrintTitleColumns: string,
    public readonly LeftMargin: number | null,
    public readonly RightMargin: number | null,
    public readonly TopMargin: number | null,
    public readonly BottomMargin: number | null,
    public readonly HeaderMargin: number | null,
    public readonly FooterMargin: number | null,
    public readonly PaperSize: number | null,
    public readonly Orientation: string,
    public readonly FirstPageNumber: number | null,
    public readonly Scale: number | null,
    public readonly FitToWidth: number | null,
    public readonly FitToHeight: number | null,
    public readonly PrintGridLines: boolean,
    public readonly PrintHeadings: boolean,
    public readonly HorizontalCentered: boolean,
    public readonly VerticalCentered: boolean,
    public readonly OddHeader: string,
    public readonly OddFooter: string,
    public readonly RowBreaks: readonly number[],
    public readonly ColumnBreaks: readonly number[],
  ) {}
}

function supportedFeaturesWorkbookMatchesOpenXmlSnapshot(): void {
  const temp = new TemporaryDirectory("supported-features-openxml");
  const workbookPath = temp.getPath("supported-features.xlsx");

  const workbook = createSupportedFeaturesWorkbook();
  const options = new SaveOptions();
  options.UseSharedStrings = true;
  workbook.Save(workbookPath, options);

  const snapshot = readWorkbookSnapshot(workbookPath);

  AssertEx.true(snapshot.UsesDate1904);
  AssertEx.true(snapshot.HasSharedStrings);
  AssertEx.equal(3, snapshot.Sheets.length);
  assertSheet(snapshot.Sheets[0], "All Features", "Visible");
  assertSheet(snapshot.Sheets[1], "Target Sheet", "Hidden");
  assertSheet(snapshot.Sheets[2], "VeryHidden Sheet", "VeryHidden");

  assertExpectedCell(workbook, snapshot.Cells, "All Features", "A1");
  assertExpectedCell(workbook, snapshot.Cells, "All Features", "B1");
  assertExpectedCell(workbook, snapshot.Cells, "All Features", "C1");
  assertExpectedCell(workbook, snapshot.Cells, "All Features", "D1");
  assertExpectedCell(workbook, snapshot.Cells, "All Features", "E1");
  assertExpectedCell(workbook, snapshot.Cells, "All Features", "F1");
  assertExpectedCell(workbook, snapshot.Cells, "All Features", "G1");
  assertExpectedCell(workbook, snapshot.Cells, "All Features", "H1");
  assertExpectedCell(workbook, snapshot.Cells, "All Features", "I2");
  assertExpectedCell(workbook, snapshot.Cells, "All Features", "J3");
  assertExpectedCell(workbook, snapshot.Cells, "Target Sheet", "C3");

  assertExpectedStyle(workbook, snapshot.Styles, "All Features", "A1");
  assertExpectedStyle(workbook, snapshot.Styles, "All Features", "C4");
  assertExpectedStyle(workbook, snapshot.Styles, "All Features", "H1");
  assertExpectedStyle(workbook, snapshot.Styles, "All Features", "I2");
  assertExpectedStyle(workbook, snapshot.Styles, "All Features", "J3");

  const worksheet = snapshot.Worksheets.get("All Features");
  AssertEx.notNull(worksheet, "Expected the 'All Features' worksheet snapshot.");
  const worksheetSnapshot = worksheet as WorksheetSnapshot;

  AssertEx.equal("A1:J4", worksheetSnapshot.Dimension);
  assertRow(worksheetSnapshot, 1, 22.5, false);
  assertRow(worksheetSnapshot, 3, null, true);
  assertColumn(worksheetSnapshot, 0, 0, 18.25, false);
  assertColumn(worksheetSnapshot, 2, 2, null, true);
  AssertEx.equal(1, worksheetSnapshot.MergeReferences.length);
  AssertEx.equal("A3:B4", worksheetSnapshot.MergeReferences[0]);

  AssertEx.equal(3, worksheetSnapshot.Hyperlinks.length);
  assertHyperlink(
    worksheetSnapshot.Hyperlinks[0],
    "H1",
    "https://example.com/docs?q=1",
    "",
    "External docs",
    "Open External Documentation",
  );
  assertHyperlink(
    worksheetSnapshot.Hyperlinks[1],
    "I2",
    "",
    "'Target Sheet'!C3",
    "Jump to target",
    "Navigate to Target Sheet",
  );
  assertHyperlink(
    worksheetSnapshot.Hyperlinks[2],
    "J3:K4",
    "mailto:test@example.com",
    "",
    "Send mail",
    "Contact Us by Email",
  );

  AssertEx.equal(3, worksheetSnapshot.Validations.length);
  assertValidation(
    getValidationSnapshot(worksheetSnapshot.Validations, "A1:A3"),
    "A1:A3",
    "list",
    "",
    "stop",
    true,
    false,
    true,
    true,
    "Status",
    "Pick a status",
    "Invalid",
    "Choose from the list",
    '"Open,Closed"',
    "",
  );
  assertValidation(
    getValidationSnapshot(worksheetSnapshot.Validations, "B2:C3 E2:E3"),
    "B2:C3 E2:E3",
    "decimal",
    "between",
    "stop",
    false,
    true,
    false,
    true,
    "",
    "",
    "Range",
    "Enter 1.5-9.5",
    "1.5",
    "9.5",
  );
  assertValidation(
    getValidationSnapshot(worksheetSnapshot.Validations, "G1"),
    "G1",
    "custom",
    "",
    "warning",
    false,
    false,
    true,
    false,
    "Code",
    "Up to 5 chars",
    "",
    "",
    "LEN(G1)<=5",
    "",
  );

  AssertEx.equal(2, worksheetSnapshot.ConditionalFormattings.length);
  assertConditionalFormatting(
    getConditionalFormattingSnapshot(worksheetSnapshot.ConditionalFormattings, "A1:A5"),
    "A1:A5",
    2,
  );
  assertConditionalFormatting(
    getConditionalFormattingSnapshot(worksheetSnapshot.ConditionalFormattings, "C1:C3"),
    "C1:C3",
    1,
  );

  const pageSetup = worksheetSnapshot.PageSetup;
  AssertEx.equal("$A$1:$J$10", pageSetup.PrintArea);
  AssertEx.equal("$1:$2", pageSetup.PrintTitleRows);
  AssertEx.equal("$A:$B", pageSetup.PrintTitleColumns);
  AssertEx.equal(0.25, pageSetup.LeftMargin ?? 0);
  AssertEx.equal(0.4, pageSetup.RightMargin ?? 0);
  AssertEx.equal(0.5, pageSetup.TopMargin ?? 0);
  AssertEx.equal(0.6, pageSetup.BottomMargin ?? 0);
  AssertEx.equal(0.2, pageSetup.HeaderMargin ?? 0);
  AssertEx.equal(0.22, pageSetup.FooterMargin ?? 0);
  AssertEx.equal(9, pageSetup.PaperSize ?? 0);
  AssertEx.equal("landscape", pageSetup.Orientation);
  AssertEx.equal(3, pageSetup.FirstPageNumber ?? 0);
  AssertEx.equal(95, pageSetup.Scale ?? 0);
  AssertEx.equal(1, pageSetup.FitToWidth ?? 0);
  AssertEx.equal(2, pageSetup.FitToHeight ?? 0);
  AssertEx.true(pageSetup.PrintGridLines);
  AssertEx.true(pageSetup.PrintHeadings);
  AssertEx.true(pageSetup.HorizontalCentered);
  AssertEx.true(pageSetup.VerticalCentered);
  AssertEx.equal("&LLeft Header&CCenter Header&RRight Header", pageSetup.OddHeader);
  AssertEx.equal("&LLeft Footer&CCenter Footer&RRight Footer", pageSetup.OddFooter);
  AssertEx.equal(2, pageSetup.RowBreaks.length);
  AssertEx.equal(4, pageSetup.RowBreaks[0]);
  AssertEx.equal(7, pageSetup.RowBreaks[1]);
  AssertEx.equal(1, pageSetup.ColumnBreaks.length);
  AssertEx.equal(2, pageSetup.ColumnBreaks[0]);

  const definedNames = readDefinedNames(workbookPath);
  AssertEx.equal(2, definedNames.length);
  assertDefinedName(definedNames[0], "GlobalAmount", "'All Features'!$A$1:$A$3", null, true, "Primary selection");
  assertDefinedName(definedNames[1], "LocalTarget", "'Target Sheet'!$C$3", 1, false, "");
}

function workbookMetadataMatchesOpenXmlSnapshot(): void {
  const temp = new TemporaryDirectory("workbook-metadata-openxml");
  const workbookPath = temp.getPath("workbook-metadata.xlsx");

  const workbook = WorkbookMetadataScenarioFactory.CreateWorkbookMetadataWorkbook();
  workbook.Save(workbookPath);

  const reader = new OpenXmlWorkbookReader(workbookPath);
  const workbookRoot = reader.workbookRoot;
  const workbookProperties = firstElement(workbookRoot, "workbookPr");
  AssertEx.equal("WorkbookCode", attribute(workbookProperties, "codeName"));
  AssertEx.true(attribute(workbookProperties, "showObjects") !== null);
  AssertEx.true(firstElement(workbookRoot, "workbookProtection") !== null);
  const bookViews = elementChildren(firstElement(workbookRoot, "bookViews"), "workbookView");
  AssertEx.equal(1, bookViews.length);
  AssertEx.equal(1, Number.parseInt(attribute(bookViews[0], "activeTab") ?? "0", 10));
  AssertEx.equal(1, Number.parseInt(attribute(bookViews[0], "firstSheet") ?? "0", 10));
  const calculationProperties = firstElement(workbookRoot, "calcPr");
  AssertEx.true(attribute(calculationProperties, "calcMode") !== null);
  AssertEx.true(attribute(calculationProperties, "refMode") !== null);

  const workbookXml = ZipPackageHelper.readEntryText(workbookPath, "xl/workbook.xml");
  const coreXml = ZipPackageHelper.readEntryText(workbookPath, "docProps/core.xml");
  const appXml = ZipPackageHelper.readEntryText(workbookPath, "docProps/app.xml");
  AssertEx.contains('showObjects="placeholders"', workbookXml);
  AssertEx.contains('calcMode="manual"', workbookXml);
  AssertEx.contains('refMode="R1C1"', workbookXml);
  AssertEx.contains("Quarterly Summary", coreXml);
  AssertEx.contains("Automation", coreXml);
  AssertEx.contains("Aspose.Cells_FOSS Tests", appXml);
  AssertEx.contains("https://example.com/base/", appXml);

  const loaded = new Workbook(workbookPath);
  WorkbookMetadataScenarioFactory.AssertWorkbookMetadata(loaded);
}

function tableWorkbookMatchesOpenXmlSnapshot(): void {
  const temp = new TemporaryDirectory("table-openxml");
  const workbookPath = temp.getPath("table-test.xlsx");

  const workbook = ListObjectScenarioFactory.CreateTableWorkbook();
  workbook.Save(workbookPath);

  const reader = new OpenXmlWorkbookReader(workbookPath);
  const worksheetPartPath = reader.sheetPartPathByIndex(0);
  const tablePartPaths = readPartPathsByRelationshipType(reader, worksheetPartPath, "table");
  AssertEx.equal(1, tablePartPaths.length);

  const root = documentRoot(reader.partXml(tablePartPaths[0]));
  if (root === null) {
    throw new Error("Table definition XML is empty.");
  }

  AssertEx.equal("Table1", attribute(root, "displayName"));
  AssertEx.equal("A1:C5", attribute(root, "ref"));
  AssertEx.equal("1", attribute(root, "totalsRowCount") ?? "0");

  const tableColumnsElement = firstElement(root, "tableColumns");
  AssertEx.notNull(tableColumnsElement, "tableColumns element must be present.");
  const tableColumns = tableColumnsElement as XmlNode;
  AssertEx.equal("3", attribute(tableColumns, "count"));

  const columns = elementChildren(tableColumns, "tableColumn");
  AssertEx.equal(3, columns.length);
  AssertEx.equal("Product", attribute(columns[0], "name"));
  AssertEx.equal("Region", attribute(columns[1], "name"));
  AssertEx.equal("Revenue", attribute(columns[2], "name"));
  AssertEx.equal("sum", attribute(columns[2], "totalsRowFunction"));
  AssertEx.equal("Total", attribute(columns[0], "totalsRowLabel"));

  const styleInfo = firstElement(root, "tableStyleInfo");
  AssertEx.notNull(styleInfo, "tableStyleInfo element must be present.");
  const tableStyleInfo = styleInfo as XmlNode;
  AssertEx.equal("TableStyleMedium2", attribute(tableStyleInfo, "name"));
  AssertEx.equal("1", attribute(tableStyleInfo, "showRowStripes"));
}

function tableWorkbookRoundtrip(): void {
  const temp = new TemporaryDirectory("table-roundtrip");
  const workbookPath = temp.getPath("table-roundtrip.xlsx");

  const workbook = ListObjectScenarioFactory.CreateTableWorkbook();
  workbook.Save(workbookPath);

  const loaded = new Workbook(workbookPath);
  ListObjectScenarioFactory.AssertTableWorkbook(loaded);
}

function pictureWorkbookMatchesOpenXmlSnapshot(): void {
  const temp = new TemporaryDirectory("picture-openxml");
  const workbookPath = temp.getPath("picture-test.xlsx");

  const workbook = PictureScenarioFactory.CreatePictureWorkbook();
  workbook.Save(workbookPath);

  const reader = new OpenXmlWorkbookReader(workbookPath);
  const worksheetPartPath = reader.sheetPartPathByIndex(0);
  const drawingPartPath = readPartPathByRelationshipType(reader, worksheetPartPath, "drawing");
  AssertEx.notNull(drawingPartPath, "DrawingsPart must be present on the worksheet.");

  const drawingRoot = documentRoot(reader.partXml(drawingPartPath as string));
  AssertEx.notNull(drawingRoot, "Drawing root must not be null.");
  const root = drawingRoot as XmlNode;

  const anchors = elementChildren(root, "twoCellAnchor");
  AssertEx.equal(2, anchors.length);

  // First picture: PayImage (pay.jpg) anchored at A1:D9
  const anchor0 = anchors[0];
  const from0 = firstElement(anchor0, "from");
  AssertEx.notNull(from0, "from element must be present on first anchor.");
  AssertEx.equal("0", textOfChild(from0 as XmlNode, "col"));
  AssertEx.equal("0", textOfChild(from0 as XmlNode, "row"));
  const to0 = firstElement(anchor0, "to");
  AssertEx.notNull(to0, "to element must be present on first anchor.");
  AssertEx.equal("5", textOfChild(to0 as XmlNode, "col"));
  AssertEx.equal("8", textOfChild(to0 as XmlNode, "row"));
  AssertEx.notNull(firstElement(anchor0, "pic"), "pic element must be present on first anchor.");

  // Second picture: ScreenImage (screen.png) anchored at A11:F17
  const anchor1 = anchors[1];
  const from1 = firstElement(anchor1, "from");
  AssertEx.notNull(from1, "from element must be present on second anchor.");
  AssertEx.equal("0", textOfChild(from1 as XmlNode, "col"));
  AssertEx.equal("10", textOfChild(from1 as XmlNode, "row"));
  AssertEx.notNull(firstElement(anchor1, "pic"), "pic element must be present on second anchor.");

  const imageParts = readPartPathsByRelationshipType(reader, drawingPartPath as string, "image");
  AssertEx.equal(2, imageParts.length);
}

function pictureWorkbookRoundtrip(): void {
  const temp = new TemporaryDirectory("picture-roundtrip");
  const workbookPath = temp.getPath("picture-roundtrip.xlsx");

  const workbook = PictureScenarioFactory.CreatePictureWorkbook();
  workbook.Save(workbookPath);

  const loaded = new Workbook(workbookPath);
  PictureScenarioFactory.AssertPictureWorkbook(loaded);
}

function shapeWorkbookMatchesOpenXmlSnapshot(): void {
  const temp = new TemporaryDirectory("shape-openxml");
  const workbookPath = temp.getPath("shape-test.xlsx");

  const workbook = ShapeScenarioFactory.CreateShapeWorkbook();
  workbook.Save(workbookPath);

  const reader = new OpenXmlWorkbookReader(workbookPath);
  const worksheetPartPath = reader.sheetPartPathByIndex(0);
  const drawingPartPath = readPartPathByRelationshipType(reader, worksheetPartPath, "drawing");
  AssertEx.notNull(drawingPartPath, "DrawingsPart must be present on the worksheet.");

  const drawingRoot = documentRoot(reader.partXml(drawingPartPath as string));
  AssertEx.notNull(drawingRoot, "Drawing root must not be null.");
  const root = drawingRoot as XmlNode;

  const anchors = elementChildren(root, "twoCellAnchor");
  AssertEx.equal(6, anchors.length);

  // Verify rect shape (anchor 0)
  const sp0 = firstElement(anchors[0], "sp");
  AssertEx.notNull(sp0, "First anchor must contain an xdr:sp element.");
  const nvSpPr0 = firstElement(sp0 as XmlNode, "nvSpPr");
  const cNvPr0 = nvSpPr0 === null ? null : firstElement(nvSpPr0, "cNvPr");
  AssertEx.equal("MyRectangle", attribute(cNvPr0, "name") ?? "");
  const spPr0 = firstElement(sp0 as XmlNode, "spPr");
  const prstGeom0 = spPr0 === null ? null : firstElement(spPr0, "prstGeom");
  AssertEx.equal("rect", attribute(prstGeom0, "prst") ?? "");

  // Verify rightArrow shape (anchor 1)
  const sp1 = firstElement(anchors[1], "sp");
  AssertEx.notNull(sp1, "Second anchor must contain an xdr:sp element.");
  const spPr1 = firstElement(sp1 as XmlNode, "spPr");
  const prstGeom1 = spPr1 === null ? null : firstElement(spPr1, "prstGeom");
  AssertEx.equal("rightArrow", attribute(prstGeom1, "prst") ?? "");

  // Verify ellipse shape (anchor 2)
  const sp2 = firstElement(anchors[2], "sp");
  AssertEx.notNull(sp2, "Third anchor must contain an xdr:sp element.");
  const spPr2 = firstElement(sp2 as XmlNode, "spPr");
  const prstGeom2 = spPr2 === null ? null : firstElement(spPr2, "prstGeom");
  AssertEx.equal("ellipse", attribute(prstGeom2, "prst") ?? "");

  // Verify all shapes have xdr:style (default fill makes them visible)
  for (let index = 0; index < 6; index++) {
    const sp = firstElement(anchors[index], "sp");
    AssertEx.notNull(sp, "Anchor " + index + " must have sp element.");
    AssertEx.notNull(
      firstElement(sp as XmlNode, "style"),
      "Shape " + index + " must have xdr:style element for visibility.",
    );
  }

  const imageParts = readPartPathsByRelationshipType(reader, drawingPartPath as string, "image");
  AssertEx.equal(0, imageParts.length);
}

function shapeWorkbookRoundtrip(): void {
  const temp = new TemporaryDirectory("shape-roundtrip");
  const workbookPath = temp.getPath("shape-roundtrip.xlsx");

  const workbook = ShapeScenarioFactory.CreateShapeWorkbook();
  workbook.Save(workbookPath);

  const loaded = new Workbook(workbookPath);
  ShapeScenarioFactory.AssertShapeWorkbook(loaded);
}

function createSupportedFeaturesWorkbook(): Workbook {
  const workbook = new Workbook();
  workbook.Settings.Date1904 = true;

  const sheet = workbook.Worksheets[0];
  sheet.Name = "All Features";

  const targetIndex = workbook.Worksheets.Add("Target Sheet");
  const targetSheet = workbook.Worksheets[targetIndex];
  targetSheet.VisibilityType = VisibilityType.Hidden;
  targetSheet.Cells.get(2, 2).PutValue("Target");

  const veryHiddenIndex = workbook.Worksheets.Add("VeryHidden Sheet");
  workbook.Worksheets[veryHiddenIndex].VisibilityType = VisibilityType.VeryHidden;

  sheet.Cells["A1"].PutValue("Hello");
  sheet.Cells["B1"].PutValue(123);
  sheet.Cells["C1"].PutValue(true);
  sheet.Cells["D1"].PutValue(12.5);
  sheet.Cells["E1"].PutValue(6.02214076e23);
  sheet.Cells["F1"].PutValue(new Date(Date.UTC(2024, 4, 6, 7, 8, 9)));
  sheet.Cells["G1"].PutValue(20);
  sheet.Cells["G1"].Formula = "=B1*2";
  sheet.Cells["A3"].PutValue("Merged");
  sheet.Cells["C4"].PutValue(42.1234);

  const primaryStyle = sheet.Cells["A1"].GetStyle();
  StyleScenarioFactory.ApplyPrimaryStyle(primaryStyle);
  sheet.Cells["A1"].SetStyle(primaryStyle);

  const customStyle = sheet.Cells["C4"].GetStyle();
  StyleScenarioFactory.ApplyCustomNumberStyle(customStyle);
  sheet.Cells["C4"].SetStyle(customStyle);

  sheet.Cells.Rows[1].Height = 22.5;
  sheet.Cells.Rows[3].IsHidden = true;
  sheet.Cells.Columns[0].Width = 18.25;
  sheet.Cells.Columns[2].IsHidden = true;
  sheet.Cells.Merge(2, 0, 2, 2);

  const hyperlinkColor = Color.FromArgb(255, 5, 99, 193);

  sheet.Cells["H1"].PutValue("Open External Documentation");
  const external = sheet.Hyperlinks[sheet.Hyperlinks.Add("H1", 1, 1, "https://example.com/docs?q=1")];
  external.TextToDisplay = "Open External Documentation";
  external.ScreenTip = "External docs";
  const h1Style = sheet.Cells["H1"].GetStyle();
  h1Style.Font.Color = hyperlinkColor;
  h1Style.Font.Underline = FontUnderlineType.Single;
  sheet.Cells["H1"].SetStyle(h1Style);

  sheet.Cells["I2"].PutValue("Navigate to Target Sheet");
  const internalLink = sheet.Hyperlinks[sheet.Hyperlinks.Add("I2", 1, 1, "'Target Sheet'!C3")];
  internalLink.TextToDisplay = "Navigate to Target Sheet";
  internalLink.ScreenTip = "Jump to target";
  const i2Style = sheet.Cells["I2"].GetStyle();
  i2Style.Font.Color = hyperlinkColor;
  i2Style.Font.Underline = FontUnderlineType.Single;
  sheet.Cells["I2"].SetStyle(i2Style);

  sheet.Cells["J3"].PutValue("Contact Us by Email");
  const rangeLink = sheet.Hyperlinks[sheet.Hyperlinks.Add("J3", 2, 2, "mailto:test@example.com")];
  rangeLink.TextToDisplay = "Contact Us by Email";
  rangeLink.ScreenTip = "Send mail";
  const j3Style = sheet.Cells["J3"].GetStyle();
  j3Style.Font.Color = hyperlinkColor;
  j3Style.Font.Underline = FontUnderlineType.Single;
  sheet.Cells["J3"].SetStyle(j3Style);

  const listValidation = sheet.Validations[sheet.Validations.Add(CellArea.CreateCellArea("A1", "A3"))];
  listValidation.Type = ValidationType.List;
  listValidation.Formula1 = '"Open,Closed"';
  listValidation.IgnoreBlank = true;
  listValidation.ShowInput = true;
  listValidation.InputTitle = "Status";
  listValidation.InputMessage = "Pick a status";
  listValidation.ShowError = true;
  listValidation.ErrorTitle = "Invalid";
  listValidation.ErrorMessage = "Choose from the list";

  const decimalValidation = sheet.Validations[sheet.Validations.Add(CellArea.CreateCellArea("B2", "C3"))];
  decimalValidation.Type = ValidationType.Decimal;
  decimalValidation.Operator = OperatorType.Between;
  decimalValidation.Formula1 = "1.5";
  decimalValidation.Formula2 = "9.5";
  decimalValidation.InCellDropDown = false;
  decimalValidation.ShowError = true;
  decimalValidation.ErrorTitle = "Range";
  decimalValidation.ErrorMessage = "Enter 1.5-9.5";
  decimalValidation.AddArea(CellArea.CreateCellArea("E2", "E3"));

  const customValidation = sheet.Validations[sheet.Validations.Add(CellArea.CreateCellArea("G1", "G1"))];
  customValidation.Type = ValidationType.Custom;
  customValidation.AlertStyle = ValidationAlertType.Warning;
  customValidation.Formula1 = "LEN(G1)<=5";
  customValidation.ShowInput = true;
  customValidation.InputTitle = "Code";
  customValidation.InputMessage = "Up to 5 chars";

  const primaryConditionalFormatting = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
  primaryConditionalFormatting.AddArea(CellArea.CreateCellArea("A1", "A5"));
  const betweenCondition = primaryConditionalFormatting.get(
    primaryConditionalFormatting.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "5", "15"),
  );
  betweenCondition.StopIfTrue = true;
  const betweenStyle = betweenCondition.Style;
  betweenStyle.Pattern = FillPattern.Solid;
  betweenStyle.ForegroundColor = Color.FromArgb(255, 255, 199, 206);
  betweenStyle.Font.IsBold = true;
  betweenStyle.Font.Color = Color.FromArgb(255, 156, 0, 6);
  betweenCondition.Style = betweenStyle;
  const expressionCondition = primaryConditionalFormatting.get(
    primaryConditionalFormatting.AddCondition(
      FormatConditionType.Expression,
      OperatorType.None,
      "MOD(A1,2)=0",
      "",
    ),
  );
  const expressionStyle = expressionCondition.Style;
  expressionStyle.Font.IsItalic = true;
  expressionStyle.Font.Color = Color.FromArgb(255, 0, 0, 255);
  expressionCondition.Style = expressionStyle;

  const secondaryConditionalFormatting = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
  secondaryConditionalFormatting.AddArea(CellArea.CreateCellArea("C1", "C3"));
  const greaterThanCondition = secondaryConditionalFormatting.get(
    secondaryConditionalFormatting.AddCondition(
      FormatConditionType.CellValue,
      OperatorType.GreaterThan,
      "10",
      "",
    ),
  );
  const greaterThanStyle = greaterThanCondition.Style;
  greaterThanStyle.Pattern = FillPattern.Solid;
  greaterThanStyle.ForegroundColor = Color.FromArgb(255, 198, 239, 206);
  greaterThanStyle.Font.Color = Color.FromArgb(255, 0, 97, 0);
  greaterThanCondition.Style = greaterThanStyle;

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
  pageSetup.PrintArea = "$A$1:$J$10";
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

  const globalDefinedName = workbook.DefinedNames.get(
    workbook.DefinedNames.Add("GlobalAmount", "'All Features'!$A$1:$A$3"),
  );
  globalDefinedName.Hidden = true;
  globalDefinedName.Comment = "Primary selection";
  workbook.DefinedNames.Add("LocalTarget", "'Target Sheet'!$C$3", 1);

  return workbook;
}

function assertExpectedCell(
  workbook: Workbook,
  actualCells: CaseInsensitiveStringMap<CellSnapshot>,
  sheetName: string,
  cellReference: string,
): void {
  const sheet = workbook.Worksheets[sheetName];
  const cell = sheet.Cells[cellReference];
  const expected = new CellSnapshot(
    sheetName,
    cellReference,
    normalizeCellType(cell.Type),
    normalizeValue(cell.Value),
    normalizeFormulaText(cell.Formula),
  );
  const key = buildCellKey(sheetName, cellReference);
  const actual = actualCells.get(key);
  AssertEx.true(actual !== undefined, `Expected Open XML cell snapshot for ${key}.`);
  const actualSnapshot = actual as CellSnapshot;

  AssertEx.equal(expected.SheetName, actualSnapshot.SheetName);
  AssertEx.equal(expected.CellReference, actualSnapshot.CellReference);
  AssertEx.equal(expected.CellType, actualSnapshot.CellType);
  AssertEx.equal(expected.Formula, actualSnapshot.Formula);

  if (cell.Value instanceof Date) {
    const expectedDateTime = cell.Value;
    const actualDateTime = new Date(actualSnapshot.Value);
    AssertEx.equal(expectedDateTime.getTime(), actualDateTime.getTime());
    return;
  }

  AssertEx.equal(expected.Value, actualSnapshot.Value);
}

function assertExpectedStyle(
  workbook: Workbook,
  actualStyles: CaseInsensitiveStringMap<StyleSnapshot>,
  sheetName: string,
  cellReference: string,
): void {
  const expected = createLibraryStyleSnapshot(
    sheetName,
    cellReference,
    workbook.Worksheets[sheetName].Cells[cellReference].GetStyle(),
  );
  const key = buildCellKey(sheetName, cellReference);
  const actual = actualStyles.get(key);
  AssertEx.true(actual !== undefined, `Expected Open XML style snapshot for ${key}.`);
  const actualSnapshot = actual as StyleSnapshot;

  AssertEx.equal(expected.SheetName, actualSnapshot.SheetName);
  AssertEx.equal(expected.CellReference, actualSnapshot.CellReference);
  AssertEx.equal(expected.FontName, actualSnapshot.FontName);
  AssertEx.equal(expected.FontSize, actualSnapshot.FontSize);
  AssertEx.equal(expected.FontBold, actualSnapshot.FontBold);
  AssertEx.equal(expected.FontItalic, actualSnapshot.FontItalic);
  AssertEx.equal(expected.FontUnderline, actualSnapshot.FontUnderline);
  AssertEx.equal(expected.FontStrikeThrough, actualSnapshot.FontStrikeThrough);
  AssertEx.equal(expected.FontColor, actualSnapshot.FontColor);
  AssertEx.equal(expected.FillPattern, actualSnapshot.FillPattern);
  AssertEx.equal(expected.FillForegroundColor, actualSnapshot.FillForegroundColor);
  AssertEx.equal(expected.FillBackgroundColor, actualSnapshot.FillBackgroundColor);
  AssertEx.equal(expected.LeftBorderStyle, actualSnapshot.LeftBorderStyle);
  AssertEx.equal(expected.LeftBorderColor, actualSnapshot.LeftBorderColor);
  AssertEx.equal(expected.RightBorderStyle, actualSnapshot.RightBorderStyle);
  AssertEx.equal(expected.RightBorderColor, actualSnapshot.RightBorderColor);
  AssertEx.equal(expected.TopBorderStyle, actualSnapshot.TopBorderStyle);
  AssertEx.equal(expected.TopBorderColor, actualSnapshot.TopBorderColor);
  AssertEx.equal(expected.BottomBorderStyle, actualSnapshot.BottomBorderStyle);
  AssertEx.equal(expected.BottomBorderColor, actualSnapshot.BottomBorderColor);
  AssertEx.equal(expected.DiagonalBorderStyle, actualSnapshot.DiagonalBorderStyle);
  AssertEx.equal(expected.DiagonalBorderColor, actualSnapshot.DiagonalBorderColor);
  AssertEx.equal(expected.DiagonalUp, actualSnapshot.DiagonalUp);
  AssertEx.equal(expected.DiagonalDown, actualSnapshot.DiagonalDown);
  AssertEx.equal(expected.HorizontalAlignment, actualSnapshot.HorizontalAlignment);
  AssertEx.equal(expected.VerticalAlignment, actualSnapshot.VerticalAlignment);
  AssertEx.equal(expected.WrapText, actualSnapshot.WrapText);
  AssertEx.equal(expected.IndentLevel, actualSnapshot.IndentLevel);
  AssertEx.equal(expected.TextRotation, actualSnapshot.TextRotation);
  AssertEx.equal(expected.ShrinkToFit, actualSnapshot.ShrinkToFit);
  AssertEx.equal(expected.ReadingOrder, actualSnapshot.ReadingOrder);
  AssertEx.equal(expected.RelativeIndent, actualSnapshot.RelativeIndent);
  AssertEx.equal(expected.IsLocked, actualSnapshot.IsLocked);
  AssertEx.equal(expected.IsHidden, actualSnapshot.IsHidden);

  if (expected.NumberFormatCode.length > 0) {
    AssertEx.equal(expected.NumberFormatCode, actualSnapshot.NumberFormatCode);
  } else {
    AssertEx.equal(expected.NumberFormatId, actualSnapshot.NumberFormatId);
  }
}

function normalizeCellType(type: CellValueType): string {
  switch (type) {
    case CellValueType.IsString:
      return "String";
    case CellValueType.IsNumeric:
      return "Number";
    case CellValueType.IsBool:
      return "Boolean";
    case CellValueType.IsDateTime:
      return "DateTime";
    case CellValueType.IsUnknown:
      return "Formula";
    default:
      return "Blank";
  }
}

function readWorkbookSnapshot(workbookPath: string): WorkbookSnapshot {
  const cells = readCellDataWithOpenXmlSdk(workbookPath);
  const styles = readCellStylesWithOpenXmlSdk(workbookPath);

  const reader = new OpenXmlWorkbookReader(workbookPath);
  const workbookRoot = reader.workbookRoot;
  const sheets: SheetSnapshot[] = [];
  const worksheets = new Map<string, WorksheetSnapshot>();

  for (const sheetElement of elementChildren(firstElement(workbookRoot, "sheets"), "sheet")) {
    const relationshipId = attribute(sheetElement, "id");
    if (relationshipId === null || relationshipId.trim().length === 0) {
      continue;
    }

    const sheetName = attribute(sheetElement, "name") ?? "";
    let worksheetPartPath: string;
    try {
      worksheetPartPath = reader.sheetPartPathByName(sheetName);
    } catch {
      continue;
    }

    sheets.push(new SheetSnapshot(sheetName, resolveVisibility(attribute(sheetElement, "state"))));
    worksheets.set(
      sheetName,
      readWorksheetSnapshot(reader, workbookRoot, sheetElement, worksheetPartPath),
    );
  }

  const usesDate1904 = parseBool(attribute(firstElement(workbookRoot, "workbookPr"), "date1904"));
  const hasSharedStrings = reader.hasPart("xl/sharedStrings.xml");
  return new WorkbookSnapshot(usesDate1904, hasSharedStrings, sheets, cells, styles, worksheets);
}

function readWorksheetSnapshot(
  reader: OpenXmlWorkbookReader,
  workbookRoot: XmlNode,
  sheetElement: XmlNode,
  worksheetPartPath: string,
): WorksheetSnapshot {
  const worksheetRoot = documentRoot(reader.partXml(worksheetPartPath));
  if (worksheetRoot === null) {
    throw new Error("Worksheet XML is empty.");
  }

  const sheetName = attribute(sheetElement, "name") ?? "";
  const sheetIndex = Number.parseInt(attribute(sheetElement, "sheetId") ?? "1", 10) - 1;

  const rows = new Map<number, RowSnapshot>();
  for (const row of elementChildren(firstElement(worksheetRoot, "sheetData"), "row")) {
    const rowIndex = tryParseInt(attribute(row, "r"));
    if (rowIndex === null) {
      continue;
    }

    rows.set(
      rowIndex - 1,
      new RowSnapshot(tryParseDouble(attribute(row, "ht")), parseBool(attribute(row, "hidden"))),
    );
  }

  const columns: ColumnSnapshot[] = [];
  for (const column of elementChildren(firstElement(worksheetRoot, "cols"), "col")) {
    const minColumnIndex = tryParseInt(attribute(column, "min"));
    const maxColumnIndex = tryParseInt(attribute(column, "max"));
    if (minColumnIndex === null || maxColumnIndex === null) {
      continue;
    }

    columns.push(
      new ColumnSnapshot(
        minColumnIndex - 1,
        maxColumnIndex - 1,
        tryParseDouble(attribute(column, "width")),
        parseBool(attribute(column, "hidden")),
      ),
    );
  }

  const merges: string[] = [];
  for (const mergeCell of elementChildren(firstElement(worksheetRoot, "mergeCells"), "mergeCell")) {
    const reference = (attribute(mergeCell, "ref") ?? "").toUpperCase();
    if (reference.length > 0) {
      merges.push(reference);
    }
  }

  const hyperlinkTargets = readHyperlinkTargets(reader, worksheetPartPath);
  const hyperlinks: HyperlinkSnapshot[] = [];
  for (const hyperlink of elementChildren(firstElement(worksheetRoot, "hyperlinks"), "hyperlink")) {
    const relationshipId = attribute(hyperlink, "id");
    const hyperlinkAddress =
      relationshipId === null ? "" : hyperlinkTargets.get(relationshipId) ?? "";
    hyperlinks.push(
      new HyperlinkSnapshot(
        (attribute(hyperlink, "ref") ?? "").toUpperCase(),
        hyperlinkAddress,
        attribute(hyperlink, "location") ?? "",
        attribute(hyperlink, "tooltip") ?? "",
        attribute(hyperlink, "display") ?? "",
      ),
    );
  }

  const validations: ValidationSnapshot[] = [];
  for (const validation of elementChildren(firstElement(worksheetRoot, "dataValidations"), "dataValidation")) {
    validations.push(
      new ValidationSnapshot(
        (attribute(validation, "sqref") ?? "").toUpperCase(),
        attribute(validation, "type") ?? "",
        attribute(validation, "operator") ?? "",
        attribute(validation, "errorStyle") ?? "stop",
        parseBool(attribute(validation, "allowBlank")),
        parseBool(attribute(validation, "showDropDown")),
        parseBool(attribute(validation, "showInputMessage")),
        parseBool(attribute(validation, "showErrorMessage")),
        attribute(validation, "promptTitle") ?? "",
        attribute(validation, "prompt") ?? "",
        attribute(validation, "errorTitle") ?? "",
        attribute(validation, "error") ?? "",
        textOf(firstElement(validation, "formula1")),
        textOf(firstElement(validation, "formula2")),
      ),
    );
  }

  const conditionalFormattings: ConditionalFormattingSnapshot[] = [];
  for (const formatting of elementChildren(worksheetRoot, "conditionalFormatting")) {
    const rules: ConditionalFormattingRuleSnapshot[] = [];
    for (const rule of elementChildren(formatting, "cfRule")) {
      rules.push(
        new ConditionalFormattingRuleSnapshot(
          attribute(rule, "type") ?? "",
          attribute(rule, "operator") ?? "",
          attribute(rule, "priority") ?? "",
          attribute(rule, "stopIfTrue") ?? "",
          attribute(rule, "dxfId") ?? "",
          elementChildren(rule, "formula").map((formula) => formula.textContent ?? ""),
        ),
      );
    }

    conditionalFormattings.push(
      new ConditionalFormattingSnapshot((attribute(formatting, "sqref") ?? "").toUpperCase(), rules),
    );
  }

  const pageSetup = readPageSetupSnapshot(workbookRoot, worksheetRoot, sheetIndex);
  return new WorksheetSnapshot(
    sheetName,
    resolveVisibility(attribute(sheetElement, "state")),
    (attribute(firstElement(worksheetRoot, "dimension"), "ref") ?? "").toUpperCase(),
    rows,
    columns,
    merges,
    hyperlinks,
    validations,
    conditionalFormattings,
    pageSetup,
  );
}

function readPageSetupSnapshot(
  workbookRoot: XmlNode,
  worksheetRoot: XmlNode,
  sheetIndex: number,
): PageSetupSnapshot {
  const margins = firstElement(worksheetRoot, "pageMargins");
  const pageSetup = firstElement(worksheetRoot, "pageSetup");
  const printOptions = firstElement(worksheetRoot, "printOptions");
  const headerFooter = firstElement(worksheetRoot, "headerFooter");
  const rowBreaks = elementChildren(firstElement(worksheetRoot, "rowBreaks"), "brk").map((element) =>
    Number.parseInt(attribute(element, "id") ?? "0", 10),
  );
  const columnBreaks = elementChildren(firstElement(worksheetRoot, "colBreaks"), "brk").map((element) =>
    Number.parseInt(attribute(element, "id") ?? "0", 10),
  );

  let printArea: string | null = null;
  let printTitleRows: string | null = null;
  let printTitleColumns: string | null = null;
  for (const definedName of elementChildren(firstElement(workbookRoot, "definedNames"), "definedName")) {
    const localSheetId = attribute(definedName, "localSheetId");
    if (localSheetId !== String(sheetIndex)) {
      continue;
    }

    const name = attribute(definedName, "name") ?? "";
    if (name === "_xlnm.Print_Area") {
      printArea = unqualifyDefinedNameValue(definedName.textContent ?? "");
    } else if (name === "_xlnm.Print_Titles") {
      for (const segment of (definedName.textContent ?? "").split(",")) {
        const unqualified = unqualifyDefinedNameValue(segment);
        if (
          unqualified.indexOf("$") >= 0 &&
          containsDigit(unqualified) &&
          unqualified.indexOf(":") >= 0 &&
          isDigit(unqualified.split("$").join("").charAt(0))
        ) {
          printTitleRows = unqualified;
        } else {
          printTitleColumns = unqualified;
        }
      }
    }
  }

  return new PageSetupSnapshot(
    printArea ?? "",
    printTitleRows ?? "",
    printTitleColumns ?? "",
    tryParseDouble(attribute(margins, "left")),
    tryParseDouble(attribute(margins, "right")),
    tryParseDouble(attribute(margins, "top")),
    tryParseDouble(attribute(margins, "bottom")),
    tryParseDouble(attribute(margins, "header")),
    tryParseDouble(attribute(margins, "footer")),
    tryParseInt(attribute(pageSetup, "paperSize")),
    attribute(pageSetup, "orientation") ?? "",
    tryParseInt(attribute(pageSetup, "firstPageNumber")),
    tryParseInt(attribute(pageSetup, "scale")),
    tryParseInt(attribute(pageSetup, "fitToWidth")),
    tryParseInt(attribute(pageSetup, "fitToHeight")),
    parseBool(attribute(printOptions, "gridLines")),
    parseBool(attribute(printOptions, "headings")),
    parseBool(attribute(printOptions, "horizontalCentered")),
    parseBool(attribute(printOptions, "verticalCentered")),
    textOf(firstElement(headerFooter, "oddHeader")),
    textOf(firstElement(headerFooter, "oddFooter")),
    rowBreaks,
    columnBreaks,
  );
}

function readDefinedNames(workbookPath: string): DefinedNameSnapshot[] {
  const reader = new OpenXmlWorkbookReader(workbookPath);
  const workbookRoot = reader.workbookRoot;
  const definedNames: DefinedNameSnapshot[] = [];

  for (const definedName of elementChildren(firstElement(workbookRoot, "definedNames"), "definedName")) {
    const name = attribute(definedName, "name") ?? "";
    if (name === "_xlnm.Print_Area" || name === "_xlnm.Print_Titles") {
      continue;
    }

    definedNames.push(
      new DefinedNameSnapshot(
        name,
        definedName.textContent ?? "",
        tryParseInt(attribute(definedName, "localSheetId")),
        parseBool(attribute(definedName, "hidden")),
        attribute(definedName, "comment") ?? "",
      ),
    );
  }

  return definedNames;
}

function assertDefinedName(
  actual: DefinedNameSnapshot,
  name: string,
  formula: string,
  localSheetIndex: number | null,
  hidden: boolean,
  comment: string,
): void {
  AssertEx.equal(name, actual.Name);
  AssertEx.equal(formula, actual.Formula);
  AssertEx.equal(localSheetIndex, actual.LocalSheetIndex);
  AssertEx.equal(hidden, actual.Hidden);
  AssertEx.equal(comment, actual.Comment);
}

function assertSheet(actual: SheetSnapshot, expectedName: string, expectedVisibility: string): void {
  AssertEx.equal(expectedName, actual.Name);
  AssertEx.equal(expectedVisibility, actual.Visibility);
}

function assertRow(
  snapshot: WorksheetSnapshot,
  rowIndex: number,
  height: number | null,
  hidden: boolean,
): void {
  const row = snapshot.Rows.get(rowIndex);
  AssertEx.true(row !== undefined, `Expected row ${rowIndex} to exist in Open XML snapshot.`);
  const rowSnapshot = row as RowSnapshot;
  AssertEx.equal(height, rowSnapshot.Height);
  AssertEx.equal(hidden, rowSnapshot.Hidden);
}

function assertColumn(
  snapshot: WorksheetSnapshot,
  minColumnIndex: number,
  maxColumnIndex: number,
  width: number | null,
  hidden: boolean,
): void {
  const found = snapshot.Columns.filter(
    (item) => item.MinColumnIndex === minColumnIndex && item.MaxColumnIndex === maxColumnIndex,
  );
  const column = found.length === 0 ? null : found[0];
  AssertEx.notNull(
    column,
    `Expected column range ${minColumnIndex}:${maxColumnIndex} to exist in Open XML snapshot.`,
  );
  const columnSnapshot = column as ColumnSnapshot;
  AssertEx.equal(width, columnSnapshot.Width);
  AssertEx.equal(hidden, columnSnapshot.Hidden);
}

function assertHyperlink(
  actual: HyperlinkSnapshot,
  area: string,
  address: string,
  location: string,
  screenTip: string,
  display: string,
): void {
  AssertEx.equal(area, actual.Area);
  AssertEx.equal(address, actual.Address);
  AssertEx.equal(location, actual.Location);
  AssertEx.equal(screenTip, actual.ScreenTip);
  AssertEx.equal(display, actual.Display);
}

function getValidationSnapshot(
  snapshots: readonly ValidationSnapshot[],
  sqref: string,
): ValidationSnapshot {
  const found = snapshots.filter((snapshot) => snapshot.Sqref === sqref);
  AssertEx.equal(1, found.length, `Expected exactly one validation snapshot for '${sqref}'.`);
  return found[0];
}

function getConditionalFormattingSnapshot(
  snapshots: readonly ConditionalFormattingSnapshot[],
  sqref: string,
): ConditionalFormattingSnapshot {
  const found = snapshots.filter((snapshot) => snapshot.Sqref === sqref);
  AssertEx.equal(1, found.length, `Expected exactly one conditional formatting snapshot for '${sqref}'.`);
  return found[0];
}

function assertValidation(
  actual: ValidationSnapshot,
  sqref: string,
  type: string,
  operatorName: string,
  errorStyle: string,
  allowBlank: boolean,
  showDropDown: boolean,
  showInputMessage: boolean,
  showErrorMessage: boolean,
  promptTitle: string,
  prompt: string,
  errorTitle: string,
  error: string,
  formula1: string,
  formula2: string,
): void {
  AssertEx.equal(sqref, actual.Sqref);
  AssertEx.equal(type, actual.Type);
  AssertEx.equal(operatorName, actual.Operator);
  AssertEx.equal(errorStyle, actual.ErrorStyle);
  AssertEx.equal(allowBlank, actual.AllowBlank);
  AssertEx.equal(showDropDown, actual.ShowDropDown);
  AssertEx.equal(showInputMessage, actual.ShowInputMessage);
  AssertEx.equal(showErrorMessage, actual.ShowErrorMessage);
  AssertEx.equal(promptTitle, actual.PromptTitle);
  AssertEx.equal(prompt, actual.Prompt);
  AssertEx.equal(errorTitle, actual.ErrorTitle);
  AssertEx.equal(error, actual.Error);
  AssertEx.equal(formula1, actual.Formula1);
  AssertEx.equal(formula2, actual.Formula2);
}

function assertConditionalFormatting(
  actual: ConditionalFormattingSnapshot,
  sqref: string,
  ruleCount: number,
): void {
  AssertEx.equal(sqref, actual.Sqref);
  AssertEx.equal(ruleCount, actual.Rules.length);
}

function resolveVisibility(state: string | null): string {
  switch (state) {
    case "hidden":
      return "Hidden";
    case "veryHidden":
      return "VeryHidden";
    default:
      return "Visible";
  }
}

function unqualifyDefinedNameValue(value: string): string {
  const trimmed = value.trim();
  const exclamationIndex = trimmed.lastIndexOf("!");
  return exclamationIndex >= 0 ? trimmed.substring(exclamationIndex + 1) : trimmed;
}

function tryParseInt(value: string | null): number | null {
  if (value === null || value.trim().length === 0) {
    return null;
  }

  const parsed = Number.parseInt(value.trim(), 10);
  return Number.isNaN(parsed) ? null : parsed;
}

function tryParseDouble(value: string | null): number | null {
  if (value === null || value.trim().length === 0) {
    return null;
  }

  const parsed = Number.parseFloat(value.trim());
  return Number.isNaN(parsed) ? null : parsed;
}

function parseBool(value: string | null): boolean {
  return value === "1" || (value ?? "").toLowerCase() === "true";
}

function containsDigit(value: string): boolean {
  for (const character of "0123456789") {
    if (value.indexOf(character) >= 0) {
      return true;
    }
  }

  return false;
}

function isDigit(value: string): boolean {
  return value.length > 0 && value >= "0" && value <= "9";
}

function textOf(node: XmlNode | null): string {
  return node === null ? "" : (node.textContent ?? "");
}

function textOfChild(parent: XmlNode, name: string): string {
  return textOf(firstElement(parent, name));
}

function readPartRelationships(
  reader: OpenXmlWorkbookReader,
  ownerPartPath: string,
): { Id: string; Type: string; Target: string; TargetPath: string; Folder: string }[] {
  const slashIndex = ownerPartPath.lastIndexOf("/");
  const folder = slashIndex >= 0 ? ownerPartPath.substring(0, slashIndex) : "";
  const fileName = slashIndex >= 0 ? ownerPartPath.substring(slashIndex + 1) : ownerPartPath;
  const relationshipsPath = (folder.length > 0 ? folder + "/" : "") + "_rels/" + fileName + ".rels";
  if (!reader.hasPart(relationshipsPath)) {
    return [];
  }

  const root = documentRoot(reader.partXml(relationshipsPath));
  if (root === null) {
    return [];
  }

  const relationships: { Id: string; Type: string; Target: string; TargetPath: string; Folder: string }[] = [];
  for (const relationship of elementChildren(root, "Relationship")) {
    const target = attribute(relationship, "Target") ?? "";
    relationships.push({
      Id: attribute(relationship, "Id") ?? "",
      Type: attribute(relationship, "Type") ?? "",
      Target: target,
      TargetPath: resolvePartPath(folder, target),
      Folder: folder,
    });
  }

  return relationships;
}

function readPartPathByRelationshipType(
  reader: OpenXmlWorkbookReader,
  ownerPartPath: string,
  relationshipType: string,
): string | null {
  const matches = readPartPathsByRelationshipType(reader, ownerPartPath, relationshipType);
  return matches.length === 0 ? null : matches[0];
}

function readPartPathsByRelationshipType(
  reader: OpenXmlWorkbookReader,
  ownerPartPath: string,
  relationshipType: string,
): string[] {
  const suffix = "/" + relationshipType;
  return readPartRelationships(reader, ownerPartPath)
    .filter((relationship) => relationship.Type.endsWith(suffix))
    .map((relationship) => relationship.TargetPath);
}

function readHyperlinkTargets(
  reader: OpenXmlWorkbookReader,
  worksheetPartPath: string,
): Map<string, string> {
  const targets = new Map<string, string>();
  for (const relationship of readPartRelationships(reader, worksheetPartPath)) {
    if (relationship.Type.endsWith("/hyperlink")) {
      targets.set(relationship.Id, relationship.Target);
    }
  }

  return targets;
}

TestRunner.runAndExit("SupportedFeatures.OpenXmlTests", [
  new TestCase(
    "supported_features_workbook_matches_openxml_snapshot",
    supportedFeaturesWorkbookMatchesOpenXmlSnapshot,
  ),
  new TestCase("workbook_metadata_matches_openxml_snapshot", workbookMetadataMatchesOpenXmlSnapshot),
  new TestCase("table_workbook_matches_openxml_snapshot", tableWorkbookMatchesOpenXmlSnapshot),
  new TestCase("table_workbook_roundtrip", tableWorkbookRoundtrip),
  new TestCase("picture_workbook_matches_openxml_snapshot", pictureWorkbookMatchesOpenXmlSnapshot),
  new TestCase("picture_workbook_roundtrip", pictureWorkbookRoundtrip),
  new TestCase("shape_workbook_matches_openxml_snapshot", shapeWorkbookMatchesOpenXmlSnapshot),
  new TestCase("shape_workbook_roundtrip", shapeWorkbookRoundtrip),
]);
