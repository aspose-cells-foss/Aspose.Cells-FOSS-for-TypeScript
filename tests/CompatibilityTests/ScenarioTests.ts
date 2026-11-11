import * as fs from "fs";
import * as path from "path";
import {
  CellArea,
  CellValueType,
  CellsException,
  Color,
  FormatConditionType,
  InvalidFileFormatException,
  OperatorType,
  SaveFormat,
  TargetModeType,
  ValidationType,
  VisibilityType,
  Workbook,
} from "../../aspose_cells";
import { documentRoot, parseXml } from "../../aspose_cells/Xml/XmlReader";
import {
  AssertEx,
  MemoryStream,
  TemporaryDirectory,
  ZipPackageHelper,
  getOrCreateOutputDirectory,
} from "../Shared/TestInfrastructure";
import {
  AutoFilterScenarioFactory,
  ConditionalFormattingScenarioFactory,
  DefinedNameScenarioFactory,
  PageSetupScenarioFactory,
  StyleScenarioFactory,
  ValidationScenarioFactory,
  WorkbookScenarioFactory,
  WorksheetScenarioFactory,
} from "../Shared/ScenarioFactories";
import { WorkbookMetadataScenarioFactory } from "../Shared/WorkbookMetadataScenarioFactory";
import { assertElementOrder, childLocalNames } from "./CompatSupport";

export function assertWorkbookDataEquals(expected: Workbook, actual: Workbook): void {
  const expectedSheet = expected.Worksheets.get(0);
  const actualSheet = actual.Worksheets.get(0);

  for (const cellName of ["A1", "B1", "C1", "D1", "E1", "F1", "G1"]) {
    const expectedCell = expectedSheet.Cells.get(cellName);
    const actualCell = actualSheet.Cells.get(cellName);
    AssertEx.equal(expectedCell.Type, actualCell.Type, "Type mismatch for " + cellName + ".");
    AssertEx.equal(expectedCell.StringValue, actualCell.StringValue, "Value mismatch for " + cellName + ".");
    AssertEx.equal(expectedCell.Formula, actualCell.Formula, "Formula mismatch for " + cellName + ".");
  }
}

export function fileAndStreamLoadPathsProduceSameValues(): void {
  const temp = new TemporaryDirectory("compat-load-paths");
  const fileName = temp.getPath("book.xlsx");
  const workbook = WorkbookScenarioFactory.CreateMixedCellWorkbook();
  workbook.Save(fileName);

  const fromFile = new Workbook(fileName);
  const fromStream = new Workbook(new MemoryStream(fs.readFileSync(fileName)));
  assertWorkbookDataEquals(fromFile, fromStream);
}

export function saveOverloadsProduceEquivalentWorkbooks(): void {
  const temp = new TemporaryDirectory("compat-save-overloads");
  const fileName = temp.getPath("book-file.xlsx");
  const workbook = WorkbookScenarioFactory.CreateMixedCellWorkbook();
  workbook.Save(fileName, SaveFormat.Xlsx);

  const stream = new MemoryStream();
  workbook.Save(stream, SaveFormat.Xlsx);
  stream.position = 0;

  const fromFile = new Workbook(fileName);
  const fromStream = new Workbook(stream);
  assertWorkbookDataEquals(fromFile, fromStream);
}

export function formulaSetterAcceptsWithOrWithoutLeadingEqual(): void {
  const workbook = new Workbook();
  const cell = workbook.Worksheets.get(0).Cells.get("A1");
  cell.PutValue(10);
  cell.Formula = "B1+C1";
  AssertEx.equal("=B1+C1", cell.Formula);
  cell.Formula = "=D1+E1";
  AssertEx.equal("=D1+E1", cell.Formula);
}

export function exceptionMappingUsesCellsExceptionTypes(): void {
  const missingSheet = AssertEx.throws(() => {
    const unused: unknown = new Workbook().Worksheets["missing"];
    return unused;
  });
  AssertEx.true(missingSheet instanceof CellsException, "Missing worksheet must raise CellsException.");

  const invalidCell = AssertEx.throws(() => {
    const unused: unknown = new Workbook().Worksheets.get(0).Cells["1A"];
    return unused;
  });
  AssertEx.true(invalidCell instanceof CellsException, "Invalid cell reference must raise CellsException.");

  const invalidWorkbook = AssertEx.throws(() => new Workbook(new MemoryStream(Buffer.from([1, 2, 3, 4]))));
  AssertEx.true(
    invalidWorkbook instanceof InvalidFileFormatException,
    "Invalid workbook payload must raise InvalidFileFormatException.",
  );
}

export function publicTypeMappingMatchesAfterRoundtrip(): void {
  const temp = new TemporaryDirectory("compat-types");
  const fileName = temp.getPath("types.xlsx");
  const workbook = WorkbookScenarioFactory.CreateMixedCellWorkbook();
  workbook.Save(fileName);

  const loaded = new Workbook(fileName);
  const cells = loaded.Worksheets.get(0).Cells;
  AssertEx.equal(CellValueType.IsString, cells.get("A1").Type);
  AssertEx.equal(CellValueType.IsNumeric, cells.get("B1").Type);
  AssertEx.equal(CellValueType.IsBool, cells.get("C1").Type);
  AssertEx.equal(CellValueType.IsNumeric, cells.get("D1").Type);
  AssertEx.equal(CellValueType.IsNumeric, cells.get("E1").Type);
  AssertEx.equal(CellValueType.IsDateTime, cells.get("F1").Type);
  AssertEx.equal(CellValueType.IsUnknown, cells.get("G1").Type);
}

export function valuePropertySetterMatchesSupportedScalarBehavior(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets.get(0);
  sheet.Cells.get("A1").Value = "alpha";
  sheet.Cells.get("B1").Value = 12;
  sheet.Cells.get("C1").Value = true;
  sheet.Cells.get("D1").Value = new Date(2024, 0, 2, 3, 4, 0);
  sheet.Cells.get("E1").Value = null;

  AssertEx.equal("alpha", sheet.Cells.get("A1").Value as string);
  AssertEx.equal(12, Number(sheet.Cells.get("B1").Value));
  AssertEx.equal(true, sheet.Cells.get("C1").Value as boolean);
  AssertEx.equal(CellValueType.IsDateTime, sheet.Cells.get("D1").Type);
  AssertEx.equal("", sheet.Cells.get("E1").DisplayStringValue);
}

export function richTextPublicApiRoundtripPreservesRuns(): void {
  const temp = new TemporaryDirectory("compat-rich-text");
  const fileName = temp.getPath("rich-text.xlsx");

  const workbook = new Workbook();
  const cell = workbook.Worksheets.get(0).Cells.get("A1");
  cell.PutValue("Hello World!");

  const first = cell.Characters(0, 5);
  first.Font.Color = Color.FromArgb(255, 48, 192, 180);

  const second = cell.Characters(5, 1);
  second.Font.Color = Color.FromArgb(255, 0, 0, 0);

  const third = cell.Characters(6, 5);
  third.Font.DoubleSize = 22;
  third.Font.Color = Color.FromArgb(255, 0, 0, 0);

  cell.SetCharacters([first, second, third]);

  AssertEx.true(cell.IsRichText(), "Cell should report rich text after SetCharacters.");
  AssertEx.equal(3, cell.GetCharacters().length);

  workbook.Save(fileName);

  const reloaded = new Workbook(fileName);
  const loadedCell = reloaded.Worksheets.get(0).Cells.get("A1");
  AssertEx.true(loadedCell.IsRichText(), "Loaded cell should keep rich text runs.");

  const runs = loadedCell.GetCharacters();
  AssertEx.equal(3, runs.length);
  AssertEx.equal(0, runs[0].StartIndex);
  AssertEx.equal(5, runs[0].Length);
  AssertEx.equal(Color.FromArgb(255, 48, 192, 180), runs[0].Font.Color);
  AssertEx.equal(5, runs[1].StartIndex);
  AssertEx.equal(1, runs[1].Length);
  AssertEx.equal(6, runs[2].StartIndex);
  AssertEx.equal(5, runs[2].Length);
  AssertEx.equal(22, runs[2].Font.DoubleSize);
}

export function compatibilityMembersFollowAsposeStyle(): void {
  const workbook = new Workbook();
  workbook.Settings.Date1904 = true;
  AssertEx.true(workbook.Settings.Date1904);

  const index = workbook.Worksheets.Add();
  workbook.Worksheets.get(index).Name = "Report";
  workbook.Worksheets.ActiveSheetName = "Report";
  AssertEx.equal(index, workbook.Worksheets.ActiveSheetIndex);

  const sheet = workbook.Worksheets.get(index);
  sheet.VisibilityType = VisibilityType.Hidden;
  AssertEx.equal(VisibilityType.Hidden, sheet.VisibilityType);

  sheet.Cells.get("A1").Value = "Docs";
  const hyperlinkIndex = sheet.Hyperlinks.Add(0, 0, 1, 1, "https://example.com/docs");
  const hyperlink = sheet.Hyperlinks.get(hyperlinkIndex);
  AssertEx.equal(TargetModeType.External, hyperlink.LinkType);
  hyperlink.Delete();
  AssertEx.equal(0, sheet.Hyperlinks.Count);

  const validationIndex = sheet.Validations.Add(CellArea.CreateCellArea("B2", "B2"));
  const validation = sheet.Validations.get(validationIndex);
  validation.Type = ValidationType.List;
  validation.Formula1 = "\"Yes,No\"";
  const inCell = sheet.Validations.GetValidationInCell(1, 1);
  AssertEx.notNull(inCell, "Validation must cover B2.");
  AssertEx.equal(ValidationType.List, inCell!.Type);

  const conditionalFormattingIndex = sheet.ConditionalFormattings.Add();
  const conditionalFormatting = sheet.ConditionalFormattings.get(conditionalFormattingIndex);
  conditionalFormatting.AddArea(CellArea.CreateCellArea("C3", "C3"));
  conditionalFormatting.AddCondition(FormatConditionType.Expression, OperatorType.None, "C3>0", "");
  AssertEx.equal(1, sheet.ConditionalFormattings.Count);

  const pageSetup = sheet.PageSetup;
  pageSetup.LeftMargin = 1.27;
  AssertEx.equal(0.5, pageSetup.LeftMarginInch);
}

function hasMember(instance: object, name: string): boolean {
  let current: object | null = instance;
  while (current !== null) {
    if (Object.prototype.hasOwnProperty.call(current, name)) {
      return true;
    }

    current = Object.getPrototypeOf(current) as object | null;
  }

  return false;
}

export function cellAreaPublicMembersFollowAsposeStyle(): void {
  const area = CellArea.CreateCellArea(0, 0, 1, 1);

  AssertEx.true(hasMember(area, "StartRow"), "CellArea must expose StartRow.");
  AssertEx.true(hasMember(area, "EndRow"), "CellArea must expose EndRow.");
  AssertEx.true(hasMember(area, "StartColumn"), "CellArea must expose StartColumn.");
  AssertEx.true(hasMember(area, "EndColumn"), "CellArea must expose EndColumn.");

  const fromReferences = CellArea.CreateCellArea("B2", "C3");
  AssertEx.equal(1, fromReferences.StartRow);
  AssertEx.equal(1, fromReferences.StartColumn);
  AssertEx.equal(2, fromReferences.EndRow);
  AssertEx.equal(2, fromReferences.EndColumn);

  const fromNumbers = CellArea.CreateCellArea(0, 0, 2, 2);
  AssertEx.equal(0, fromNumbers.StartRow);
  AssertEx.equal(0, fromNumbers.StartColumn);
  AssertEx.equal(2, fromNumbers.EndRow);
  AssertEx.equal(2, fromNumbers.EndColumn);
}

export function worksheetViewMembersFollowSupportedPatterns(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets.get(0);

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
}

export function worksheetProtectionMembersFollowSupportedPatterns(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets.get(0);

  sheet.Protect();
  sheet.Protection.Objects = true;
  sheet.Protection.Scenarios = true;
  sheet.Protection.AutoFilter = true;
  sheet.Protection.SelectLockedCells = true;
  sheet.Protection.SelectUnlockedCells = true;

  AssertEx.true(sheet.Protection.IsProtected);
  AssertEx.true(sheet.Protection.Objects);
  AssertEx.true(sheet.Protection.Scenarios);
  AssertEx.true(sheet.Protection.AutoFilter);
  AssertEx.true(sheet.Protection.SelectLockedCells);
  AssertEx.true(sheet.Protection.SelectUnlockedCells);
}

export function autofilterMembersFollowSupportedPatterns(): void {
  const workbook = AutoFilterScenarioFactory.CreateAutoFilterWorkbook();
  AutoFilterScenarioFactory.AssertAutoFilter(workbook);
}

export function worksheetSerializationUsesSchemaChildOrder(): void {
  const temp = new TemporaryDirectory("compat-worksheet-order");
  const fileName = temp.getPath("worksheet-order.xlsx");

  const workbook = new Workbook();
  const sheet = workbook.Worksheets.get(0);
  sheet.Name = "Ordered";
  sheet.Protect();
  sheet.Protection.AutoFilter = true;
  sheet.Cells.get("A1").PutValue("Header");
  sheet.Cells.get("A2").PutValue("Value");
  sheet.AutoFilter.Range = "A1:B2";
  sheet.Cells.Merge(0, 0, 1, 2);
  sheet.Hyperlinks.Add("A2", 1, 1, "https://example.com");

  const validationIndex = sheet.Validations.Add(CellArea.CreateCellArea("B2", "B2"));
  const validation = sheet.Validations.get(validationIndex);
  validation.Type = ValidationType.List;
  validation.Formula1 = "\"Yes,No\"";

  const formattingIndex = sheet.ConditionalFormattings.Add();
  const formatting = sheet.ConditionalFormattings.get(formattingIndex);
  formatting.AddArea(CellArea.CreateCellArea("A2", "A2"));
  formatting.AddCondition(FormatConditionType.Expression, OperatorType.None, "A2<>\"\"", "");

  workbook.Save(fileName);

  const worksheetXml = ZipPackageHelper.readEntryText(fileName, "xl/worksheets/sheet1.xml");
  const root = documentRoot(parseXml(worksheetXml));
  const childNames = childLocalNames(root);

  assertElementOrder(childNames, "sheetData", "sheetProtection");
  assertElementOrder(childNames, "sheetProtection", "autoFilter");
  assertElementOrder(childNames, "autoFilter", "mergeCells");
  assertElementOrder(childNames, "mergeCells", "conditionalFormatting");
  assertElementOrder(childNames, "conditionalFormatting", "dataValidations");
  assertElementOrder(childNames, "dataValidations", "hyperlinks");
}

export function packageContentTypesUseWorkbookOverride(): void {
  const temp = new TemporaryDirectory("compat-content-types");
  const fileName = temp.getPath("content-types.xlsx");

  const workbook = WorksheetScenarioFactory.CreateWorksheetSettingsWorkbook();
  workbook.Save(fileName);

  const contentTypesXml = ZipPackageHelper.readEntryText(fileName, "[Content_Types].xml");
  AssertEx.contains(
    'Extension="xml" ContentType="application/xml"',
    contentTypesXml,
    "Generic .xml parts must use the application/xml default content type.",
  );
  AssertEx.contains(
    'PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"',
    contentTypesXml,
    "The workbook part must use an explicit workbook content-type override.",
  );
}

export function fileAndStreamRoundtripPreserveAutofilter(): void {
  const temp = new TemporaryDirectory("compat-autofilter");
  const fileName = temp.getPath("autofilter.xlsx");
  const workbook = AutoFilterScenarioFactory.CreateAutoFilterWorkbook();
  workbook.Save(fileName);

  const stream = new MemoryStream();
  workbook.Save(stream, SaveFormat.Xlsx);
  stream.position = 0;

  const fromFile = new Workbook(fileName);
  const fromStream = new Workbook(stream);

  AutoFilterScenarioFactory.AssertAutoFilter(fromFile);
  AutoFilterScenarioFactory.AssertAutoFilter(fromStream);
}

export function fileAndStreamRoundtripPreserveDefinedNames(): void {
  const temp = new TemporaryDirectory("compat-defined-names");
  const fileName = temp.getPath("defined-names.xlsx");
  const workbook = DefinedNameScenarioFactory.CreateDefinedNamesWorkbook();
  workbook.Save(fileName);

  const stream = new MemoryStream();
  workbook.Save(stream, SaveFormat.Xlsx);
  stream.position = 0;

  const fromFile = new Workbook(fileName);
  const fromStream = new Workbook(stream);

  DefinedNameScenarioFactory.AssertDefinedNames(fromFile);
  DefinedNameScenarioFactory.AssertDefinedNames(fromStream);
}

export function fileAndStreamRoundtripPreserveStyles(): void {
  const temp = new TemporaryDirectory("compat-styles");
  const fileName = temp.getPath("styled-file.xlsx");
  const workbook = StyleScenarioFactory.CreateStyledWorkbook();
  workbook.Save(fileName);

  const stream = new MemoryStream();
  workbook.Save(stream, SaveFormat.Xlsx);
  stream.position = 0;

  const fromFile = new Workbook(fileName);
  const fromStream = new Workbook(stream);

  StyleScenarioFactory.AssertPrimaryStyle(fromFile.Worksheets.get(0).Cells.get("A1").GetStyle());
  StyleScenarioFactory.AssertPrimaryStyle(fromStream.Worksheets.get(0).Cells.get("A1").GetStyle());
  StyleScenarioFactory.AssertCustomNumberStyle(fromFile.Worksheets.get(0).Cells.get("B2").GetStyle());
  StyleScenarioFactory.AssertCustomNumberStyle(fromStream.Worksheets.get(0).Cells.get("B2").GetStyle());
  AssertEx.equal(CellValueType.IsNull, fromFile.Worksheets.get(0).Cells.get("B2").Type);
  AssertEx.equal(CellValueType.IsNull, fromStream.Worksheets.get(0).Cells.get("B2").Type);
}

export function fileAndStreamRoundtripPreserveWorksheetSettings(): void {
  const temp = new TemporaryDirectory("compat-worksheet-settings");
  const fileName = temp.getPath("worksheet-settings.xlsx");
  const workbook = WorksheetScenarioFactory.CreateWorksheetSettingsWorkbook();
  workbook.Save(fileName);

  const stream = new MemoryStream();
  workbook.Save(stream, SaveFormat.Xlsx);
  stream.position = 0;

  const fromFile = new Workbook(fileName);
  const fromStream = new Workbook(stream);

  WorksheetScenarioFactory.AssertWorksheetSettings(fromFile);
  WorksheetScenarioFactory.AssertWorksheetSettingsScenarioHasVisibleSheet(fromFile);
  WorksheetScenarioFactory.AssertWorksheetSettings(fromStream);
  WorksheetScenarioFactory.AssertWorksheetSettingsScenarioHasVisibleSheet(fromStream);
}

export function fileAndStreamRoundtripPreserveDataValidations(): void {
  const temp = new TemporaryDirectory("compat-validations");
  const fileName = temp.getPath("validations.xlsx");
  const workbook = ValidationScenarioFactory.CreateValidationWorkbook();
  workbook.Save(fileName);

  const stream = new MemoryStream();
  workbook.Save(stream, SaveFormat.Xlsx);
  stream.position = 0;

  const fromFile = new Workbook(fileName);
  const fromStream = new Workbook(stream);

  ValidationScenarioFactory.AssertValidations(fromFile);
  ValidationScenarioFactory.AssertValidations(fromStream);
}

export function fileAndStreamRoundtripPreserveConditionalFormattings(): void {
  const temp = new TemporaryDirectory("compat-conditional-formatting");
  const fileName = temp.getPath("conditional-formatting.xlsx");
  const workbook = ConditionalFormattingScenarioFactory.CreateConditionalFormattingWorkbook();
  workbook.Save(fileName);

  const stream = new MemoryStream();
  workbook.Save(stream, SaveFormat.Xlsx);
  stream.position = 0;

  const fromFile = new Workbook(fileName);
  const fromStream = new Workbook(stream);

  ConditionalFormattingScenarioFactory.AssertConditionalFormattings(fromFile);
  ConditionalFormattingScenarioFactory.AssertConditionalFormattings(fromStream);
}

export function fileAndStreamRoundtripPreserveAdvancedConditionalFormattings(): void {
  const temp = new TemporaryDirectory("compat-advanced-conditional-formatting");
  const fileName = temp.getPath("advanced-conditional-formatting.xlsx");
  const workbook = ConditionalFormattingScenarioFactory.CreateAdvancedConditionalFormattingWorkbook();
  workbook.Save(fileName);

  const stream = new MemoryStream();
  workbook.Save(stream, SaveFormat.Xlsx);
  stream.position = 0;

  const fromFile = new Workbook(fileName);
  const fromStream = new Workbook(stream);

  ConditionalFormattingScenarioFactory.AssertAdvancedConditionalFormattings(fromFile);
  ConditionalFormattingScenarioFactory.AssertAdvancedConditionalFormattings(fromStream);
}

export function fileAndStreamRoundtripPreservePageSetup(): void {
  const temp = new TemporaryDirectory("compat-page-setup");
  const fileName = temp.getPath("page-setup.xlsx");
  const workbook = PageSetupScenarioFactory.CreatePageSetupWorkbook();
  workbook.Save(fileName);

  const stream = new MemoryStream();
  workbook.Save(stream, SaveFormat.Xlsx);
  stream.position = 0;

  const fromFile = new Workbook(fileName);
  const fromStream = new Workbook(stream);

  PageSetupScenarioFactory.AssertPageSetup(fromFile);
  PageSetupScenarioFactory.AssertPageSetup(fromStream);
}

export function workbookMetadataMembersFollowSupportedPatterns(): void {
  const workbook = WorkbookMetadataScenarioFactory.CreateWorkbookMetadataWorkbook();
  WorkbookMetadataScenarioFactory.AssertWorkbookMetadata(workbook);
  AssertEx.equal("Data", workbook.Worksheets.get(workbook.Properties.View.ActiveTab).Name);
}

export function fileAndStreamRoundtripPreserveWorkbookMetadata(): void {
  const temp = new TemporaryDirectory("compat-workbook-metadata");
  const fileName = temp.getPath("workbook-metadata.xlsx");
  const workbook = WorkbookMetadataScenarioFactory.CreateWorkbookMetadataWorkbook();
  workbook.Save(fileName);

  const stream = new MemoryStream();
  workbook.Save(stream, SaveFormat.Xlsx);
  stream.position = 0;

  const fromFile = new Workbook(fileName);
  const fromStream = new Workbook(stream);

  WorkbookMetadataScenarioFactory.AssertWorkbookMetadata(fromFile);
  WorkbookMetadataScenarioFactory.AssertWorkbookMetadata(fromStream);
}

export function commentProgrammaticRoundtripPreservesCount(): void {
  const outputDir = getOrCreateOutputDirectory("comment-roundtrips");
  const outputPath = path.join(outputDir, "comment-count-out.xlsx");

  const workbook = new Workbook();
  const sheet = workbook.Worksheets.get(0);
  sheet.Comments.Add(0, 0).Note = "First";
  sheet.Comments.Add(1, 1).Note = "Second";
  sheet.Comments.Add(2, 0).Note = "Third";
  workbook.Save(outputPath);

  const reloaded = new Workbook(outputPath);
  AssertEx.equal(
    3,
    reloaded.Worksheets.get(0).Comments.Count,
    "Comment count must be preserved after save/load roundtrip",
  );
}

export function commentProgrammaticRoundtripPreservesNoteAndAuthor(): void {
  const outputDir = getOrCreateOutputDirectory("comment-roundtrips");
  const outputPath = path.join(outputDir, "comment-content-out.xlsx");

  const workbook = new Workbook();
  const sheet = workbook.Worksheets.get(0);

  const comment = sheet.Comments.Add(3, 5);
  comment.Note = "Important: review before merge";
  comment.Author = "Carol";
  comment.IsVisible = true;
  comment.Width = 180;
  comment.Height = 90;
  workbook.Save(outputPath);

  const reloaded = new Workbook(outputPath);
  const loadedSheet = reloaded.Worksheets.get(0);

  AssertEx.equal(1, loadedSheet.Comments.Count);
  const loaded = loadedSheet.Comments.get("F4");
  AssertEx.notNull(loaded, "Comment at F4 must survive roundtrip");
  AssertEx.equal("Important: review before merge", loaded!.Note);
  AssertEx.equal("Carol", loaded!.Author);
  AssertEx.true(loaded!.IsVisible, "IsVisible must be preserved");
  AssertEx.equal(180, loaded!.Width, "Width must be preserved");
  AssertEx.equal(90, loaded!.Height, "Height must be preserved");
}
