import { Color, FillPattern, FormatConditionType, OperatorType, Workbook } from "../../aspose_cells";
import { attribute, documentRoot, elementChildren } from "../../aspose_cells/Xml/XmlReader";
import { AssertEx, TemporaryDirectory, TestCase, TestRunner, ZipPackageHelper } from "../Shared/TestInfrastructure";
import { ConditionalFormattingScenarioFactory } from "../Shared/ScenarioFactories";
import { OpenXmlPackageBuilder, OpenXmlWorkbookReader } from "../Shared/OpenXmlPackage";

const STYLES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const STYLES_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

class ConditionalFormattingRuleSnapshot {
  public constructor(
    public readonly Type: string,
    public readonly Operator: string,
    public readonly DxfId: string,
    public readonly Priority: string,
    public readonly StopIfTrue: string,
    public readonly Formulas: readonly string[],
  ) {}
}

class ConditionalFormattingSnapshot {
  public constructor(
    public readonly Sqref: string,
    public readonly Rules: readonly ConditionalFormattingRuleSnapshot[],
  ) {}
}

function conditionalFormattingsMatchOpenXmlSdk(): void {
  const temp = new TemporaryDirectory("conditional-formatting-openxml-write");
  const path = temp.getPath("conditional-formatting.xlsx");

  const workbook = ConditionalFormattingScenarioFactory.CreateConditionalFormattingWorkbook();
  workbook.Save(path);

  const loaded = new Workbook(path);
  ConditionalFormattingScenarioFactory.AssertConditionalFormattings(loaded);

  const snapshot = readSnapshots(path, "Conditional Formatting");
  AssertEx.equal(2, snapshot.length);
  assertSnapshot(getSnapshot(snapshot, "A1:A5"), "A1:A5", 2);
  assertSnapshot(getSnapshot(snapshot, "C1:C3"), "C1:C3", 1);
}

function advancedConditionalFormattingsMatchOpenXmlSdk(): void {
  const temp = new TemporaryDirectory("conditional-formatting-openxml-advanced-write");
  const path = temp.getPath("advanced-conditional-formatting.xlsx");

  const workbook = ConditionalFormattingScenarioFactory.CreateAdvancedConditionalFormattingWorkbook();
  workbook.Save(path);

  const loaded = new Workbook(path);
  ConditionalFormattingScenarioFactory.AssertAdvancedConditionalFormattings(loaded);

  const worksheetXml = ZipPackageHelper.readEntryText(path, "xl/worksheets/sheet1.xml");
  AssertEx.contains('type="containsText"', worksheetXml);
  AssertEx.contains('text="error"', worksheetXml);
  AssertEx.contains('type="timePeriod"', worksheetXml);
  AssertEx.contains('timePeriod="today"', worksheetXml);
  AssertEx.contains('type="duplicateValues"', worksheetXml);
  AssertEx.contains('type="uniqueValues"', worksheetXml);
  AssertEx.contains('type="colorScale"', worksheetXml);
  AssertEx.contains('type="dataBar"', worksheetXml);
  AssertEx.contains('type="iconSet"', worksheetXml);
}

function openXmlCreatedConditionalFormattingsLoadInLibrary(): void {
  const temp = new TemporaryDirectory("conditional-formatting-openxml-read");
  const path = temp.getPath("openxml-conditional-formatting.xlsx");

  createOpenXmlWorkbook(path);

  const loaded = new Workbook(path);
  AssertEx.equal(1, loaded.Worksheets[0].ConditionalFormattings.Count);
  const collection = loaded.Worksheets[0].ConditionalFormattings.get(0);
  AssertEx.equal(2, collection.Count);
  AssertEx.equal(1, collection.RangeCount);
  AssertEx.equal(FormatConditionType.CellValue, collection.get(0).Type);
  AssertEx.equal(OperatorType.GreaterThan, collection.get(0).Operator);
  AssertEx.equal("10", collection.get(0).Formula1);
  AssertEx.equal(1, collection.get(0).Priority);
  AssertEx.true(collection.get(0).StopIfTrue);
  AssertEx.equal(FillPattern.Solid, collection.get(0).Style.Pattern);
  AssertEx.equal(Color.FromArgb(255, 255, 235, 156), collection.get(0).Style.ForegroundColor);
  AssertEx.equal(FormatConditionType.Expression, collection.get(1).Type);
  AssertEx.equal('A1="Y"', collection.get(1).Formula1);
  AssertEx.equal(2, collection.get(1).Priority);
}

function readSnapshots(workbookPath: string, sheetName: string): ConditionalFormattingSnapshot[] {
  const reader = new OpenXmlWorkbookReader(workbookPath);
  const worksheetXml = reader.sheetXmlByName(sheetName);
  const root = documentRoot(worksheetXml);
  const snapshots: ConditionalFormattingSnapshot[] = [];

  for (const formatting of elementChildren(root, "conditionalFormatting")) {
    const rules: ConditionalFormattingRuleSnapshot[] = [];
    for (const rule of elementChildren(formatting, "cfRule")) {
      rules.push(
        new ConditionalFormattingRuleSnapshot(
          attribute(rule, "type") ?? "",
          attribute(rule, "operator") ?? "",
          attribute(rule, "dxfId") ?? "",
          attribute(rule, "priority") ?? "",
          attribute(rule, "stopIfTrue") ?? "",
          elementChildren(rule, "formula").map((formula) => formula.textContent ?? ""),
        ),
      );
    }

    snapshots.push(new ConditionalFormattingSnapshot((attribute(formatting, "sqref") ?? "").toUpperCase(), rules));
  }

  return snapshots;
}

function getSnapshot(snapshots: readonly ConditionalFormattingSnapshot[], sqref: string): ConditionalFormattingSnapshot {
  const found = snapshots.filter((snapshot) => snapshot.Sqref === sqref);
  AssertEx.equal(1, found.length);
  return found[0];
}

function assertSnapshot(snapshot: ConditionalFormattingSnapshot, sqref: string, ruleCount: number): void {
  AssertEx.equal(sqref, snapshot.Sqref);
  AssertEx.equal(ruleCount, snapshot.Rules.length);
}

function createOpenXmlWorkbook(path: string): void {
  const worksheetXml =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<dimension ref="A1:A3"/>' +
    '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Y</t></is></c></row></sheetData>' +
    '<conditionalFormatting sqref="A1:A3">' +
    '<cfRule type="cellIs" operator="greaterThan" dxfId="0" priority="1" stopIfTrue="1"><formula>10</formula></cfRule>' +
    '<cfRule type="expression" priority="2"><formula>A1="Y"</formula></cfRule>' +
    "</conditionalFormatting>" +
    "</worksheet>";

  const stylesXml =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>" +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '<dxfs count="1"><dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFFEB9C"/></patternFill></fill></dxf></dxfs>' +
    "</styleSheet>";

  const builder = new OpenXmlPackageBuilder();
  builder.addWorksheet(worksheetXml, "rId1");
  builder.setWorkbookXml(
    '<?xml version="1.0" encoding="utf-8"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Conditional Formatting" sheetId="1" r:id="rId1"/></sheets>' +
      "</workbook>",
  );
  builder.addPart({
    partPath: "xl/styles.xml",
    xml: stylesXml,
    contentType: STYLES_CONTENT_TYPE,
    relationshipType: STYLES_RELATIONSHIP,
    relationshipId: "rId2",
  });
  builder.save(path);
}

TestRunner.runAndExit("ConditionalFormatting.OpenXmlTests", [
  new TestCase("conditional_formattings_match_openxml_sdk", conditionalFormattingsMatchOpenXmlSdk),
  new TestCase("advanced_conditional_formattings_match_openxml_sdk", advancedConditionalFormattingsMatchOpenXmlSdk),
  new TestCase(
    "openxml_created_conditional_formattings_load_in_library",
    openXmlCreatedConditionalFormattingsLoadInLibrary,
  ),
]);
