import { Workbook } from "../../aspose_cells";
import {
  attribute,
  documentRoot,
  elementChildren,
  firstElement,
  XmlNode,
} from "../../aspose_cells/Xml/XmlReader";
import { AssertEx, TemporaryDirectory, TestCase, TestRunner } from "../Shared/TestInfrastructure";
import { PageSetupScenarioFactory } from "../Shared/ScenarioFactories";
import { createOpenXmlWorkbook, OpenXmlWorkbookReader } from "../Shared/OpenXmlPackage";

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

function pageSetupMatchesOpenXmlSdk(): void {
  const temp = new TemporaryDirectory("page-setup-openxml-write");
  const path = temp.getPath("page-setup.xlsx");

  const workbook = PageSetupScenarioFactory.CreatePageSetupWorkbook();
  workbook.Save(path);

  const loaded = new Workbook(path);
  PageSetupScenarioFactory.AssertPageSetup(loaded);

  const snapshot = readPageSetupSnapshot(path, "Print Sheet");
  AssertEx.equal("$A$1:$C$10", snapshot.PrintArea);
  AssertEx.equal("$1:$2", snapshot.PrintTitleRows);
  AssertEx.equal("$A:$B", snapshot.PrintTitleColumns);
  AssertEx.equal(0.25, snapshot.LeftMargin ?? 0);
  AssertEx.equal(0.4, snapshot.RightMargin ?? 0);
  AssertEx.equal(0.5, snapshot.TopMargin ?? 0);
  AssertEx.equal(0.6, snapshot.BottomMargin ?? 0);
  AssertEx.equal(0.2, snapshot.HeaderMargin ?? 0);
  AssertEx.equal(0.22, snapshot.FooterMargin ?? 0);
  AssertEx.equal(9, snapshot.PaperSize ?? 0);
  AssertEx.equal("landscape", snapshot.Orientation);
  AssertEx.equal(3, snapshot.FirstPageNumber ?? 0);
  AssertEx.equal(95, snapshot.Scale ?? 0);
  AssertEx.equal(1, snapshot.FitToWidth ?? 0);
  AssertEx.equal(2, snapshot.FitToHeight ?? 0);
  AssertEx.true(snapshot.PrintGridLines);
  AssertEx.true(snapshot.PrintHeadings);
  AssertEx.true(snapshot.HorizontalCentered);
  AssertEx.true(snapshot.VerticalCentered);
  AssertEx.equal("&LLeft Header&CCenter Header&RRight Header", snapshot.OddHeader);
  AssertEx.equal("&LLeft Footer&CCenter Footer&RRight Footer", snapshot.OddFooter);
  AssertEx.equal(2, snapshot.RowBreaks.length);
  AssertEx.equal(4, snapshot.RowBreaks[0]);
  AssertEx.equal(7, snapshot.RowBreaks[1]);
  AssertEx.equal(1, snapshot.ColumnBreaks.length);
  AssertEx.equal(2, snapshot.ColumnBreaks[0]);
}

function openXmlCreatedPageSetupLoadsInLibrary(): void {
  const temp = new TemporaryDirectory("page-setup-openxml-read");
  const path = temp.getPath("openxml-page-setup.xlsx");

  createOpenXmlPageSetupWorkbook(path);

  const loaded = new Workbook(path);
  PageSetupScenarioFactory.AssertPageSetup(loaded);

  const snapshot = readPageSetupSnapshot(path, "Print Sheet");
  AssertEx.equal("$A$1:$C$10", snapshot.PrintArea);
  AssertEx.equal("$1:$2", snapshot.PrintTitleRows);
  AssertEx.equal("$A:$B", snapshot.PrintTitleColumns);
}

function readPageSetupSnapshot(workbookPath: string, sheetName: string): PageSetupSnapshot {
  const reader = new OpenXmlWorkbookReader(workbookPath);
  const workbookRoot = reader.workbookRoot;
  const sheets = elementChildren(firstElement(workbookRoot, "sheets"), "sheet");
  const sheetElement = sheets.filter((element) => (attribute(element, "name") ?? "") === sheetName)[0];
  if (sheetElement === undefined) {
    throw new Error(`Worksheet '${sheetName}' was not found.`);
  }

  const sheetId = Number.parseInt(attribute(sheetElement, "sheetId") ?? "1", 10) - 1;
  const worksheetXml = reader.sheetXmlByName(sheetName);
  const root = documentRoot(worksheetXml);
  if (root === null) {
    throw new Error("Worksheet XML is empty.");
  }

  const margins = firstElement(root, "pageMargins");
  const pageSetup = firstElement(root, "pageSetup");
  const printOptions = firstElement(root, "printOptions");
  const headerFooter = firstElement(root, "headerFooter");
  const rowBreaks = elementChildren(firstElement(root, "rowBreaks"), "brk").map((element) =>
    Number.parseInt(attribute(element, "id") ?? "0", 10),
  );
  const columnBreaks = elementChildren(firstElement(root, "colBreaks"), "brk").map((element) =>
    Number.parseInt(attribute(element, "id") ?? "0", 10),
  );

  let printArea: string | null = null;
  let printTitleRows: string | null = null;
  let printTitleColumns: string | null = null;
  for (const definedName of elementChildren(firstElement(workbookRoot, "definedNames"), "definedName")) {
    const localSheetId = attribute(definedName, "localSheetId");
    if (localSheetId !== String(sheetId)) {
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
          searchDigits(unqualified) &&
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
    tryParseFloat(attribute(margins, "left")),
    tryParseFloat(attribute(margins, "right")),
    tryParseFloat(attribute(margins, "top")),
    tryParseFloat(attribute(margins, "bottom")),
    tryParseFloat(attribute(margins, "header")),
    tryParseFloat(attribute(margins, "footer")),
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
    textOfElement(headerFooter, "oddHeader"),
    textOfElement(headerFooter, "oddFooter"),
    rowBreaks,
    columnBreaks,
  );
}

function searchDigits(value: string): boolean {
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

function textOfElement(parent: XmlNode | null, name: string): string {
  const child = firstElement(parent, name);
  return child === null ? "" : (child.textContent ?? "");
}

function createOpenXmlPageSetupWorkbook(path: string): void {
  createOpenXmlWorkbook(
    path,
    "Print Sheet",
    '<?xml version="1.0" encoding="utf-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:C10"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Title</t></is></c></row><row r="10"><c r="C10"><v>42</v></c></row></sheetData><printOptions headings="1" gridLines="1" gridLinesSet="1" horizontalCentered="1" verticalCentered="1"/><pageMargins left="0.25" right="0.4" top="0.5" bottom="0.6" header="0.2" footer="0.22"/><pageSetup paperSize="9" scale="95" fitToWidth="1" fitToHeight="2" firstPageNumber="3" useFirstPageNumber="1" orientation="landscape"/><headerFooter><oddHeader>&amp;LLeft Header&amp;CCenter Header&amp;RRight Header</oddHeader><oddFooter>&amp;LLeft Footer&amp;CCenter Footer&amp;RRight Footer</oddFooter></headerFooter><rowBreaks count="2" manualBreakCount="2"><brk id="4" max="16383" man="1"/><brk id="7" max="16383" man="1"/></rowBreaks><colBreaks count="1" manualBreakCount="1"><brk id="2" max="1048575" man="1"/></colBreaks></worksheet>',
    {
      workbookXml:
        '<?xml version="1.0" encoding="utf-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Print Sheet" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">\'Print Sheet\'!$A$1:$C$10</definedName><definedName name="_xlnm.Print_Titles" localSheetId="0">\'Print Sheet\'!$1:$2,\'Print Sheet\'!$A:$B</definedName></definedNames></workbook>',
    },
  );
}

function unqualifyDefinedNameValue(value: string): string {
  const trimmed = value.trim();
  const exclamationIndex = trimmed.lastIndexOf("!");
  return exclamationIndex >= 0 ? trimmed.substring(exclamationIndex + 1) : trimmed;
}

function tryParseInt(value: string | null): number | null {
  if (value === null) {
    return null;
  }

  const parsed = Number.parseInt(value, 10);
  return Number.isNaN(parsed) ? null : parsed;
}

function tryParseFloat(value: string | null): number | null {
  if (value === null) {
    return null;
  }

  const parsed = Number.parseFloat(value);
  return Number.isNaN(parsed) ? null : parsed;
}

function parseBool(value: string | null): boolean {
  return value === "1" || (value ?? "").toLowerCase() === "true";
}

TestRunner.runAndExit("PageSetup.OpenXmlTests", [
  new TestCase("page_setup_matches_openxml_sdk", pageSetupMatchesOpenXmlSdk),
  new TestCase("openxml_created_page_setup_loads_in_library", openXmlCreatedPageSetupLoadsInLibrary),
]);
