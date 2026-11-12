import { HtmlSaveOptions } from "../../aspose_cells";
import {
  HtmlDocument,
  LegacyWorkbook,
  workbookToHtml,
  worksheetToHtml,
} from "../../aspose_cells/html";
import { AssertEx } from "../Shared/TestInfrastructure";

export function htmlWorksheetExportEscapesTextAndPreservesBlankCells(): void {
  const workbook = new LegacyWorkbook();
  const sheet = workbook.worksheets.get(0)!;
  sheet.putValue("A1", "R&D <North>");
  sheet.putValue("C1", '"Quoted"');
  sheet.putValue("A2", 12.5);
  sheet.putValue("B2", true);

  const html = worksheetToHtml(sheet);

  AssertEx.contains("R&amp;D &lt;North&gt;", html);
  AssertEx.contains("&quot;Quoted&quot;", html);
  AssertEx.contains("<td></td>", html);
  AssertEx.contains("<td>12.5</td>", html);
  AssertEx.contains("<td>true</td>", html);
}

export function htmlWorkbookExportContainsEveryWorksheet(): void {
  const workbook = new LegacyWorkbook();
  workbook.worksheets.get(0)!.putValue("A1", "First sheet");
  workbook.worksheets.addWorksheet("Second").putValue("A1", "Second sheet");

  const html = workbookToHtml(workbook);

  AssertEx.equal(2, (html.match(/<table/g) ?? []).length);
  AssertEx.contains("First sheet", html);
  AssertEx.contains("Second sheet", html);
}

export function htmlSaveOptionsControlTableStyling(): void {
  const workbook = new LegacyWorkbook();
  const sheet = workbook.worksheets.get(0)!;
  sheet.putValue("A1", "Value");

  const defaultHtml = worksheetToHtml(sheet, new HtmlSaveOptions());
  AssertEx.contains('<table style="border-collapse: collapse">', defaultHtml);

  const withoutStyles = new HtmlSaveOptions();
  withoutStyles.includeStyles = false;
  AssertEx.true(worksheetToHtml(sheet, withoutStyles).startsWith("<table>"));

  const withoutCollapse = new HtmlSaveOptions();
  withoutCollapse.borderCollapse = false;
  AssertEx.true(worksheetToHtml(sheet, withoutCollapse).startsWith("<table>"));
}

export function htmlDocumentParseConvertsScalarValuesAndEntities(): void {
  const document = HtmlDocument.parse(`
    <table>
      <tr><td>A &amp; B</td><td>42</td><td>true</td><td>2024-05-06</td><td>&lt;North&gt;</td></tr>
      <tr><td>001</td><td>3.5</td><td>false</td><td></td><td>Text</td></tr>
    </table>
  `);
  const workbook = document.toWorkbook();
  const sheet = workbook.worksheets.get(0)!;

  AssertEx.equal(1, workbook.worksheets.length);
  AssertEx.equal("A & B", sheet.getCellByRef("A1")?.value);
  AssertEx.equal(42, sheet.getCellByRef("B1")?.value);
  AssertEx.equal(true, sheet.getCellByRef("C1")?.value);
  AssertEx.equal("<North>", sheet.getCellByRef("E1")?.value);
  AssertEx.equal("001", sheet.getCellByRef("A2")?.value);
  AssertEx.equal(3.5, sheet.getCellByRef("B2")?.value);
  AssertEx.equal(false, sheet.getCellByRef("C2")?.value);
  AssertEx.equal("Text", sheet.getCellByRef("E2")?.value);

  const dateValue = sheet.getCellByRef("D1")?.value;
  AssertEx.true(dateValue instanceof Date);
  const date = dateValue as Date;
  AssertEx.equal(2024, date.getFullYear());
  AssertEx.equal(4, date.getMonth());
  AssertEx.equal(6, date.getDate());
}

export function htmlDocumentParsePreservesDimensionsHiddenRowsAndColspan(): void {
  const document = HtmlDocument.parse(`
    <table>
      <colgroup>
        <col style="width:75pt">
        <col style="width:60pt">
        <col style="width:45pt">
      </colgroup>
      <tr style="height:24pt"><td colspan="2">Merged</td><td>End</td></tr>
      <tr style="display:none"><td>Hidden</td><td>10</td><td>20</td></tr>
    </table>
  `);
  const sheet = document.toWorkbook().worksheets.get(0)!;

  AssertEx.equal(100 / 7.5, sheet.getColumnWidth(0));
  AssertEx.equal(80 / 7.5, sheet.getColumnWidth(1));
  AssertEx.equal(60 / 7.5, sheet.getColumnWidth(2));
  AssertEx.equal(24, sheet.getRowHeight(0));
  AssertEx.true(sheet.isRowHidden(1));
  AssertEx.equal("Merged", sheet.getCellByRef("A1")?.value);
  AssertEx.equal("End", sheet.getCellByRef("C1")?.value);
  AssertEx.equal("Hidden", sheet.getCellByRef("A2")?.value);
  AssertEx.equal(10, sheet.getCellByRef("B2")?.value);
  AssertEx.equal(20, sheet.getCellByRef("C2")?.value);
}

export function htmlExportImportRoundtripPreservesScalarValues(): void {
  const workbook = new LegacyWorkbook();
  const sheet = workbook.worksheets.get(0)!;
  sheet.putValue("A1", "Alpha & Beta");
  sheet.putValue("B1", 7);
  sheet.putValue("C1", false);
  sheet.putValue("A2", "Second row");
  sheet.putValue("B2", -2.25);
  sheet.putValue("C2", true);

  const html = workbookToHtml(workbook);
  const loaded = HtmlDocument.parse(html).toWorkbook();
  const loadedSheet = loaded.worksheets.get(0)!;

  AssertEx.equal("Alpha & Beta", loadedSheet.getCellByRef("A1")?.value);
  AssertEx.equal(7, loadedSheet.getCellByRef("B1")?.value);
  AssertEx.equal(false, loadedSheet.getCellByRef("C1")?.value);
  AssertEx.equal("Second row", loadedSheet.getCellByRef("A2")?.value);
  AssertEx.equal(-2.25, loadedSheet.getCellByRef("B2")?.value);
  AssertEx.equal(true, loadedSheet.getCellByRef("C2")?.value);
}
