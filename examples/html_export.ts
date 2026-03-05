import {
  Workbook,
  HtmlDocument,
  workbookToHtml,
  worksheetToHtml,
} from "../aspose_cells/html";

export async function testHtmlExport() {
  console.log("Testing HTML export...");

  const workbook = new Workbook();
  const worksheet = workbook.worksheet;

  worksheet.putValue("A1", "Name");
  worksheet.putValue("B1", "Age");
  worksheet.putValue("C1", "City");
  worksheet.putValue("A2", "Alice");
  worksheet.putValue("B2", 25);
  worksheet.putValue("C2", "New York");
  worksheet.putValue("A3", "Bob");
  worksheet.putValue("B3", 30);
  worksheet.putValue("C3", "London");

  const html = worksheetToHtml(worksheet);
  console.log("HTML:\n", html);

  await workbook.save("outputfiles/test_html_export.xlsx");
  console.log("Saved to outputfiles/test_html_export.xlsx");
}

export async function testHtmlImport() {
  console.log("Testing HTML import...");

  const html = `
<table>
  <thead>
    <tr><th>Name</th><th>Age</th></tr>
  </thead>
  <tbody>
    <tr><td>Alice</td><td>25</td></tr>
    <tr><td>Bob</td><td>30</td></tr>
  </tbody>
</table>
`;

  const doc = HtmlDocument.parse(html);
  console.log("Tables:", doc.tables.length);
  console.log("Rows:", doc.tables[0]?.rows);

  const workbook = doc.toWorkbook();
  console.log("Worksheets:", workbook.worksheetCount);
  console.log("A1:", workbook.worksheet.getCell(0, 0)?.value);
}

// if (import.meta.main) {
//   await testHtmlExport()
//   await testHtmlImport()
// }
