import { Workbook } from "../aspose_cells";

export function testhtmlExport() {
  console.log("Testing html export...");

  const workbook = new Workbook();
  const worksheet = workbook.worksheet;

  worksheet.putValue("A1", "Name");
  worksheet.putValue("B1", "Age");
  worksheet.putValue("A2", "Alice");
  worksheet.putValue("B2", 25);
  worksheet.putValue("A3", "Bob");
  worksheet.putValue("B3", 30);

  workbook.save("outputfiles/test_html_export.html");
  console.log("Saved to outputfiles/test_html_export.html");
}

export async function testhtmlImport() {
  console.log("Testing html import...");

  const workbook = await Workbook.load("outputfiles/test_html_export.html");
  const worksheet = workbook.worksheet;

  console.log("A1:", worksheet.getCell(0, 0)?.value);
  console.log("B1:", worksheet.getCell(0, 1)?.value);
  console.log("A2:", worksheet.getCell(1, 0)?.value);
}

export async function testJsonExport() {
  console.log("Testing JSON export...");

  const workbook = new Workbook();
  const worksheet = workbook.worksheet;

  worksheet.putValue("A1", "Name");
  worksheet.putValue("B1", "Age");
  worksheet.putValue("A2", "Alice");
  worksheet.putValue("B2", 25);
  worksheet.putValue("A3", "Bob");
  worksheet.putValue("B3", 30);

  const json = workbook.toJson();
  console.log("JSON:", json);

  await workbook.save("outputfiles/test_json_export.json");
  console.log("Saved to outputfiles/test_json_export.json");
}

export async function testMarkdownExport() {
  console.log("Testing Markdown export...");

  const workbook = new Workbook();
  const worksheet = workbook.worksheet;

  worksheet.putValue("A1", "Name");
  worksheet.putValue("B1", "Age");
  worksheet.putValue("A2", "Alice");
  worksheet.putValue("B2", 25);
  worksheet.putValue("A3", "Bob");
  worksheet.putValue("B3", 30);

  const markdown = workbook.toMarkdown();
  console.log("Markdown:\n", markdown);

  await workbook.save("outputfiles/test_markdown_export.md");
  console.log("Saved to outputfiles/test_markdown_export.md");
}

// if (import.meta.main) {
//   await testhtmlExport()
//   await testhtmlImport()
//   await testJsonExport()
//   await testMarkdownExport()
// }
