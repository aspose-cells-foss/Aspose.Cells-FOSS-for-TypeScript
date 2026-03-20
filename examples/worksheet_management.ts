import { Workbook, Worksheet, Cell, Style } from "../aspose_cells";

export async function testWorksheetManagement() {
  console.log("Testing worksheet management...");

  const workbook = new Workbook();

  const ws1 = workbook.worksheets.addWorksheet();
  console.log("Created worksheet:", ws1.name);

  const ws2 = workbook.worksheets.addWorksheet("CustomSheet1");
  console.log("Created worksheet:", ws2.name);

  console.log("Total worksheets:", workbook.worksheets.length);

  workbook.worksheets.removeWorksheet(1);
  console.log("After delete:", workbook.worksheets.length);

  const wsMain = workbook.worksheets.addWorksheet("MainSheet");
  wsMain.putValue("A1", "Main Worksheet");
  wsMain.putValue("A2", "This is the primary worksheet");

  const wsData = workbook.worksheets.addWorksheet("DataSheet");
  wsData.putValue("A1", "Data Worksheet");
  wsData.putValue("A2", "Contains data tables");

  const wsReport = workbook.worksheets.addWorksheet("ReportSheet");
  wsReport.putValue("A1", "Report Worksheet");
  wsReport.putValue("A2", "Contains reports");

  console.log(
    "Worksheets:",
    workbook.worksheets.worksheets.map((w) => w.name),
  );

  await workbook.save("outputfiles/test_worksheet_management.xlsx");
  console.log("Saved to outputfiles/test_worksheet_management.xlsx");

  const loaded = await Workbook.load(
    "outputfiles/test_worksheet_management.xlsx",
  );
  console.log(
    "Loaded worksheets:",
    loaded.worksheets.worksheets.map((w) => w.name),
  );
}

export async function testWorksheetRename() {
  const workbook = new Workbook();

  workbook.worksheets[0]!.name = "RenamedSheet";
  console.log("Renamed to:", workbook.worksheets[0]!.name);

  await workbook.save("outputfiles/test_worksheet_rename.xlsx");
  console.log("Saved to outputfiles/test_worksheet_rename.xlsx");
}

export async function testWorksheetCopy() {
  const workbook = new Workbook();

  const ws1 = workbook.worksheets[0]!;
  ws1.putValue("A1", "Original Data");
  ws1.putValue("A2", "Row 2");

  const ws2 = workbook.worksheets.addWorksheet("CopySheet");
  ws2.putValue("A1", "Copied Data");

  console.log(
    "Worksheets:",
    workbook.worksheets.worksheets.map((w) => w.name),
  );

  await workbook.save("outputfiles/test_worksheet_copy.xlsx");
  console.log("Saved to outputfiles/test_worksheet_copy.xlsx");
}

export async function testWorksheetAccess() {
  const workbook = new Workbook();

  workbook.worksheets.addWorksheet("FirstSheet");
  workbook.worksheets.addWorksheet("SecondSheet");
  workbook.worksheets.addWorksheet("ThirdSheet");

  console.log("Worksheet 0:", workbook.worksheets[0]?.name);
  console.log("Worksheet 1:", workbook.worksheets[1]?.name);
  console.log("Worksheet 2:", workbook.worksheets[2]?.name);

  await workbook.save("outputfiles/test_worksheet_access.xlsx");
}

if (import.meta.main) {
  await testWorksheetManagement();
  await testWorksheetRename();
  await testWorksheetCopy();
  await testWorksheetAccess();
}
