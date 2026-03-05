import { Workbook, Style } from "../aspose_cells";

export async function testCellProtection() {
  console.log("Testing cell protection...");

  const workbook = new Workbook();
  const worksheet = workbook.worksheet;

  const style = new Style();
  style.setLocked(true);
  style.setHidden(false);

  const cell = worksheet.getCell2("A1");
  cell.putValue("Protected Cell");
  cell.setStyle(style);

  worksheet.putValue("A2", "Unprotected Cell");

  await workbook.save("outputfiles/test_cell_protection.xlsx");
  console.log("Saved to outputfiles/test_cell_protection.xlsx");
}

export async function testWorkbookProtection() {
  console.log("Testing workbook protection...");

  const workbook = new Workbook();
  workbook.protect(true, "password");

  await workbook.save("outputfiles/test_workbook_protection.xlsx");
  console.log("Saved to outputfiles/test_workbook_protection.xlsx");
}

if (import.meta.main) {
  await testCellProtection();
  await testWorkbookProtection();
}
