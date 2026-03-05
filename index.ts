//import { testhtmlExport } from "./examples/export"
import { Workbook } from "./aspose_cells";
console.log("Testing  export...");

const workbook = new Workbook();
const worksheet = workbook.worksheet;

worksheet.putValue("A1", "Name");
worksheet.putValue("B1", "Age");
worksheet.putValue("A2", "Alice");
worksheet.putValue("B2", 25);
worksheet.putValue("A3", "Bob");
worksheet.putValue("B3", 30);

workbook.save("outputfiles/test_export.xlsx");
console.log("Saved to outputfiles/test_export.xlsx");
