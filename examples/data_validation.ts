import { Workbook, Worksheet, DataValidation } from "../aspose_cells";

export async function testDataValidation() {
  console.log("Testing data validation...");

  const workbook = new Workbook();
  const worksheet = workbook.worksheet;

  worksheet.putValue("A1", "Option 1");
  worksheet.putValue("A2", "Option 2");
  worksheet.putValue("A3", "Option 3");

  const validation = new DataValidation();
  validation.type = "list";
  validation.formula1 = '"Option1,Option2,Option3"';
  worksheet.addDataValidation(validation, "B1:B10");

  await workbook.save("outputfiles/test_data_validation.xlsx");
  console.log("Saved to outputfiles/test_data_validation.xlsx");
}

export async function testNumberValidation() {
  const workbook = new Workbook();
  const worksheet = workbook.worksheet;

  worksheet.putValue("A1", 10);
  worksheet.putValue("A2", 20);
  worksheet.putValue("A3", 30);

  const validation = new DataValidation();
  validation.type = "whole";
  validation.operator = "between";
  validation.formula1 = "1";
  validation.formula2 = "100";
  worksheet.addDataValidation(validation, "B1:B5");

  await workbook.save("outputfiles/test_number_validation.xlsx");
  console.log("Saved to outputfiles/test_number_validation.xlsx");
}

if (import.meta.main) {
  await testDataValidation();
  await testNumberValidation();
}
