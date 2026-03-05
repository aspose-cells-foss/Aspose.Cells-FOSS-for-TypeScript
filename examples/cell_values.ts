import { Workbook, Cell, SaveFormat } from "../aspose_cells";

export async function testCellValues() {
  console.log("Testing cell values...");

  const workbook = new Workbook();
  const worksheet = workbook.worksheets[0];

  const intValues = [0, 1, -1, 42, 1000, -999];
  for (let i = 0; i < intValues.length; i++) {
    worksheet.putValue(`A${i + 1}`, intValues[i]);
  }

  const doubleValues = [0.0, 1.5, -2.7, 3.14159, 0.0001, -999.999];
  for (let i = 0; i < doubleValues.length; i++) {
    worksheet.putValue(`B${i + 1}`, doubleValues[i]);
  }

  const stringValues = [
    "Hello World",
    "Test String",
    "123",
    "3.14",
    "",
    "Special chars: !@#$%^&*()",
    "Unicode: 你好世界",
    "Multi\nline\nstring",
  ];
  for (let i = 0; i < stringValues.length; i++) {
    worksheet.putValue(`C${i + 1}`, stringValues[i]);
  }

  const formulas = [
    "=SUM(A1:A5)",
    "=A1+B1",
    '=IF(A1>0,"Positive","Non-positive")',
    "=VLOOKUP(A1,B1:C10,2,FALSE)",
    "=AVERAGE(A1:A10)",
    "=MAX(A1:A5)",
    "=MIN(A1:A5)",
    "=COUNT(A1:A10)",
  ];
  for (let i = 0; i < formulas.length; i++) {
    const cell = worksheet.getCell2(`D${i + 1}`);
    cell.setFormula(formulas[i]);
  }

  worksheet.putValue("A1", 42);
  worksheet.putValue("A2", 3.14159);
  worksheet.putValue("A3", "Hello");
  const cellA4 = worksheet.getCell2("A4");
  cellA4.setFormula("=SUM(A1:A2)");

  console.log("A1:", worksheet.getCell(0, 0)?.value);
  console.log("A2:", worksheet.getCell(1, 0)?.value);
  console.log("A3:", worksheet.getCell(2, 0)?.value);
  console.log("A4 formula:", worksheet.getCell(3, 0)?.formula);

  await workbook.save("outputfiles/test_cell_values.xlsx");
  console.log("Saved to outputfiles/test_cell_values.xlsx");
}

export async function testMixedValues() {
  const workbook = new Workbook();
  const worksheet = workbook.worksheets[0];

  worksheet.putValue("A1", 42);
  worksheet.putValue("A2", 3.14159);
  worksheet.putValue("A3", "Hello World");
  const cellA4 = worksheet.getCell2("A4");
  cellA4.setFormula("=SUM(A1:A2)");
  worksheet.putValue("A5", -100);
  worksheet.putValue("A6", 2.71828);
  worksheet.putValue("A7", "");
  const cellA8 = worksheet.getCell2("A8");
  cellA8.setFormula("=A1+A2");
  worksheet.putValue("A9", "Test String");
  worksheet.putValue("A10", 0);

  await workbook.save("outputfiles/test_mixed_values.xlsx");
  console.log("Saved to outputfiles/test_mixed_values.xlsx");

  const loaded = await Workbook.load("outputfiles/test_mixed_values.xlsx");
  const ws = loaded.worksheets[0];
  console.log("Loaded A1:", ws.getCell(0, 0)?.value);
  console.log("Loaded A3:", ws.getCell(2, 0)?.value);
  console.log("Loaded A4 formula:", ws.getCell(3, 0)?.formula);
}

export async function testEdgeCases() {
  const workbook = new Workbook();
  const worksheet = workbook.worksheets[0];

  const cell1 = worksheet.getCell2("A1");
  console.log("None value:", cell1.value);

  worksheet.putValue("A2", "");
  console.log("Empty string:", worksheet.getCell(1, 0)?.value);

  worksheet.putValue("A3", 999999999999);
  console.log("Large number:", worksheet.getCell(2, 0)?.value);

  worksheet.putValue("A4", 0.0000001);
  console.log("Small decimal:", worksheet.getCell(3, 0)?.value);

  worksheet.putValue("A5", 1.23e-10);
  console.log("Scientific notation:", worksheet.getCell(4, 0)?.value);

  await workbook.save("outputfiles/test_edge_cases.xlsx");
  console.log("Saved to outputfiles/test_edge_cases.xlsx");
}

if (import.meta.main) {
  await testCellValues();
  await testMixedValues();
  await testEdgeCases();
}
