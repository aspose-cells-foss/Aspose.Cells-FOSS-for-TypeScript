import { Workbook, AutoFilter } from "../aspose_cells";

export async function testAutoFilter() {
  console.log("Testing auto filter...");

  const workbook = new Workbook();
  const worksheet = workbook.worksheet;

  worksheet.putValue("A1", "Name");
  worksheet.putValue("B1", "Age");
  worksheet.putValue("C1", "City");

  worksheet.putValue("A2", "Alice");
  worksheet.putValue("B2", "25");
  worksheet.putValue("C2", "New York");

  worksheet.putValue("A3", "Bob");
  worksheet.putValue("B3", "30");
  worksheet.putValue("C3", "London");

  worksheet.putValue("A4", "Charlie");
  worksheet.putValue("B4", "35");
  worksheet.putValue("C4", "Paris");

  worksheet.setAutoFilter("A1:C4");

  await workbook.save("outputfiles/test_auto_filter.xlsx");
  console.log("Saved to outputfiles/test_auto_filter.xlsx");
}

export async function testAutoFilterWithData() {
  const workbook = new Workbook();
  const worksheet = workbook.worksheet;

  for (let i = 0; i < 10; i++) {
    worksheet.putValue(`A${i + 1}`, `Item ${i + 1}`);
    worksheet.putValue(`B${i + 1}`, Math.floor(Math.random() * 100));
    worksheet.putValue(`C${i + 1}`, ["Red", "Green", "Blue"][i % 3]);
  }

  worksheet.setAutoFilter("A1:C10");

  await workbook.save("outputfiles/test_auto_filter_data.xlsx");
  console.log("Saved to outputfiles/test_auto_filter_data.xlsx");
}

if (import.meta.main) {
  await testAutoFilter();
  await testAutoFilterWithData();
}
