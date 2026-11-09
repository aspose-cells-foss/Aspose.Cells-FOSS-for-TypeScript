import { LegacyWorkbook, LegacyStyle } from "../aspose_cells";

export async function testFontSettings() {
  console.log("Testing font settings...");

  const workbook = new LegacyWorkbook();
  const worksheet = workbook.worksheets[0]!;
  const style = new LegacyStyle();

  style.setFontName("Arial");
  style.setFontSize(14);
  style.setBold(true);
  style.setItalic(true);
  style.setFontColor("FF0000");

  const cell = worksheet.getCell2("A1");
  cell.putValue("Styled Text");
  cell.setStyle(style);

  await workbook.save("outputfiles/test_font_settings.xlsx");
  console.log("Saved to outputfiles/test_font_settings.xlsx");
}

export async function testFillSettings() {
  console.log("Testing fill settings...");

  const workbook = new LegacyWorkbook();
  const worksheet = workbook.worksheets[0]!;
  const style = new LegacyStyle();

  style.setForegroundColor("FFFF00");
  style.setBackgroundColor("FF0000");

  const cell = worksheet.getCell2("A1");
  cell.putValue("Colored LegacyCell");
  cell.setStyle(style);

  await workbook.save("outputfiles/test_fill_settings.xlsx");
  console.log("Saved to outputfiles/test_fill_settings.xlsx");
}

export async function testBorderSettings() {
  console.log("Testing border settings...");

  const workbook = new LegacyWorkbook();
  const worksheet = workbook.worksheets[0]!;
  const style = new LegacyStyle();

  style.getBorder().left = { style: "thick", color: "FF0000" };
  style.getBorder().right = { style: "thick", color: "FF0000" };
  style.getBorder().top = { style: "thick", color: "FF0000" };
  style.getBorder().bottom = { style: "thick", color: "FF0000" };

  const cell = worksheet.getCell2("A1");
  cell.putValue("Bordered LegacyCell");
  cell.setStyle(style);

  await workbook.save("outputfiles/test_border_settings.xlsx");
  console.log("Saved to outputfiles/test_border_settings.xlsx");
}

export async function testAlignment() {
  console.log("Testing alignment...");

  const workbook = new LegacyWorkbook();
  const worksheet = workbook.worksheets[0]!;
  const style = new LegacyStyle();

  style.setHorizontalAlignment("center");
  style.setVerticalAlignment("center");
  style.setWrapText(true);

  const cell = worksheet.getCell2("A1");
  cell.putValue("Centered Text\nWith Wrap");
  cell.setStyle(style);

  worksheet.setColumnWidth(0, 20);
  worksheet.setRowHeight(0, 40);

  await workbook.save("outputfiles/test_alignment.xlsx");
  console.log("Saved to outputfiles/test_alignment.xlsx");
}

if (import.meta.main) {
  await testFontSettings();
  await testFillSettings();
  await testBorderSettings();
  await testAlignment();
}
