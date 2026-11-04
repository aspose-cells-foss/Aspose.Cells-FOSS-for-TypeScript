import { Workbook } from "./aspose_cells";
 

async function main() {
  const workbook = await   Workbook.load("sample_expect.html");
  await workbook.save("outputfiles/test_new2.pdf");
  await workbook.save("outputfiles/test_new2.xlsx");
  console.log("Done");
}

main();
