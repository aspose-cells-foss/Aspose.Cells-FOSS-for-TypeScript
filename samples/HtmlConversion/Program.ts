import { HtmlSaveOptions } from "../../aspose_cells";
import { HtmlDocument, LegacyWorkbook } from "../../aspose_cells/html";
import { sampleInputPath, sampleOutputPath } from "../Shared";

async function main(): Promise<void> {
  const inputPath = sampleInputPath("sample.xlsx");
  const htmlPath = sampleOutputPath("html-conversion.html");
  const roundtripPath = sampleOutputPath("html-conversion-roundtrip.xlsx");

  const workbook = await LegacyWorkbook.load(inputPath);
  const options = new HtmlSaveOptions();
  options.ExportImagesAsBase64 = true;
  await workbook.save(htmlPath, options);

  console.log("XLSX input: " + inputPath);
  console.log("HTML output: " + htmlPath);

  const document = await HtmlDocument.load(htmlPath);
  await document.toExcel(roundtripPath);

  console.log("Roundtrip XLSX output: " + roundtripPath);
}

main().catch((exception: unknown) => {
  console.error(exception);
  process.exitCode = 1;
});
