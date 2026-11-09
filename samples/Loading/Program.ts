import { LoadOptions, Workbook, WorkbookLoadException } from "../../aspose_cells";
import { sampleInputPath } from "../Shared";

const options = new LoadOptions();
options.TryRepairPackage = true;
options.TryRepairXml = true;

try {
  const inputPath = sampleInputPath("sample.xlsx");
  const workbook = new Workbook(inputPath, options);
  console.log("Loaded: " + inputPath);
} catch (exception) {
  if (exception instanceof WorkbookLoadException) {
    console.log(exception.message);
  } else {
    throw exception;
  }
}
