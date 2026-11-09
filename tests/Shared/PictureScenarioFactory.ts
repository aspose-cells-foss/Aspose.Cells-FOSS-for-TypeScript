import * as fs from "fs";
import * as path from "path";
import { ImageType, Workbook } from "../../aspose_cells";
import { AssertEx, resolveRepositoryRoot } from "./TestInfrastructure";

/** Scenario factory for picture workbooks. */
export class PictureScenarioFactory {
  public static LoadJpeg(): Uint8Array {
    return new Uint8Array(fs.readFileSync(PictureScenarioFactory.ResolveInputFile("pay.jpg")));
  }

  public static LoadPng(): Uint8Array {
    return new Uint8Array(fs.readFileSync(PictureScenarioFactory.ResolveInputFile("screen.png")));
  }

  public static CreatePictureWorkbook(): Workbook {
    const workbook = new Workbook();
    const sheet = workbook.Worksheets[0];
    sheet.Name = "Images";

    sheet.Cells["A1"].PutValue("Name");
    sheet.Cells["B1"].PutValue("Value");
    sheet.Cells["A2"].PutValue("Alpha");
    sheet.Cells["B2"].PutValue(42);

    const jpegIndex = sheet.Pictures.Add(0, 0, 8, 5, PictureScenarioFactory.LoadJpeg());
    sheet.Pictures.get(jpegIndex).Name = "PayImage";

    const pngIndex = sheet.Pictures.Add(10, 0, 16, 5, PictureScenarioFactory.LoadPng());
    sheet.Pictures.get(pngIndex).Name = "ScreenImage";

    return workbook;
  }

  public static AssertPictureWorkbook(workbook: Workbook): void {
    const sheet = workbook.Worksheets[0];
    AssertEx.equal(2, sheet.Pictures.Count);

    const jpeg = sheet.Pictures.get(0);
    AssertEx.equal("PayImage", jpeg.Name);
    AssertEx.equal(0, jpeg.UpperLeftRow);
    AssertEx.equal(0, jpeg.UpperLeftColumn);
    AssertEx.equal(8, jpeg.LowerRightRow);
    AssertEx.equal(5, jpeg.LowerRightColumn);
    AssertEx.equal(ImageType.Jpeg, jpeg.ImageType);
    AssertEx.equal(199797, jpeg.Data.length);

    const png = sheet.Pictures.get(1);
    AssertEx.equal("ScreenImage", png.Name);
    AssertEx.equal(10, png.UpperLeftRow);
    AssertEx.equal(0, png.UpperLeftColumn);
    AssertEx.equal(16, png.LowerRightRow);
    AssertEx.equal(5, png.LowerRightColumn);
    AssertEx.equal(ImageType.Png, png.ImageType);
    AssertEx.equal(30171, png.Data.length);
  }

  private static ResolveInputFile(fileName: string): string {
    return path.join(resolveRepositoryRoot(), "Input", fileName);
  }
}
