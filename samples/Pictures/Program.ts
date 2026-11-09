import { ImageType, Workbook } from "../../aspose_cells";
import { sampleOutputPath } from "../Shared";

function createSamplePngImage(): Uint8Array {
  const chunks: Buffer[] = [];
  const pngHeader = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  chunks.push(pngHeader);

  const ihdr = Buffer.from([0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00]);
  chunks.push(writeChunk("IHDR", ihdr));

  const idat = Buffer.from([0x78, 0x9c, 0x62, 0x00, 0x02, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0d]);
  chunks.push(writeChunk("IDAT", idat));

  chunks.push(writeChunk("IEND", Buffer.alloc(0)));

  return new Uint8Array(Buffer.concat(chunks));
}

function writeChunk(type: string, data: Buffer): Buffer {
  const lengthBytes = Buffer.alloc(4);
  lengthBytes.writeUInt32BE(data.length, 0);

  const typeBytes = Buffer.from(type, "ascii");

  const crcBytes = Buffer.alloc(4);
  crcBytes.writeUInt32BE(calculateCrc(typeBytes, data), 0);

  return Buffer.concat([lengthBytes, typeBytes, data, crcBytes]);
}

function calculateCrc(typeBytes: Buffer, data: Buffer): number {
  let crc = 0xffffffff;
  const buffer = Buffer.concat([typeBytes, data]);

  for (let i = 0; i < buffer.length; i++) {
    crc ^= buffer[i];
    for (let j = 0; j < 8; j++) {
      if ((crc & 1) !== 0) {
        crc = (crc >>> 1) ^ 0xedb88320;
      } else {
        crc = crc >>> 1;
      }
    }
  }

  return (~crc) >>> 0;
}

const outputPath = sampleOutputPath("pictures-sample.xlsx");

const workbook = new Workbook();
const sheet = workbook.Worksheets[0];
sheet.Name = "Pictures";

sheet.Cells["A1"].PutValue("Product");
sheet.Cells["B1"].PutValue("Description");
sheet.Cells["A2"].PutValue("Item 1");
sheet.Cells["B2"].PutValue("Sample product with logo");
sheet.Cells["A4"].PutValue("Item 2");
sheet.Cells["B4"].PutValue("Another product");

const sampleImageData = createSamplePngImage();

const picture1Index = sheet.Pictures.Add(1, 2, 3, 3, sampleImageData);
const picture1 = sheet.Pictures[picture1Index];
picture1.Name = "Product Logo 1";

const picture2Index = sheet.Pictures.Add(3, 2, 5, 3, sampleImageData);
const picture2 = sheet.Pictures[picture2Index];
picture2.Name = "Product Logo 2";
picture2.UpperLeftColumn = 4;
picture2.LowerRightColumn = 5;

workbook.Save(outputPath);

const loaded = new Workbook(outputPath);
const loadedSheet = loaded.Worksheets["Pictures"];

console.log("Saved: " + outputPath);
console.log("Picture count: " + loadedSheet.Pictures.Count);
console.log("First picture: " + loadedSheet.Pictures[0].Name + " (Type: " + ImageType[loadedSheet.Pictures[0].ImageType] + ")");
console.log("First picture anchor: R" + loadedSheet.Pictures[0].UpperLeftRow + "C" + loadedSheet.Pictures[0].UpperLeftColumn + " to R" + loadedSheet.Pictures[0].LowerRightRow + "C" + loadedSheet.Pictures[0].LowerRightColumn);
console.log("Second picture: " + loadedSheet.Pictures[1].Name + " (Type: " + ImageType[loadedSheet.Pictures[1].ImageType] + ")");
