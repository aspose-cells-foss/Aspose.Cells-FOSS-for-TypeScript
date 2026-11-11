import { readFileSync } from "fs";
import { CellsException } from "./CellsException";
import { Picture } from "./Picture";
import { PictureModel } from "./Core/PictureModel";

// Implemented by serializer module.
/** Represents a collection of pictures anchored to a worksheet. */
export class PictureCollection {
  private readonly models: PictureModel[];

  public constructor(models: PictureModel[]) {
    this.models = models;
  }

  public get Count(): number {
    return this.models.length;
  }

  public get(index: number): Picture {
    if (index < 0 || index >= this.models.length) {
      throw new CellsException(`Picture index ${index} is out of range.`);
    }

    return new Picture(this.models[index]);
  }

  public at(index: number): Picture {
    return this.get(index);
  }

  public Add(
    upperLeftRow: number,
    upperLeftColumn: number,
    lowerRightRow: number,
    lowerRightColumn: number,
    imageData: Uint8Array,
  ): number;
  public Add(
    upperLeftRow: number,
    upperLeftColumn: number,
    lowerRightRow: number,
    lowerRightColumn: number,
    stream: { toBuffer(): Buffer },
  ): number;
  public Add(
    upperLeftRow: number,
    upperLeftColumn: number,
    lowerRightRow: number,
    lowerRightColumn: number,
    fileName: string,
  ): number;
  public Add(
    upperLeftRow: number,
    upperLeftColumn: number,
    lowerRightRow: number,
    lowerRightColumn: number,
    source: Uint8Array | { toBuffer(): Buffer } | string,
  ): number {
    if (upperLeftRow < 0 || upperLeftColumn < 0) {
      throw new CellsException("Picture anchor indices must be non-negative.");
    }

    if (lowerRightRow < upperLeftRow || lowerRightColumn < upperLeftColumn) {
      throw new CellsException("Picture anchor end must not precede the start.");
    }

    const data =
      typeof source === "string"
        ? new Uint8Array(readFileSync(source))
        : source instanceof Uint8Array
          ? source
          : new Uint8Array(source.toBuffer());

    if (data.length === 0) {
      throw new CellsException("Image data must be non-null and non-empty.");
    }

    const model = new PictureModel();
    model.Name = `Picture ${this.models.length + 1}`;
    model.UpperLeftRow = upperLeftRow;
    model.UpperLeftColumn = upperLeftColumn;
    model.LowerRightRow = lowerRightRow;
    model.LowerRightColumn = lowerRightColumn;
    model.ImageData = data;
    if (typeof source === "string") {
      const extension = source.substring(source.lastIndexOf(".") + 1).toLowerCase();
      if (extension.length > 0) {
        model.ImageExtension = extension === "jpg" ? "jpeg" : extension;
      }
    } else {
      model.ImageExtension = PictureCollection.detectExtension(data);
    }

    this.models.push(model);
    return this.models.length - 1;
  }

  private static detectExtension(data: Uint8Array): string {
    if (data.length < 4) {
      return "bin";
    }

    if (data[0] === 0xff && data[1] === 0xd8) {
      return "jpeg";
    }

    if (data[0] === 0x89 && data[1] === 0x50 && data[2] === 0x4e && data[3] === 0x47) {
      return "png";
    }

    if (data[0] === 0x47 && data[1] === 0x49 && data[2] === 0x46) {
      return "gif";
    }

    if (data[0] === 0x42 && data[1] === 0x4d) {
      return "bmp";
    }

    return "bin";
  }

  public RemoveAt(index: number): void {
    if (index < 0 || index >= this.models.length) {
      throw new CellsException(`Picture index ${index} is out of range.`);
    }

    this.models.splice(index, 1);
  }

  public *[Symbol.iterator](): IterableIterator<Picture> {
    for (let index = 0; index < this.models.length; index++) {
      yield this.get(index);
    }
  }
}
