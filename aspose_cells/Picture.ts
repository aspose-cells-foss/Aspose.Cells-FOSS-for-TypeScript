import { CellsException } from "./CellsException";
import { ImageType } from "./enums";
import { PictureModel } from "./Core/PictureModel";

// Implemented by serializer module.
/** Represents a picture anchored to a worksheet. */
export class Picture {
  public constructor(private readonly model: PictureModel) {}

  public get Model(): PictureModel {
    return this.model;
  }

  public get Name(): string {
    return this.model.Name;
  }

  public set Name(value: string) {
    this.model.Name = value ?? "";
  }

  public get UpperLeftRow(): number {
    return this.model.UpperLeftRow;
  }

  public set UpperLeftRow(value: number) {
    if (value < 0) {
      throw new CellsException("UpperLeftRow must be non-negative.");
    }

    this.model.UpperLeftRow = value;
  }

  public get UpperLeftColumn(): number {
    return this.model.UpperLeftColumn;
  }

  public set UpperLeftColumn(value: number) {
    if (value < 0) {
      throw new CellsException("UpperLeftColumn must be non-negative.");
    }

    this.model.UpperLeftColumn = value;
  }

  public get LowerRightRow(): number {
    return this.model.LowerRightRow;
  }

  public set LowerRightRow(value: number) {
    this.model.LowerRightRow = value;
  }

  public get LowerRightColumn(): number {
    return this.model.LowerRightColumn;
  }

  public set LowerRightColumn(value: number) {
    this.model.LowerRightColumn = value;
  }

  public get ImageType(): ImageType {
    switch (this.model.ImageExtension.toLowerCase()) {
      case "jpg":
      case "jpeg":
        return ImageType.Jpeg;
      case "png":
        return ImageType.Png;
      case "gif":
        return ImageType.Gif;
      case "bmp":
        return ImageType.Bmp;
      default:
        return ImageType.Unknown;
    }
  }

  public get Data(): Uint8Array {
    return this.model.ImageData;
  }

  public set Data(value: Uint8Array) {
    if (value === null || value === undefined || value.length === 0) {
      throw new CellsException("Picture data must be non-null and non-empty.");
    }

    this.model.ImageData = value;
    this.model.ImageExtension = Picture.detectExtension(value);
  }

  public get ImageData(): Uint8Array {
    return this.model.ImageData;
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
}
