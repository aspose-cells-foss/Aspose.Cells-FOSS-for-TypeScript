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
    this.model.UpperLeftRow = value;
  }

  public get UpperLeftColumn(): number {
    return this.model.UpperLeftColumn;
  }

  public set UpperLeftColumn(value: number) {
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

  public get ImageData(): Uint8Array {
    return this.model.ImageData;
  }
}
