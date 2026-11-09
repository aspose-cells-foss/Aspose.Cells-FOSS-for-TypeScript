export class PictureModel {
  public Name: string;
  public UpperLeftRow: number = 0;
  public UpperLeftColumn: number = 0;
  public UpperLeftColumnOffset: number = 0;
  public UpperLeftRowOffset: number = 0;
  public LowerRightRow: number = 0;
  public LowerRightColumn: number = 0;
  public LowerRightColumnOffset: number = 0;
  public LowerRightRowOffset: number = 0;
  public ExtentCx: number = 0;
  public ExtentCy: number = 0;
  public ImageExtension: string;
  public ImageData: Uint8Array;
  public OriginalRId: string | null = null;

  public constructor() {
    this.Name = "";
    this.ImageExtension = "jpeg";
    this.ImageData = new Uint8Array(0);
  }
}
