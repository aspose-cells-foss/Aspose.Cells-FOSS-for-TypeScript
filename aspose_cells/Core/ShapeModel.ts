export class ShapeModel {
  public Name: string;
  public DrawingObjectId: number;
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
  public GeometryType: string;
  public RawStyleXml: string | null = null;
  public RawTxBodyXml: string | null = null;
  public StartConnectionShapeId: number;
  public StartConnectionSite: number;
  public EndConnectionShapeId: number;
  public EndConnectionSite: number;
  public RawElementXml: string | null = null;

  public constructor() {
    this.Name = "";
    this.GeometryType = "rect";
    this.DrawingObjectId = -1;
    this.StartConnectionShapeId = -1;
    this.StartConnectionSite = -1;
    this.EndConnectionShapeId = -1;
    this.EndConnectionSite = -1;
  }
}
