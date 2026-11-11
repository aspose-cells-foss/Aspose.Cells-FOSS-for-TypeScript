import { ChartType } from "../ChartType";

export class ChartModel {
  public Name: string = "";
  public ChartType: ChartType = ChartType.Unknown;
  /** Cell range formula used by the chart series (e.g. "Sheet1!$B$1:$B$5"). */
  public DataRange: string = "";
  /** Series entries detected in the chart definition. */
  public readonly Series: ChartSeriesModel[] = [];
  public UpperLeftRow: number = 0;
  public UpperLeftColumn: number = 0;
  public UpperLeftRowOffset: number = 0;
  public UpperLeftColumnOffset: number = 0;
  public LowerRightRow: number = 0;
  public LowerRightColumn: number = 0;
  public LowerRightRowOffset: number = 0;
  public LowerRightColumnOffset: number = 0;
  public ExtentCx: number = 0;
  public ExtentCy: number = 0;
  public RawChartXml: string | null = null;
  public RawGraphicFrameXml: string | null = null;
  public OriginalRId: string | null = null;
  public IsChartEx: boolean = false;
  public readonly CompanionFiles: ChartCompanionFile[] = [];
}

/** Series entry detected in a chart definition file. */
export class ChartSeriesModel {
  public Index: number = 0;
  public Order: number = 0;
  /** Value range formula of the series (c:val/c:numRef/c:f). */
  public ValuesFormula: string = "";
  /** Category range formula of the series (c:cat), when present. */
  public CategoriesFormula: string = "";
}

export class ChartCompanionFile {
  public RelationshipId: string | null = null;
  public RelationshipType: string | null = null;
  public FileName: string | null = null;
  public RawContent: string | null = null;
  public BinaryContent: Uint8Array | null = null;
}
