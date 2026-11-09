import { ColorValue } from "./ColorValue";
import { FormatConditionType } from "../FormatConditionType";
import { OperatorType } from "../OperatorType";
import { StyleValue } from "./StyleValue";

export class FormatConditionModel {
  public Type: FormatConditionType = FormatConditionType.CellValue;
  public Operator: OperatorType = OperatorType.None;
  public Formula1: string | null = null;
  public Formula2: string | null = null;
  public TimePeriod: string | null = null;
  public Duplicate: boolean = true;
  public Top: boolean = true;
  public Percent: boolean = false;
  public Rank: number = 0;
  public Above: boolean = true;
  public StandardDeviation: number = 0;
  public ColorScaleCount: number = 2;
  public MinColor: ColorValue = new ColorValue();
  public MidColor: ColorValue = new ColorValue();
  public MaxColor: ColorValue = new ColorValue();
  public BarColor: ColorValue = new ColorValue();
  public NegativeBarColor: ColorValue = new ColorValue();
  public ShowBorder: boolean = false;
  public DataBarGradient: boolean = true;
  public DataBarMinLength: number = 0.1;
  public DataBarMaxLength: number = 0.9;
  public DataBarAutoMinimum: boolean = false;
  public Direction: string | null = null;
  public BarLength: string | null = null;
  public IconSetType: string | null = null;
  public ReverseIcons: boolean = false;
  public ShowIconOnly: boolean = false;
  public Priority: number = 0;
  public StopIfTrue: boolean = false;
  public Style: StyleValue;

  public constructor() {
    this.Style = StyleValue.Default.clone();
  }
}
