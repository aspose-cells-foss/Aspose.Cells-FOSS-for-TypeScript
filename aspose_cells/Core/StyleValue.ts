import { AlignmentValue } from "./AlignmentValue";
import { BordersValue } from "./BordersValue";
import { ColorValue } from "./ColorValue";
import { FillPatternKind } from "./FillPatternKind";
import { FontValue } from "./FontValue";
import { NumberFormatValue } from "./NumberFormatValue";
import { ProtectionValue } from "./ProtectionValue";

export class StyleValue {
  public static get Default(): StyleValue {
    return new StyleValue();
  }

  public Font: FontValue = new FontValue();
  public Pattern: FillPatternKind = FillPatternKind.None;
  public ForegroundColor: ColorValue = new ColorValue();
  public BackgroundColor: ColorValue = new ColorValue();
  public Borders: BordersValue = new BordersValue();
  public Alignment: AlignmentValue = new AlignmentValue();
  public Protection: ProtectionValue = new ProtectionValue();
  public NumberFormat: NumberFormatValue = new NumberFormatValue();
  public QuotePrefix: boolean = false;

  public clone(): StyleValue {
    const clone = new StyleValue();
    clone.Font = this.Font.clone();
    clone.Pattern = this.Pattern;
    clone.ForegroundColor = this.ForegroundColor;
    clone.BackgroundColor = this.BackgroundColor;
    clone.Borders = this.Borders.clone();
    clone.Alignment = this.Alignment.clone();
    clone.Protection = this.Protection.clone();
    clone.NumberFormat = this.NumberFormat.clone();
    clone.QuotePrefix = this.QuotePrefix;
    return clone;
  }
}
