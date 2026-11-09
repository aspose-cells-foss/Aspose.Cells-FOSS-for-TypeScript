import { Color } from "./Color";
import { BorderStyleType } from "./enums";
import { BorderSideValue } from "./Core/BorderSideValue";
import { BorderStyle as CoreBorderStyle } from "./Core/BorderStyle";
import { toCoreColor } from "./ColorConvert";

/** Represents a single border side. */
export class Border {
  public LineStyle: BorderStyleType = BorderStyleType.None;
  public Color: Color = Color.Empty;

  /** Creates a copy of the current instance. */
  public clone(): Border {
    const clone = new Border();
    clone.LineStyle = this.LineStyle;
    clone.Color = this.Color;
    return clone;
  }

  public ToCore(): BorderSideValue {
    const value = new BorderSideValue();
    value.Style = this.LineStyle as unknown as CoreBorderStyle;
    value.Color = toCoreColor(this.Color);
    return value;
  }

  public static FromCore(value: BorderSideValue): Border {
    const border = new Border();
    border.LineStyle = value.Style as unknown as BorderStyleType;
    border.Color = Color.FromCore(value.Color);
    return border;
  }
}
