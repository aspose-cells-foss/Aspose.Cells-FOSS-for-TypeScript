import { ColorValue } from "./ColorValue";
import { BorderStyle } from "./BorderStyle";

export class BorderSideValue {
  public Style: BorderStyle = BorderStyle.None;
  public Color: ColorValue = new ColorValue();

  public clone(): BorderSideValue {
    const clone = new BorderSideValue();
    clone.Style = this.Style;
    clone.Color = this.Color;
    return clone;
  }
}
