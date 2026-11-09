import { ColorValue } from "./ColorValue";
import { FontUnderlineType } from "../FontUnderlineType";
import { FontSchemeType } from "./FontSchemeType";

export class FontValue {
  public Name: string = "Calibri";
  public Size: number = 11;
  public Bold: boolean = false;
  public Italic: boolean = false;
  public Underline: FontUnderlineType = FontUnderlineType.None;
  public StrikeThrough: boolean = false;
  public Color: ColorValue = new ColorValue();
  public Family: number | null = null;
  public Scheme: FontSchemeType = FontSchemeType.None;

  public clone(): FontValue {
    const clone = new FontValue();
    clone.Name = this.Name;
    clone.Size = this.Size;
    clone.Bold = this.Bold;
    clone.Italic = this.Italic;
    clone.Underline = this.Underline;
    clone.StrikeThrough = this.StrikeThrough;
    clone.Color = this.Color;
    clone.Family = this.Family;
    clone.Scheme = this.Scheme;
    return clone;
  }
}
