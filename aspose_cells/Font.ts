import { Color } from "./Color";
import { FontUnderlineType } from "./enums";
import { FontValue } from "./Core/FontValue";
import { FontUnderlineType as CoreFontUnderlineType } from "./FontUnderlineType";
import { toCoreColor } from "./ColorConvert";

/** Represents font settings. */
export class Font {
  public Name: string = "Calibri";
  public Size: number = 11;
  public Underline: FontUnderlineType = FontUnderlineType.None;
  public IsBold: boolean = false;
  public IsItalic: boolean = false;
  public IsStrikeout: boolean = false;
  public Color: Color = Color.Empty;

  /** Aspose-compatible alias of Size. */
  public get DoubleSize(): number {
    return this.Size;
  }

  public set DoubleSize(value: number) {
    this.Size = value;
  }

  /** Gets or sets the 32-bit ARGB color value. */
  public get ArgbColor(): number {
    return (this.Color.A << 24) | (this.Color.R << 16) | (this.Color.G << 8) | this.Color.B;
  }

  public set ArgbColor(value: number) {
    this.Color = Color.FromArgb(
      (value >> 24) & 0xff,
      (value >> 16) & 0xff,
      (value >> 8) & 0xff,
      value & 0xff,
    );
  }

  /** Checks whether this font equals another font instance. */
  public Equals(font: Font | null | undefined): boolean {
    if (font === null || font === undefined) {
      return false;
    }

    return (
      this.Name === font.Name &&
      this.Size === font.Size &&
      this.IsBold === font.IsBold &&
      this.IsItalic === font.IsItalic &&
      this.Underline === font.Underline &&
      this.IsStrikeout === font.IsStrikeout &&
      this.Color.Equals(font.Color)
    );
  }

  /** Creates a copy of the current instance. */
  public clone(): Font {
    const clone = new Font();
    clone.Name = this.Name;
    clone.Size = this.Size;
    clone.IsBold = this.IsBold;
    clone.IsItalic = this.IsItalic;
    clone.Underline = this.Underline;
    clone.IsStrikeout = this.IsStrikeout;
    clone.Color = this.Color;
    return clone;
  }

  public ToString(): string {
    return `Font(Name=${this.Name}, Size=${this.Size}, Bold=${this.IsBold}, Italic=${this.IsItalic}, Underline=${this.Underline}, Strikeout=${this.IsStrikeout}, Argb=${this.ArgbColor})`;
  }

  public toString(): string {
    return this.ToString();
  }

  public ToCore(): FontValue {
    const value = new FontValue();
    value.Name = this.Name;
    value.Size = this.Size;
    value.Bold = this.IsBold;
    value.Italic = this.IsItalic;
    value.Underline = this.Underline as unknown as CoreFontUnderlineType;
    value.StrikeThrough = this.IsStrikeout;
    value.Color = toCoreColor(this.Color);
    return value;
  }

  public static FromCore(value: FontValue): Font {
    const font = new Font();
    font.Name = value.Name;
    font.Size = value.Size;
    font.IsBold = value.Bold;
    font.IsItalic = value.Italic;
    font.Underline = value.Underline as unknown as FontUnderlineType;
    font.IsStrikeout = value.StrikeThrough;
    font.Color = Color.FromCore(value.Color);
    return font;
  }
}
