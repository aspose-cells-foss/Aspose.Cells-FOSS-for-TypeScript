import { Font } from "./Font";
import { Color } from "./Color";
import { FontUnderlineType as ApiFontUnderlineType } from "./enums";
import { RichTextRunValue } from "./Core/RichTextRunValue";
import { FontValue } from "./Core/FontValue";
import { FontUnderlineType as CoreFontUnderlineType } from "./FontUnderlineType";
import { toCoreColor } from "./ColorConvert";

/** Represents a range of characters within a cell's rich text. */
export class FontSetting {
  public readonly StartIndex: number;
  public readonly Length: number;
  public Font: Font;

  public constructor(startIndex: number, length: number) {
    if (startIndex < 0) {
      throw new RangeError("startIndex");
    }

    if (length < 0) {
      throw new RangeError("length");
    }

    this.StartIndex = startIndex;
    this.Length = length;
    this.Font = new Font();
  }

  /** Creates a copy of the current instance. */
  public clone(): FontSetting {
    const clone = new FontSetting(this.StartIndex, this.Length);
    clone.Font = this.Font === null || this.Font === undefined ? new Font() : this.Font.clone();
    return clone;
  }

  public ToCore(): RichTextRunValue {
    const value = new RichTextRunValue();
    value.StartIndex = this.StartIndex;
    value.Length = this.Length;

    const font = new FontValue();
    const source = this.Font;
    if (source === null || source === undefined) {
      font.Name = "Calibri";
      font.Size = 11;
      font.Color = toCoreColor(Color.Empty);
    } else {
      font.Name = source.Name;
      font.Size = source.Size;
      font.Bold = source.IsBold;
      font.Italic = source.IsItalic;
      font.Underline = source.Underline as unknown as CoreFontUnderlineType;
      font.StrikeThrough = source.IsStrikeout;
      font.Color = toCoreColor(source.Color);
    }

    value.Font = font;
    return value;
  }

  public static FromCore(run: RichTextRunValue): FontSetting {
    const setting = new FontSetting(run.StartIndex, run.Length);
    if (run.Font !== null && run.Font !== undefined) {
      setting.Font.Name = run.Font.Name;
      setting.Font.Size = run.Font.Size;
      setting.Font.IsBold = run.Font.Bold;
      setting.Font.IsItalic = run.Font.Italic;
      setting.Font.Underline = run.Font.Underline as unknown as ApiFontUnderlineType;
      setting.Font.IsStrikeout = run.Font.StrikeThrough;
      setting.Font.Color = Color.FromCore(run.Font.Color);
    }

    return setting;
  }
}
