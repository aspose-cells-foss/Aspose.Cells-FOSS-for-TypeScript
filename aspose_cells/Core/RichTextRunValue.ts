import { FontValue } from "./FontValue";

export class RichTextRunValue {
  public StartIndex: number = 0;
  public Length: number = 0;
  public Font: FontValue = new FontValue();

  public clone(): RichTextRunValue {
    const clone = new RichTextRunValue();
    clone.StartIndex = this.StartIndex;
    clone.Length = this.Length;
    clone.Font = this.Font == null ? new FontValue() : this.Font.clone();
    return clone;
  }
}
