import { HorizontalAlignment } from "./HorizontalAlignment";
import { VerticalAlignment } from "./VerticalAlignment";

export class AlignmentValue {
  public Horizontal: HorizontalAlignment = HorizontalAlignment.General;
  public Vertical: VerticalAlignment = VerticalAlignment.Bottom;
  public WrapText: boolean = false;
  public IndentLevel: number = 0;
  public TextRotation: number = 0;
  public ShrinkToFit: boolean = false;
  public ReadingOrder: number = 0;
  public RelativeIndent: number = 0;

  public clone(): AlignmentValue {
    const clone = new AlignmentValue();
    clone.Horizontal = this.Horizontal;
    clone.Vertical = this.Vertical;
    clone.WrapText = this.WrapText;
    clone.IndentLevel = this.IndentLevel;
    clone.TextRotation = this.TextRotation;
    clone.ShrinkToFit = this.ShrinkToFit;
    clone.ReadingOrder = this.ReadingOrder;
    clone.RelativeIndent = this.RelativeIndent;
    return clone;
  }
}
