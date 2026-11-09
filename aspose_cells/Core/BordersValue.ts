import { BorderSideValue } from "./BorderSideValue";

export class BordersValue {
  public Left: BorderSideValue = new BorderSideValue();
  public Right: BorderSideValue = new BorderSideValue();
  public Top: BorderSideValue = new BorderSideValue();
  public Bottom: BorderSideValue = new BorderSideValue();
  public Diagonal: BorderSideValue = new BorderSideValue();
  public DiagonalUp: boolean = false;
  public DiagonalDown: boolean = false;

  public clone(): BordersValue {
    const clone = new BordersValue();
    clone.Left = this.Left.clone();
    clone.Right = this.Right.clone();
    clone.Top = this.Top.clone();
    clone.Bottom = this.Bottom.clone();
    clone.Diagonal = this.Diagonal.clone();
    clone.DiagonalUp = this.DiagonalUp;
    clone.DiagonalDown = this.DiagonalDown;
    return clone;
  }
}
