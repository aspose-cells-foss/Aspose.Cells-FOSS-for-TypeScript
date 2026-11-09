import { Border } from "./Border";
import { BordersValue } from "./Core/BordersValue";

/** Represents the borders of a cell. */
export class Borders {
  public Left: Border = new Border();
  public Right: Border = new Border();
  public Top: Border = new Border();
  public Bottom: Border = new Border();
  public Diagonal: Border = new Border();
  public DiagonalUp: boolean = false;
  public DiagonalDown: boolean = false;

  /** Creates a copy of the current instance. */
  public clone(): Borders {
    const clone = new Borders();
    clone.Left = this.Left.clone();
    clone.Right = this.Right.clone();
    clone.Top = this.Top.clone();
    clone.Bottom = this.Bottom.clone();
    clone.Diagonal = this.Diagonal.clone();
    clone.DiagonalUp = this.DiagonalUp;
    clone.DiagonalDown = this.DiagonalDown;
    return clone;
  }

  public ToCore(): BordersValue {
    const value = new BordersValue();
    value.Left = this.Left.ToCore();
    value.Right = this.Right.ToCore();
    value.Top = this.Top.ToCore();
    value.Bottom = this.Bottom.ToCore();
    value.Diagonal = this.Diagonal.ToCore();
    value.DiagonalUp = this.DiagonalUp;
    value.DiagonalDown = this.DiagonalDown;
    return value;
  }

  public static FromCore(value: BordersValue): Borders {
    const borders = new Borders();
    borders.Left = Border.FromCore(value.Left);
    borders.Right = Border.FromCore(value.Right);
    borders.Top = Border.FromCore(value.Top);
    borders.Bottom = Border.FromCore(value.Bottom);
    borders.Diagonal = Border.FromCore(value.Diagonal);
    borders.DiagonalUp = value.DiagonalUp;
    borders.DiagonalDown = value.DiagonalDown;
    return borders;
  }
}
