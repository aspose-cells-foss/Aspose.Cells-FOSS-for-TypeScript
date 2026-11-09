export class NumberFormatValue {
  public Number: number = 0;
  public Custom: string | null = null;

  public clone(): NumberFormatValue {
    const clone = new NumberFormatValue();
    clone.Number = this.Number;
    clone.Custom = this.Custom;
    return clone;
  }
}
