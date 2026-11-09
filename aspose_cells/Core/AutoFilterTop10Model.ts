export class AutoFilterTop10Model {
  public Enabled: boolean = false;
  public Top: boolean = true;
  public Percent: boolean = false;
  public Value: number | null = null;
  public FilterValue: number | null = null;

  public Clear(): void {
    this.Enabled = false;
    this.Top = true;
    this.Percent = false;
    this.Value = null;
    this.FilterValue = null;
  }
}
