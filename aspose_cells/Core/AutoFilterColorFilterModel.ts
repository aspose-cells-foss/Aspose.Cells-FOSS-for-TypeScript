export class AutoFilterColorFilterModel {
  public Enabled: boolean = false;
  public DifferentialStyleId: number | null = null;
  public CellColor: boolean = false;

  public Clear(): void {
    this.Enabled = false;
    this.DifferentialStyleId = null;
    this.CellColor = false;
  }
}
