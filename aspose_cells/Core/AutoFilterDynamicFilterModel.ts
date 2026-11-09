export class AutoFilterDynamicFilterModel {
  public Enabled: boolean = false;
  public Type: string = "";
  public Value: number | null = null;
  public MaxValue: number | null = null;

  public Clear(): void {
    this.Enabled = false;
    this.Type = "";
    this.Value = null;
    this.MaxValue = null;
  }
}
