export class CalculationPropertiesModel {
  public CalculationId: number | null = null;
  public CalculationMode: string = "";
  public FullCalculationOnLoad: boolean = false;
  public ReferenceMode: string = "";
  public Iterate: boolean = false;
  public IterateCount: number | null = null;
  public IterateDelta: number | null = null;
  public FullPrecision: boolean | null = null;
  public CalculationCompleted: boolean | null = null;
  public CalculationOnSave: boolean | null = null;
  public ConcurrentCalculation: boolean | null = null;
  public ForceFullCalculation: boolean = false;

  public CopyFrom(source: CalculationPropertiesModel): void {
    this.CalculationId = source.CalculationId;
    this.CalculationMode = source.CalculationMode;
    this.FullCalculationOnLoad = source.FullCalculationOnLoad;
    this.ReferenceMode = source.ReferenceMode;
    this.Iterate = source.Iterate;
    this.IterateCount = source.IterateCount;
    this.IterateDelta = source.IterateDelta;
    this.FullPrecision = source.FullPrecision;
    this.CalculationCompleted = source.CalculationCompleted;
    this.CalculationOnSave = source.CalculationOnSave;
    this.ConcurrentCalculation = source.ConcurrentCalculation;
    this.ForceFullCalculation = source.ForceFullCalculation;
  }

  public HasStoredState(): boolean {
    return (
      this.CalculationId !== null ||
      (this.CalculationMode != null && this.CalculationMode.length > 0) ||
      this.FullCalculationOnLoad ||
      (this.ReferenceMode != null && this.ReferenceMode.length > 0) ||
      this.Iterate ||
      this.IterateCount !== null ||
      this.IterateDelta !== null ||
      this.FullPrecision !== null ||
      this.CalculationCompleted !== null ||
      this.CalculationOnSave !== null ||
      this.ConcurrentCalculation !== null ||
      this.ForceFullCalculation
    );
  }
}
