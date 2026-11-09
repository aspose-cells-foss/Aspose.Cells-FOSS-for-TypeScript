export class WorksheetProtectionModel {
  public IsProtected: boolean = false;
  public Objects: boolean = false;
  public Scenarios: boolean = false;
  public FormatCells: boolean = false;
  public FormatColumns: boolean = false;
  public FormatRows: boolean = false;
  public InsertColumns: boolean = false;
  public InsertRows: boolean = false;
  public InsertHyperlinks: boolean = false;
  public DeleteColumns: boolean = false;
  public DeleteRows: boolean = false;
  public SelectLockedCells: boolean = false;
  public Sort: boolean = false;
  public AutoFilter: boolean = false;
  public PivotTables: boolean = false;
  public SelectUnlockedCells: boolean = false;
  public PasswordHash: string | null = null;
  public AlgorithmName: string | null = null;
  public HashValue: string | null = null;
  public SaltValue: string | null = null;
  public SpinCount: string | null = null;

  public Clear(): void {
    this.IsProtected = false;
    this.Objects = false;
    this.Scenarios = false;
    this.FormatCells = false;
    this.FormatColumns = false;
    this.FormatRows = false;
    this.InsertColumns = false;
    this.InsertRows = false;
    this.InsertHyperlinks = false;
    this.DeleteColumns = false;
    this.DeleteRows = false;
    this.SelectLockedCells = false;
    this.Sort = false;
    this.AutoFilter = false;
    this.PivotTables = false;
    this.SelectUnlockedCells = false;
    this.PasswordHash = null;
    this.AlgorithmName = null;
    this.HashValue = null;
    this.SaltValue = null;
    this.SpinCount = null;
  }

  public HasStoredState(): boolean {
    return (
      this.IsProtected ||
      this.Objects ||
      this.Scenarios ||
      this.FormatCells ||
      this.FormatColumns ||
      this.FormatRows ||
      this.InsertColumns ||
      this.InsertRows ||
      this.InsertHyperlinks ||
      this.DeleteColumns ||
      this.DeleteRows ||
      this.SelectLockedCells ||
      this.Sort ||
      this.AutoFilter ||
      this.PivotTables ||
      this.SelectUnlockedCells ||
      WorksheetProtectionModel.hasText(this.PasswordHash) ||
      WorksheetProtectionModel.hasText(this.AlgorithmName) ||
      WorksheetProtectionModel.hasText(this.HashValue) ||
      WorksheetProtectionModel.hasText(this.SaltValue) ||
      WorksheetProtectionModel.hasText(this.SpinCount)
    );
  }

  private static hasText(value: string | null): boolean {
    return value != null && value.trim().length > 0;
  }
}
