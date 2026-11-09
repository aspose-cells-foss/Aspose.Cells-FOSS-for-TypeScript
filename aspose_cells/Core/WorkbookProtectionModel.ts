export class WorkbookProtectionModel {
  public LockStructure: boolean = false;
  public LockWindows: boolean = false;
  public LockRevision: boolean = false;
  public WorkbookPassword: string = "";
  public RevisionsPassword: string = "";

  public CopyFrom(source: WorkbookProtectionModel): void {
    this.LockStructure = source.LockStructure;
    this.LockWindows = source.LockWindows;
    this.LockRevision = source.LockRevision;
    this.WorkbookPassword = source.WorkbookPassword;
    this.RevisionsPassword = source.RevisionsPassword;
  }

  public HasStoredState(): boolean {
    return (
      this.LockStructure ||
      this.LockWindows ||
      this.LockRevision ||
      (this.WorkbookPassword != null && this.WorkbookPassword.length > 0) ||
      (this.RevisionsPassword != null && this.RevisionsPassword.length > 0)
    );
  }
}
