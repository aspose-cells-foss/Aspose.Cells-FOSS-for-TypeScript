import { WorkbookProtectionModel } from "./Core/WorkbookProtectionModel";

/** Represents workbook protection settings. */
export class WorkbookProtection {
  private readonly model: WorkbookProtectionModel;

  public constructor(model: WorkbookProtectionModel) {
    this.model = model;
  }

  public get LockStructure(): boolean {
    return this.model.LockStructure;
  }

  public set LockStructure(value: boolean) {
    this.model.LockStructure = value;
  }

  public get LockWindows(): boolean {
    return this.model.LockWindows;
  }

  public set LockWindows(value: boolean) {
    this.model.LockWindows = value;
  }

  public get LockRevision(): boolean {
    return this.model.LockRevision;
  }

  public set LockRevision(value: boolean) {
    this.model.LockRevision = value;
  }

  public get WorkbookPassword(): string {
    return this.model.WorkbookPassword;
  }

  public set WorkbookPassword(value: string) {
    this.model.WorkbookPassword = value ?? "";
  }

  public get RevisionsPassword(): string {
    return this.model.RevisionsPassword;
  }

  public set RevisionsPassword(value: string) {
    this.model.RevisionsPassword = value ?? "";
  }

  public get IsProtected(): boolean {
    return this.model.HasStoredState();
  }
}
