import { LoadIssue } from "./LoadIssue";

/** Represents diagnostics collected while loading a workbook. */
export class LoadDiagnostics {
  private readonly issues: LoadIssue[] = [];

  /** Gets the collected load issues. */
  public get Issues(): readonly LoadIssue[] {
    return this.issues;
  }

  /** Gets whether any issue required a repair. */
  public get HasRepairs(): boolean {
    return this.issues.some((issue) => issue.RepairApplied);
  }

  /** Gets whether any issue may have caused data loss. */
  public get HasDataLossRisk(): boolean {
    return this.issues.some((issue) => issue.DataLossRisk);
  }

  /** Adds a load issue. */
  public Add(issue: LoadIssue): void {
    this.issues.push(issue);
  }
}
