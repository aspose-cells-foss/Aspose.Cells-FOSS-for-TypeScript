import { DiagnosticSeverity } from "./enums";

/** Represents a single load issue. */
export class LoadIssue {
  public readonly Code: string;
  public readonly Severity: DiagnosticSeverity;
  public readonly Message: string;
  public readonly RepairApplied: boolean;
  public readonly DataLossRisk: boolean;
  public PartUri: string | null = null;
  public SheetName: string | null = null;
  public CellRef: string | null = null;
  public RowIndex: number | null = null;

  public constructor(
    code: string,
    severity: DiagnosticSeverity,
    message: string,
    repairApplied: boolean = false,
    dataLossRisk: boolean = false,
  ) {
    if (code === null || code === undefined) {
      throw new Error("code");
    }

    if (message === null || message === undefined) {
      throw new Error("message");
    }

    this.Code = code;
    this.Severity = severity;
    this.Message = message;
    this.RepairApplied = repairApplied;
    this.DataLossRisk = dataLossRisk;
  }
}
