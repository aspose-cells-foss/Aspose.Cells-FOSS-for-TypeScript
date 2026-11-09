import { DiagnosticSeverity } from "../DiagnosticSeverity";

export class DiagnosticEntry {
  public Code: string = "";
  public Severity: DiagnosticSeverity = DiagnosticSeverity.Warning;
  public Message: string = "";
  public RepairApplied: boolean = false;
  public DataLossRisk: boolean = false;
}
