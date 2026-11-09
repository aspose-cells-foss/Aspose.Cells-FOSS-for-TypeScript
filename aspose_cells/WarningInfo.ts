import { DiagnosticSeverity } from "./enums";

/** Represents warning information passed to a warning callback. */
export class WarningInfo {
  public Code: string = "";
  public Severity: DiagnosticSeverity = DiagnosticSeverity.Warning;
  public Message: string = "";
  public DataLossRisk: boolean = false;
  public PartUri: string | null = null;
  public SheetName: string | null = null;
  public CellRef: string | null = null;
  public RowIndex: number | null = null;
}
