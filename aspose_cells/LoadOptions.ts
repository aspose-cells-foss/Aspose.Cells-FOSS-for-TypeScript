import { LoadFormat } from "./enums";
import { IWarningCallback } from "./IWarningCallback";

/** Specifies how a workbook should be loaded. */
export class LoadOptions {
  public LoadFormat: LoadFormat = LoadFormat.Auto;
  public StrictMode: boolean = false;
  public TryRepairPackage: boolean = true;
  public TryRepairXml: boolean = true;
  public PreserveUnsupportedParts: boolean = true;
  public WarningCallback: IWarningCallback | null = null;
}
