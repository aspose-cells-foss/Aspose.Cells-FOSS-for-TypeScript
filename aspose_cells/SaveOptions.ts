import { SaveFormat } from "./enums";

/** Specifies how a workbook should be saved. */
export class SaveOptions {
  public SaveFormat: SaveFormat = SaveFormat.Xlsx;
  public UseSharedStrings: boolean = true;
  public ValidateBeforeSave: boolean = true;
  public CompactStyles: boolean = true;
  public PreserveRecoveryMetadata: boolean = false;
}
