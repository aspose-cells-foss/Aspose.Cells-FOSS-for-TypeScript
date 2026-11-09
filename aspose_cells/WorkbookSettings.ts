import { CellsException } from "./CellsException";
import { DateSystem } from "./Core/DateSystem";
import { DisplayCulture, WorkbookSettingsModel } from "./Core/WorkbookSettingsModel";

/** Represents workbook-level settings that affect date handling and display formatting. */
export class WorkbookSettings {
  private readonly model: WorkbookSettingsModel;

  public constructor(model: WorkbookSettingsModel) {
    this.model = model;
  }

  /** Gets or sets whether the workbook uses the 1904 date system. */
  public get Date1904(): boolean {
    return this.model.DateSystem === DateSystem.Mac1904;
  }

  public set Date1904(value: boolean) {
    this.model.DateSystem = value ? DateSystem.Mac1904 : DateSystem.Windows1900;
  }

  /** Gets or sets the culture used for display-string formatting. */
  public get Culture(): DisplayCulture {
    return { Name: this.model.DisplayCulture.Name };
  }

  public set Culture(value: DisplayCulture | string) {
    if (value === null || value === undefined) {
      throw new CellsException("Culture must not be null.");
    }

    const name = typeof value === "string" ? value : value.Name;
    this.model.DisplayCulture = { Name: name };
  }
}
