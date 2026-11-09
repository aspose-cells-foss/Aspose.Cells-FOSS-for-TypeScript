import { DateSystem } from "./DateSystem";

/** Minimal stand-in for System.Globalization.CultureInfo, which only contributes its name to the model. */
export interface DisplayCulture {
  Name: string;
}

export class WorkbookSettingsModel {
  public DateSystem: DateSystem = DateSystem.Windows1900;
  public DisplayCulture: DisplayCulture = { Name: "" };
}
