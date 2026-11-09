import { WarningInfo } from "./WarningInfo";

/** Defines a callback that receives load warnings. */
export interface IWarningCallback {
  Warning(warningInfo: WarningInfo): void;
}
