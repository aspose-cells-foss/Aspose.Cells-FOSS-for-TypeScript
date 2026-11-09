import { StyleValue } from "./StyleValue";

export class ColumnRangeModel {
  public MinColumnIndex: number = 0;
  public MaxColumnIndex: number = 0;
  public Width: number | null = null;
  public Hidden: boolean = false;
  public StyleIndex: number | null = null;
  public Style: StyleValue | null = null;
}
