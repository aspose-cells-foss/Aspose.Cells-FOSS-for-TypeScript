import { StyleValue } from "./StyleValue";

export class RowModel {
  public Height: number | null = null;
  public CustomHeight: boolean = false;
  public Hidden: boolean = false;
  public StyleIndex: number | null = null;
  public Style: StyleValue | null = null;
}
