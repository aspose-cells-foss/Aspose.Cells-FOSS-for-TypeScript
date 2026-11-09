import { CellValueKind } from "./CellValueKind";
import { RichTextRunValue } from "./RichTextRunValue";
import { StyleValue } from "./StyleValue";

export class CellRecord {
  public Value: unknown | null = null;
  public Kind: CellValueKind = CellValueKind.Blank;
  public Formula: string | null = null;
  public Style: StyleValue = StyleValue.Default;
  public RichTextRuns: RichTextRunValue[] | null = null;
  public IsExplicitlyStored: boolean = false;
}
