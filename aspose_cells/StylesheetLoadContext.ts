import { StyleValue } from "./Core/StyleValue";

export class StylesheetLoadContext {
  public DefaultCellStyle: StyleValue = StyleValue.Default.clone();
  public readonly CellFormats: StyleValue[] = [StyleValue.Default.clone()];
  public readonly DifferentialFormats: StyleValue[] = [];
  public RawTableStylesXml: string | null = null;
  public readonly RawDifferentialFormatsXml: string[] = [];
  public readonly DateStyleIndexes: Set<number> = new Set<number>();
}
