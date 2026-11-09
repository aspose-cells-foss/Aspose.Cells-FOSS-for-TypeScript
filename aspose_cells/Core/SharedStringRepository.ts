import { SharedStringEntry } from "../SharedStringEntry";

export class SharedStringRepository {
  private readonly _indices: Map<string, number> = new Map<string, number>();
  private readonly _entries: SharedStringEntry[] = [];

  public get Values(): readonly string[] {
    const values: string[] = new Array<string>(this._entries.length);
    for (let index = 0; index < this._entries.length; index++) {
      values[index] = this._entries[index]!.Text ?? "";
    }

    return values;
  }

  public get Entries(): readonly SharedStringEntry[] {
    return this._entries;
  }

  public Clear(): void {
    this._indices.clear();
    this._entries.length = 0;
  }

  public TryGetValue(index: number): [boolean, string] {
    if (index >= 0 && index < this._entries.length) {
      return [true, this._entries[index]!.Text ?? ""];
    }

    return [false, ""];
  }

  public TryGetEntry(index: number): [boolean, SharedStringEntry | null] {
    if (index >= 0 && index < this._entries.length) {
      return [true, this._entries[index]!.clone()];
    }

    return [false, null];
  }

  public Intern(value: string | null): number;
  public Intern(entry: SharedStringEntry | null): number;
  public Intern(valueOrEntry: string | SharedStringEntry | null): number {
    if (typeof valueOrEntry === "string" || valueOrEntry == null) {
      const entry = new SharedStringEntry();
      entry.Text = valueOrEntry ?? "";
      return this.Intern(entry);
    }

    const safeEntry = valueOrEntry.clone();
    safeEntry.Text = safeEntry.Text ?? "";

    const key = SharedStringRepository.BuildKey(safeEntry);
    const existing = this._indices.get(key);
    if (existing !== undefined) {
      return existing;
    }

    const index = this._entries.length;
    this._entries.push(safeEntry);
    this._indices.set(key, index);
    return index;
  }

  private static BuildKey(entry: SharedStringEntry): string {
    const runs = entry.Runs;
    if (runs == null || runs.length === 0) {
      return "T|" + (entry.Text ?? "");
    }

    let builder = "R|";
    builder += entry.Text ?? "";
    for (let index = 0; index < runs.length; index++) {
      const run = runs[index]!;
      builder += "|";
      builder += run.StartIndex.toString();
      builder += ":";
      builder += run.Length.toString();
      builder += ":";
      builder += run.Font.Name ?? "";
      builder += ":";
      builder += SharedStringRepository.formatSize(run.Font.Size);
      builder += ":";
      builder += run.Font.Bold ? "1" : "0";
      builder += ":";
      builder += run.Font.Italic ? "1" : "0";
      builder += ":";
      builder += run.Font.Underline.toString();
      builder += ":";
      builder += run.Font.StrikeThrough ? "1" : "0";
      builder += ":";
      builder += SharedStringRepository.formatHex(run.Font.Color.A);
      builder += SharedStringRepository.formatHex(run.Font.Color.R);
      builder += SharedStringRepository.formatHex(run.Font.Color.G);
      builder += SharedStringRepository.formatHex(run.Font.Color.B);
      builder += ":";
      builder += run.Font.Color.ThemeIndex !== null ? run.Font.Color.ThemeIndex.toString() : "";
      builder += ":";
      builder += run.Font.Color.Tint !== null ? run.Font.Color.Tint.toString() : "";
      builder += ":";
      builder += run.Font.Color.Indexed !== null ? run.Font.Color.Indexed.toString() : "";
      builder += ":";
      builder += run.Font.Family !== null ? run.Font.Family.toString() : "";
      builder += ":";
      builder += run.Font.Scheme.toString();
    }

    return builder;
  }

  /** Mirrors the invariant "0.####" number formatting used for the rich-text key. */
  private static formatSize(value: number): string {
    const text = value.toFixed(4);
    const trimmed = text.includes(".") ? text.replace(/0+$/, "").replace(/\.$/, "") : text;
    return trimmed === "-0" ? "0" : trimmed;
  }

  private static formatHex(value: number): string {
    return (value & 0xff).toString(16).toUpperCase().padStart(2, "0");
  }
}
