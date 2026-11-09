import { DiagnosticEntry } from "./DiagnosticEntry";

export class DiagnosticBag {
  private readonly _entries: DiagnosticEntry[] = [];

  public get Entries(): readonly DiagnosticEntry[] {
    return this._entries;
  }

  public Add(entry: DiagnosticEntry): void {
    this._entries.push(entry);
  }
}
