import { Worksheet } from "./worksheet";
import type { Style } from "./types";

export class WorksheetCollection {
  private _worksheets: Worksheet[] = [];
  private _styles: Map<number, Style> = new Map();

  constructor(createDefault: boolean = true) {
    if (createDefault) {
      this._worksheets.push(new Worksheet("Sheet1", 0));
    }
  }

  [Symbol.iterator](): Iterator<Worksheet> {
    return this._worksheets[Symbol.iterator]();
  }

  [index: number]: Worksheet | undefined;

  get worksheets(): Worksheet[] {
    return this._worksheets;
  }

  get length(): number {
    return this._worksheets.length;
  }

  get worksheetCount(): number {
    return this._worksheets.length;
  }

  get worksheet(): Worksheet {
    return this._worksheets[0];
  }

  get styles(): Map<number, Style> {
    return this._styles;
  }

  get(index: number): Worksheet | undefined {
    return this._worksheets[index];
  }

  getByName(name: string): Worksheet | undefined {
    return this._worksheets.find(w => w.name === name);
  }

  map<T>(callback: (value: Worksheet, index: number) => T): T[] {
    return this._worksheets.map(callback);
  }

  filter(callback: (value: Worksheet, index: number) => boolean): Worksheet[] {
    return this._worksheets.filter(callback);
  }

  forEach(callback: (value: Worksheet, index: number) => void): void {
    this._worksheets.forEach(callback);
  }

  find(
    callback: (value: Worksheet, index: number) => boolean,
  ): Worksheet | undefined {
    return this._worksheets.find(callback);
  }

  indexOf(searchElement: Worksheet, fromIndex?: number): number {
    return this._worksheets.indexOf(searchElement, fromIndex);
  }

  addWorksheet(name?: string): Worksheet {
    const index = this._worksheets.length;
    const sheetName = name ?? `Sheet${index + 1}`;
    const sheet = new Worksheet(sheetName, index);
    this._worksheets.push(sheet);
    return sheet;
  }

  removeWorksheet(index: number): boolean {
    if (index >= 0 && index < this._worksheets.length) {
      this._worksheets.splice(index, 1);
      return true;
    }
    return false;
  }

  moveWorksheet(fromIndex: number, toIndex: number): boolean {
    if (fromIndex < 0 || fromIndex >= this._worksheets.length) return false;
    if (toIndex < 0 || toIndex >= this._worksheets.length) return false;
    const sheet = this._worksheets.splice(fromIndex, 1)[0];
    this._worksheets.splice(toIndex, 0, sheet);
    return true;
  }

  getStyle(index: number): Style | undefined {
    return this._styles.get(index);
  }

  setStyle(index: number, style: Style): void {
    this._styles.set(index, style);
  }

  clear(): void {
    this._worksheets.length = 0;
  }

  getNumFmt(numFmtId: string | null): string {
    if (!numFmtId) return "General";
    const id = parseInt(numFmtId, 10);
    const formats: { [key: number]: string } = {
      0: "General",
      1: "0",
      2: "0.00",
      3: "#,##0",
      4: "#,##0.00",
      5: "$#,##0_);($#,##0)",
      6: "$#,##0.00_);($#,##0.00)",
      7: "$#,##0.00",
      8: "$#,##0.00",
      9: "0%",
      10: "0.00%",
      11: "0.00E+00",
      14: "yyyy-mm-dd",
      41: "yyyy\\-mm\\-dd",
    };
    return formats[id] || "General";
  }
}
