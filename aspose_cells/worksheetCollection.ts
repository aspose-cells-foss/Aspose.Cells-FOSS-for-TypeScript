import { LegacyWorksheet } from "./worksheet";
import type { Style } from "./types";

export class LegacyWorksheetCollection {
  private _worksheets: LegacyWorksheet[] = [];
  private _styles: Map<number, Style> = new Map();

  constructor(createDefault: boolean = true) {
    if (createDefault) {
      this._worksheets.push(new LegacyWorksheet("Sheet1", 0));
    }
  }

  [Symbol.iterator](): Iterator<LegacyWorksheet> {
    return this._worksheets[Symbol.iterator]();
  }

  [index: number]: LegacyWorksheet | undefined;

  get worksheets(): LegacyWorksheet[] {
    return this._worksheets;
  }

  get length(): number {
    return this._worksheets.length;
  }

  get worksheetCount(): number {
    return this._worksheets.length;
  }

  get worksheet(): LegacyWorksheet {
    return this._worksheets[0];
  }

  get styles(): Map<number, Style> {
    return this._styles;
  }

  get(index: number): LegacyWorksheet | undefined {
    return this._worksheets[index];
  }

  getByName(name: string): LegacyWorksheet | undefined {
    return this._worksheets.find(w => w.name === name);
  }

  map<T>(callback: (value: LegacyWorksheet, index: number) => T): T[] {
    return this._worksheets.map(callback);
  }

  filter(callback: (value: LegacyWorksheet, index: number) => boolean): LegacyWorksheet[] {
    return this._worksheets.filter(callback);
  }

  forEach(callback: (value: LegacyWorksheet, index: number) => void): void {
    this._worksheets.forEach(callback);
  }

  find(
    callback: (value: LegacyWorksheet, index: number) => boolean,
  ): LegacyWorksheet | undefined {
    return this._worksheets.find(callback);
  }

  indexOf(searchElement: LegacyWorksheet, fromIndex?: number): number {
    return this._worksheets.indexOf(searchElement, fromIndex);
  }

  addWorksheet(name?: string): LegacyWorksheet {
    const index = this._worksheets.length;
    const sheetName = name ?? `Sheet${index + 1}`;
    const sheet = new LegacyWorksheet(sheetName, index);
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

import { CellsException } from "./CellsException";
import { createIndexerProxy } from "./IndexerProxy";
import { Workbook } from "./workbook";
import { Worksheet } from "./worksheet";
import { WorksheetModel } from "./Core/WorksheetModel";

/** Encapsulates the workbook's worksheets and active-sheet state. */
export class WorksheetCollection {
  [key: string]: any;

  private readonly workbook: Workbook;

  public constructor(workbook: Workbook) {
    this.workbook = workbook;
  }

  /** Gets the worksheet at the specified zero-based index or with the specified name. */
  public get(index: number): Worksheet;
  public get(name: string): Worksheet;
  public get(indexOrName: number | string): Worksheet {
    return typeof indexOrName === "number"
      ? this.getByIndex(indexOrName)
      : this.getByName(indexOrName);
  }

  /** Alias of get for index access. */
  public at(index: number): Worksheet {
    return this.getByIndex(index);
  }

  /** Gets the total number of worksheets in the workbook. */
  public get Count(): number {
    return this.workbook.Model.Worksheets.length;
  }

  /** Gets or sets the zero-based index of the active worksheet. */
  public get ActiveSheetIndex(): number {
    return this.workbook.Model.ActiveSheetIndex;
  }

  public set ActiveSheetIndex(value: number) {
    if (value < 0 || value >= this.workbook.Model.Worksheets.length) {
      throw new CellsException("ActiveSheetIndex must refer to an existing worksheet.");
    }

    this.workbook.Model.ActiveSheetIndex = value;
  }

  /** Gets or sets the name of the active worksheet. */
  public get ActiveSheetName(): string {
    return this.workbook.Model.Worksheets[this.ActiveSheetIndex].Name;
  }

  public set ActiveSheetName(value: string) {
    const index = this.indexOf(value);
    if (index < 0) {
      throw new CellsException(`Worksheet '${value}' was not found.`);
    }

    this.ActiveSheetIndex = index;
  }

  /** Adds a worksheet using the next generated default name. */
  public Add(): number;
  /** Adds a worksheet with the specified name and returns its index. */
  public Add(sheetName: string): number;
  public Add(sheetName?: string): number {
    const name = sheetName ?? this.generateDefaultSheetName();
    if (name === null || name === undefined || name.trim().length === 0) {
      throw new CellsException("Worksheet name must be non-empty.");
    }

    this.workbook.EnsureUniqueSheetName(name);
    this.workbook.Model.Worksheets.push(new WorksheetModel(name));
    return this.workbook.Model.Worksheets.length - 1;
  }

  /** Removes the worksheet with the specified name. */
  public RemoveAt(sheetName: string): void;
  /** Removes the worksheet at the specified zero-based index. */
  public RemoveAt(index: number): void;
  public RemoveAt(indexOrName: number | string): void {
    const index = typeof indexOrName === "number" ? indexOrName : this.indexOf(indexOrName);
    if (index < 0 || index >= this.workbook.Model.Worksheets.length) {
      throw new CellsException(
        typeof indexOrName === "string"
          ? `Worksheet '${indexOrName}' was not found.`
          : "Worksheet index was out of range.",
      );
    }

    if (this.workbook.Model.Worksheets.length === 1) {
      throw new CellsException("A workbook must contain at least one worksheet.");
    }

    this.workbook.Model.Worksheets.splice(index, 1);
    if (this.workbook.Model.ActiveSheetIndex > index) {
      this.workbook.Model.ActiveSheetIndex--;
    } else if (this.workbook.Model.ActiveSheetIndex >= this.workbook.Model.Worksheets.length) {
      this.workbook.Model.ActiveSheetIndex = this.workbook.Model.Worksheets.length - 1;
    }

    const firstSheet = this.workbook.Model.Properties.View.FirstSheet;
    if (firstSheet !== null && firstSheet !== undefined) {
      if (this.workbook.Model.Worksheets.length === 0) {
        this.workbook.Model.Properties.View.FirstSheet = 0;
      } else if (firstSheet >= this.workbook.Model.Worksheets.length) {
        this.workbook.Model.Properties.View.FirstSheet = this.workbook.Model.Worksheets.length - 1;
      }
    }
  }

  public *[Symbol.iterator](): IterableIterator<Worksheet> {
    for (const model of this.workbook.Model.Worksheets) {
      yield new Worksheet(this.workbook, model);
    }
  }

  private getByIndex(index: number): Worksheet {
    if (index < 0 || index >= this.workbook.Model.Worksheets.length) {
      throw new CellsException("Worksheet index was out of range.");
    }

    return new Worksheet(this.workbook, this.workbook.Model.Worksheets[index]);
  }

  private getByName(name: string): Worksheet {
    const index = this.indexOf(name);
    if (index < 0) {
      throw new CellsException(`Worksheet '${name}' was not found.`);
    }

    return new Worksheet(this.workbook, this.workbook.Model.Worksheets[index]);
  }

  private indexOf(name: string): number {
    if (name === null || name === undefined) {
      throw new Error("name");
    }

    for (let index = 0; index < this.workbook.Model.Worksheets.length; index++) {
      if (this.workbook.Model.Worksheets[index].Name.toLowerCase() === name.toLowerCase()) {
        return index;
      }
    }

    return -1;
  }

  private generateDefaultSheetName(): string {
    let suffix = 1;
    while (true) {
      const candidate = `Sheet${suffix}`;
      if (this.indexOf(candidate) < 0) {
        return candidate;
      }

      suffix++;
    }
  }
}

/**
 * Creates a worksheet collection that supports C#-style bracket indexing:
 * worksheets[0] and worksheets["Sheet1"].
 */
export function createWorksheetCollection(workbook: Workbook): WorksheetCollection {
  const collection = new WorksheetCollection(workbook);
  return createIndexerProxy(collection, (key) =>
    typeof key === "number" ? collection.get(key) : collection.get(key),
  );
}
