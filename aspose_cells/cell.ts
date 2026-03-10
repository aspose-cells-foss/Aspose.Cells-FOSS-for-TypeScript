import type { CellValue, Style } from "./types";
import { cellRef as ref, escapeXml } from "./util";

function dateToExcelSerial(date: Date): number {
  const excelEpoch = new Date(1899, 11, 30);
  const diff = date.getTime() - excelEpoch.getTime();
  return Math.floor(diff / (1000 * 60 * 60 * 24));
}

export class Cell {
  static sharedStringMap: Map<string, number> = new Map();
  static setSharedStrings(strings: string[]) {
    Cell.sharedStringMap = new Map();
    strings.forEach((s, i) => Cell.sharedStringMap.set(s, i));
  }
  private _row: number;
  private _col: number;
  private _value: CellValue = null;
  private _formula?: string;
  private _style?: Style;
  private _styleIndex?: number;
  private _hyperlink?: string;

  constructor(row: number, col: number, value?: CellValue) {
    this._row = row;
    this._col = col;
    if (value !== undefined) this._value = value;
  }

  get row() {
    return this._row;
  }

  get col() {
    return this._col;
  }

  get ref() {
    return ref(this._row, this._col);
  }

  get value(): CellValue {
    return this._value;
  }

  set value(v: CellValue) {
    this._value = v;
  }

  putValue(value: CellValue) {
    this._value = value;
    return this;
  }

  get formula(): string | undefined {
    return this._formula;
  }

  set formula(f: string | undefined) {
    this._formula = f;
  }

  setFormula(formula: string) {
    this._formula = formula;
    return this;
  }

  get style(): Style | undefined {
    return this._style;
  }

  set style(s: Style | undefined) {
    this._style = s;
  }

  setStyle(style: Style) {
    this._style = style;
    return this;
  }

  setStyleIndex(index: number) {
    this._styleIndex = index;
  }

  get styleIndex(): number | undefined {
    return this._styleIndex;
  }

  get hyperlink(): string | undefined {
    return this._hyperlink;
  }

  setHyperlink(url: string) {
    this._hyperlink = url;
    return this;
  }

  toXml(): string {
    const coords = ref(this._row, this._col);
    let attrs = `r="${coords}"`;
    let valueXml = "";
    let typeAttr = "";

    if (this._value !== null && this._value !== undefined) {
      if (typeof this._value === "string") {
        const idx = Cell.sharedStringMap.get(this._value);
        if (idx !== undefined) {
          typeAttr = 't="s"';
          valueXml = `<v>${idx}</v>`;
        } else {
          typeAttr = 't="inlineStr"';
          valueXml = `<is><t>${escapeXml(this._value)}</t></is>`;
        }
      } else if (typeof this._value === "number") {
        valueXml = `<v>${this._value}</v>`;
      } else if (typeof this._value === "boolean") {
        valueXml = `<v>${this._value ? 1 : 0}</v>`;
      } else if (this._value instanceof Date) {
        const serial = dateToExcelSerial(this._value);
        valueXml = `<v>${serial}</v>`;
      }
    }

    if (this._formula) {
      valueXml = `<f>${escapeXml(this._formula)}</f>${valueXml}`;
    }

    const styleAttr =
      this._styleIndex !== undefined && this._styleIndex !== 0
        ? ` s="${this._styleIndex}"`
        : "";

    let attrsStr = " " + attrs;
    if (typeAttr) attrsStr += " " + typeAttr;
    if (styleAttr) attrsStr += styleAttr;

    return `<c${attrsStr}>${valueXml}</c>`;
  }
}

export class Cells {
  private cells: Map<string, Cell> = new Map();

  get(key: string): Cell | undefined {
    return this.cells.get(key);
  }

  set(key: string, cell: Cell) {
    this.cells.set(key, cell);
    return this;
  }

  getOrCreate(row: number, col: number): Cell {
    const key = ref(row, col);
    let cell = this.cells.get(key);
    if (!cell) {
      cell = new Cell(row, col);
      this.cells.set(key, cell);
    }
    return cell;
  }

  get count() {
    return this.cells.size;
  }

  *[Symbol.iterator]() {
    for (const cell of this.cells.values()) {
      yield cell;
    }
  }

  values() {
    return this.cells.values();
  }

  toXml(): string {
    const rows = new Map<number, Cell[]>();

    for (const cell of this.cells.values()) {
      const rowCells = rows.get(cell.row) || [];
      rowCells.push(cell);
      rows.set(cell.row, rowCells);
    }

    let xml = "";
    for (const [rowNum, rowCells] of rows) {
      xml += `<row r="${rowNum + 1}">`;
      const sortedCells = rowCells.sort((a, b) => a.col - b.col);
      for (const cell of sortedCells) {
        xml += cell.toXml();
      }
      xml += "</row>";
    }

    return xml;
  }
}
