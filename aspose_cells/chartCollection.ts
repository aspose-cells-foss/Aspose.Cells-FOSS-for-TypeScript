import { Chart } from "./chart";
import type { ChartType, ShapeInfo } from "./types";
import { Worksheet } from "./worksheet";

export class ChartCollection {
  private _charts: Chart[] = [];
  private _sheetIndex: number = 0;
  private _worksheet?: Worksheet;

  constructor(sheetIndex: number = 0, worksheet?: Worksheet) {
    this._sheetIndex = sheetIndex;
    this._worksheet = worksheet;
  }

  setWorksheet(worksheet: Worksheet): void {
    this._worksheet = worksheet;
  }

  [index: number]: Chart | undefined;

  get length(): number {
    return this._charts.length;
  }

  get count(): number {
    return this._charts.length;
  }

  get(index: number): Chart | undefined {
    return this._charts[index];
  }

  getByName(name: string): Chart | undefined {
    return this._charts.find((c) => c.name === name);
  }

  add(type: ChartType, topRow: number, leftColumn: number, bottomRow: number, rightColumn: number): number {
    const chart = new Chart(type, this._sheetIndex);
    chart.fromRow = topRow;
    chart.fromCol = leftColumn;
    chart.toRow = bottomRow;
    chart.toCol = rightColumn;
    chart.name = `Chart ${this._charts.length + 1}`;
    this._charts.push(chart);
    if (this._worksheet) {
      this._worksheet.addShape(chart as unknown as ShapeInfo);
    }
    return this._charts.length - 1;
  }

  addWithDataRange(
    type: ChartType,
    dataRange: string,
    isVertical: boolean,
    topRow: number,
    leftColumn: number,
    rightRow: number,
    bottomColumn: number
  ): number {
    const chart = new Chart(type, this._sheetIndex);
    chart.setDataRange(dataRange, isVertical);
    if (this._worksheet) {
      chart.setWorksheet(this._worksheet);
    }
    chart.fromRow = topRow;
    chart.fromCol = leftColumn;
    chart.toRow = rightRow;
    chart.toCol = bottomColumn;
    chart.name = `Chart ${this._charts.length + 1}`;
    this._charts.push(chart);
    if (this._worksheet) {
      this._worksheet.addShape(chart as unknown as ShapeInfo);
    }
    return this._charts.length - 1;
  }

  addWithOffset(
    type: ChartType,
    topRow: number,
    top: number,
    leftColumn: number,
    left: number,
    height: number,
    width: number
  ): number {
    const chart = new Chart(type, this._sheetIndex);
    chart.setPositionWithOffset(topRow, top, leftColumn, left, height, width);
    chart.name = `Chart ${this._charts.length + 1}`;
    this._charts.push(chart);
    if (this._worksheet) {
      this._worksheet.addShape(chart as unknown as ShapeInfo);
    }
    return this._charts.length - 1;
  }

  removeAt(index: number): void {
    if (index >= 0 && index < this._charts.length) {
      this._charts.splice(index, 1);
    }
  }

  clear(): void {
    this._charts.length = 0;
  }

  toArray(): Chart[] {
    return [...this._charts];
  }

  getChartInfos(): Chart[] {
    return this._charts;
  }
}