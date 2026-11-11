import { LegacyChart } from "./chart";
import type { ChartType as LegacyChartType, ShapeInfo } from "./types";
import { LegacyWorksheet } from "./worksheet";

export class LegacyChartCollection {
  private _charts: LegacyChart[] = [];
  private _sheetIndex: number = 0;
  private _worksheet?: LegacyWorksheet;

  constructor(sheetIndex: number = 0, worksheet?: LegacyWorksheet) {
    this._sheetIndex = sheetIndex;
    this._worksheet = worksheet;
  }

  setWorksheet(worksheet: LegacyWorksheet): void {
    this._worksheet = worksheet;
  }

  [index: number]: LegacyChart | undefined;

  get length(): number {
    return this._charts.length;
  }

  get count(): number {
    return this._charts.length;
  }

  get(index: number): LegacyChart | undefined {
    return this._charts[index];
  }

  getByName(name: string): LegacyChart | undefined {
    return this._charts.find((c) => c.name === name);
  }

  add(type: LegacyChartType, topRow: number, leftColumn: number, bottomRow: number, rightColumn: number): number {
    const chart = new LegacyChart(type, this._sheetIndex);
    chart.fromRow = topRow;
    chart.fromCol = leftColumn;
    chart.toRow = bottomRow;
    chart.toCol = rightColumn;
    chart.name = `LegacyChart ${this._charts.length + 1}`;
    this._charts.push(chart);
    if (this._worksheet) {
      this._worksheet.addShape(chart as unknown as ShapeInfo);
    }
    return this._charts.length - 1;
  }

  addWithDataRange(
    type: LegacyChartType,
    dataRange: string,
    isVertical: boolean,
    topRow: number,
    leftColumn: number,
    rightRow: number,
    bottomColumn: number
  ): number {
    const chart = new LegacyChart(type, this._sheetIndex);
    chart.setDataRange(dataRange, isVertical);
    if (this._worksheet) {
      chart.setWorksheet(this._worksheet);
    }
    chart.fromRow = topRow;
    chart.fromCol = leftColumn;
    chart.toRow = rightRow;
    chart.toCol = bottomColumn;
    chart.name = `LegacyChart ${this._charts.length + 1}`;
    this._charts.push(chart);
    if (this._worksheet) {
      this._worksheet.addShape(chart as unknown as ShapeInfo);
    }
    return this._charts.length - 1;
  }

  addWithOffset(
    type: LegacyChartType,
    topRow: number,
    top: number,
    leftColumn: number,
    left: number,
    height: number,
    width: number
  ): number {
    const chart = new LegacyChart(type, this._sheetIndex);
    chart.setPositionWithOffset(topRow, top, leftColumn, left, height, width);
    chart.name = `LegacyChart ${this._charts.length + 1}`;
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

  toArray(): LegacyChart[] {
    return [...this._charts];
  }

  getChartInfos(): LegacyChart[] {
    return this._charts;
  }
}

import { CellsException } from "./CellsException";
import { Chart } from "./chart";
import { ChartType } from "./enums";
import { ChartModel } from "./Core/ChartModel";
import { ChartXmlTemplates } from "./ChartXmlTemplates";

// Implemented by serializer module.
/** Represents a collection of charts on a worksheet. */
export class ChartCollection {
  private readonly models: ChartModel[];

  public constructor(models: ChartModel[]) {
    this.models = models;
  }

  public get Count(): number {
    return this.models.length;
  }

  public get(index: number): Chart {
    if (index < 0 || index >= this.models.length) {
      throw new CellsException(`Chart index ${index} is out of range.`);
    }

    return new Chart(this.models[index]);
  }

  public at(index: number): Chart {
    return this.get(index);
  }

  public Add(
    type: ChartType,
    dataRange: string,
    upperLeftRow: number,
    upperLeftColumn: number,
    lowerRightRow: number,
    lowerRightColumn: number,
  ): number {
    if (dataRange === null || dataRange === undefined || dataRange.trim().length === 0) {
      throw new CellsException("Chart data range must be non-empty.");
    }

    if (upperLeftRow < 0 || upperLeftColumn < 0) {
      throw new CellsException("Chart anchor indices must be non-negative.");
    }

    if (lowerRightRow < upperLeftRow || lowerRightColumn < upperLeftColumn) {
      throw new CellsException("Chart anchor end must not precede the start.");
    }

    const model = new ChartModel();
    const chartName = `Chart ${this.models.length + 1}`;
    model.Name = chartName;
    model.ChartType = type as unknown as import("./ChartType").ChartType;
    model.DataRange = dataRange;
    // Modern chartex types throw from the template builder, matching ChartCollection.Add in .NET.
    model.RawChartXml = ChartXmlTemplates.Build(model.ChartType, dataRange, chartName);
    model.UpperLeftRow = upperLeftRow;
    model.UpperLeftColumn = upperLeftColumn;
    model.LowerRightRow = lowerRightRow;
    model.LowerRightColumn = lowerRightColumn;
    this.models.push(model);
    return this.models.length - 1;
  }

  public *[Symbol.iterator](): IterableIterator<Chart> {
    for (let index = 0; index < this.models.length; index++) {
      yield this.get(index);
    }
  }
}
