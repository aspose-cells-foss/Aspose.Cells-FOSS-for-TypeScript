import type { ChartType, ChartSeries, ChartInfo } from "./types";
import type { Worksheet } from "./worksheet";

export class Chart {
  chartType: ChartType;
  title?: string;
  series: ChartSeries[] = [];
  legendPosition?: "left" | "right" | "top" | "bottom";
  categoryAxisPosition?: "left" | "right" | "top" | "bottom";
  valueAxisPosition?: "left" | "right" | "top" | "bottom";
  plotArea?: { x: number; y: number; width: number; height: number };
  fromRow: number = 0;
  fromCol: number = 0;
  toRow: number = 0;
  toCol: number = 0;
  fromRowOff?: number;
  fromColOff?: number;
  toRowOff?: number;
  toColOff?: number;
  originalX?: number;
  originalY?: number;
  sheetIndex: number = 0;
  type: string = "chart";
  name: string = "";
  originalChartXml?: string;
  originalStyleXml?: string;
  originalColorsXml?: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  dataRange?: string;
  dataIsVertical: boolean = true;
  private _worksheet?: Worksheet;

  constructor(type: ChartType, sheetIndex: number = 0) {
    this.chartType = type;
    this.sheetIndex = sheetIndex;
  }

  setPosition(topRow: number, leftColumn: number, bottomRow: number, rightColumn: number): Chart {
    this.fromRow = topRow;
    this.fromCol = leftColumn;
    this.toRow = bottomRow;
    this.toCol = rightColumn;
    return this;
  }

  setPositionWithOffset(topRow: number, top: number, leftColumn: number, left: number, height: number, width: number): Chart {
    this.fromRow = topRow;
    this.fromRowOff = top;
    this.fromCol = leftColumn;
    this.fromColOff = left;
    this.toRow = topRow + height - 1;
    this.toCol = leftColumn + width - 1;
    this.toRowOff = 0;
    this.toColOff = 0;
    return this;
  }

  setDataRange(dataRange: string, isVertical: boolean): Chart {
    this.dataRange = dataRange;
    this.dataIsVertical = isVertical;
    return this;
  }

  setWorksheet(worksheet: Worksheet): void {
    this._worksheet = worksheet;
    this.sheetIndex = worksheet.index;
    this.rebuildSeriesFromWorksheet();
  }

  private rebuildSeriesFromWorksheet(): void {
    if (!this.dataRange || !this._worksheet) return;

    const range = this.dataRange.replace(/^[^!]*!/, "");
    const match = range.match(/^(\$?[A-Z]+)(\$?\d+):(\$?[A-Z]+)(\$?\d+)$/);
    if (!match) return;

    const startCol = match[1].replace("$", "");
    const startRow = parseInt(match[2].replace("$", ""), 10) - 1;
    const endCol = match[3].replace("$", "");
    const endRow = parseInt(match[4].replace("$", ""), 10) - 1;

    const startColIdx = startCol.charCodeAt(0) - 65;
    const endColIdx = endCol.charCodeAt(0) - 65;
    const numCols = endColIdx - startColIdx + 1;
    const numRows = endRow - startRow + 1;

    const hasHeader = startRow === 0;
    const dataStartRow = hasHeader ? startRow + 1 : startRow;
    const categoryStartColIdx = startColIdx;
    const seriesStartColIdx = hasHeader ? startColIdx + 1 : startColIdx;
    const seriesNumCols = hasHeader ? numCols - 1 : numCols;

    const categories: string[] = [];
    for (let r = dataStartRow; r <= endRow; r++) {
      const cell = this._worksheet.getCell(r, categoryStartColIdx);
      categories.push(String(cell?.value ?? `Item ${r}`));
    }

    this.series = [];
    for (let c = 0; c < seriesNumCols; c++) {
      const colLetter = String.fromCharCode(65 + seriesStartColIdx + c);
      const nameCell = hasHeader 
        ? this._worksheet.getCell(startRow, seriesStartColIdx + c)
        : this._worksheet.getCell(startRow, seriesStartColIdx + c);
      const values: number[] = [];
      for (let r = dataStartRow; r <= endRow; r++) {
        const cell = this._worksheet.getCell(r, seriesStartColIdx + c);
        const val = cell?.value;
        if (typeof val === "number") {
          values.push(val);
        } else if (typeof val === "string") {
          values.push(parseFloat(val) || 0);
        } else {
          values.push(0);
        }
      }
      this.series.push({
        name: String(nameCell?.value ?? `${colLetter}${dataStartRow + 1}`),
        categories: [...categories],
        values: values,
      });
    }
  }

  setTitle(title: string): Chart {
    this.title = title;
    return this;
  }

  setLegendPosition(position: "left" | "right" | "top" | "bottom"): Chart {
    this.legendPosition = position;
    return this;
  }
}