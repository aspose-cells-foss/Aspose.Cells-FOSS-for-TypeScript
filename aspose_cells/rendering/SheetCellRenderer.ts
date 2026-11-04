import type PDFDocument from "pdfkit";
import { RenderContext } from "./RenderContext";
import { SheetLayout } from "./SheetLayout";
import { PageLayout } from "./PageLayout";

export class SheetCellRenderer {
  constructor(
    private doc: InstanceType<typeof PDFDocument>,
    private context: RenderContext,
    private layout: SheetLayout,
  ) {}

  renderCellGrid(page: PageLayout): void {
    const { startRow, endRow, startColumn, endColumn } = page;

    for (let r = startRow; r <= endRow && r <= this.layout.lastRow; r++) {
      if (this.layout.rowHidden[r]) continue;
      for (let c = startColumn; c <= endColumn && c <= this.layout.lastColumn; c++) {
        if (this.layout.columnHidden[c]) continue;
        this.renderCellText(r, c);
      }
    }
  }

  private renderCellText(row: number, col: number): void {
    const cellKey = `${row},${col}`;
    const cell = this.layout.sheet.cells.get(cellKey);
    if (!cell) return;
    if (cell.value === null || cell.value === undefined) return;
    if (this.layout.isMergeCovered(row, col)) return;

    const text = String(cell.value);
    if (!text) return;

    const x = this.layout.columnStartPt[col];
    const y = this.layout.rowStartPt[row];
    const w = this.layout.columnWidthPt[col];
    const h = this.layout.rowHeightPt[row];

    this.doc.font("Helvetica");
    this.doc.fontSize(11);
    this.doc.fillColor("#000000");

    const textX = x + SheetLayout.HorizontalPaddingPt;
    const textY = y + SheetLayout.VerticalPaddingPt + 2;
    const availableW = w - 2 * SheetLayout.HorizontalPaddingPt;

    this.doc.text(text, textX, textY, {
      width: availableW,
      ellipsis: true,
    });
  }
}
