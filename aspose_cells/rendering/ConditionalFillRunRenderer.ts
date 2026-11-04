import type PDFDocument from "pdfkit";
import { SheetLayout } from "./SheetLayout";
import { RgbaColor } from "./ColorValue";

/**
 * 条件格式填充渲染器。在 PageLayout 范围的各列内，连续行合并为一次矩形绘制。
 * 依赖 ConditionalFormatEvaluator 和 TableStyleResolver，目前为可选传入（any 占位）。
 */
export class ConditionalFillRunRenderer {
  static draw(
    doc: InstanceType<typeof PDFDocument>,
    page: { startRow: number; endRow: number; startColumn: number; endColumn: number; sheet: SheetLayout },
    formatting: any, // ConditionalFormatEvaluator
    tableStyles: any, // TableStyleResolver
  ): number {
    const layout = page.sheet;
    let count = 0;
    for (let column = page.startColumn; column <= page.endColumn; column++) {
      let firstRow = -1;
      let runColor: RgbaColor | null = null;
      for (let row = page.startRow; row <= page.endRow + 1; row++) {
        let color: RgbaColor | null = null;
        let hasFill = false;
        if (row <= page.endRow) {
          hasFill = ConditionalFillRunRenderer.tryGetFill(layout, formatting, tableStyles, row, column);
          if (hasFill && formatting && formatting.tryGet) {
            const effect = formatting.tryGet(row, column);
            if (effect && effect.hasFill) {
              color = effect.fill as RgbaColor;
            }
          }
        }
        if (firstRow >= 0 && (!hasFill || !ConditionalFillRunRenderer.colorEqual(runColor!, color!))) {
          const left = layout.columnStartPt[column] - layout.columnStartPt[page.startColumn];
          const right = layout.columnStartPt[column + 1] - layout.columnStartPt[page.startColumn];
          const top = layout.rowStartPt[firstRow] - layout.rowStartPt[page.startRow];
          const bottom = layout.rowStartPt[row] - layout.rowStartPt[page.startRow];
          doc.save();
          doc.fillColor((runColor as RgbaColor).toPdfKitColor());
          doc.rect(left, top, right - left, bottom - top);
          doc.fill();
          doc.restore();
          count++;
          firstRow = -1;
        }
        if (hasFill && firstRow < 0) {
          firstRow = row;
          runColor = color!;
        }
      }
    }
    return count;
  }

  private static colorEqual(a: RgbaColor, b: RgbaColor): boolean {
    if (!a || !b) return a === b;
    return a.r === b.r && a.g === b.g && a.b === b.b && a.a === b.a;
  }

  private static tryGetFill(layout: SheetLayout, formatting: any, tableStyles: any, row: number, column: number): boolean {
    if (layout.rowHidden[row] || layout.columnHidden[column]) return false;
    if (layout.isMergeCovered(row, column)) return false;
    if (!formatting) return false;
    if (!formatting.tryGet) return false;
    const effect = formatting.tryGet(row, column);
    if (!effect || !effect.hasFill) return false;
    if (effect.fill && effect.fill.a === 0) return false;
    return true;
  }
}
