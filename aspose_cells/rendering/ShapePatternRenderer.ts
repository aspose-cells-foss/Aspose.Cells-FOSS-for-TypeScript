import type PDFDocument from "pdfkit";
import { Rect, PdfPath } from "./ShapeTypes";
import { ParsedShape } from "./ShapeData";

/** DrawingML pct5 图案填充：16x16 tile 内两个 2x2 dot，按 pdfkit 单位适配。 */
export class ShapePatternRenderer {
  static draw(doc: InstanceType<typeof PDFDocument>, rect: Rect, outline: PdfPath, shape: ParsedShape): void {
    if (shape.fillPattern !== "pct5") return;
    // DrawingML pct5: 在 16x16 tile 中有两个 2x2 dots，按 192 dpi 渲染
    const tile = 6;
    const dot = 0.75;
    doc.save();
    // 用 clip 限定在形状轮廓内
    outline.draw(doc);
    doc.clip();

    doc.fillColor(shape.patternForeground.toPdfKitColor());
    for (let y = rect.top; y < rect.bottom; y += tile) {
      for (let x = rect.left; x < rect.right; x += tile) {
        doc.rect(x, y, dot, dot);
        doc.fill();
        doc.rect(x + tile / 2, y + tile / 2, dot, dot);
        doc.fill();
      }
    }
    doc.restore();
  }
}
