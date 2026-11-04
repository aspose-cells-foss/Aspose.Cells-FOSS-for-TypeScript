import type PDFDocument from "pdfkit";
import { Rect } from "./ShapeTypes";
import { ParsedShape } from "./ShapeData";

/** 形状内的自动换行文本。按空格分词，宽度通过 pdfkit widthOfString 测定。 */
export class ShapeWrappedTextRenderer {
  static draw(doc: InstanceType<typeof PDFDocument>, rect: Rect, shape: ParsedShape): void {
    if (!shape.text) return;
    const availableWidth = Math.max(1, rect.width - shape.textInsetLeftPt * 2);
    const lines: string[] = [];
    const words = shape.text.split(/\s+/).filter(w => w.length > 0);
    let current = "";
    for (const word of words) {
      const candidate = current.length === 0 ? word : current + " " + word;
      if (current.length > 0 && doc.widthOfString(candidate) > availableWidth) {
        lines.push(current);
        current = word;
      } else {
        current = candidate;
      }
    }
    if (current.length > 0) lines.push(current);
    if (lines.length === 0) return;

    doc.save();
    doc.rect(rect.left, rect.top, rect.width, rect.height);
    doc.clip();
    doc.fontSize(shape.textSizePt);
    doc.fillColor(shape.textColor.toPdfKitColor());
    if (shape.textBold) doc.font("Helvetica-Bold"); else doc.font("Helvetica");

    const lineHeight = shape.textSizePt * 1.2;
    // 近似文本高度 = 最后一行的下降 + 行间距
    const textHeight = lines.length * lineHeight - shape.textSizePt * 0.3;
    let baseline: number;
    if (shape.textVAnchor === "t") baseline = rect.top + shape.textInsetTopPt;
    else if (shape.textVAnchor === "b") baseline = rect.bottom - textHeight;
    else baseline = rect.midY - textHeight / 2;

    for (const line of lines) {
      const lineWidth = doc.widthOfString(line);
      let x: number;
      if (shape.textHAlign === "l") x = rect.left + shape.textInsetLeftPt;
      else if (shape.textHAlign === "r") x = rect.right - shape.textInsetLeftPt - lineWidth;
      else x = rect.midX - lineWidth / 2;
      doc.text(line, x, baseline);
      baseline += lineHeight;
    }
    doc.restore();
  }
}
