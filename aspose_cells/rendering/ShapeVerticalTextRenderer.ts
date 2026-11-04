import type PDFDocument from "pdfkit";
import { Rect } from "./ShapeTypes";
import { ParsedShape } from "./ShapeData";
import { RgbaColor } from "./ColorValue";

/**
 * 垂直文字渲染器。pdfkit 不提供直接的 FontMetrics，需要通过 fontOf 和 widthOfString 近似。
 */
export class ShapeVerticalTextRenderer {
  static isVerticalText(shape: ParsedShape | null): boolean {
    if (!shape || !shape.textVerticalType) return false;
    return shape.textVerticalType === "vert"
      || shape.textVerticalType === "vert270"
      || shape.textVerticalType === "wordArtVert"
      || shape.textVerticalType === "wordArtVertRtl";
  }

  static draw(doc: InstanceType<typeof PDFDocument>, rect: Rect, shape: ParsedShape): void {
    if (shape.textVerticalType === "wordArtVert" || shape.textVerticalType === "wordArtVertRtl") {
      ShapeVerticalTextRenderer.drawWordArtVerticalText(doc, rect, shape);
      return;
    }
    const clockwise = shape.textVerticalType === "vert270" || shape.textVerticalType === "wordArtVertRtl";
    doc.save();
    doc.translate(rect.midX, rect.midY);
    doc.rotate(clockwise ? 90 : -90);

    doc.fontSize(shape.textSizePt);
    doc.fillColor(shape.textColor.toPdfKitColor());
    if (shape.textBold) doc.font("Helvetica-Bold"); else doc.font("Helvetica");

    const textWidth = doc.widthOfString(shape.text ?? "");
    const drawX = -textWidth / 2;
    const drawY = -shape.textSizePt / 3; // 近似 baseline

    doc.text(shape.text ?? "", drawX, drawY, { align: "center" });
    doc.restore();
  }

  private static drawWordArtVerticalText(doc: InstanceType<typeof PDFDocument>, rect: Rect, shape: ParsedShape): void {
    if (!shape.text) return;

    const charAdvance = Math.max(shape.textSizePt, shape.textSizePt * 1.2);
    const availableHeight = Math.max(charAdvance, rect.height - 2 * shape.textInsetTopPt);
    const rows = Math.max(1, Math.floor(availableHeight / charAdvance));
    const columns = Math.ceil(shape.text.length / rows);
    const totalHeight = charAdvance * Math.min(rows, shape.text.length);
    const currentY = rect.midY - totalHeight * 0.5;

    doc.save();
    doc.rect(rect.left, rect.top, rect.width, rect.height);
    doc.clip();
    doc.fontSize(shape.textSizePt);
    doc.fillColor(shape.textColor.toPdfKitColor());
    if (shape.textBold) doc.font("Helvetica-Bold"); else doc.font("Helvetica");

    for (let i = 0; i < shape.text.length; i++) {
      const ch = shape.text[i];
      let column = Math.floor(i / rows);
      if (shape.textVerticalType === "wordArtVertRtl") column = columns - 1 - column;
      const centerX = rect.midX + (column - (columns - 1) * 0.5) * charAdvance;
      const width = doc.widthOfString(ch);
      const drawX = centerX - width * 0.5;
      const drawY = currentY + (i % rows) * charAdvance;
      doc.text(ch, drawX, drawY);
    }
    doc.restore();
  }
}
