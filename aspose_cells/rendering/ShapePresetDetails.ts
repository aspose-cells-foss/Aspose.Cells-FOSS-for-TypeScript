import type PDFDocument from "pdfkit";
import { Rect, PdfPath } from "./ShapeTypes";
import { ParsedShape } from "./ShapeData";
import { RgbaColor } from "./ColorValue";
import { ShapeExtendedGeometry } from "./ShapeExtendedGeometry";

/**
 * 特定预设几何的额外绘制：bevel 立体面填充、额外叠加线条等。
 */
export class ShapePresetDetails {
  static drawFill(doc: InstanceType<typeof PDFDocument>, r: Rect, shape: ParsedShape): boolean {
    if (shape.geometry !== "bevel") return false;
    const inset = ShapePresetDetails.bevelInset(shape, r);
    const x1 = r.left + inset;
    const x2 = r.right - inset;
    const y1 = r.top + inset;
    const y2 = r.bottom - inset;

    // Disjoint faces: 在透明面之下不绘制基底矩形
    ShapePresetDetails.face(doc, shape.fillColor, [x1, y1, x2, y1, x2, y2, x1, y2]);
    ShapePresetDetails.face(doc, ShapePresetDetails.faceColor(shape.fillColor, 0.2), [r.left, r.top, r.right, r.top, x2, y1, x1, y1]);
    ShapePresetDetails.face(doc, ShapePresetDetails.faceColor(shape.fillColor, -0.2), [r.left, r.bottom, x1, y2, x2, y2, r.right, r.bottom]);
    ShapePresetDetails.face(doc, ShapePresetDetails.faceColor(shape.fillColor, 0.4), [r.left, r.top, x1, y1, x1, y2, r.left, r.bottom]);
    ShapePresetDetails.face(doc, ShapePresetDetails.faceColor(shape.fillColor, -0.4), [r.right, r.top, r.right, r.bottom, x2, y2, x2, y1]);
    return true;
  }

  static drawLines(doc: InstanceType<typeof PDFDocument>, r: Rect, shape: ParsedShape): void {
    if (!shape.hasLine) return;
    doc.save();
    doc.strokeColor(shape.lineColor.toPdfKitColor());
    doc.lineWidth(Math.max(0.5, shape.lineWidthPt));

    if (shape.geometry === "flowChartMagneticDrum") {
      // 简化：画矩形
      doc.rect(r.right - r.width / 3, r.top, r.width / 3, r.height);
      doc.stroke();
    } else if (shape.geometry === "accentBorderCallout3") {
      const defaults = [0.1875, -0.08333, 0.1875, -0.16667, 1, -0.16667, 1.12963, -0.08333];
      const points: Array<{ x: number; y: number }> = [];
      for (let i = 0; i < 4; i++) {
        points.push({
          x: r.left + r.width * ShapeExtendedGeometry.adjustment(shape, i * 2 + 1, defaults[i * 2 + 1]),
          y: r.top + r.height * ShapeExtendedGeometry.adjustment(shape, i * 2, defaults[i * 2]),
        });
      }
      doc.moveTo(points[0].x, r.top);
      doc.lineTo(points[0].x, r.bottom);
      doc.stroke();
      doc.moveTo(points[0].x, points[0].y);
      for (let i = 1; i < points.length; i++) {
        doc.lineTo(points[i].x, points[i].y);
      }
      doc.stroke();
    } else if (shape.geometry === "bevel") {
      const inset = ShapePresetDetails.bevelInset(shape, r);
      const inner = new Rect(r.left + inset, r.top + inset, r.right - inset, r.bottom - inset);
      doc.rect(inner.left, inner.top, inner.width, inner.height);
      doc.stroke();
      doc.moveTo(r.left, r.top); doc.lineTo(inner.left, inner.top);
      doc.moveTo(r.right, r.top); doc.lineTo(inner.right, inner.top);
      doc.moveTo(r.right, r.bottom); doc.lineTo(inner.right, inner.bottom);
      doc.moveTo(r.left, r.bottom); doc.lineTo(inner.left, inner.bottom);
      doc.stroke();
    } else {
      doc.restore();
      return;
    }
    doc.restore();
  }

  private static bevelInset(shape: ParsedShape, r: Rect): number {
    return Math.min(r.width, r.height)
      * Math.max(0, Math.min(0.5, ShapeExtendedGeometry.adjustment(shape, 0, 0.125)));
  }

  private static face(doc: InstanceType<typeof PDFDocument>, color: RgbaColor, points: number[]): void {
    doc.save();
    doc.fillColor(color.toPdfKitColor());
    doc.moveTo(points[0], points[1]);
    for (let i = 2; i < points.length; i += 2) doc.lineTo(points[i], points[i + 1]);
    doc.closePath();
    doc.fill();
    doc.restore();
  }

  private static faceColor(color: RgbaColor, amount: number): RgbaColor {
    return new RgbaColor(
      ShapePresetDetails.channel(color.r, amount),
      ShapePresetDetails.channel(color.g, amount),
      ShapePresetDetails.channel(color.b, amount),
      color.a,
    );
  }

  private static channel(value: number, amount: number): number {
    return Math.round(amount >= 0 ? value + (255 - value) * amount : value * (1 + amount));
  }
}
