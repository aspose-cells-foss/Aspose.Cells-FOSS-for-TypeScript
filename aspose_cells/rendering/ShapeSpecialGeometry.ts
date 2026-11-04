import { Rect, PdfPath } from "./ShapeTypes";
import { ParsedShape } from "./ShapeData";
import { ShapeCustomGeometry } from "./ShapeCustomGeometry";
import { ShapeCurveGeometry } from "./ShapeCurveGeometry";
import { ShapeExtendedGeometry } from "./ShapeExtendedGeometry";

/**
 * 特殊几何：custom geometry、irregularSeal1、wedgeRectCallout、snip1Rect/snip2SameRect。
 * 其余交给 ShapeExtendedGeometry。
 */
export class ShapeSpecialGeometry {
  static build(shape: ParsedShape, rect: Rect): PdfPath | null {
    if (shape.customGeometry) {
      return ShapeCustomGeometry.build(shape.customGeometry, rect);
    }

    if (shape.geometry === "irregularSeal1") {
      // DrawingML irregularSeal1 顶点 (21600 x 21600 坐标系)
      const vertices = [
        10800, 5800, 14522, 0, 14155, 5325, 18380, 4457,
        16702, 7315, 21097, 8137, 17607, 10475, 21600, 13290, 16837, 12942,
        18145, 18095, 14020, 14457, 13247, 19737, 10532, 14935, 8485, 21600,
        7715, 15627, 4762, 17617, 5667, 13937, 135, 14587, 3722, 11775,
        0, 8615, 4627, 7617, 370, 2295, 7312, 6320, 8352, 2295,
      ];
      const path = new PdfPath();
      for (let i = 0; i < vertices.length; i += 2) {
        const x = rect.left + rect.width * vertices[i] / 21600;
        const y = rect.top + rect.height * vertices[i + 1] / 21600;
        if (i === 0) path.moveTo(x, y);
        else path.lineTo(x, y);
      }
      path.close();
      return path;
    }

    if (shape.geometry === "wedgeRectCallout") {
      return ShapeSpecialGeometry.buildCallout(shape.adjusts, rect);
    }

    if (shape.geometry === "snip1Rect" || shape.geometry === "snip2SameRect") {
      const top = Math.min(rect.width, rect.height)
        * Math.max(0, Math.min(0.5, ShapeCurveGeometry.adjustment(shape, 0, 0.16667)));
      const both = shape.geometry === "snip2SameRect";
      const bottom = both
        ? Math.min(rect.width, rect.height) * Math.max(0, Math.min(0.5, ShapeCurveGeometry.adjustment(shape, 1, 0)))
        : 0;
      const path = new PdfPath();
      path.moveTo(rect.left + (both ? top : 0), rect.top);
      path.lineTo(rect.right - top, rect.top);
      path.lineTo(rect.right, rect.top + top);
      path.lineTo(rect.right, rect.bottom - bottom);
      path.lineTo(rect.right - bottom, rect.bottom);
      path.lineTo(rect.left + bottom, rect.bottom);
      path.lineTo(rect.left, rect.bottom - bottom);
      path.lineTo(rect.left, rect.top + (both ? top : 0));
      path.close();
      return path;
    }

    return ShapeExtendedGeometry.build(shape, rect);
  }

  private static buildCallout(adjusts: number[] | null, r: Rect): PdfPath {
    const dx = adjusts && adjusts.length > 0 ? adjusts[0] : -0.20833;
    const dy = adjusts && adjusts.length > 1 ? adjusts[1] : 0.625;
    const tipX = r.midX + r.width * dx;
    const tipY = r.midY + r.height * dy;
    const tip = { x: tipX, y: tipY };
    const x1 = r.left + r.width * (dx > 0 ? 7 : 2) / 12;
    const x2 = r.left + r.width * (dx > 0 ? 10 : 5) / 12;
    const y1 = r.top + r.height * (dy > 0 ? 7 : 2) / 12;
    const y2 = r.top + r.height * (dy > 0 ? 10 : 5) / 12;
    const vertical = Math.abs(dy) > Math.abs(dx);
    const path = new PdfPath();
    path.moveTo(r.left, r.top);
    if (vertical && dy <= 0) {
      path.lineTo(x1, r.top); path.lineTo(tip.x, tip.y); path.lineTo(x2, r.top);
    }
    path.lineTo(r.right, r.top);
    if (!vertical && dx > 0) {
      path.lineTo(r.right, y1); path.lineTo(tip.x, tip.y); path.lineTo(r.right, y2);
    }
    path.lineTo(r.right, r.bottom);
    if (vertical && dy > 0) {
      path.lineTo(x2, r.bottom); path.lineTo(tip.x, tip.y); path.lineTo(x1, r.bottom);
    }
    path.lineTo(r.left, r.bottom);
    if (!vertical && dx <= 0) {
      path.lineTo(r.left, y2); path.lineTo(tip.x, tip.y); path.lineTo(r.left, y1);
    }
    path.close();
    return path;
  }
}
