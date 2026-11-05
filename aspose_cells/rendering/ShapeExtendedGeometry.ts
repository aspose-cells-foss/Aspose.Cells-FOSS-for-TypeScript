import { Rect, PdfPath } from "./ShapeTypes";
import { ParsedShape } from "./ShapeData";
import { ShapeCurveGeometry } from "./ShapeCurveGeometry";

/**
 * Extended preset geometries: leftArrowCallout, mathNotEqual, flowChartMagneticDrum, etc.
 * Also provides the polygon() and adjustment() helpers.
 */
export class ShapeExtendedGeometry {
  static build(shape: ParsedShape, r: Rect): PdfPath | null {
    if (r.width <= 0 || r.height <= 0) return null;
    if (shape.geometry === "leftArrowCallout") return ShapeExtendedGeometry.leftArrowCallout(shape, r);
    if (shape.geometry === "mathNotEqual") return ShapeExtendedGeometry.notEqual(shape, r);
    if (shape.geometry === "flowChartMagneticDrum") {
      const path = new PdfPath();
      path.moveTo(r.left + r.width / 6, r.top);
      path.lineTo(r.right - r.width / 6, r.top);
      path.arcTo(new Rect(r.right - r.width / 3, r.top, r.right, r.bottom), -90, 180, false);
      path.lineTo(r.left + r.width / 6, r.bottom);
      path.arcTo(new Rect(r.left, r.top, r.left + r.width / 3, r.bottom), 90, 180, false);
      path.close();
      return path;
    }
    return null;
  }

  static adjustment(shape: ParsedShape, index: number, fallback: number): number {
    const value = ShapeCurveGeometry.adjustment(shape, index, fallback);
    if (isNaN(value) || !isFinite(value)) return fallback;
    return value;
  }

  private static clamp(value: number, maximum: number): number {
    return Math.max(0, Math.min(maximum, value));
  }

  private static leftArrowCallout(shape: ParsedShape, r: Rect): PdfPath {
    const ss = Math.min(r.width, r.height);
    const a2 = ShapeExtendedGeometry.clamp(ShapeExtendedGeometry.adjustment(shape, 1, 0.25), r.height / (2 * ss));
    const a1 = ShapeExtendedGeometry.clamp(ShapeExtendedGeometry.adjustment(shape, 0, 0.25), a2 * 2);
    const a3 = ShapeExtendedGeometry.clamp(ShapeExtendedGeometry.adjustment(shape, 2, 0.25), r.width / ss);
    const a4 = ShapeExtendedGeometry.clamp(ShapeExtendedGeometry.adjustment(shape, 3, 0.64977), 1 - a3 * ss / r.width);
    const x1 = r.left + ss * a3;
    const x2 = r.right - r.width * a4;
    const y1 = r.midY - ss * a2;
    const y2 = r.midY - ss * a1 / 2;
    const y3 = r.midY + ss * a1 / 2;
    const y4 = r.midY + ss * a2;
    return ShapeExtendedGeometry.polygon([
      r.left, r.midY,
      x1, y1, x1, y2, x2, y2, x2, r.top,
      r.right, r.top, r.right, r.bottom, x2, r.bottom, x2, y3, x1, y3, x1, y4,
    ]);
  }

  private static notEqual(shape: ParsedShape, r: Rect): PdfPath {
    const w = r.width;
    const h = r.height;
    const a1 = ShapeExtendedGeometry.clamp(ShapeExtendedGeometry.adjustment(shape, 0, 0.2352), 0.5);
    const a3 = ShapeExtendedGeometry.clamp(ShapeExtendedGeometry.adjustment(shape, 2, 0.1176), 1 - 2 * a1);
    const angle = Math.max(42, Math.min(66, ShapeExtendedGeometry.adjustment(shape, 1, 66))) / 0.6 - 90;
    const dx = h / 2 * Math.tan(angle * Math.PI / 180);
    const length = Math.sqrt(dx * dx + h * h / 4);
    const thickness = h * a1;
    const bandWidth = length * thickness / (h / 2);
    const x1 = w * (0.5 - 0.36745);
    const x8 = w - x1;
    const y2 = h * (0.5 - a3 / 2);
    const y3 = h - y2;
    const y1 = y2 - thickness;
    const y4 = y3 + thickness;
    const x7 = w / 2 + dx - bandWidth / 2;
    const x6 = x7 - dx * y1 / (h / 2);
    const x5 = x7 - dx * y2 / (h / 2);
    const x4 = x7 - dx * y3 / (h / 2);
    const x3 = x7 - dx * y4 / (h / 2);
    const dx7 = thickness * h / 2 / length;
    const dy = thickness * dx / length;
    const rx = angle > 0 ? x7 + dx7 : x7 + bandWidth;
    const lx = angle > 0 ? x7 : x7 + bandWidth - dx7;
    const ry = angle > 0 ? dy : 0;
    const ly = angle > 0 ? 0 : -dy;
    const pts = [x1, y1, x6, y1, lx, ly, rx, ry, x6 + bandWidth, y1,
      x8, y1, x8, y2, x5 + bandWidth, y2, x4 + bandWidth, y3, x8, y3, x8, y4,
      x3 + bandWidth, y4, w - lx, h - ly, w - rx, h - ry, x3, y4, x1, y4, x1, y3, x4, y3, x5, y2, x1, y2];
    for (let i = 0; i < pts.length; i += 2) {
      pts[i] += r.left;
      pts[i + 1] += r.top;
    }
    return ShapeExtendedGeometry.polygon(pts);
  }

  static polygon(points: number[]): PdfPath {
    const path = new PdfPath();
    path.moveTo(points[0], points[1]);
    for (let i = 2; i < points.length; i += 2) path.lineTo(points[i], points[i + 1]);
    path.close();
    return path;
  }
}
