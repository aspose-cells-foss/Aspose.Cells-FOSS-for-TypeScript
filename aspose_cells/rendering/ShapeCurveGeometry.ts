import { Point } from "./ShapeTypes";
import { ParsedShape } from "./ShapeData";

/** 曲线连接器几何计算（curvedConnector3 / curvedConnector4）。 */
export class ShapeCurveGeometry {
  static build(start: Point, end: Point, shape: ParsedShape): Point[] {
    if (shape.geometry === "curvedConnector4") {
      const x2 = ShapeCurveGeometry.adjustment(shape, 0, 0.5);
      const y4 = ShapeCurveGeometry.adjustment(shape, 1, 0.5);
      const x3 = (1 + x2) / 2;
      return ShapeCurveGeometry.map([
        new Point(0, 0),
        new Point(x2 / 2, 0),
        new Point(x2, y4 / 4),
        new Point(x2, y4 / 2),
        new Point(x2, y4 * 0.75),
        new Point((x2 + x3) / 2, y4),
        new Point(x3, y4),
        new Point((x3 + 1) / 2, y4),
        new Point(1, (1 + y4) / 2),
        new Point(1, 1),
      ], start, end, shape);
    }
    if (shape.geometry !== "curvedConnector3") {
      return [start, new Point((start.x + end.x) / 2, start.y), new Point(end.x, (start.y + end.y) / 2), end];
    }
    const adj = shape.adjusts && shape.adjusts.length > 0 ? shape.adjusts[0] : 0.5;
    const normalized = [
      new Point(0, 0),
      new Point(adj / 2, 0),
      new Point(adj, 0.25),
      new Point(adj, 0.5),
      new Point(adj, 0.75),
      new Point((1 + adj) / 2, 1),
      new Point(1, 1),
    ];
    return ShapeCurveGeometry.map(normalized, start, end, shape);
  }

  static adjustment(shape: ParsedShape, index: number, fallback: number): number {
    if (shape.adjusts && shape.adjusts.length > index) return shape.adjusts[index];
    return fallback;
  }

  /** 将归一化坐标映射到实际起止点之间。 */
  static map(normalized: Point[], start: Point, end: Point, shape: ParsedShape): Point[] {
    const vertical = Math.abs(Math.sin(shape.rotationDeg * Math.PI / 180)) > 0.707;
    for (let i = 0; i < normalized.length; i++) {
      const nx = vertical ? normalized[i].y : normalized[i].x;
      const ny = vertical ? normalized[i].x : normalized[i].y;
      normalized[i] = new Point(
        start.x + (end.x - start.x) * nx,
        start.y + (end.y - start.y) * ny,
      );
    }
    return normalized;
  }
}
