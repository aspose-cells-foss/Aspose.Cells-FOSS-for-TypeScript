import { Point } from "./ShapeTypes";
import { ParsedShape } from "./ShapeData";
import { ShapeCurveGeometry } from "./ShapeCurveGeometry";

/** 折线连接器几何计算（bentConnector2/3/4, straightConnector）。 */
export class ShapePolylineGeometry {
  static build(start: Point, end: Point, shape: ParsedShape): Point[] {
    const a1 = ShapeCurveGeometry.adjustment(shape, 0, 0.5);
    const a2 = ShapeCurveGeometry.adjustment(shape, 1, 0.5);
    let points: Point[];
    switch (shape.geometry) {
      case "bentConnector2":
        points = [new Point(0, 0), new Point(1, 0), new Point(1, 1)];
        break;
      case "bentConnector3":
        points = [new Point(0, 0), new Point(a1, 0), new Point(a1, 1), new Point(1, 1)];
        break;
      case "bentConnector4":
        points = [new Point(0, 0), new Point(a1, 0), new Point(a1, a2), new Point(1, a2), new Point(1, 1)];
        break;
      default:
        return [start, end];
    }
    return ShapeCurveGeometry.map(points, start, end, shape);
  }
}
