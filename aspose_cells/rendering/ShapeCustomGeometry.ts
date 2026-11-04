import { Rect, Point, PdfPath } from "./ShapeTypes";

/** 命名空间常量。 */
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

function attrOf(el: Element | null, name: string): string | null {
  if (!el) return null;
  const a = el.getAttribute(name);
  return a ?? null;
}

function children(el: Element | null): Element[] {
  const result: Element[] = [];
  if (!el) return result;
  for (let i = 0; i < el.childNodes.length; i++) {
    const c = el.childNodes[i];
    if (c.nodeType === 1) result.push(c as Element);
  }
  return result;
}

function tryNumber(el: Element | null, name: string): number | null {
  if (!el) return null;
  const s = attrOf(el, name);
  if (!s) return null;
  const v = parseFloat(s);
  if (isNaN(v) || !isFinite(v)) return null;
  return v;
}

/**
 * 自定义自由曲线几何。保留在 DrawingML XML 中的数值路径；guide formulas 和 arcs 暂不支持。
 */
export class ShapeCustomGeometry {
  static build(geometry: Element | null, rect: Rect): PdfPath | null {
    if (!geometry) return null;
    let pathLst: Element | null = null;
    for (const c of children(geometry)) {
      if (c.namespaceURI === A && c.localName === "pathLst") { pathLst = c; break; }
    }
    if (!pathLst) return null;

    const result = new PdfPath();
    for (const pathEl of children(pathLst)) {
      if (pathEl.namespaceURI !== A || pathEl.localName !== "path") continue;
      const width = tryNumber(pathEl, "w");
      const height = tryNumber(pathEl, "h");
      if (!width || !height || width <= 0 || height <= 0) return null;

      for (const command of children(pathEl)) {
        const points: Point[] = [];
        for (const pt of children(command)) {
          if (pt.namespaceURI !== A || pt.localName !== "pt") continue;
          const x = tryNumber(pt, "x");
          const y = tryNumber(pt, "y");
          if (x === null || y === null) return null;
          points.push(new Point(
            rect.left + x / width * rect.width,
            rect.top + y / height * rect.height,
          ));
        }
        const name = command.localName;
        if (name === "moveTo" && points.length === 1) result.moveTo(points[0].x, points[0].y);
        else if (name === "lnTo" && points.length === 1) result.lineTo(points[0].x, points[0].y);
        else if (name === "quadBezTo" && points.length === 2) result.quadTo(points[0].x, points[0].y, points[1].x, points[1].y);
        else if (name === "cubicBezTo" && points.length === 3) result.cubicTo(
          points[0].x, points[0].y,
          points[1].x, points[1].y,
          points[2].x, points[2].y,
        );
        else if (name === "close" && points.length === 0) result.close();
        else return null;
      }
    }
    return result;
  }
}
