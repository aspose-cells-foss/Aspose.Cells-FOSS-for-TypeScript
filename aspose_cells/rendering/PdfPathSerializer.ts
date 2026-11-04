/**
 * Lightweight 2-D path used to serialise glyph outlines into PDF content
 * streams.  Designed as an interface so the Type3 font builder stays
 * backend-agnostic (SkiaSharp, Canvas 2D, @napi-rs/canvas – any path API
 * can be wrapped).
 */
export interface PdfPathPoint {
  x: number;
  y: number;
}

export enum PdfPathVerb {
  Move = 0,
  Line = 1,
  Quad = 2,
  Cubic = 3,
  Close = 4,
  Done = 5,
}

export interface PdfPath {
  IsEmpty: boolean;
  Bounds: { left: number; top: number; right: number; bottom: number };
  /** Iterator over path segments.  `points` is a 4-element array populated by each call. */
  CreateIterator(): {
    Next(points: PdfPathPoint[]): PdfPathVerb;
  };
}

/**
 * Serialises a {@link PdfPath} into a PDF graphics-state sub-path string.
 * Quadratic bezier segments are converted to their cubic equivalents so
 * the output uses only `m`, `l`, `c`, `v`, `y` and `h` operators.
 */
export function SerializePath(path: PdfPath | null): string {
  if (!path) return "";
  const out: string[] = [];
  const points: PdfPathPoint[] = [{ x: 0, y: 0 }, { x: 0, y: 0 }, { x: 0, y: 0 }, { x: 0, y: 0 }];
  let currentX = 0,
    currentY = 0;
  let hasCurrent = false;
  const iter = path.CreateIterator();
  while (true) {
    const verb = iter.Next(points);
    if (verb === PdfPathVerb.Done) break;
    if (verb === PdfPathVerb.Move) {
      out.push(fmt(points[0].x), fmt(points[0].y), "m");
      currentX = points[0].x;
      currentY = points[0].y;
      hasCurrent = true;
    } else if (verb === PdfPathVerb.Line) {
      out.push(fmt(points[1].x), fmt(points[1].y), "l");
      currentX = points[1].x;
      currentY = points[1].y;
      hasCurrent = true;
    } else if (verb === PdfPathVerb.Quad) {
      const c1x = points[0].x + ((points[1].x - points[0].x) * 2) / 3;
      const c1y = points[0].y + ((points[1].y - points[0].y) * 2) / 3;
      const c2x = points[2].x + ((points[1].x - points[2].x) * 2) / 3;
      const c2y = points[2].y + ((points[1].y - points[2].y) * 2) / 3;
      appendCubic(out, c1x, c1y, c2x, c2y, points[2].x, points[2].y);
      currentX = points[2].x;
      currentY = points[2].y;
      hasCurrent = true;
    } else if (verb === PdfPathVerb.Cubic) {
      appendCubic(out, points[1].x, points[1].y, points[2].x, points[2].y, points[3].x, points[3].y);
      currentX = points[3].x;
      currentY = points[3].y;
      hasCurrent = true;
    } else if (verb === PdfPathVerb.Close) {
      out.push("h");
      hasCurrent = false;
    }
  }
  return out.join(" ") + "\n";

  function appendCubic(o: string[], c1x: number, c1y: number, c2x: number, c2y: number, x: number, y: number) {
    if (hasCurrent && sameCoord(c1x, currentX) && sameCoord(c1y, currentY)) {
      o.push(fmt(c2x), fmt(c2y), fmt(x), fmt(y), "v");
    } else if (sameCoord(c2x, x) && sameCoord(c2y, y)) {
      o.push(fmt(c1x), fmt(c1y), fmt(x), fmt(y), "y");
    } else {
      o.push(fmt(c1x), fmt(c1y), fmt(c2x), fmt(c2y), fmt(x), fmt(y), "c");
    }
    currentX = x;
    currentY = y;
    hasCurrent = true;
  }
}

function sameCoord(a: number, b: number): boolean {
  return Math.abs(a - b) <= 0.05;
}

function fmt(value: number): string {
  const rounded = Math.round(value * 10) / 10;
  if (rounded === 0) return "0";
  return rounded.toString().replace(/^-0\.?0*$/, "0");
}
