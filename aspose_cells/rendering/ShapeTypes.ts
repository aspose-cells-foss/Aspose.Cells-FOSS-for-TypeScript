import type PDFDocument from "pdfkit";

/** 2D point, replacing SkiaSharp's SKPoint. */
export class Point {
  constructor(
    public x: number,
    public y: number,
  ) {}

  static readonly zero = new Point(0, 0);

  static copy(p: Point): Point {
    return new Point(p.x, p.y);
  }
}

/** Rectangle, replacing SkiaSharp's SKRect. All coordinates use a top-left origin. */
export class Rect {
  constructor(
    public left: number,
    public top: number,
    public right: number,
    public bottom: number,
  ) {}

  get width(): number {
    return this.right - this.left;
  }

  get height(): number {
    return this.bottom - this.top;
  }

  get midX(): number {
    return (this.left + this.right) / 2;
  }

  get midY(): number {
    return (this.top + this.bottom) / 2;
  }

  static fromSize(left: number, top: number, width: number, height: number): Rect {
    return new Rect(left, top, left + width, top + height);
  }

  static copy(r: Rect): Rect {
    return new Rect(r.left, r.top, r.right, r.bottom);
  }
}

/**
 * Lightweight path builder, replacing SkiaSharp's SKPath.
 * Accumulates commands and executes them on a pdfkit PDFDocument via draw().
 */
export class PdfPath {
  private commands: Array<{ op: string; args: number[] }> = [];

  // Building-time current point, mirroring SKPath. Needed when an arc has to
  // connect to the preceding segment.
  private currentX = 0;
  private currentY = 0;
  private hasCurrent = false;
  private subpathX = 0;
  private subpathY = 0;

  moveTo(x: number, y: number): PdfPath {
    this.commands.push({ op: "M", args: [x, y] });
    this.currentX = x;
    this.currentY = y;
    this.subpathX = x;
    this.subpathY = y;
    this.hasCurrent = true;
    return this;
  }

  lineTo(x: number, y: number): PdfPath {
    this.commands.push({ op: "L", args: [x, y] });
    this.currentX = x;
    this.currentY = y;
    this.hasCurrent = true;
    return this;
  }

  cubicTo(x1: number, y1: number, x2: number, y2: number, x3: number, y3: number): PdfPath {
    this.commands.push({ op: "C", args: [x1, y1, x2, y2, x3, y3] });
    this.currentX = x3;
    this.currentY = y3;
    this.hasCurrent = true;
    return this;
  }

  /**
   * Quadratic bezier: (x1, y1) is the control point and (x2, y2) the end
   * point; the start point is the current point (SKPath.QuadTo semantics).
   */
  quadTo(x1: number, y1: number, x2: number, y2: number): PdfPath {
    if (!this.hasCurrent) this.moveTo(0, 0);
    this.commands.push({ op: "Q", args: [x1, y1, x2, y2] });
    this.currentX = x2;
    this.currentY = y2;
    this.hasCurrent = true;
    return this;
  }

  close(): PdfPath {
    this.commands.push({ op: "Z", args: [] });
    if (this.hasCurrent) {
      this.currentX = this.subpathX;
      this.currentY = this.subpathY;
    }
    return this;
  }

  addRect(r: Rect): PdfPath {
    this.moveTo(r.left, r.top);
    this.lineTo(r.right, r.top);
    this.lineTo(r.right, r.bottom);
    this.lineTo(r.left, r.bottom);
    this.close();
    return this;
  }

  addRoundRect(r: Rect, rx: number, ry: number): PdfPath {
    const w = r.width;
    const h = r.height;
    if (rx <= 0 || ry <= 0 || w <= 0 || h <= 0) {
      this.addRect(r);
      return this;
    }
    // SKPath.AddRoundRect scales the corner radii down uniformly when they
    // exceed half of the rectangle, which is what SkRRect does as well.
    const scale = Math.min(1, w / (2 * rx), h / (2 * ry));
    const xr = rx * scale;
    const yr = ry * scale;
    const k = 0.5522847498; // circle-to-bezier kappa factor, 4/3 * tan(pi/8)
    const kx = xr * k;
    const ky = yr * k;
    this.moveTo(r.left + xr, r.top);
    this.lineTo(r.right - xr, r.top);
    this.cubicTo(r.right - xr + kx, r.top, r.right, r.top + yr - ky, r.right, r.top + yr);
    this.lineTo(r.right, r.bottom - yr);
    this.cubicTo(r.right, r.bottom - yr + ky, r.right - xr + kx, r.bottom, r.right - xr, r.bottom);
    this.lineTo(r.left + xr, r.bottom);
    this.cubicTo(r.left + xr - kx, r.bottom, r.left, r.bottom - yr + ky, r.left, r.bottom - yr);
    this.lineTo(r.left, r.top + yr);
    this.cubicTo(r.left, r.top + yr - ky, r.left + xr - kx, r.top, r.left + xr, r.top);
    this.close();
    return this;
  }

  addOval(r: Rect): PdfPath {
    // Ellipse approximated by four cubic segments.
    const cx = r.midX;
    const cy = r.midY;
    const rx = r.width / 2;
    const ry = r.height / 2;
    const k = 0.5522847498;
    this.moveTo(cx + rx, cy);
    this.cubicTo(cx + rx, cy + ry * k, cx + rx * k, cy + ry, cx, cy + ry);
    this.cubicTo(cx - rx * k, cy + ry, cx - rx, cy + ry * k, cx - rx, cy);
    this.cubicTo(cx - rx, cy - ry * k, cx - rx * k, cy - ry, cx, cy - ry);
    this.cubicTo(cx + rx * k, cy - ry, cx + rx, cy - ry * k, cx + rx, cy);
    this.close();
    return this;
  }

  /**
   * Appends an elliptical arc, mirroring SKPath.ArcTo(rect, startAngle,
   * sweepAngle, forceMoveTo). Angles are in degrees with 0 pointing right and
   * positive angles sweeping clockwise in the y-down coordinate space. The
   * arc is decomposed into cubic bezier segments of at most 90 degrees each
   * using kappa = 4/3 * tan(theta/4).
   */
  arcTo(r: Rect, startAngle: number, sweepAngle: number, forceMoveTo: boolean): PdfPath {
    const cx = r.midX;
    const cy = r.midY;
    const rx = r.width / 2;
    const ry = r.height / 2;
    if (rx <= 0 || ry <= 0 || sweepAngle === 0) return this;

    // Skia limits the sweep to a full turn.
    const sweep = Math.max(-360, Math.min(360, sweepAngle));
    const startRad = startAngle * Math.PI / 180;
    const startX = cx + rx * Math.cos(startRad);
    const startY = cy + ry * Math.sin(startRad);

    if (forceMoveTo || !this.hasCurrent) {
      this.moveTo(startX, startY);
    } else if (Math.abs(this.currentX - startX) > 1e-6 || Math.abs(this.currentY - startY) > 1e-6) {
      this.lineTo(startX, startY);
    }

    const segments = Math.max(1, Math.ceil(Math.abs(sweep) / 90));
    const stepRad = sweep * Math.PI / 180 / segments;
    const k = 4 / 3 * Math.tan(stepRad / 4);
    let a0 = startRad;
    for (let i = 0; i < segments; i++) {
      const a1 = a0 + stepRad;
      const p0x = cx + rx * Math.cos(a0);
      const p0y = cy + ry * Math.sin(a0);
      const p1x = cx + rx * Math.cos(a1);
      const p1y = cy + ry * Math.sin(a1);
      const c1x = p0x - rx * Math.sin(a0) * k;
      const c1y = p0y + ry * Math.cos(a0) * k;
      const c2x = p1x + rx * Math.sin(a1) * k;
      const c2y = p1y - ry * Math.cos(a1) * k;
      this.cubicTo(c1x, c1y, c2x, c2y, p1x, p1y);
      a0 = a1;
    }
    return this;
  }

  /** Draws the accumulated commands onto a pdfkit PDFDocument. */
  draw(doc: InstanceType<typeof PDFDocument>): void {
    // Track the current point so a quadratic command can be converted to its
    // cubic equivalent (pdfkit has no native quadratic operator).
    let curX = 0;
    let curY = 0;
    let startX = 0;
    let startY = 0;
    let hasCurrent = false;
    for (const cmd of this.commands) {
      switch (cmd.op) {
        case "M":
          curX = cmd.args[0];
          curY = cmd.args[1];
          startX = curX;
          startY = curY;
          hasCurrent = true;
          doc.moveTo(curX, curY);
          break;
        case "L":
          curX = cmd.args[0];
          curY = cmd.args[1];
          hasCurrent = true;
          doc.lineTo(curX, curY);
          break;
        case "C":
          curX = cmd.args[4];
          curY = cmd.args[5];
          hasCurrent = true;
          doc.bezierCurveTo(cmd.args[0], cmd.args[1], cmd.args[2], cmd.args[3], cmd.args[4], cmd.args[5]);
          break;
        case "Q": {
          const ctrlX = cmd.args[0], ctrlY = cmd.args[1];
          const endX = cmd.args[2], endY = cmd.args[3];
          if (!hasCurrent) {
            curX = 0;
            curY = 0;
            startX = 0;
            startY = 0;
            hasCurrent = true;
            doc.moveTo(curX, curY);
          }
          // De Casteljau degree elevation: C1 = P0 + 2/3 (P1 - P0),
          // C2 = P2 + 2/3 (P1 - P2).
          const c1x = curX + (ctrlX - curX) * 2 / 3;
          const c1y = curY + (ctrlY - curY) * 2 / 3;
          const c2x = endX + (ctrlX - endX) * 2 / 3;
          const c2y = endY + (ctrlY - endY) * 2 / 3;
          doc.bezierCurveTo(c1x, c1y, c2x, c2y, endX, endY);
          curX = endX;
          curY = endY;
          break;
        }
        case "Z":
          doc.closePath();
          if (hasCurrent) {
            curX = startX;
            curY = startY;
          }
          break;
      }
    }
  }

  isEmpty(): boolean {
    return this.commands.length === 0;
  }
}
