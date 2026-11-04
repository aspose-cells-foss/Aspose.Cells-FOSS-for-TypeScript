import type PDFDocument from "pdfkit";

/** 2D 点，替代 SkiaSharp 的 SKPoint。 */
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

/** 矩形，替代 SkiaSharp 的 SKRect。所有坐标以左上角为原点。 */
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
 * 轻量路径构建器，替代 SkiaSharp 的 SKPath。
 * 累积命令，最后通过 draw() 方法在 pdfkit PDFDocument 上执行。
 */
export class PdfPath {
  private commands: Array<{ op: string; args: number[] }> = [];

  moveTo(x: number, y: number): PdfPath {
    this.commands.push({ op: "M", args: [x, y] });
    return this;
  }

  lineTo(x: number, y: number): PdfPath {
    this.commands.push({ op: "L", args: [x, y] });
    return this;
  }

  cubicTo(x1: number, y1: number, x2: number, y2: number, x3: number, y3: number): PdfPath {
    this.commands.push({ op: "C", args: [x1, y1, x2, y2, x3, y3] });
    return this;
  }

  quadTo(x1: number, y1: number, x2: number, y2: number): PdfPath {
    this.commands.push({ op: "Q", args: [x1, y1, x2, y2] });
    return this;
  }

  close(): PdfPath {
    this.commands.push({ op: "Z", args: [] });
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
    // 简化：用矩形近似
    this.addRect(r);
    return this;
  }

  addOval(r: Rect): PdfPath {
    // 用贝塞尔曲线近似椭圆
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
    return this;
  }

  /** 将累积的命令绘制到 pdfkit PDFDocument 上。 */
  draw(doc: InstanceType<typeof PDFDocument>): void {
    for (const cmd of this.commands) {
      switch (cmd.op) {
        case "M":
          doc.moveTo(cmd.args[0], cmd.args[1]);
          break;
        case "L":
          doc.lineTo(cmd.args[0], cmd.args[1]);
          break;
        case "C":
          doc.bezierCurveTo(cmd.args[0], cmd.args[1], cmd.args[2], cmd.args[3], cmd.args[4], cmd.args[5]);
          break;
        case "Q":
          // pdfkit 没有直接的二次贝塞尔，转为三次近似
          const x0 = cmd.args[0], y0 = cmd.args[1];
          const x1 = cmd.args[2], y1 = cmd.args[3];
          const cx1 = x0 + (x1 - x0) * 2 / 3;
          const cy1 = y0 + (y1 - y0) * 2 / 3;
          const cx2 = (x0 + 2 * x1) / 3;
          const cy2 = (y0 + 2 * y1) / 3;
          doc.bezierCurveTo(cx1, cy1, cx2, cy2, x1, y1);
          break;
        case "Z":
          doc.closePath();
          break;
      }
    }
  }

  isEmpty(): boolean {
    return this.commands.length === 0;
  }
}
