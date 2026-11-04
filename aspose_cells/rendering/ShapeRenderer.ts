import type PDFDocument from "pdfkit";
import { Rect, Point, PdfPath } from "./ShapeTypes";
import { ParsedShape } from "./ShapeData";
import { RenderContext } from "./RenderContext";
import { PageLayout } from "./PageLayout";
import { SheetLayout } from "./SheetLayout";
import { ShapeSpecialGeometry } from "./ShapeSpecialGeometry";
import { ShapePresetDetails } from "./ShapePresetDetails";
import { ShapePatternRenderer } from "./ShapePatternRenderer";
import { ShapeReflectionRenderer } from "./ShapeReflectionRenderer";
import { ShapePolylineGeometry } from "./ShapePolylineGeometry";
import { ShapeCurveGeometry } from "./ShapeCurveGeometry";
import { ShapeVerticalTextRenderer } from "./ShapeVerticalTextRenderer";
import { ShapeWrappedTextRenderer } from "./ShapeWrappedTextRenderer";
import { RgbaColor } from "./ColorValue";

/**
 * 形状渲染器核心：将 ParsedShape 绘制到矩形内。处理预设几何（rect/ellipse/方块箭头/加减号…）、
 * straight/curved/bent 连接器（带箭头）、居中文字。未知几何降级为圆角/普通矩形。
 */
export class ShapeRenderer {
  private doc: InstanceType<typeof PDFDocument>;
  private context: RenderContext;

  private static readonly MinArrowScalePt = 1.1;

  constructor(
    doc: InstanceType<typeof PDFDocument>,
    context: RenderContext,
    layout?: SheetLayout,
  ) {
    this.doc = doc;
    this.context = context;
  }

  draw(rect: Rect, shape: ParsedShape): void;
  draw(page: PageLayout | null, rect: Rect, shape: ParsedShape): void;
  draw(page: PageLayout | null, rect: Rect, shape: ParsedShape, _shapeModel?: any, _shapes?: any[]): void;
  draw(pageOrRect: PageLayout | null | Rect, rectOrShape: Rect | ParsedShape, shapeOrUndefined?: ParsedShape, _shapeModel?: any, _shapes?: any[]): void {
    let page: PageLayout | null;
    let rect: Rect;
    let shape: ParsedShape;
    if (pageOrRect instanceof Rect) {
      page = null;
      rect = pageOrRect;
      shape = rectOrShape as ParsedShape;
    } else {
      page = pageOrRect;
      rect = rectOrShape as Rect;
      shape = shapeOrUndefined as ParsedShape;
    }

    if (ShapeRenderer.isConnector(shape.geometry)) {
      this.drawConnector(shape, rect);
      return;
    }

    if (Math.abs(shape.rotationDeg) > 0.01) {
      this.doc.save();
      this.doc.translate(rect.midX, rect.midY);
      this.doc.rotate(shape.rotationDeg);
      this.doc.translate(-rect.midX, -rect.midY);
      this.drawFilledShape(shape, rect);
      this.doc.restore();
    } else {
      this.drawFilledShape(shape, rect);
    }
  }

  /** 渲染页面上所有形状。目前为空实现，待数据模型完善。 */
  renderAll(_page: PageLayout): void {
    // TODO: 从 SheetLayout 提取形状数据并逐个 draw
  }

  private static isConnector(geometry: string): boolean {
    return geometry === "line"
      || geometry.startsWith("straightConnector")
      || geometry.startsWith("bentConnector")
      || geometry.startsWith("curvedConnector");
  }

  // ———————— 连接器 ————————

  private drawConnector(shape: ParsedShape, rect: Rect): void {
    if (!shape.hasLine) return;

    if (shape.geometry.startsWith("curvedConnector")) {
      this.drawCurvedConnector(shape, rect);
      return;
    }

    const width = Math.max(0.5, shape.lineWidthPt);
    const initialStart = ShapeRenderer.transformNormalizedPoint(rect, shape, 0, 0);
    const initialEnd = ShapeRenderer.transformNormalizedPoint(rect, shape, 1, 1);
    const points = ShapePolylineGeometry.build(initialStart, initialEnd, shape);
    if (!points || points.length < 2) return;

    const start = points[0];
    const end = points[points.length - 1];
    const dx = end.x - start.x;
    const dy = end.y - start.y;
    const len = Math.sqrt(dx * dx + dy * dy);
    if (len < 0.01) return;

    const arrowScale = Math.max(width, ShapeRenderer.MinArrowScalePt);

    // 倒影
    ShapeReflectionRenderer.draw(this.doc, points, shape);

    const drawPoints = ShapeRenderer.copyPoints(points);
    if (ShapeRenderer.hasArrow(shape.headEnd)) {
      ShapeRenderer.shortenPolylineStart(drawPoints, ShapeRenderer.filledArrow(shape.headEnd)
        ? arrowScale * ShapeRenderer.arrowLineHideMultiplier(shape.headEnd) : 0);
    }
    if (ShapeRenderer.hasArrow(shape.tailEnd)) {
      ShapeRenderer.shortenPolylineEnd(drawPoints, ShapeRenderer.filledArrow(shape.tailEnd)
        ? arrowScale * ShapeRenderer.arrowLineHideMultiplier(shape.tailEnd) : 0);
    }

    this.doc.save();
    this.doc.strokeColor(shape.lineColor.toPdfKitColor());
    this.doc.lineWidth(width);
    this.doc.lineCap("butt");
    this.doc.lineJoin("miter");
    this.doc.moveTo(drawPoints[0].x, drawPoints[0].y);
    for (let i = 1; i < drawPoints.length; i++) {
      this.doc.lineTo(drawPoints[i].x, drawPoints[i].y);
    }
    this.doc.stroke();
    this.doc.restore();

    const headDirection = ShapeRenderer.segmentDirection(points[1], points[0]);
    const tailDirection = ShapeRenderer.segmentDirection(points[points.length - 2], points[points.length - 1]);
    ShapeRenderer.drawArrowhead(this.doc, shape.headEnd, start, headDirection.x, headDirection.y, width, shape.lineColor);
    ShapeRenderer.drawArrowhead(this.doc, shape.tailEnd, end, tailDirection.x, tailDirection.y, width, shape.lineColor);
  }

  private drawCurvedConnector(shape: ParsedShape, rect: Rect): void {
    const width = Math.max(0.5, shape.lineWidthPt);
    const arrowScale = Math.max(width, ShapeRenderer.MinArrowScalePt);

    const start = ShapeRenderer.transformNormalizedPoint(rect, shape, 0, 0);
    const end = ShapeRenderer.transformNormalizedPoint(rect, shape, 1, 1);
    const curve = ShapeCurveGeometry.build(start, end, shape);
    const control1 = curve[1];
    const control2 = curve[curve.length - 2];

    const endDirection = ShapeRenderer.segmentDirection(control2, end);
    const headDirection = ShapeRenderer.segmentDirection(control1, start);

    let visibleStart = start;
    let visibleEnd = end;
    if (ShapeRenderer.hasArrow(shape.headEnd)) {
      visibleStart = new Point(
        start.x + headDirection.x * arrowScale * ShapeRenderer.arrowLineHideMultiplier(shape.headEnd),
        start.y + headDirection.y * arrowScale * ShapeRenderer.arrowLineHideMultiplier(shape.headEnd),
      );
    }
    if (ShapeRenderer.hasArrow(shape.tailEnd)) {
      visibleEnd = new Point(
        end.x - endDirection.x * arrowScale * ShapeRenderer.arrowLineHideMultiplier(shape.tailEnd),
        end.y - endDirection.y * arrowScale * ShapeRenderer.arrowLineHideMultiplier(shape.tailEnd),
      );
    }

    this.doc.save();
    this.doc.strokeColor(shape.lineColor.toPdfKitColor());
    this.doc.lineWidth(width);
    this.doc.lineCap("butt");
    this.doc.lineJoin("round");
    this.doc.moveTo(visibleStart.x, visibleStart.y);
    const c2 = [...curve];
    c2[c2.length - 1] = visibleEnd;
    for (let i = 1; i < c2.length; i += 3) {
      this.doc.bezierCurveTo(c2[i].x, c2[i].y, c2[i + 1].x, c2[i + 1].y, c2[i + 2].x, c2[i + 2].y);
    }
    this.doc.stroke();
    this.doc.restore();

    ShapeRenderer.drawArrowhead(this.doc, shape.headEnd, start, headDirection.x, headDirection.y, width, shape.lineColor);
    ShapeRenderer.drawArrowhead(this.doc, shape.tailEnd, end, endDirection.x, endDirection.y, width, shape.lineColor);
  }

  // ———————— 箭头 ————————

  private static hasArrow(type: string | null): boolean {
    return !!type && type !== "none";
  }

  private static filledArrow(type: string | null): boolean {
    return type === "triangle" || type === "stealth" || type === "diamond" || type === "oval";
  }

  private static drawArrowhead(doc: InstanceType<typeof PDFDocument>, type: string | null, tip: Point, dirX: number, dirY: number, lineWidth: number, color: RgbaColor): void {
    if (!ShapeRenderer.hasArrow(type)) return;
    const t = type!;
    const px = -dirY;
    const py = dirX;
    const arrowScale = t === "triangle" ? Math.max(lineWidth, 1) : Math.max(lineWidth, ShapeRenderer.MinArrowScalePt);
    const arrowLen = arrowScale * ShapeRenderer.arrowLengthMultiplier(t);
    const halfW = arrowScale * ShapeRenderer.arrowHalfWidthMultiplier(t);
    const baseX = tip.x - dirX * arrowLen;
    const baseY = tip.y - dirY * arrowLen;

    if (!ShapeRenderer.filledArrow(t)) {
      // 开放 "arrow"：V 形笔画
      doc.save();
      doc.strokeColor(color.toPdfKitColor());
      doc.lineWidth(lineWidth);
      doc.lineCap("round");
      doc.moveTo(tip.x, tip.y);
      doc.lineTo(baseX + px * halfW, baseY + py * halfW);
      doc.moveTo(tip.x, tip.y);
      doc.lineTo(baseX - px * halfW, baseY - py * halfW);
      doc.stroke();
      doc.restore();
      return;
    }

    doc.save();
    doc.fillColor(color.toPdfKitColor());
    doc.moveTo(tip.x, tip.y);
    doc.lineTo(baseX + px * halfW, baseY + py * halfW);
    doc.lineTo(baseX - px * halfW, baseY - py * halfW);
    doc.closePath();
    doc.fill();
    doc.restore();
  }

  private static arrowLengthMultiplier(type: string | null): number {
    if (type === "triangle") return 6;
    if (type === "stealth") return 0.9;
    return 1.05;
  }

  private static arrowHalfWidthMultiplier(type: string | null): number {
    if (type === "triangle") return 3;
    if (type === "stealth") return 1.05;
    return 1.15;
  }

  private static arrowLineHideMultiplier(type: string | null): number {
    if (type === "triangle") return 0;
    return ShapeRenderer.arrowLengthMultiplier(type);
  }

  // ———————— 填充形状 ————————

  private drawFilledShape(shape: ParsedShape, rect: Rect): void {
    const path = this.buildGeometry(shape, rect);
    if (shape.hasFill) {
      if (!ShapePresetDetails.drawFill(this.doc, rect, shape)) {
        this.doc.save();
        this.doc.fillColor(shape.fillColor.toPdfKitColor());
        path.draw(this.doc);
        this.doc.fill();
        this.doc.restore();
        ShapePatternRenderer.draw(this.doc, rect, path, shape);
      }
    }

    if (shape.hasLine) {
      this.doc.save();
      this.doc.strokeColor(shape.lineColor.toPdfKitColor());
      this.doc.lineWidth(Math.max(0.5, shape.lineWidthPt));
      this.doc.lineJoin("miter");
      path.draw(this.doc);
      this.doc.stroke();
      this.doc.restore();
    }

    ShapePresetDetails.drawLines(this.doc, rect, shape);
    this.drawText(shape, rect);
  }

  // ———————— 几何构建 ————————

  private buildGeometry(shape: ParsedShape, r: Rect): PdfPath {
    const special = ShapeSpecialGeometry.build(shape, r);
    if (special) return special;

    const geometry = shape.geometry;
    const path = new PdfPath();
    const l = r.left, t = r.top, right = r.right, b = r.bottom;
    const w = r.width, h = r.height;
    const cx = r.midX, cy = r.midY;

    switch (geometry) {
      case "corner":
        ShapeRenderer.cornerShape(path, r, shape.adjusts);
        break;
      case "ellipse":
        path.addOval(r);
        break;
      case "roundRect":
        path.addRoundRect(r, Math.min(w, h) * 0.16, Math.min(w, h) * 0.16);
        break;
      case "triangle": {
        const apexAdj = shape.adjusts && shape.adjusts.length > 0 ? shape.adjusts[0] : 0.5;
        path.moveTo(l + w * apexAdj, t);
        path.lineTo(right, b);
        path.lineTo(l, b);
        path.close();
        break;
      }
      case "rtTriangle":
        path.moveTo(l, t); path.lineTo(l, b); path.lineTo(right, b); path.close();
        break;
      case "diamond":
        path.moveTo(cx, t); path.lineTo(right, cy); path.lineTo(cx, b); path.lineTo(l, cy); path.close();
        break;
      case "rightArrow":
        ShapeRenderer.blockArrowRight(path, r); break;
      case "leftArrow":
        ShapeRenderer.blockArrowLeft(path, r); break;
      case "upArrow":
        ShapeRenderer.blockArrowUp(path, r); break;
      case "downArrow":
        ShapeRenderer.blockArrowDown(path, r); break;
      case "mathPlus":
      case "plus":
        ShapeRenderer.plusShape(path, r); break;
      default:
        path.addRect(r);
        break;
    }
    return path;
  }

  private static cornerShape(path: PdfPath, r: Rect, adj: number[] | null): void {
    const a1 = adj && adj.length > 0 ? adj[0] : 0.5;
    const a2 = adj && adj.length > 1 ? adj[1] : 0.5;
    const ss = Math.min(r.width, r.height);
    const dx1 = ss * a2;
    const dy1 = ss * a1;
    path.moveTo(r.left, r.top);
    path.lineTo(r.left + dx1, r.top);
    path.lineTo(r.left + dx1, r.bottom - dy1);
    path.lineTo(r.right, r.bottom - dy1);
    path.lineTo(r.right, r.bottom);
    path.lineTo(r.left, r.bottom);
    path.close();
  }

  private static plusShape(path: PdfPath, r: Rect): void {
    const d = Math.min(r.width, r.height) * 0.2352;
    const half = d / 2;
    const cx = r.midX, cy = r.midY;
    const hx = r.width * 0.36745;
    const hy = r.height * 0.36745;
    const l = cx - hx, right = cx + hx, t = cy - hy, b = cy + hy;

    path.moveTo(cx - half, t);
    path.lineTo(cx + half, t);
    path.lineTo(cx + half, cy - half);
    path.lineTo(right, cy - half);
    path.lineTo(right, cy + half);
    path.lineTo(cx + half, cy + half);
    path.lineTo(cx + half, b);
    path.lineTo(cx - half, b);
    path.lineTo(cx - half, cy + half);
    path.lineTo(l, cy + half);
    path.lineTo(l, cy - half);
    path.lineTo(cx - half, cy - half);
    path.close();
  }

  private static blockArrowRight(path: PdfPath, r: Rect): void {
    const headW = Math.min(r.width, r.height) * 0.5;
    const shaftH = r.height * 0.5;
    const hx = r.right - headW;
    const sy0 = r.midY - shaftH / 2;
    const sy1 = r.midY + shaftH / 2;
    path.moveTo(r.left, sy0);
    path.lineTo(hx, sy0); path.lineTo(hx, r.top); path.lineTo(r.right, r.midY);
    path.lineTo(hx, r.bottom); path.lineTo(hx, sy1); path.lineTo(r.left, sy1);
    path.close();
  }

  private static blockArrowLeft(path: PdfPath, r: Rect): void {
    const headW = r.width * 0.4;
    const shaftH = r.height * 0.5;
    const hx = r.left + headW;
    const sy0 = r.midY - shaftH / 2;
    const sy1 = r.midY + shaftH / 2;
    path.moveTo(r.right, sy0);
    path.lineTo(hx, sy0); path.lineTo(hx, r.top); path.lineTo(r.left, r.midY);
    path.lineTo(hx, r.bottom); path.lineTo(hx, sy1); path.lineTo(r.right, sy1);
    path.close();
  }

  private static blockArrowUp(path: PdfPath, r: Rect): void {
    const headH = r.height * 0.4;
    const shaftW = r.width * 0.5;
    const sx0 = r.midX - shaftW / 2;
    const sx1 = r.midX + shaftW / 2;
    const hy = r.top + headH;
    path.moveTo(sx0, r.bottom);
    path.lineTo(sx0, hy); path.lineTo(r.left, hy); path.lineTo(r.midX, r.top);
    path.lineTo(r.right, hy); path.lineTo(sx1, hy); path.lineTo(sx1, r.bottom);
    path.close();
  }

  private static blockArrowDown(path: PdfPath, r: Rect): void {
    const headH = r.height * 0.4;
    const shaftW = r.width * 0.5;
    const sx0 = r.midX - shaftW / 2;
    const sx1 = r.midX + shaftW / 2;
    const hy = r.bottom - headH;
    path.moveTo(sx0, r.top);
    path.lineTo(sx0, hy); path.lineTo(r.left, hy); path.lineTo(r.midX, r.bottom);
    path.lineTo(r.right, hy); path.lineTo(sx1, hy); path.lineTo(sx1, r.top);
    path.close();
  }

  // ———————— 文字 ————————

  private drawText(shape: ParsedShape, rect: Rect): void {
    if (!shape.text) return;

    if (ShapeVerticalTextRenderer.isVerticalText(shape)) {
      ShapeVerticalTextRenderer.draw(this.doc, rect, shape);
      return;
    }

    const textWidth = this.doc.widthOfString(shape.text);
    const width = Math.max(1, rect.width - shape.textInsetLeftPt * 2);

    if (shape.textWrap && textWidth > width) {
      ShapeWrappedTextRenderer.draw(this.doc, rect, shape);
      return;
    }

    this.doc.save();
    this.doc.fontSize(shape.textSizePt);
    this.doc.fillColor(shape.textColor.toPdfKitColor());
    if (shape.textBold) this.doc.font("Helvetica-Bold"); else this.doc.font("Helvetica");

    // 水平对齐
    let x: number;
    if (shape.textHAlign === "l") x = rect.left + shape.textInsetLeftPt;
    else if (shape.textHAlign === "r") x = rect.right - shape.textInsetLeftPt - textWidth;
    else x = rect.midX - textWidth / 2;

    // 垂直锚点（近似 baseline）
    let baseline: number;
    const asc = shape.textSizePt * 0.85; // 近似 ascent
    const desc = shape.textSizePt * 0.15; // 近似 descent
    if (shape.textVAnchor === "t") baseline = rect.top + shape.textInsetTopPt + asc;
    else if (shape.textVAnchor === "b") baseline = rect.bottom - shape.textInsetTopPt - desc;
    else baseline = rect.midY + (asc - desc) / 2;

    this.doc.text(shape.text, x, baseline - asc);
    this.doc.restore();
  }

  // ———————— 点变换 & 辅助 ————————

  private static transformNormalizedPoint(rect: Rect, shape: ParsedShape, x: number, y: number): Point {
    if (shape.flipH) x = 1 - x;
    if (shape.flipV) y = 1 - y;
    const transformed = ShapeRenderer.rotateNormalizedPoint(x, y, shape.rotationDeg);
    return new Point(
      rect.left + transformed.x * rect.width,
      rect.top + transformed.y * rect.height,
    );
  }

  private static rotateNormalizedPoint(x: number, y: number, rotationDeg: number): Point {
    if (Math.abs(rotationDeg) <= 0.01) return new Point(x, y);
    const radians = rotationDeg * Math.PI / 180;
    const dx = x - 0.5;
    const dy = y - 0.5;
    const rotatedX = dx * Math.cos(radians) - dy * Math.sin(radians);
    const rotatedY = dx * Math.sin(radians) + dy * Math.cos(radians);
    return new Point(0.5 + rotatedX, 0.5 + rotatedY);
  }

  private static copyPoints(points: Point[]): Point[] {
    return points.map(p => new Point(p.x, p.y));
  }

  private static shortenPolylineStart(points: Point[], shortenBy: number): void {
    if (shortenBy <= 0 || !points || points.length < 2) return;
    const direction = ShapeRenderer.segmentDirection(points[1], points[0]);
    points[0] = new Point(points[0].x + direction.x * shortenBy, points[0].y + direction.y * shortenBy);
  }

  private static shortenPolylineEnd(points: Point[], shortenBy: number): void {
    if (shortenBy <= 0 || !points || points.length < 2) return;
    const last = points.length - 1;
    const direction = ShapeRenderer.segmentDirection(points[last - 1], points[last]);
    points[last] = new Point(points[last].x - direction.x * shortenBy, points[last].y - direction.y * shortenBy);
  }

  private static segmentDirection(from: Point, to: Point): Point {
    const dx = to.x - from.x;
    const dy = to.y - from.y;
    const length = Math.sqrt(dx * dx + dy * dy);
    if (length <= 0.01) return new Point(1, 0);
    return new Point(dx / length, dy / length);
  }
}
