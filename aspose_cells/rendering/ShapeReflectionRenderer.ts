import type PDFDocument from "pdfkit";
import { Point } from "./ShapeTypes";
import { ParsedShape } from "./ShapeData";
import { RgbaColor } from "./ColorValue";

/** 连接器下方的垂直渐变倒影。其他倒影变换暂不处理。 */
export class ShapeReflectionRenderer {
  static draw(doc: InstanceType<typeof PDFDocument>, points: Point[], shape: ParsedShape): void {
    const reflection = shape.reflection;
    if (!reflection || points.length < 2) return;

    // sy < 0 表示垂直下方倒影
    const sy = ShapeReflectionRenderer.number(reflection, "sy", -100000);
    const dir = ShapeReflectionRenderer.number(reflection, "dir", 5400000);
    if (sy >= 0 || Math.abs(dir - 5400000) > 1) return;

    let top = points[0].y;
    let bottom = top;
    for (let i = 1; i < points.length; i++) {
      top = Math.min(top, points[i].y);
      bottom = Math.max(bottom, points[i].y);
    }
    const height = bottom - top;
    if (height <= 0) return;

    const distance = ShapeReflectionRenderer.number(reflection, "dist", 0) / 12700;
    const endPosition = Math.max(0.001, Math.min(0.999, ShapeReflectionRenderer.number(reflection, "endPos", 100000) / 100000));

    const startA = ShapeReflectionRenderer.alpha(reflection, "stA", 100000);
    const endA = ShapeReflectionRenderer.alpha(reflection, "endA", 0);
    const startColor = new RgbaColor(shape.lineColor.r, shape.lineColor.g, shape.lineColor.b, startA);
    const endColor = new RgbaColor(shape.lineColor.r, shape.lineColor.g, shape.lineColor.b, endA);
    const transparent = new RgbaColor(shape.lineColor.r, shape.lineColor.g, shape.lineColor.b, 0);

    // 构建倒影路径（Y 轴翻转）
    const reflected: Point[] = [];
    for (const p of points) {
      reflected.push(new Point(p.x, 2 * bottom + distance - p.y));
    }

    doc.save();
    // 设置线性渐变
    const grad = doc.linearGradient(0, bottom + distance, 0, bottom + distance + height);
    grad.stop(0, startColor.toPdfKitColor());
    grad.stop(endPosition, endColor.toPdfKitColor());
    grad.stop(1, transparent.toPdfKitColor());

    doc.lineWidth(Math.max(0.5, shape.lineWidthPt));

    doc.moveTo(reflected[0].x, reflected[0].y);
    for (let i = 1; i < reflected.length; i++) {
      doc.lineTo(reflected[i].x, reflected[i].y);
    }
    doc.stroke(grad);
    doc.restore();
  }

  private static alpha(el: Element | null, name: string, fallback: number): number {
    if (!el) return ShapeReflectionRenderer.scaleAlpha(fallback);
    const v = ShapeReflectionRenderer.number(el, name, fallback);
    return ShapeReflectionRenderer.scaleAlpha(v);
  }

  private static scaleAlpha(value: number): number {
    return Math.round(Math.max(0, Math.min(100000, value)) * 255 / 100000);
  }

  private static number(el: Element | null, name: string, fallback: number): number {
    if (!el) return fallback;
    const a = el.getAttribute(name);
    if (!a) return fallback;
    const v = parseFloat(a);
    if (isNaN(v) || !isFinite(v)) return fallback;
    return v;
  }
}
