
/**
 * 平滑路径构建器。
 * 使用 Catmull-Rom 样条曲线为散点图和折线图生成平滑路径。
 */
export class ScatterSmoothPathBuilder {
    /**
     * 向 PDF 文档路径添加平滑曲线。
     * 使用 Catmull-Rom 到 Bezier 转换。
     */
    static Append(
        doc: any,
        points: Array<{ x: number; y: number }>
    ): void {
        if (points.length < 2) return;

        if (points.length === 2) {
            doc.moveTo(points[0].x, points[0].y);
            doc.lineTo(points[1].x, points[1].y);
            return;
        }

        doc.moveTo(points[0].x, points[0].y);

        for (let i = 0; i < points.length - 1; i++) {
            const p0 = points[i - 1] ?? points[i];
            const p1 = points[i];
            const p2 = points[i + 1];
            const p3 = points[i + 2] ?? p2;

            // Catmull-Rom 控制点 -> Bezier 控制点
            const cp1x = p1.x + (p2.x - p0.x) / 6;
            const cp1y = p1.y + (p2.y - p0.y) / 6;
            const cp2x = p2.x - (p3.x - p1.x) / 6;
            const cp2y = p2.y - (p3.y - p1.y) / 6;

            doc.bezierCurveTo(cp1x, cp1y, cp2x, cp2y, p2.x, p2.y);
        }
    }
}
