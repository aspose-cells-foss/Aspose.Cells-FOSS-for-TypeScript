import type { ParsedChart, ChartSeries } from "./ChartData.js";

/**
 * 3D 圆柱体柱体渲染器。
 * 将柱体绘制为近似的圆柱体截面。
 */
export class Column3DCylinderRenderer {
    /**
     * 绘制圆柱体柱体的一组面。
     */
    static Draw(
        doc: any,
        frontPoints: Array<{ x: number; y: number }>,
        backPoints: Array<{ x: number; y: number }>,
        series: ChartSeries
    ): void {
        if (frontPoints.length < 2) return;

        const color = series.Color;
        const h = (v: number) => Math.round(v).toString(16).padStart(2, "0");
        const baseHex = `#${h(color.red)}${h(color.green)}${h(color.blue)}`;

        doc.save();
        doc.fillColor(baseHex);

        // 侧面近似：用多段条带
        const segments = 8;
        for (let i = 0; i < segments; i++) {
            const t1 = i / segments;
            const t2 = (i + 1) / segments;
            const ptTop = Column3DCylinderRenderer.Interpolate(frontPoints[0], frontPoints[1], t1);
            const ptBot = Column3DCylinderRenderer.Interpolate(frontPoints[2], frontPoints[3], t1);
            const ptTop2 = Column3DCylinderRenderer.Interpolate(frontPoints[0], frontPoints[1], t2);
            const ptBot2 = Column3DCylinderRenderer.Interpolate(frontPoints[2], frontPoints[3], t2);

            const shade = 0.75 + 0.25 * Math.cos((t1 + t2) * Math.PI);
            const r = Math.round(color.red * shade);
            const g = Math.round(color.green * shade);
            const b = Math.round(color.blue * shade);

            doc.fillColor(`#${h(r)}${h(g)}${h(b)}`);
            doc.moveTo(ptTop.x, ptTop.y);
            doc.lineTo(ptTop2.x, ptTop2.y);
            doc.lineTo(ptBot2.x, ptBot2.y);
            doc.lineTo(ptBot.x, ptBot.y);
            doc.closePath();
            doc.fill();
        }

        doc.restore();
    }

    private static Interpolate(
        a: { x: number; y: number },
        b: { x: number; y: number },
        t: number
    ): { x: number; y: number } {
        return { x: a.x + (b.x - a.x) * t, y: a.y + (b.y - a.y) * t };
    }
}
