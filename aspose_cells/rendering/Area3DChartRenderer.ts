import type { ParsedChart, Rect } from "./ChartData.js";
import { ChartStackingMath } from "./ChartStackingMath.js";
import { Area3DDateAxisMath } from "./Area3DDateAxisMath.js";
import { Area3DWallFillRenderer } from "./Area3DWallFillRenderer.js";

/**
 * Area 3D 图表渲染器。
 * 将面积图序列绘制为 3D 投影的填充曲面。
 */
export class Area3DChartRenderer {
    /**
     * 绘制 Area 3D 图表。
     */
    Draw(
        doc: any,
        area: Rect,
        plot: Rect,
        chart: ParsedChart,
        min: number,
        max: number,
        major: number
    ): void {
        if (chart == null || chart.Series.length === 0 || max <= min) return;

        const count = Math.max(chart.Categories.length, chart.CategoryValues.length);
        const w = plot.right - plot.left;
        const h = plot.bottom - plot.top;
        const depthScale = Math.max(0.25, Math.min(2, (chart.DepthPercent ?? 100) / 100));
        const depthOffset = h * 0.04 * depthScale;

        // 绘制墙壁背景
        Column3DWallFills(doc, plot, depthOffset);

        // 绘制每个序列的面积带
        for (let s = chart.Series.length - 1; s >= 0; s--) {
            const series = chart.Series[s];
            const displayValues = ChartStackingMath.BuildDisplayValues(chart, s);
            const d = (s + 0.5) / chart.Series.length;
            const depthX = depthOffset * d;

            // 构建点路径
            const frontPoints: Array<{ x: number; y: number }> = [];
            const backPoints: Array<{ x: number; y: number }> = [];

            for (let i = 0; i < displayValues.length; i++) {
                if (!displayValues[i].hasValue) continue;
                let x: number;
                if (chart.HasDateCategoryAxis) {
                    x = Area3DDateAxisMath.CategoryPosition(chart, i, count, 1900);
                } else {
                    x = count > 0 ? i / count : 0;
                }
                const y = Math.max(0, Math.min(1, (displayValues[i].value - min) / (max - min)));
                frontPoints.push({
                    x: plot.left + x * w,
                    y: plot.bottom - y * h
                });
                backPoints.push({
                    x: plot.left + x * w + depthX,
                    y: plot.bottom - y * h - depthOffset * 0.6 * d
                });
            }

            if (frontPoints.length < 2) continue;

            Area3DChartRenderer.DrawAreaBand(doc, frontPoints, backPoints, series);
        }
    }

    private static DrawAreaBand(
        doc: any,
        front: Array<{ x: number; y: number }>,
        back: Array<{ x: number; y: number }>,
        series: { Color: { red: number; green: number; blue: number; alpha: number } }
    ): void {
        const c = series.Color;
        const h = (v: number) => Math.round(v).toString(16).padStart(2, "0");
        const baseHex = `#${h(c.red)}${h(c.green)}${h(c.blue)}`;

        doc.save();
        doc.fillColor(baseHex);

        // 顶面轮廓
        doc.moveTo(front[0].x, front[0].y);
        for (let i = 1; i < front.length; i++) doc.lineTo(front[i].x, front[i].y);
        for (let i = back.length - 1; i >= 0; i--) doc.lineTo(back[i].x, back[i].y);
        doc.closePath();
        doc.fill();
        doc.restore();
    }
}

function Column3DWallFills(
    doc: any,
    plot: Rect,
    depthOffset: number
): void {
    // 后墙
    const backCorners = [
        { x: plot.left, y: plot.top },
        { x: plot.right, y: plot.top },
        { x: plot.right + depthOffset, y: plot.top - depthOffset * 0.6 },
        { x: plot.left + depthOffset, y: plot.top - depthOffset * 0.6 }
    ];
    Area3DWallFillRenderer.Fill(doc, backCorners, "#E8E8E8");

    // 侧墙
    const sideCorners = [
        { x: plot.left, y: plot.top },
        { x: plot.left, y: plot.bottom },
        { x: plot.left + depthOffset, y: plot.bottom - depthOffset * 0.6 },
        { x: plot.left + depthOffset, y: plot.top - depthOffset * 0.6 }
    ];
    Area3DWallFillRenderer.Fill(doc, sideCorners, "#F0F0F0");
}
