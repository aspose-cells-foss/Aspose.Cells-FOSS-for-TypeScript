import type { ParsedChart, Rect } from "./ChartData.js";
import { Line3DProjection } from "./Line3DProjection.js";
import { ChartStackingMath } from "./ChartStackingMath.js";
import { Area3DDateAxisMath } from "./Area3DDateAxisMath.js";

/**
 * Line 3D 图表渲染器 - 用投影带（ribbon）绘制折线。
 */
export class Line3DChartRenderer {
    Draw(
        doc: any,
        _page: any,
        plot: Rect,
        chart: ParsedChart,
        min: number,
        max: number,
        major: number
    ): void {
        if (chart == null || chart.Series.length === 0 || max <= min) return;
        const count = Math.max(chart.Categories.length, chart.CategoryValues.length);
        const projection = new Line3DProjection(plot, chart, count);
        Line3DChartRenderer.DrawBox(doc, projection, chart, min, max, major);
        for (let s = chart.Series.length - 1; s >= 0; s--) {
            Line3DChartRenderer.DrawRibbons(doc, projection, chart, s, min, max);
        }
    }

    private static DrawBox(
        doc: any,
        p: Line3DProjection,
        chart: ParsedChart,
        min: number,
        max: number,
        major: number
    ): void {
        doc.save();
        doc.strokeColor(chart.ValueGridlineColor.toPdfKitColor());
        doc.lineWidth(Math.max(0.5, chart.ValueGridlineWidthPt));
        if (major > 0) {
            const steps = Math.round((max - min) / major);
            for (let i = 0; i <= steps; i++) {
                const value = Math.min(max, min + i * major);
                const y = (value - min) / (max - min);
                const front = p.Project(0, y, 0);
                const backLeft = p.Project(0, y, 1);
                const backRight = p.Project(1, y, 1);
                doc.moveTo(front.x, front.y);
                doc.lineTo(backLeft.x, backLeft.y);
                doc.lineTo(backRight.x, backRight.y);
                doc.stroke();
            }
        }
        doc.restore();
    }

    private static DrawRibbons(
        doc: any,
        p: Line3DProjection,
        chart: ParsedChart,
        seriesIndex: number,
        min: number,
        max: number
    ): void {
        const series = chart.Series[seriesIndex];
        const displayValues = ChartStackingMath.BuildDisplayValues(chart, seriesIndex);
        const count = Math.max(chart.Categories.length, chart.CategoryValues.length);
        const pts: Array<{ x: number; y: number }> = [];
        for (let i = 0; i < displayValues.length; i++) {
            if (!displayValues[i].hasValue) continue;
            const x = Area3DDateAxisMath.CategoryPosition(chart, i, count, 1900);
            const y = Math.max(0, Math.min(1, (displayValues[i].value - min) / (max - min)));
            pts.push({ x, y });
        }
        if (pts.length < 2) return;
        pts.sort((a, b) => a.x - b.x);

        const c = series.Color;
        const h = (v: number) => Math.round(v).toString(16).padStart(2, "0");
        const shade = (factor: number) => `#${h(c.red * factor)}${h(c.green * factor)}${h(c.blue * factor)}`;

        const faces: Array<Array<{ x: number; y: number }>> = [];
        for (let i = 1; i < pts.length; i++) {
            faces.push(p.RibbonCorners(pts[i - 1], pts[i], seriesIndex));
        }

        doc.save();

        // Underpaint the continuous silhouette so antialiasing on
        // shared face edges cannot leave a white crack at a bend.
        doc.fillColor(shade(0.76));
        doc.moveTo(faces[0][0].x, faces[0][0].y);
        for (let i = 0; i < faces.length; i++) {
            doc.lineTo(faces[i][1].x, faces[i][1].y);
        }
        for (let i = faces.length - 1; i >= 0; i--) {
            doc.lineTo(faces[i][2].x, faces[i][2].y);
        }
        doc.lineTo(faces[0][3].x, faces[0][3].y);
        doc.closePath();
        doc.fill();

        for (let i = 0; i < faces.length; i++) {
            const corners = faces[i];
            doc.fillColor(shade(p.RibbonShade(pts[i], pts[i + 1], corners)));
            doc.moveTo(corners[0].x, corners[0].y);
            doc.lineTo(corners[1].x, corners[1].y);
            doc.lineTo(corners[2].x, corners[2].y);
            doc.lineTo(corners[3].x, corners[3].y);
            doc.closePath();
            doc.fill();
        }

        doc.restore();
    }
}
