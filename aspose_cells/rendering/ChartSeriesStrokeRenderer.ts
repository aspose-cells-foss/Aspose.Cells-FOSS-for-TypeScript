import type { ParsedChart, ChartSeries } from "./ChartData.js";

/**
 * 绘制序列的线条描边（带虚线样式处理）。
 */
export class ChartSeriesStrokeRenderer {
    /**
     * 使用序列的样式绘制一条路径。
     */
    static Draw(
        doc: any,
        points: Array<{ x: number; y: number }>,
        series: ChartSeries
    ): void {
        if (series == null || !series.HasVisibleLine || points.length < 2) {
            return;
        }

        const lineWidth = Math.max(0.5, series.LineWidthPt);
        const color = series.LineColor.toPdfKitColor();

        doc.save();
        doc.lineWidth(lineWidth);
        doc.strokeColor(color);
        doc.lineCap("round");
        doc.lineJoin("round");

        // 处理虚线样式
        if (series.LineDashStyle && series.LineDashStyle !== "solid" && series.LineDashStyle !== "auto") {
            const dashArray = ChartSeriesStrokeRenderer.DashArrayFor(series.LineDashStyle);
            if (dashArray) {
                doc.dash(dashArray[0], { space: dashArray[1] });
            }
        }

        doc.moveTo(points[0].x, points[0].y);
        for (let i = 1; i < points.length; i++) {
            doc.lineTo(points[i].x, points[i].y);
        }
        doc.stroke();
        doc.restore();
    }

    private static DashArrayFor(style: string): [number, number] | null {
        switch (style) {
            case "dash":
            case "sysDash":
                return [4, 2];
            case "dot":
            case "sysDot":
                return [1, 2];
            case "dashDot":
                return [4, 2];
            case "longDash":
                return [6, 2];
            default:
                return null;
        }
    }
}
