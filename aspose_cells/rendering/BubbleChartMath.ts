import type { ParsedChart, ChartSeries } from "./ChartData.js";
import { ScatterChartAxisMath } from "./ScatterChartAxisMath.js";

/**
 * 气泡图的数学计算和渲染。
 * 计算气泡大小并绘制散点。
 */
export class BubbleChartMath {
    /**
     * 绘制气泡图的所有序列。
     */
    static DrawSeries(
        doc: any,
        plot: { left: number; top: number; right: number; bottom: number },
        chart: ParsedChart,
        min: number,
        max: number
    ): void {
        if (chart == null) return;

        const { min: xMin, max: xMax } = ScatterChartAxisMath.ResolveXRange(chart);

        for (let s = 0; s < chart.Series.length; s++) {
            const series = chart.Series[s];
            const count = Math.min(series.XValues.length, series.Values.length, series.BubbleSizes?.length ?? Infinity);

            for (let i = 0; i < count; i++) {
                if (!series.XValues[i].hasValue || !series.Values[i].hasValue) continue;

                const x = ScatterChartAxisMath.ToScreen(series.XValues[i].value, xMin, xMax, plot.left, plot.right);
                const y = BubbleChartMath.ValueToY(series.Values[i].value, min, max, plot);
                const size = BubbleChartMath.ComputeBubbleSize(series, i, chart.Series.length);

                if (series.Bubble3D) {
                    BubbleChartMath.DrawBubbleSphere(doc, x, y, size, series);
                } else {
                    BubbleChartMath.DrawBubble(doc, x, y, size, series);
                }
            }
        }
    }

    private static DrawBubble(
        doc: any,
        x: number,
        y: number,
        size: number,
        series: ChartSeries
    ): void {
        doc.save();
        doc.fillColor(series.Color.toPdfKitColor());
        doc.circle(x, y, size * 0.5);
        doc.fill();
        doc.restore();
    }

    private static DrawBubbleSphere(
        doc: any,
        x: number,
        y: number,
        size: number,
        series: ChartSeries
    ): void {
        BubbleSphereRenderer.Draw(doc, { x, y }, size * 0.5, series.Color.toPdfKitColor());
    }

    static DrawMarker(
        doc: any,
        point: { x: number; y: number },
        radius: number,
        series: ChartSeries
    ): void {
        doc.save();
        doc.fillColor(series.MarkerFillColor.toPdfKitColor());
        doc.circle(point.x, point.y, radius);
        doc.fill();
        doc.restore();
    }

    private static ComputeBubbleSize(series: ChartSeries, index: number, seriesCount: number): number {
        const baseSize = 12; // PDF points
        if (series.BubbleSizes && index < series.BubbleSizes.length && series.BubbleSizes[index].hasValue) {
            // Excel 的气泡大小是百分比，这里映射到 PDF points
            const pct = Math.max(5, Math.min(300, series.BubbleSizes[index].value));
            return baseSize * (pct / 100);
        }
        return baseSize;
    }

    private static ValueToY(
        value: number,
        min: number,
        max: number,
        plot: { top: number; bottom: number }
    ): number {
        if (max <= min) return plot.bottom;
        const ratio = (value - min) / (max - min);
        return plot.bottom - ratio * (plot.bottom - plot.top);
    }
}

import { BubbleSphereRenderer } from "./BubbleSphereRenderer.js";
