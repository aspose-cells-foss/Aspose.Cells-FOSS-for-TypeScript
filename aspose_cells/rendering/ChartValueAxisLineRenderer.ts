import type { ParsedChart } from "./ChartData.js";
import { ChartLegacyGridStyle } from "./ChartLegacyGridStyle.js";

/**
 * 绘制数值轴的轴线。
 * 处理轴线、主刻度、次刻度的显式样式。
 */
export class ChartValueAxisLineRenderer {
    /**
     * 绘制数值轴线和刻度。
     */
    static Draw(
        doc: any,
        plot: { left: number; top: number; right: number; bottom: number },
        chart: ParsedChart,
        min: number,
        max: number,
        major: number
    ): void {
        if (chart == null || major <= 0 || !chart.HasValueAxisLine) {
            return;
        }

        const lineWidth = Math.max(0.5, chart.ValueAxisLineWidthPt);
        const color = chart.ValueAxisLineColor.toPdfKitColor();

        doc.save();
        doc.lineWidth(lineWidth);
        doc.strokeColor(color);

        // 绘制主轴线（左侧）
        doc.moveTo(plot.left, plot.bottom);
        doc.lineTo(plot.left, plot.top);
        doc.stroke();

        // 主刻度
        const majorTick = chart.ValueAxisMajorTickMark;
        if (majorTick && majorTick !== "none") {
            const steps = Math.round((max - min) / major);
            for (let i = 0; i <= steps; i++) {
                const value = min + i * major;
                if (value > max + major * 0.001) break;
                const y = ChartValueAxisLineRenderer.ValueToY(value, min, max, plot);

                if (majorTick === "out" || majorTick === "cross") {
                    doc.moveTo(plot.left, y);
                    doc.lineTo(plot.left - 3, y);
                    doc.stroke();
                }
                if (majorTick === "in" || majorTick === "cross") {
                    doc.moveTo(plot.left, y);
                    doc.lineTo(plot.left + 3, y);
                    doc.stroke();
                }
            }
        }

        doc.restore();
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
