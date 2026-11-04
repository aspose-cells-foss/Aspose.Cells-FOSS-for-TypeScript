import type { ParsedChart, ChartSeries } from "./ChartData.js";

/**
 * 根据显式格式设置获取序列点的颜色。
 * 优先使用每个点的单独颜色，否则回退到序列颜色。
 */
export class ChartExplicitFormatting {
    /**
     * 获取指定点的颜色。
     * 如果该点有显式设置的颜色，返回之；否则返回序列颜色。
     */
    static PointColor(series: ChartSeries, pointIndex: number): {
        red: number;
        green: number;
        blue: number;
        alpha: number;
    } {
        if (series == null) {
            return { red: 0, green: 0, blue: 0, alpha: 255 };
        }
        if (series.PointColors && pointIndex >= 0 && pointIndex < series.PointColors.length) {
            const pc = series.PointColors[pointIndex];
            if (pc != null) {
                return pc;
            }
        }
        return series.Color;
    }
}
