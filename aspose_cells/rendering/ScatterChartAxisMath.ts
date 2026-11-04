import type { ParsedChart } from "./ChartData.js";

/**
 * 散点图 X 轴数学工具。
 * 处理 X 轴范围计算、刻度间隔和屏幕坐标映射。
 */
export class ScatterChartAxisMath {
    /**
     * 解析散点图 X 轴范围。
     */
    static ResolveXRange(
        chart: ParsedChart
    ): { min: number; max: number; major: number } {
        if (chart == null) {
            return { min: 0, max: 1, major: 1 };
        }

        let min = Infinity;
        let max = -Infinity;

        // 显式设置优先
        if (chart.ScatterXAxisMin != null && chart.ScatterXAxisMin.hasValue) {
            min = chart.ScatterXAxisMin.value;
        } else {
            for (const series of chart.Series) {
                for (const v of series.XValues) {
                    if (v.hasValue) {
                        min = Math.min(min, v.value);
                        max = Math.max(max, v.value);
                    }
                }
            }
        }

        if (min === Infinity) {
            min = 0;
            max = 1;
        }
        if (max <= min) {
            max = min + 1;
        }

        if (chart.ScatterXAxisMax != null && chart.ScatterXAxisMax.hasValue) {
            max = chart.ScatterXAxisMax.value;
        }

        let major = chart.ScatterXAxisMajorUnit ?? 0;
        if (major <= 0) {
            major = ScatterChartAxisMath.AutoMajorUnit(min, max);
        }

        return { min, max, major };
    }

    /**
     * 自动计算合适的刻度间隔。
     */
    static AutoMajorUnit(min: number, max: number): number {
        const range = max - min;
        if (range <= 0) return 1;
        const raw = range / 5;
        const exp = Math.pow(10, Math.floor(Math.log10(raw)));
        const f = raw / exp;
        let niceF: number;
        if (f <= 1.5) niceF = 1;
        else if (f <= 3) niceF = 2;
        else if (f <= 7) niceF = 5;
        else niceF = 10;
        return niceF * exp;
    }

    /**
     * 将数值转换为屏幕 X 坐标。
     */
    static ToScreen(
        value: number,
        min: number,
        max: number,
        left: number,
        right: number
    ): number {
        if (max <= min) return left;
        const ratio = (value - min) / (max - min);
        return left + ratio * (right - left);
    }
}
