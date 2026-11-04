import type { ParsedChart, ChartSeries } from "./ChartData.js";

/**
 * 计算图表堆叠值的数学工具类。
 * 处理普通、堆叠和百分比堆叠三种模式下的数据显示值计算。
 */
export class ChartStackingMath {
    /**
     * 构建指定序列的显示值列表。
     * 根据图表的堆叠模式返回原始值、累加值或百分比值。
     */
    static BuildDisplayValues(
        chart: ParsedChart,
        seriesIndex: number
    ): Array<{ hasValue: boolean; value: number }> {
        if (chart == null || seriesIndex < 0 || seriesIndex >= chart.Series.length) {
            return [];
        }

        const series = chart.Series[seriesIndex];
        const values: Array<{ hasValue: boolean; value: number }> = [];
        const count = ChartStackingMath.MaxPointCount(chart);

        let positiveTotals: number[] | null = null;
        let negativeTotals: number[] | null = null;

        if (chart.IsPercentStacked) {
            positiveTotals = new Array(count).fill(0);
            negativeTotals = new Array(count).fill(0);
            for (let s = 0; s < chart.Series.length; s++) {
                const other = chart.Series[s];
                for (let i = 0; i < Math.min(count, other.Values.length); i++) {
                    if (!other.Values[i].hasValue) continue;
                    const v = other.Values[i].value;
                    if (v >= 0) {
                        positiveTotals[i] += v;
                    } else {
                        negativeTotals[i] += Math.abs(v);
                    }
                }
            }
        }

        const positiveStack = new Array(count).fill(0);
        const negativeStack = new Array(count).fill(0);

        for (let s = 0; s <= seriesIndex; s++) {
            const current = chart.Series[s];
            for (let i = 0; i < Math.min(count, current.Values.length); i++) {
                if (!current.Values[i].hasValue) continue;
                let v = current.Values[i].value;

                if (chart.IsPercentStacked && positiveTotals && negativeTotals) {
                    if (v >= 0) {
                        if (positiveTotals[i] <= 0) continue;
                        v /= positiveTotals[i];
                    } else {
                        if (negativeTotals[i] <= 0) continue;
                        v /= negativeTotals[i];
                    }
                }

                if (s === seriesIndex) {
                    let startValue: number;
                    let endValue: number;
                    if (v >= 0) {
                        startValue = positiveStack[i];
                        endValue = startValue + v;
                    } else {
                        startValue = negativeStack[i];
                        endValue = startValue + v;
                    }
                    values.push({ hasValue: true, value: endValue });
                }

                if (v >= 0) {
                    positiveStack[i] += v;
                } else {
                    negativeStack[i] += v;
                }
            }
        }

        return values;
    }

    private static MaxPointCount(chart: ParsedChart): number {
        let count = Math.max(chart.Categories.length, chart.CategoryValues.length);
        for (const series of chart.Series) {
            count = Math.max(count, series.Values.length);
        }
        return count;
    }
}
