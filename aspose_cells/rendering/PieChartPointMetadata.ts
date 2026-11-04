import type { ParsedChart } from "./ChartData.js";

/**
 * 饼图点元数据。
 * 存储每个扇区的起始角度、扫掠角度、是否分离等信息。
 */
export interface PieSliceInfo {
    startAngle: number;  // 度
    sweepAngle: number;  // 度（正值）
    isExploded: boolean;
    explodeDistance: number;  // PDF points
    value: number;
    label: string;
    colorIndex: number;
}

/**
 * 饼图/环形图的点元数据计算工具。
 */
export class PieChartPointMetadata {
    /**
     * 计算所有扇区的元数据。
     */
    static ComputeSlices(chart: ParsedChart): PieSliceInfo[] {
        if (chart == null || chart.Series.length === 0) {
            return [];
        }
        const series = chart.Series[0];
        const values = series.Values;
        let sum = 0;
        for (const v of values) {
            if (v.hasValue) sum += Math.abs(v.value);
        }
        if (sum <= 0) return [];

        const slices: PieSliceInfo[] = [];
        let startAngle = chart.PieRotationDeg ?? -90;

        for (let i = 0; i < values.length; i++) {
            if (!values[i].hasValue) continue;
            const sweep = (Math.abs(values[i].value) / sum) * 360;
            const explodeDist = series.ExplodePoints?.includes(i)
                ? (series.ExplodeDistancePt ?? 0)
                : 0;
            slices.push({
                startAngle,
                sweepAngle: sweep,
                isExploded: explodeDist > 0,
                explodeDistance: explodeDist,
                value: values[i].value,
                label: chart.Categories[i] ?? "",
                colorIndex: i
            });
            startAngle += sweep;
        }

        return slices;
    }
}
