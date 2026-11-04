import type { ParsedChart } from "./ChartData.js";
import { ChartDateAxisMath } from "./ChartDateAxisMath.js";

/**
 * 日期轴的布局信息。
 * 存储标签旋转角度、刻度间隔等。
 */
export interface DateAxisLayout {
    /** 标签旋转角度（度）。正值表示逆时针。 */
    RotationDeg: number;
    /** 主刻度间隔（月数或日数，取决于单位）。 */
    MajorUnit: number;
    /** 刻度单位类型：months, days, years, hours 等。 */
    TimeUnit: string;
    /** 是否需要隐藏部分刻度标签以避免重叠。 */
    NeedsThinning: boolean;
}

/**
 * 计算日期轴的标签旋转和布局参数。
 */
export class ChartDateAxisLayout {
    /**
     * 解析日期轴的完整布局信息。
     */
    static Resolve(
        chart: ParsedChart,
        dateSystem: number,
        availableWidth: number
    ): DateAxisLayout {
        if (chart == null || !chart.HasDateCategoryAxis) {
            return { RotationDeg: 0, MajorUnit: 1, TimeUnit: "days", NeedsThinning: false };
        }

        // 从图表属性获取设置，或根据数据范围自动计算
        const layout: DateAxisLayout = {
            RotationDeg: Math.abs(chart.CategoryAxisTextRotationDeg) || 0,
            MajorUnit: chart.DateAxisMajorUnit ?? 1,
            TimeUnit: chart.CategoryBaseTimeUnit ?? "days",
            NeedsThinning: false
        };

        // 自动旋转：如果标签太密集则旋转 45 度
        if (layout.RotationDeg === 0 && availableWidth > 0) {
            const pointCount = Math.max(
                chart.Categories.length,
                chart.CategoryValues.length
            );
            if (pointCount > 0 && availableWidth / pointCount < 42) {
                layout.RotationDeg = 45;
                layout.NeedsThinning = true;
            }
        }

        return layout;
    }
}
