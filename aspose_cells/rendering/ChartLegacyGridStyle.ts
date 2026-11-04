import type { ParsedChart } from "./ChartData.js";

/**
 * 处理旧版 Excel 网格线样式（虚线、点线等）。
 * 在 PDFKit 中通过 dash 方法实现。
 */
export class ChartLegacyGridStyle {
    /**
     * 将图表的网格线样式转换为 PDFKit 可用的 dash 参数。
     * 返回 null 表示实线。
     */
    static GetDashArray(chart: ParsedChart, isMajor: boolean): number[] | null {
        if (chart == null) return null;

        const dashStyle = isMajor
            ? chart.ValueGridlineDashStyle
            : chart.MinorValueGridlineDashStyle;

        if (!dashStyle || dashStyle === "solid" || dashStyle === "auto") {
            return null;
        }

        switch (dashStyle) {
            case "dash":
            case "sysDash":
                return [4, 2];
            case "dot":
            case "sysDot":
                return [1, 2];
            case "dashDot":
            case "sysDashDot":
                return [4, 2, 1, 2];
            case "dashDotDot":
            case "sysDashDotDot":
                return [4, 2, 1, 2, 1, 2];
            case "longDash":
                return [6, 2];
            default:
                return null;
        }
    }
}
