import type { ParsedChart } from "./ChartData.js";

/**
 * 折线图的自动布局计算。
 * 处理标题位置、绘图区域调整等。
 */
export class LineChartAutomaticLayout {
    /**
     * 根据图表类型和尺寸，计算绘图区域矩形。
     */
    static Plot(
        area: { left: number; top: number; right: number; bottom: number },
        chart: ParsedChart,
        axisLabelMaxWidth: number,
        rotationDeg: number,
        initialRect: { left: number; top: number; right: number; bottom: number }
    ): { left: number; top: number; right: number; bottom: number } {
        if (chart == null) return initialRect;

        let result = { ...initialRect };

        // 有旋转标签时，给底部更多空间
        if (Math.abs(rotationDeg) > 0.1) {
            result.bottom = Math.min(result.bottom + 8, area.bottom - 2);
        }

        // 保证绘图区域不超出整体区域
        result.left = Math.max(result.left, area.left + 2);
        result.top = Math.max(result.top, area.top + 2);
        result.right = Math.min(result.right, area.right - 2);
        result.bottom = Math.min(result.bottom, area.bottom - 2);

        return result;
    }
}
