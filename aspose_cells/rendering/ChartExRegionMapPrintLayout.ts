import type { ParsedChart } from "./ChartData.js";

/**
 * Excel 扩展区域地图的打印布局计算。
 * 处理地图图表在 PDF 页面中的位置和尺寸。
 */
export class ChartExRegionMapPrintLayout {
    /**
     * 计算地图图表的打印区域矩形。
     */
    static Resolve(
        pageArea: { left: number; top: number; right: number; bottom: number },
        chart: ParsedChart
    ): { left: number; top: number; right: number; bottom: number } {
        if (chart == null) {
            return pageArea;
        }

        // 默认地图在页面中央，占据可用区域的 80%
        const padding = Math.min(pageArea.right - pageArea.left, pageArea.bottom - pageArea.top) * 0.1;

        const left = pageArea.left + padding;
        const top = pageArea.top + padding * 1.2;
        const right = pageArea.right - padding;
        const bottom = pageArea.bottom - padding * 0.8;

        return { left, top, right, bottom };
    }
}
