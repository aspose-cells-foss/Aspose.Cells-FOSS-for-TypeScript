import type { ParsedChart, Rect } from "./ChartData.js";
import { Column3DSceneRenderer } from "./Column3DSceneRenderer.js";

/**
 * Column 3D（纵向柱体图 3D）渲染器。
 * 组合场景渲染器完成绘制。
 */
export class Column3DChartRenderer {
    /**
     * 绘制 Column 3D 图表。
     */
    Draw(
        doc: any,
        _area: Rect,
        chart: ParsedChart,
        min: number,
        max: number,
        major: number
    ): void {
        // Column 3D 使用场景渲染器
        Column3DSceneRenderer.prototype.Draw.call(
            Column3DSceneRenderer.prototype,
            doc as any,
            { left: 0, top: 0, right: 0, bottom: 0 } as any,
            chart,
            min,
            max,
            major
        );
    }
}
