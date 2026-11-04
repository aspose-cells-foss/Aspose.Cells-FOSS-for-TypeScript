import type { ParsedChart, Rect } from "./ChartData.js";
import { Column3DSceneRenderer } from "./Column3DSceneRenderer.js";

/**
 * Bar 3D 渲染器（横向柱体 3D）。
 * 与 Column 3D 共享场景渲染逻辑。
 */
export class Bar3DChartRenderer {
    Draw(
        doc: any,
        _area: Rect,
        plot: Rect,
        chart: ParsedChart,
        min: number,
        max: number,
        major: number
    ): void {
        new Column3DSceneRenderer().Draw(doc, plot, chart, min, max, major);
    }
}
