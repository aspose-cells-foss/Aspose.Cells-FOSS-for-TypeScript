import type { ParsedChart, ChartSeries, Rect } from "./ChartData.js";
import { ToPdfKitColor } from "./ChartData.js";

/**
 * Renders Excel pattern (hatch) fills used by chart elements.
 * PDFKit has no native tiling shader, so the tiles are stamped as vector rects.
 */
export class ChartPatternRenderer {
    /**
     * 12 x 12 bitmap of the Excel "weave" preset;
     * '#' marks foreground pixels and '.' marks background pixels.
     */
    private static readonly WeaveRows: string[] = [
        "...#....##..", "...#....##..", ".##.##....#.", "#.....#....#",
        "#.....#....#", "....##.#..#.", "...#....##..", "...#....##..",
        ".##.##......", "#.....#....#", "#.....#....#", ".##....#..#."
    ];

    /**
     * Draws the chart area "weave" pattern.
     * @returns true when the pattern was drawn, false when the caller should fall back to the normal fill.
     */
    static DrawChartArea(doc: any, rect: Rect, chart: ParsedChart): boolean {
        if (doc == null || chart == null) return false;

        const width = rect.right - rect.left;
        const height = rect.bottom - rect.top;
        if (width <= 0 || height <= 0) return false;
        if (!chart.HasChartAreaPattern || chart.ChartAreaPatternPreset !== "weave") return false;

        const foreground = chart.ChartAreaPatternForegroundColor.toPdfKitColor();
        const background = chart.ChartAreaPatternBackgroundColor.toPdfKitColor();
        // The 12 x 12 tile is drawn at half scale, so one tile covers 6 pt.
        const tile = 12 * 0.5;
        const pixel = 0.5;

        doc.save();

        doc.fillColor(background);
        doc.rect(rect.left, rect.top, width, height);
        doc.fill();

        // Clip the stamped tiles to the chart area.
        doc.rect(rect.left, rect.top, width, height);
        doc.clip();

        doc.fillColor(foreground);
        const startX = Math.floor(rect.left / tile) * tile;
        const startY = Math.floor(rect.top / tile) * tile;
        for (let ty = startY; ty < rect.bottom; ty += tile) {
            for (let tx = startX; tx < rect.right; tx += tile) {
                for (let row = 0; row < ChartPatternRenderer.WeaveRows.length; row++) {
                    const line = ChartPatternRenderer.WeaveRows[row];
                    let col = 0;
                    while (col < line.length) {
                        if (line[col] !== "#") {
                            col++;
                            continue;
                        }
                        const runStart = col;
                        while (col < line.length && line[col] === "#") col++;
                        doc.rect(tx + runStart * pixel, ty + row * pixel, (col - runStart) * pixel, pixel);
                    }
                }
            }
        }
        doc.fill();

        doc.restore();
        return true;
    }

    /**
     * Draws the "ltUpDiag" pattern inside a rectangle (series legend swatches).
     * @returns true when the pattern was drawn, false when the caller should fall back to the normal fill.
     */
    static DrawRect(doc: any, rect: Rect, series: ChartSeries, fallbackColor: string): boolean {
        if (doc == null || series == null) return false;

        const width = rect.right - rect.left;
        const height = rect.bottom - rect.top;
        if (width <= 0 || height <= 0) return false;
        if (!series.HasPatternFill || series.PatternPreset !== "ltUpDiag") return false;

        const background = series.PatternBackgroundColor.alpha > 0
            ? series.PatternBackgroundColor.toPdfKitColor()
            : fallbackColor;
        const foreground = series.PatternForegroundColor.alpha > 0
            ? series.PatternForegroundColor
            : series.Color;

        doc.save();

        doc.fillColor(background);
        doc.rect(rect.left, rect.top, width, height);
        doc.fill();

        // Clip the diagonal hatch to the rectangle.
        doc.rect(rect.left, rect.top, width, height);
        doc.clip();

        const run = Math.max(1, height);
        doc.lineWidth(0.35);
        doc.strokeColor(ToPdfKitColor({
            red: foreground.red,
            green: foreground.green,
            blue: foreground.blue,
            alpha: 150
        }));
        for (let x = rect.left - run; x <= rect.right; x += 2.2) {
            doc.moveTo(x, rect.bottom);
            doc.lineTo(x + run, rect.top);
        }
        doc.stroke();

        doc.restore();
        return true;
    }
}
