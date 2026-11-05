import type { ParsedChart, Rect } from "./ChartData.js";

/**
 * Pie 3D chart renderer.
 * Draws the projected top and visible front rim of a three-dimensional pie.
 */
export class Pie3DChartRenderer {
    /**
     * Draws a Pie 3D chart and returns the projected ellipse area.
     */
    static Draw(
        doc: any,
        area: Rect,
        chart: ParsedChart,
        values: Array<{ hasValue: boolean; value: number }>,
        colors: Array<{ red: number; green: number; blue: number; alpha: number }>,
        sum: number
    ): Rect {
        const perspective = chart.RightAngleAxes ? 0
            : Math.sin((chart.PerspectiveDeg ?? 30) * Math.PI / 180) * 0.68;
        const p = Math.max(0, Math.min(0.6, perspective));

        const flatView = 1 - Math.min(1, p / 0.34);
        const projectedRadiusX = (area.right - area.left) * (0.45 - 0.03 * flatView);
        const radiusX = projectedRadiusX * Math.sqrt(1 - p * p);
        const tilt = Math.sin(Math.max(0, Math.min(70, chart.RotationXDeg ?? 15)) * Math.PI / 180);
        const radiusY = (area.bottom - area.top) * (0.15 + 0.065 * flatView)
            * Math.max(0.2, tilt / Math.sin(15 * Math.PI / 180));

        const centerX = (area.left + area.right) / 2;
        const centerY = area.top + (area.bottom - area.top) * (0.303 + 0.107 * flatView);

        const depth = (area.bottom - area.top) * 0.193
            * Math.max(0.25, Math.min(2, (chart.DepthPercent ?? 100) / 100));

        const ellipse = {
            left: centerX - projectedRadiusX,
            top: centerY - radiusY / (1 + p),
            right: centerX + projectedRadiusX,
            bottom: centerY + radiusY / (1 - p)
        };

        if (sum <= 0) return ellipse;

        const start = -90 + (chart.RotationYDeg ?? 0);
        const sliceStarts: number[] = [];
        let current = start;
        for (let i = 0; i < values.length; i++) {
            sliceStarts[i] = current;
            if (!values[i].hasValue) continue;

            const sweep = (Math.abs(values[i].value) / sum) * 360;
            Pie3DChartRenderer.DrawVisibleRim(
                doc, centerX, centerY, radiusX, radiusY, p, depth,
                current, sweep, colors[i] ?? colors[0]
            );
            current += sweep;
        }

        for (let i = values.length - 1; i >= 0; i--) {
            if (!values[i].hasValue) continue;

            const sweep = (Math.abs(values[i].value) / sum) * 360;
            const c = colors[i] ?? colors[0];
            doc.save();
            doc.fillColor(Pie3DChartRenderer.Darken(c, 0.92));
            doc.moveTo(centerX, centerY);
            Pie3DChartRenderer.AddProjectedArc(
                doc, centerX, centerY, radiusX, radiusY, p, depth,
                sliceStarts[i], sliceStarts[i] + sweep, false, false
            );
            doc.closePath();
            doc.fill();
            doc.restore();
        }

        return ellipse;
    }

    private static DrawVisibleRim(
        doc: any,
        centerX: number, centerY: number, radiusX: number, radiusY: number,
        p: number, depth: number,
        start: number, sweep: number,
        color: { red: number; green: number; blue: number; alpha: number }
    ): void {
        const silhouetteAngle = Math.asin(p) * 180 / Math.PI;
        const visibleStart = Math.max(silhouetteAngle, start);
        const visibleEnd = Math.min(180 - silhouetteAngle, start + sweep);
        if (visibleEnd <= visibleStart) return;

        doc.save();
        Pie3DChartRenderer.AddProjectedArc(
            doc, centerX, centerY, radiusX, radiusY, p, depth,
            visibleStart, visibleEnd, false, true
        );
        Pie3DChartRenderer.AddProjectedArc(
            doc, centerX, centerY, radiusX, radiusY, p, depth,
            visibleEnd, visibleStart, true, false
        );
        doc.closePath();

        // The rim is lit from the upper left: bright near the left silhouette,
        // falling off towards the right silhouette.
        const gradient = doc.linearGradient(
            centerX - radiusX, 0,
            centerX + radiusX, 0
        );
        gradient.stop(0, Pie3DChartRenderer.Darken(color, 0.95));
        gradient.stop(0.5, Pie3DChartRenderer.Darken(color, 0.70));
        gradient.stop(1, Pie3DChartRenderer.Darken(color, 0.35));
        doc.fill(gradient);
        doc.restore();
    }

    private static AddProjectedArc(
        doc: any,
        centerX: number, centerY: number, radiusX: number, radiusY: number,
        p: number, depth: number,
        start: number, end: number, lower: boolean, moveToStart: boolean
    ): void {
        const segments = Math.max(1, Math.ceil(Math.abs(end - start) / 2));
        for (let i = 0; i <= segments; i++) {
            const angle = start + (end - start) * i / segments;
            const point = Pie3DChartRenderer.ProjectedPoint(
                centerX, centerY, radiusX, radiusY, p, depth, angle, lower
            );
            if (i === 0 && moveToStart) {
                doc.moveTo(point.x, point.y);
            } else {
                doc.lineTo(point.x, point.y);
            }
        }
    }

    private static ProjectedPoint(
        centerX: number, centerY: number, radiusX: number, radiusY: number,
        p: number, depth: number, angleDeg: number, lower: boolean
    ): { x: number; y: number } {
        const radians = angleDeg * Math.PI / 180;
        const sine = Math.sin(radians);
        const divisor = 1 - p * sine;
        let x = centerX + radiusX * Math.cos(radians) / divisor;
        const y = centerY + (radiusY * sine + (lower ? depth : 0)) / divisor;
        if (lower) {
            x = centerX + (x - centerX) * 0.978;
        }
        return { x, y };
    }

    private static Darken(
        color: { red: number; green: number; blue: number; alpha: number },
        factor: number
    ): string {
        const h = (v: number) => Math.round(v * factor).toString(16).padStart(2, "0");
        return `#${h(color.red)}${h(color.green)}${h(color.blue)}`;
    }
}
