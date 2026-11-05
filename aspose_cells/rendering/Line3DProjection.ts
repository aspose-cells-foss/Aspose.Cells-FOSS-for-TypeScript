import type { ParsedChart } from "./ChartData.js";

/**
 * Line3D perspective projection.
 * Category, value and depth coordinates all use the same perspective camera.
 * Ribbon width is a depth interval, never a screen-space stroke width.
 */
export class Line3DProjection {
    private readonly sinX: number;
    private readonly cosX: number;
    private readonly sinY: number;
    private readonly cosY: number;
    private readonly height: number;
    private readonly depth: number;
    private readonly perspective: number;
    private readonly cameraOffsetX: number;
    private readonly cameraOffsetY: number;
    private readonly seriesCount: number;
    private readonly ribbonDepth: number;
    private readonly bounds: { left: number; top: number; right: number; bottom: number };
    private readonly cameraBounds: { left: number; top: number; right: number; bottom: number };

    constructor(
        area: { left: number; top: number; right: number; bottom: number },
        chart: ParsedChart,
        categoryCount: number
    ) {
        const rx = (chart.RotationXDeg ?? 0) * Math.PI / 180;
        const ry = (chart.RotationYDeg ?? 0) * Math.PI / 180;
        this.sinX = Math.sin(rx);
        this.cosX = Math.cos(rx);
        this.sinY = Math.sin(ry);
        this.cosY = Math.cos(ry);
        this.seriesCount = Math.max(1, chart.Series.length);
        this.depth = this.seriesCount / Math.max(1, categoryCount)
            * Math.max(0.01, (chart.DepthPercent ?? 100) / 100);
        // A deep series box must not flatten the value dimension or move
        // its near plane towards the camera singularity. Normalize both
        // height and viewing distance by the horizontal extent of the box.
        const extent = Math.max(1, this.depth);
        const plotWidth = Math.max(1, area.right - area.left);
        this.height = Math.max(0.05, (area.bottom - area.top) / plotWidth * 0.76)
            * Math.max(1, this.depth * 0.875);
        this.ribbonDepth = 1 / (this.seriesCount
            * (1 + Math.max(0, chart.GapDepthPercent ?? 150) / 100));
        const framing = chart.RightAngleAxes ? 0
            : Math.max(0, Math.min(1, ((chart.PerspectiveDeg ?? 30) - 30) / 30));
        let perspective = chart.RightAngleAxes ? 0
            : Math.tan(Math.min(120, Math.max(0, chart.PerspectiveDeg ?? 30)) * Math.PI / 360)
                / (0.475 * extent * (1 + 0.22 * framing));
        // Excel's stronger perspective is an off-axis view, with the
        // vanishing point inside the plot rather than above its near edge.
        this.cameraOffsetX = 0.21 * framing;
        this.cameraOffsetY = -0.26 * framing;
        const radius = Math.abs(this.cosX * this.sinY) * 0.5
            + Math.abs(this.cosX * this.cosY) * this.depth * 0.5
            + Math.abs(this.sinX) * this.height * 0.5;
        perspective = Math.min(perspective, 0.9 / Math.max(0.001, radius));
        this.perspective = perspective;

        // Reserve the label bands around the complete projected box, not
        // around independently positioned walls and data series.
        this.bounds = {
            left: area.left + plotWidth * (0.048 + framing * 0.109),
            top: area.top + (area.bottom - area.top) * (0.045 - framing * 0.072),
            right: area.left + plotWidth * (0.876 - framing * 0.131),
            bottom: area.top + (area.bottom - area.top) * (0.965 - framing * 0.05)
        };

        let minX = Number.MAX_VALUE;
        let minY = Number.MAX_VALUE;
        let maxX = -Number.MAX_VALUE;
        let maxY = -Number.MAX_VALUE;
        for (let x = 0; x <= 1; x++) {
            for (let y = 0; y <= 1; y++) {
                for (let z = 0; z <= 1; z++) {
                    const point = this.CameraPoint(x, y, z);
                    minX = Math.min(minX, point.x);
                    minY = Math.min(minY, point.y);
                    maxX = Math.max(maxX, point.x);
                    maxY = Math.max(maxY, point.y);
                }
            }
        }

        this.cameraBounds = { left: minX, top: minY, right: maxX, bottom: maxY };
    }

    /**
     * Projects category [0,1], value [0,1] and depth [0,1] into the reserved bounds.
     */
    Project(category: number, value: number, depth: number): { x: number; y: number } {
        const point = this.CameraPoint(category, value, depth);
        return {
            x: this.bounds.left + (point.x - this.cameraBounds.left)
                / Math.max(0.000001, this.cameraBounds.right - this.cameraBounds.left)
                * (this.bounds.right - this.bounds.left),
            y: this.bounds.top + (point.y - this.cameraBounds.top)
                / Math.max(0.000001, this.cameraBounds.bottom - this.cameraBounds.top)
                * (this.bounds.bottom - this.bounds.top)
        };
    }

    /**
     * Returns the normalized depth-plane center of a series.
     */
    SeriesDepth(index: number): number {
        return (index + 0.5) / this.seriesCount;
    }

    /**
     * Computes the four projected corners of one ribbon face.
     */
    RibbonCorners(
        prev: { x: number; y: number },
        curr: { x: number; y: number },
        seriesIndex: number
    ): Array<{ x: number; y: number }> {
        const center = this.SeriesDepth(seriesIndex);
        const front = center - this.ribbonDepth * 0.5;
        const back = center + this.ribbonDepth * 0.5;
        return [
            this.Project(prev.x, prev.y, front),
            this.Project(curr.x, curr.y, front),
            this.Project(curr.x, curr.y, back),
            this.Project(prev.x, prev.y, back)
        ];
    }

    /**
     * Computes the lighting factor of a ribbon face.
     */
    RibbonShade(
        prev: { x: number; y: number },
        curr: { x: number; y: number },
        corners: Array<{ x: number; y: number }>
    ): number {
        const dx = curr.x - prev.x;
        const dy = (curr.y - prev.y) * this.height;
        const length = Math.sqrt(dx * dx + dy * dy);
        if (length < 0.000001) {
            return 0.76;
        }

        const edgeX = corners[1].x - corners[0].x;
        const edgeY = corners[1].y - corners[0].y;
        const depthX = corners[3].x - corners[0].x;
        const depthY = corners[3].y - corners[0].y;
        // A steep rising segment can expose the underside of the ribbon.
        // The projected winding distinguishes it from an upward-facing face.
        if (edgeX * depthY - edgeY * depthX > 0) {
            return 0.40;
        }

        const normalX = -dy / length;
        return Math.max(0.4, Math.min(0.81, 0.76 + 0.25 * normalX));
    }

    private CameraPoint(category: number, value: number, depth: number): { x: number; y: number } {
        const x = category - 0.5;
        const y = (value - 0.5) * this.height;
        const z = (depth - 0.5) * this.depth;
        const yawDepth = -this.sinY * x + this.cosY * z;
        const cameraDepth = this.cosX * yawDepth - this.sinX * y;
        const divisor = 1 + this.perspective * cameraDepth;
        return {
            x: (this.cosY * x + this.sinY * z + this.cameraOffsetX) / divisor,
            y: (-(this.cosX * y + this.sinX * yawDepth) + this.cameraOffsetY) / divisor
        };
    }
}
