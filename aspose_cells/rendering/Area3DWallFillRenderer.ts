

/**
 * Area 3D 墙壁填充渲染器。
 */
export class Area3DWallFillRenderer {
    static Fill(
        doc: any,
        corners: Array<{ x: number; y: number }>,
        colorHex: string
    ): void {
        if (corners.length < 3) return;
        doc.save();
        doc.fillColor(colorHex);
        doc.moveTo(corners[0].x, corners[0].y);
        for (let i = 1; i < corners.length; i++) doc.lineTo(corners[i].x, corners[i].y);
        doc.closePath();
        doc.fill();
        doc.restore();
    }
}
