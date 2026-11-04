/**
 * 3D 柱体几何工具函数。
 * 计算各种柱体形状（box, cylinder, cone, pyramid, tube）的参数。
 */
export class Column3DGeometry {
    static ShapeType(shape: string): string {
        switch (shape?.toLowerCase()) {
            case "box":
            case "rectprism":
            default:
                return "box";
            case "cylinder":
                return "cylinder";
            case "cone":
                return "cone";
            case "pyramid":
                return "pyramid";
            case "tube":
                return "tube";
        }
    }

    static TopWidthRatio(shape: string, heightRatio: number): number {
        const type = Column3DGeometry.ShapeType(shape);
        if (type === "cone" || type === "pyramid") {
            return Math.max(0, 1 - heightRatio);
        }
        return 1;
    }
}
