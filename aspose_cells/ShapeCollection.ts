import { CellsException } from "./CellsException";
import { AutoShapeType } from "./enums";
import { Shape } from "./Shape";
import { ShapeModel } from "./Core/ShapeModel";

// Implemented by serializer module.
/** Represents a collection of drawing shapes on a worksheet. */
export class ShapeCollection {
  private readonly models: ShapeModel[];

  public constructor(models: ShapeModel[]) {
    this.models = models;
  }

  public get Count(): number {
    return this.models.length;
  }

  public get(index: number): Shape {
    if (index < 0 || index >= this.models.length) {
      throw new CellsException(`Shape index ${index} is out of range.`);
    }

    return new Shape(this.models[index]);
  }

  public at(index: number): Shape {
    return this.get(index);
  }

  public Add(
    upperLeftRow: number,
    upperLeftColumn: number,
    lowerRightRow: number,
    lowerRightColumn: number,
    shapeType: AutoShapeType,
  ): number {
    if (upperLeftRow < 0 || upperLeftColumn < 0) {
      throw new CellsException("Shape anchor indices must be non-negative.");
    }

    if (lowerRightRow < upperLeftRow || lowerRightColumn < upperLeftColumn) {
      throw new CellsException("Shape anchor end must not precede the start.");
    }

    const model = new ShapeModel();
    model.Name = `Shape ${this.models.length + 1}`;
    model.DrawingObjectId = this.models.length + 2;
    model.UpperLeftRow = upperLeftRow;
    model.UpperLeftColumn = upperLeftColumn;
    model.LowerRightRow = lowerRightRow;
    model.LowerRightColumn = lowerRightColumn;
    model.GeometryType = ShapeCollection.toGeometryType(shapeType);
    this.models.push(model);
    return this.models.length - 1;
  }

  public RemoveAt(index: number): void {
    if (index < 0 || index >= this.models.length) {
      throw new CellsException(`Shape index ${index} is out of range.`);
    }

    this.models.splice(index, 1);
  }

  public *[Symbol.iterator](): IterableIterator<Shape> {
    for (let index = 0; index < this.models.length; index++) {
      yield this.get(index);
    }
  }

  private static toGeometryType(shapeType: AutoShapeType): string {
    switch (shapeType) {
      case AutoShapeType.Rectangle:
        return "rect";
      case AutoShapeType.RoundedRectangle:
        return "roundRect";
      case AutoShapeType.Ellipse:
        return "ellipse";
      case AutoShapeType.Triangle:
        return "triangle";
      case AutoShapeType.RightTriangle:
        return "rtTriangle";
      case AutoShapeType.Diamond:
        return "diamond";
      case AutoShapeType.Pentagon:
        return "pentagon";
      case AutoShapeType.Hexagon:
        return "hexagon";
      case AutoShapeType.Octagon:
        return "octagon";
      case AutoShapeType.Plus:
        return "plus";
      case AutoShapeType.Cube:
        return "cube";
      case AutoShapeType.Cylinder:
        return "cyl";
      case AutoShapeType.Heart:
        return "heart";
      case AutoShapeType.Lightning:
        return "lightningBolt";
      case AutoShapeType.Sun:
        return "sun";
      case AutoShapeType.Moon:
        return "moon";
      case AutoShapeType.Cloud:
        return "cloud";
      case AutoShapeType.RightArrow:
        return "rightArrow";
      case AutoShapeType.LeftArrow:
        return "leftArrow";
      case AutoShapeType.UpArrow:
        return "upArrow";
      case AutoShapeType.DownArrow:
        return "downArrow";
      case AutoShapeType.LeftRightArrow:
        return "leftRightArrow";
      case AutoShapeType.UpDownArrow:
        return "upDownArrow";
      case AutoShapeType.Star4Point:
        return "star4";
      case AutoShapeType.Star5Point:
        return "star5";
      case AutoShapeType.Star6Point:
        return "star6";
      case AutoShapeType.Star7Point:
        return "star7";
      case AutoShapeType.Star8Point:
        return "star8";
      case AutoShapeType.Star10Point:
        return "star10";
      case AutoShapeType.Star12Point:
        return "star12";
      case AutoShapeType.Star16Point:
        return "star16";
      case AutoShapeType.Star24Point:
        return "star24";
      case AutoShapeType.Star32Point:
        return "star32";
      case AutoShapeType.TextBox:
        return "rect";
      case AutoShapeType.MathPlus:
        return "mathPlus";
      case AutoShapeType.StraightConnector:
        return "straightConnector1";
      case AutoShapeType.BentConnector:
        return "bentConnector3";
      case AutoShapeType.CurvedConnector:
        return "curvedConnector3";
      default:
        return "rect";
    }
  }
}
