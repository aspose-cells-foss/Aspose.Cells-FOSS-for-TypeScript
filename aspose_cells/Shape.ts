import { AutoShapeType } from "./enums";
import { ShapeModel } from "./Core/ShapeModel";

// Implemented by serializer module.
/** Represents a drawing shape anchored to a worksheet. */
export class Shape {
  public constructor(private readonly model: ShapeModel) {}

  public get Model(): ShapeModel {
    return this.model;
  }

  public get Name(): string {
    return this.model.Name;
  }

  public set Name(value: string) {
    this.model.Name = value ?? "";
  }

  public get UpperLeftRow(): number {
    return this.model.UpperLeftRow;
  }

  public set UpperLeftRow(value: number) {
    this.model.UpperLeftRow = value;
  }

  public get UpperLeftColumn(): number {
    return this.model.UpperLeftColumn;
  }

  public set UpperLeftColumn(value: number) {
    this.model.UpperLeftColumn = value;
  }

  public get LowerRightRow(): number {
    return this.model.LowerRightRow;
  }

  public set LowerRightRow(value: number) {
    this.model.LowerRightRow = value;
  }

  public get LowerRightColumn(): number {
    return this.model.LowerRightColumn;
  }

  public set LowerRightColumn(value: number) {
    this.model.LowerRightColumn = value;
  }

  public get GeometryType(): string {
    return this.model.GeometryType;
  }

  public set GeometryType(value: string) {
    this.model.GeometryType = value === null || value === undefined || value.length === 0 ? "rect" : value;
  }

  public get AutoShapeType(): AutoShapeType {
    return Shape.geometryToAutoShapeType(this.model.GeometryType);
  }

  public set AutoShapeType(value: AutoShapeType) {
    this.model.GeometryType = Shape.autoShapeTypeToGeometry(value);
  }

  private static autoShapeTypeToGeometry(type: AutoShapeType): string {
    switch (type) {
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

  private static geometryToAutoShapeType(prst: string): AutoShapeType {
    if (prst === null || prst === undefined || prst.length === 0) {
      return AutoShapeType.Unknown;
    }

    switch (prst) {
      case "rect":
        return AutoShapeType.Rectangle;
      case "roundRect":
        return AutoShapeType.RoundedRectangle;
      case "ellipse":
        return AutoShapeType.Ellipse;
      case "triangle":
        return AutoShapeType.Triangle;
      case "rtTriangle":
        return AutoShapeType.RightTriangle;
      case "diamond":
        return AutoShapeType.Diamond;
      case "pentagon":
        return AutoShapeType.Pentagon;
      case "hexagon":
        return AutoShapeType.Hexagon;
      case "octagon":
        return AutoShapeType.Octagon;
      case "plus":
        return AutoShapeType.Plus;
      case "cube":
        return AutoShapeType.Cube;
      case "cyl":
        return AutoShapeType.Cylinder;
      case "heart":
        return AutoShapeType.Heart;
      case "lightningBolt":
        return AutoShapeType.Lightning;
      case "sun":
        return AutoShapeType.Sun;
      case "moon":
        return AutoShapeType.Moon;
      case "cloud":
        return AutoShapeType.Cloud;
      case "rightArrow":
        return AutoShapeType.RightArrow;
      case "leftArrow":
        return AutoShapeType.LeftArrow;
      case "upArrow":
        return AutoShapeType.UpArrow;
      case "downArrow":
        return AutoShapeType.DownArrow;
      case "leftRightArrow":
        return AutoShapeType.LeftRightArrow;
      case "upDownArrow":
        return AutoShapeType.UpDownArrow;
      case "star4":
        return AutoShapeType.Star4Point;
      case "star5":
        return AutoShapeType.Star5Point;
      case "star6":
        return AutoShapeType.Star6Point;
      case "star7":
        return AutoShapeType.Star7Point;
      case "star8":
        return AutoShapeType.Star8Point;
      case "star10":
        return AutoShapeType.Star10Point;
      case "star12":
        return AutoShapeType.Star12Point;
      case "star16":
        return AutoShapeType.Star16Point;
      case "star24":
        return AutoShapeType.Star24Point;
      case "star32":
        return AutoShapeType.Star32Point;
      case "mathPlus":
        return AutoShapeType.MathPlus;
      case "straightConnector1":
        return AutoShapeType.StraightConnector;
      case "bentConnector2":
      case "bentConnector3":
      case "bentConnector4":
      case "bentConnector5":
        return AutoShapeType.BentConnector;
      case "curvedConnector2":
      case "curvedConnector3":
      case "curvedConnector4":
      case "curvedConnector5":
        return AutoShapeType.CurvedConnector;
      default:
        return AutoShapeType.Unknown;
    }
  }
}
