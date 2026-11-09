import { AutoShapeType, Workbook } from "../../aspose_cells";
import { AssertEx } from "./TestInfrastructure";

/** Scenario factory for drawing shape workbooks. */
export class ShapeScenarioFactory {
  public static CreateShapeWorkbook(): Workbook {
    const workbook = new Workbook();
    const sheet = workbook.Worksheets[0];
    sheet.Name = "Shapes";

    sheet.Cells["A1"].PutValue("Drawing Objects Demo");

    // Column A-D: left column of shapes
    const rectIndex = sheet.Shapes.Add(2, 0, 6, 4, AutoShapeType.Rectangle);
    sheet.Shapes.get(rectIndex).Name = "MyRectangle";

    const arrowIndex = sheet.Shapes.Add(8, 0, 11, 4, AutoShapeType.RightArrow);
    sheet.Shapes.get(arrowIndex).Name = "MyRightArrow";

    const ellipseIndex = sheet.Shapes.Add(13, 0, 17, 4, AutoShapeType.Ellipse);
    sheet.Shapes.get(ellipseIndex).Name = "MyEllipse";

    // Column F-J: right column of shapes
    const starIndex = sheet.Shapes.Add(2, 5, 7, 9, AutoShapeType.Star5Point);
    sheet.Shapes.get(starIndex).Name = "MyStar5";

    const downArrowIndex = sheet.Shapes.Add(9, 5, 12, 9, AutoShapeType.DownArrow);
    sheet.Shapes.get(downArrowIndex).Name = "MyDownArrow";

    const star12Index = sheet.Shapes.Add(14, 5, 19, 9, AutoShapeType.Star12Point);
    sheet.Shapes.get(star12Index).Name = "MyStar12";

    return workbook;
  }

  public static AssertShapeWorkbook(workbook: Workbook): void {
    const sheet = workbook.Worksheets[0];
    AssertEx.equal(6, sheet.Shapes.Count);

    const rect = sheet.Shapes.get(0);
    AssertEx.equal("MyRectangle", rect.Name);
    AssertEx.equal(2, rect.UpperLeftRow);
    AssertEx.equal(0, rect.UpperLeftColumn);
    AssertEx.equal(6, rect.LowerRightRow);
    AssertEx.equal(4, rect.LowerRightColumn);
    AssertEx.equal(AutoShapeType.Rectangle, rect.AutoShapeType);

    const arrow = sheet.Shapes.get(1);
    AssertEx.equal("MyRightArrow", arrow.Name);
    AssertEx.equal(AutoShapeType.RightArrow, arrow.AutoShapeType);

    const ellipse = sheet.Shapes.get(2);
    AssertEx.equal("MyEllipse", ellipse.Name);
    AssertEx.equal(AutoShapeType.Ellipse, ellipse.AutoShapeType);

    const star = sheet.Shapes.get(3);
    AssertEx.equal("MyStar5", star.Name);
    AssertEx.equal(AutoShapeType.Star5Point, star.AutoShapeType);

    const downArrow = sheet.Shapes.get(4);
    AssertEx.equal("MyDownArrow", downArrow.Name);
    AssertEx.equal(AutoShapeType.DownArrow, downArrow.AutoShapeType);

    const star12 = sheet.Shapes.get(5);
    AssertEx.equal("MyStar12", star12.Name);
    AssertEx.equal(AutoShapeType.Star12Point, star12.AutoShapeType);
  }
}
