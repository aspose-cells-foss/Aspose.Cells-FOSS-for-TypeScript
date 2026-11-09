import { AutoShapeType, Workbook } from "../../aspose_cells";
import { sampleOutputPath } from "../Shared";

const outputPath = sampleOutputPath("shapes-sample.xlsx");

const workbook = new Workbook();
const sheet = workbook.Worksheets[0];
sheet.Name = "Shapes";

sheet.Cells["A1"].PutValue("Shape");
sheet.Cells["B1"].PutValue("Description");

const rectIndex = sheet.Shapes.Add(0, 2, 2, 3, AutoShapeType.Rectangle);
const rectShape = sheet.Shapes[rectIndex];
rectShape.Name = "Rectangle Box";

const roundedRectIndex = sheet.Shapes.Add(0, 4, 2, 5, AutoShapeType.RoundedRectangle);
const roundedRectShape = sheet.Shapes[roundedRectIndex];
roundedRectShape.Name = "Rounded Rectangle";

const ellipseIndex = sheet.Shapes.Add(0, 6, 2, 7, AutoShapeType.Ellipse);
const ellipseShape = sheet.Shapes[ellipseIndex];
ellipseShape.Name = "Oval Shape";

const triangleIndex = sheet.Shapes.Add(3, 2, 5, 3, AutoShapeType.Triangle);
const triangleShape = sheet.Shapes[triangleIndex];
triangleShape.Name = "Triangle";

const rightTriangleIndex = sheet.Shapes.Add(3, 4, 5, 5, AutoShapeType.RightTriangle);
const rightTriangleShape = sheet.Shapes[rightTriangleIndex];
rightTriangleShape.Name = "Right Triangle";

const diamondIndex = sheet.Shapes.Add(3, 6, 5, 7, AutoShapeType.Diamond);
const diamondShape = sheet.Shapes[diamondIndex];
diamondShape.Name = "Diamond";

const pentagonIndex = sheet.Shapes.Add(6, 2, 8, 3, AutoShapeType.Pentagon);
const pentagonShape = sheet.Shapes[pentagonIndex];
pentagonShape.Name = "Pentagon";

const hexagonIndex = sheet.Shapes.Add(6, 4, 8, 5, AutoShapeType.Hexagon);
const hexagonShape = sheet.Shapes[hexagonIndex];
hexagonShape.Name = "Hexagon";

const octagonIndex = sheet.Shapes.Add(6, 6, 8, 7, AutoShapeType.Octagon);
const octagonShape = sheet.Shapes[octagonIndex];
octagonShape.Name = "Octagon";

const star5Index = sheet.Shapes.Add(9, 2, 11, 3, AutoShapeType.Star5Point);
const star5Shape = sheet.Shapes[star5Index];
star5Shape.Name = "5-Point Star";

const star8Index = sheet.Shapes.Add(9, 4, 11, 5, AutoShapeType.Star8Point);
const star8Shape = sheet.Shapes[star8Index];
star8Shape.Name = "8-Point Star";

const star12Index = sheet.Shapes.Add(9, 6, 11, 7, AutoShapeType.Star12Point);
const star12Shape = sheet.Shapes[star12Index];
star12Shape.Name = "12-Point Star";

const rightArrowIndex = sheet.Shapes.Add(12, 2, 14, 3, AutoShapeType.RightArrow);
const rightArrowShape = sheet.Shapes[rightArrowIndex];
rightArrowShape.Name = "Right Arrow";

const leftArrowIndex = sheet.Shapes.Add(12, 4, 14, 5, AutoShapeType.LeftArrow);
const leftArrowShape = sheet.Shapes[leftArrowIndex];
leftArrowShape.Name = "Left Arrow";

const upDownArrowIndex = sheet.Shapes.Add(12, 6, 14, 7, AutoShapeType.UpDownArrow);
const upDownArrowShape = sheet.Shapes[upDownArrowIndex];
upDownArrowShape.Name = "Up-Down Arrow";

const heartIndex = sheet.Shapes.Add(15, 2, 17, 3, AutoShapeType.Heart);
const heartShape = sheet.Shapes[heartIndex];
heartShape.Name = "Heart";

const lightningIndex = sheet.Shapes.Add(15, 4, 17, 5, AutoShapeType.Lightning);
const lightningShape = sheet.Shapes[lightningIndex];
lightningShape.Name = "Lightning Bolt";

const cloudIndex = sheet.Shapes.Add(15, 6, 17, 7, AutoShapeType.Cloud);
const cloudShape = sheet.Shapes[cloudIndex];
cloudShape.Name = "Cloud";

const sunIndex = sheet.Shapes.Add(18, 2, 20, 3, AutoShapeType.Sun);
const sunShape = sheet.Shapes[sunIndex];
sunShape.Name = "Sun";

const moonIndex = sheet.Shapes.Add(18, 4, 20, 5, AutoShapeType.Moon);
const moonShape = sheet.Shapes[moonIndex];
moonShape.Name = "Moon";

const plusIndex = sheet.Shapes.Add(18, 6, 20, 7, AutoShapeType.Plus);
const plusShape = sheet.Shapes[plusIndex];
plusShape.Name = "Plus Sign";

const cubeIndex = sheet.Shapes.Add(21, 2, 23, 3, AutoShapeType.Cube);
const cubeShape = sheet.Shapes[cubeIndex];
cubeShape.Name = "Cube";

const cylinderIndex = sheet.Shapes.Add(21, 4, 23, 5, AutoShapeType.Cylinder);
const cylinderShape = sheet.Shapes[cylinderIndex];
cylinderShape.Name = "Cylinder";

const mathPlusIndex = sheet.Shapes.Add(21, 6, 23, 7, AutoShapeType.MathPlus);
const mathPlusShape = sheet.Shapes[mathPlusIndex];
mathPlusShape.Name = "Math Plus";

workbook.Save(outputPath);

const loaded = new Workbook(outputPath);
const loadedSheet = loaded.Worksheets["Shapes"];

console.log("Saved: " + outputPath);
console.log("Shape count: " + loadedSheet.Shapes.Count);
console.log();
console.log("Basic Shapes:");
console.log("  " + loadedSheet.Shapes[0].Name + " (" + AutoShapeType[loadedSheet.Shapes[0].AutoShapeType] + ")");
console.log("  " + loadedSheet.Shapes[1].Name + " (" + AutoShapeType[loadedSheet.Shapes[1].AutoShapeType] + ")");
console.log("  " + loadedSheet.Shapes[2].Name + " (" + AutoShapeType[loadedSheet.Shapes[2].AutoShapeType] + ")");
console.log();
console.log("Triangles & Polygons:");
console.log("  " + loadedSheet.Shapes[3].Name + " (" + AutoShapeType[loadedSheet.Shapes[3].AutoShapeType] + ")");
console.log("  " + loadedSheet.Shapes[4].Name + " (" + AutoShapeType[loadedSheet.Shapes[4].AutoShapeType] + ")");
console.log("  " + loadedSheet.Shapes[5].Name + " (" + AutoShapeType[loadedSheet.Shapes[5].AutoShapeType] + ")");
console.log("  " + loadedSheet.Shapes[6].Name + " (" + AutoShapeType[loadedSheet.Shapes[6].AutoShapeType] + ")");
console.log("  " + loadedSheet.Shapes[7].Name + " (" + AutoShapeType[loadedSheet.Shapes[7].AutoShapeType] + ")");
console.log("  " + loadedSheet.Shapes[8].Name + " (" + AutoShapeType[loadedSheet.Shapes[8].AutoShapeType] + ")");
console.log();
console.log("Stars:");
console.log("  " + loadedSheet.Shapes[9].Name + " (" + AutoShapeType[loadedSheet.Shapes[9].AutoShapeType] + ")");
console.log("  " + loadedSheet.Shapes[10].Name + " (" + AutoShapeType[loadedSheet.Shapes[10].AutoShapeType] + ")");
console.log("  " + loadedSheet.Shapes[11].Name + " (" + AutoShapeType[loadedSheet.Shapes[11].AutoShapeType] + ")");
console.log();
console.log("Arrows:");
console.log("  " + loadedSheet.Shapes[12].Name + " (" + AutoShapeType[loadedSheet.Shapes[12].AutoShapeType] + ")");
console.log("  " + loadedSheet.Shapes[13].Name + " (" + AutoShapeType[loadedSheet.Shapes[13].AutoShapeType] + ")");
console.log("  " + loadedSheet.Shapes[14].Name + " (" + AutoShapeType[loadedSheet.Shapes[14].AutoShapeType] + ")");
console.log();
console.log("Symbols & Icons:");
console.log("  " + loadedSheet.Shapes[15].Name + " (" + AutoShapeType[loadedSheet.Shapes[15].AutoShapeType] + ")");
console.log("  " + loadedSheet.Shapes[16].Name + " (" + AutoShapeType[loadedSheet.Shapes[16].AutoShapeType] + ")");
console.log("  " + loadedSheet.Shapes[17].Name + " (" + AutoShapeType[loadedSheet.Shapes[17].AutoShapeType] + ")");
console.log("  " + loadedSheet.Shapes[18].Name + " (" + AutoShapeType[loadedSheet.Shapes[18].AutoShapeType] + ")");
console.log("  " + loadedSheet.Shapes[19].Name + " (" + AutoShapeType[loadedSheet.Shapes[19].AutoShapeType] + ")");
console.log("  " + loadedSheet.Shapes[20].Name + " (" + AutoShapeType[loadedSheet.Shapes[20].AutoShapeType] + ")");
console.log();
console.log("3D & Math:");
console.log("  " + loadedSheet.Shapes[21].Name + " (" + AutoShapeType[loadedSheet.Shapes[21].AutoShapeType] + ")");
console.log("  " + loadedSheet.Shapes[22].Name + " (" + AutoShapeType[loadedSheet.Shapes[22].AutoShapeType] + ")");
console.log("  " + loadedSheet.Shapes[23].Name + " (" + AutoShapeType[loadedSheet.Shapes[23].AutoShapeType] + ")");
