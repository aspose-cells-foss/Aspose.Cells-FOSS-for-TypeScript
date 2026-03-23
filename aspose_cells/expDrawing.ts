import { ShapeInfo, ShapeFill, AllConnectorShapeInfo, ImageInfo, ChartInfo } from "./types";
import { escapeXml } from "./util";

export class ExpDrawing {
  static generateDrawingXml(shapes: ShapeInfo[], sheetImages: ImageInfo[] = [], imageMap: Map<string, number> = new Map()): string {
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">`;

    const sortedShapes = [...shapes].sort((a, b) => {
      const aZ = (a as any).zIndex || 0;
      const bZ = (b as any).zIndex || 0;
      if (aZ !== bZ) return aZ - bZ;
      const aCol = (a as any).fromCol ?? 0;
      const bCol = (b as any).fromCol ?? 0;
      return aCol - bCol;
    });

    let rid = 1;
    let chartIndex = 1;
    for (let i = 0; i < sortedShapes.length; i++) {
      const shape = sortedShapes[i];
      
      if ((shape as any).type === "chart" || (shape as any).chartType) {
        xml += `<xdr:twoCellAnchor>`;
        xml += ExpDrawing.chartToDrawingXml(shape as any, rid++, chartIndex++);
        xml += `<xdr:clientData/></xdr:twoCellAnchor>`;
      } else if ((shape as any).type === "picture") {
        xml += `<xdr:twoCellAnchor>`;
        xml += ExpDrawing.pictureToDrawingXml(shape as any, rid++);
        xml += `<xdr:clientData/></xdr:twoCellAnchor>`;
      } else {
        xml += `<xdr:twoCellAnchor>`;
        xml += ExpDrawing.shapeToDrawingXml(shape, i);
        xml += `</xdr:twoCellAnchor>`;
      }
    }

    xml += "</xdr:wsDr>";
    return xml;
  }

  private static chartToDrawingXml(shape: any, rid: number, chartIndex: number): string {
    const fromCol = shape.fromCol ?? 0;
    const fromColOff = shape.fromColOff ?? 0;
    const fromRow = shape.fromRow ?? 0;
    const fromRowOff = shape.fromRowOff ?? 0;
    const toCol = shape.toCol ?? 0;
    const toColOff = shape.toColOff ?? 0;
    const toRow = shape.toRow ?? 0;
    const toRowOff = shape.toRowOff ?? 0;

    const x = shape.originalX ?? shape.x ?? 0;
    const y = shape.originalY ?? shape.y ?? 0;
    const cx = shape.width ?? 0;
    const cy = shape.height ?? 0;

    const uniqueId = ExpDrawing.generateUniqueId();
    const extLst = `<a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="${uniqueId}"/></a:ext></a:extLst>`;

    return `<xdr:from><xdr:col>${fromCol}</xdr:col><xdr:colOff>${fromColOff}</xdr:colOff><xdr:row>${fromRow}</xdr:row><xdr:rowOff>${fromRowOff}</xdr:rowOff></xdr:from><xdr:to><xdr:col>${toCol}</xdr:col><xdr:colOff>${toColOff}</xdr:colOff><xdr:row>${toRow}</xdr:row><xdr:rowOff>${toRowOff}</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${rid + 2}" name="${escapeXml(shape.name || "Chart")}">${extLst}</xdr:cNvPr><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame>`;
  }

  private static pictureToDrawingXml(shape: any, sortedIdx: number): string {
    const fromCol = shape.fromCol ?? 0;
    const fromColOff = shape.fromColOff ?? 0;
    const fromRow = shape.fromRow ?? 0;
    const fromRowOff = shape.fromRowOff ?? 0;
    const toCol = shape.toCol ?? 0;
    const toColOff = shape.toColOff ?? 0;
    const toRow = shape.toRow ?? 0;
    const toRowOff = shape.toRowOff ?? 0;

    const cNvId = sortedIdx + 3;
    const embedRid = sortedIdx + 1;

    const cx = shape.width ?? 64 * 9144;
    const cy = shape.height ?? 64 * 9144;

    const uniqueId = ExpDrawing.generateUniqueId();
    const extLst = `<a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="${uniqueId}"/></a:ext></a:extLst>`;

    return `<xdr:from><xdr:col>${fromCol}</xdr:col><xdr:colOff>${fromColOff}</xdr:colOff><xdr:row>${fromRow}</xdr:row><xdr:rowOff>${fromRowOff}</xdr:rowOff></xdr:from><xdr:to><xdr:col>${toCol}</xdr:col><xdr:colOff>${toColOff}</xdr:colOff><xdr:row>${toRow}</xdr:row><xdr:rowOff>${toRowOff}</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${cNvId}" name="${escapeXml(shape.name || "Picture")}">${extLst}</xdr:cNvPr><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId${embedRid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="${shape.x ?? 0}" y="${shape.y ?? 0}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"/></xdr:spPr></xdr:pic>`;
  }

  static generateUniqueId(): string {
    return "{" + ExpDrawing.generateGuid() + "}";
  }

  static generateGuid(): string {
    return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
      const r = (Math.random() * 16) | 0;
      const v = c === "x" ? r : (r & 0x3) | 0x8;
      return v.toString(16).toUpperCase();
    });
  }

  static shapeToDrawingXml(shape: ShapeInfo, index: number): string {
    const fromCol = shape.fromCol;
    const fromColOff = shape.fromColOff ?? 0;
    const fromRow = shape.fromRow;
    const fromRowOff = shape.fromRowOff ?? 0;
    const toCol = shape.toCol;
    const toColOff = shape.toColOff ?? 0;
    const toRow = shape.toRow;
    const toRowOff = shape.toRowOff ?? 0;

    const type = shape.type;
    const typeLower = type.toLowerCase();
    const isConnector = typeLower === "line" || typeLower.includes("connector");

    const shapeType = ExpDrawing.getDrawingShapeType(type);
    const fillXml = ExpDrawing.shapeFillToXml(shape.fill, "#ffffff");
    const lineWidth = shape.lineWidth || 12700;

    const x = shape.x ?? 0;
    const y = shape.y ?? 0;
    const cx = shape.width ?? 0;
    const cy = shape.height ?? 0;

    const flipV = typeLower === "line" ? ' flipV="1"' : "";
    const uniqueId = ExpDrawing.generateUniqueId();
    const extLst = `<a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="${uniqueId}"/></a:ext></a:extLst>`;

    let anchor = `<xdr:from><xdr:col>${fromCol}</xdr:col><xdr:colOff>${fromColOff}</xdr:colOff><xdr:row>${fromRow}</xdr:row><xdr:rowOff>${fromRowOff}</xdr:rowOff></xdr:from><xdr:to><xdr:col>${toCol}</xdr:col><xdr:colOff>${toColOff}</xdr:colOff><xdr:row>${toRow}</xdr:row><xdr:rowOff>${toRowOff}</xdr:rowOff></xdr:to>`;

    const lnStyle =
      lineWidth && lineWidth !== 12700 ? `<a:ln w="${lineWidth}"/>` : "";

    const tailEnd = ExpDrawing.isConnectorShape(shape) && shape.hasArrowEnd
      ? '<a:ln><a:tailEnd type="triangle"/></a:ln>'
      : "";

    if (isConnector) {
      anchor += `<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="${index + 3}" name="${escapeXml(shape.name)}">${extLst}</xdr:cNvPr><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr><a:xfrm${flipV}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${shapeType}"><a:avLst/></a:prstGeom>${lnStyle}${tailEnd}</xdr:spPr><xdr:style>`;
      anchor += `<a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></xdr:style>`;
      anchor += "</xdr:cxnSp>";
    } else {
      anchor += `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${index + 3}" name="${escapeXml(shape.name)}">${extLst}</xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${shapeType}"><a:avLst/></a:prstGeom>${fillXml}</xdr:spPr><xdr:style>`;
      anchor += `<a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>`;
      anchor += `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="t"/><a:lstStyle/><a:p><a:pPr algn="l"/><a:endParaRPr lang="en-US" sz="1100"/></a:p></xdr:txBody>`;
      anchor += "</xdr:sp>";
    }

    anchor += "<xdr:clientData/>";

    return anchor;
  }

  static getDrawingShapeType(type: string): string {
    const typeMap: Record<string, string> = {
      rect: "rect",
      rectangle: "rect",
      ellipse: "ellipse",
      oval: "ellipse",
      line: "line",
      straightConnector1: "straightConnector1",
      straightConnector2: "straightConnector2",
      straightConnector3: "straightConnector3",
      straightConnector: "straightConnector1",
      bentConnector1: "bentConnector1",
      bentConnector2: "bentConnector2",
      bentConnector3: "bentConnector3",
      bentConnector4: "bentConnector4",
      bentConnector5: "bentConnector5",
      elbowConnector: "bentConnector3",
      curvedConnector1: "curvedConnector1",
      curvedConnector2: "curvedConnector2",
      curvedConnector3: "curvedConnector3",
      curvedConnector4: "curvedConnector4",
      curvedConnector5: "curvedConnector5",
      triangle: "triangle",
      rightTriangle: "rightTriangle",
      parallelogram: "parallelogram",
      trapezoid: "trapezoid",
      diamond: "diamond",
      snip1Rect: "snip1Rect",
      singleCornerSnipped: "snip1Rect",
      snip2Same: "snip2Same",
      snip2Diag: "snip2Diag",
      snip4Same: "snip4Same",
      snip4Diag: "snip4Diag",
      round1Rect: "round1Rect",
      round2Same: "round2Same",
      round2Diag: "round2Diag",
      round4Same: "round4Same",
      round4Diag: "round4Diag",
      pie: "pie",
      pieWedge: "pieWedge",
      quarterCircle: "quarterCircle",
      partialCircle: "quarterCircle",
      halfCircle: "halfCircle",
      threeQuarterCircle: "threeQuarterCircle",
      brace: "brace",
      bracket: "bracket",
      arrow: "arrow",
      rightArrow: "rightArrow",
      leftArrow: "leftArrow",
      upArrow: "upArrow",
      downArrow: "downArrow",
      leftRightArrow: "leftRightArrow",
      upDownArrow: "upDownArrow",
      quadrilateralArrow: "quadrilateralArrow",
      pentagon: "pentagon",
      hexagon: "hexagon",
      heptagon: "heptagon",
      octagon: "octagon",
      decagon: "decagon",
      dodecagon: "dodecagon",
      star4: "star4",
      star5: "star5",
      star6: "star6",
      star7: "star7",
      star8: "star8",
      star10: "star10",
      star12: "star12",
      star16: "star16",
      star24: "star24",
      star32: "star32",
      plus: "plus",
      mathPlus: "mathPlus",
      mathMinus: "mathMinus",
      mathMultiply: "mathMultiply",
      mathDivide: "mathDivide",
      mathEqual: "mathEqual",
      mathNotEqual: "mathNotEqual",
      ribbon: "ribbon",
      ribbon2: "ribbon2",
      chevron: "chevron",
      pentagonArrow: "pentagonArrow",
      chord: "chord",
      cloud: "cloud",
      cloudCallout: "cloudCallout",
      sun: "sun",
      moon: "moon",
      doubleWave: "doubleWave",
      irregularSeal1: "irregularSeal1",
      irregularSeal2: "irregularSeal2",
      irregularSeal3: "irregularSeal3",
      irregularSeal4: "irregularSeal4",
      irregularSeal5: "irregularSeal5",
      irregularSeal6: "irregularSeal6",
      irregularSeal7: "irregularSeal7",
      irregularSeal8: "irregularSeal8",
      irregularSeal9: "irregularSeal9",
      irregularSeal10: "irregularSeal10",
      explosion: "explosion",
      lightningBolt: "lightningBolt",
      heart: "heart",
      pictureFrame: "pictureFrame",
      tetris: "tetris",
      cube: "cube",
      cylinder: "cylinder",
      cylinderSolid: "cylinderSolid",
      prism: "prism",
      prismRight: "prismRight",
      flowChartProcess: "flowChartProcess",
      nonIsoscelesTrapezoid: "nonIsoscelesTrapezoid",
      starBurst: "starBurst",
      smileyFace: "smileyFace",
      Donut: "donut",
    };
    return typeMap[type] || "rect";
  }

  static colorToScheme(color: string): string {
    const colorMap: Record<string, string> = {
      "#4472c4": '<a:schemeClr val="accent1"/>',
      "#ed7d31": '<a:schemeClr val="accent2"/>',
      "#a5a5a5": '<a:schemeClr val="accent3"/>',
      "#ffc000": '<a:schemeClr val="accent4"/>',
      "#5b9bd5": '<a:schemeClr val="accent5"/>',
      "#70ad47": '<a:schemeClr val="accent6"/>',
    };
    if (colorMap[color.toLowerCase()]) {
      return colorMap[color.toLowerCase()];
    }
    return `<a:srgbClr val="${color.replace("#", "")}"/>`;
  }

  static schemeToColor(schemeColor: string): string {
    const colorMap: Record<string, string> = {
      accent1: "4472c4",
      accent2: "ed7d31",
      accent3: "a5a5a5",
      accent4: "ffc000",
      accent5: "5b9bd5",
      accent6: "70ad47",
    };
    return colorMap[schemeColor] || schemeColor;
  }

  static shapeFillToXml(fill?: ShapeFill, fallbackColor?: string): string {
    if (!fill) {
      const color = ExpDrawing.colorToScheme(fallbackColor || "#ffffff");
      return `<a:solidFill>${color}</a:solidFill>`;
    }

    if (fill.type === "none") {
      return "";
    }

    if (fill.type === "solid" && fill.color) {
      let colorXml: string;
      if (fill.isSchemeColor) {
        colorXml = `<a:schemeClr val="${fill.color}"/>`;
      } else {
        colorXml = ExpDrawing.colorToScheme(fill.color);
      }
      return `<a:solidFill>${colorXml}</a:solidFill>`;
    }

    if (
      fill.type === "gradient" &&
      fill.gradientStops &&
      fill.gradientStops.length > 0
    ) {
      let gsLst = "";
      for (const stop of fill.gradientStops) {
        let colorXml: string;
        if (stop.isSchemeColor) {
          colorXml = `<a:schemeClr val="${stop.color}">`;
          if (stop.lumMod !== undefined) {
            colorXml += `<a:lumMod val="${stop.lumMod}"/>`;
          }
          if (stop.lumOff !== undefined) {
            colorXml += `<a:lumOff val="${stop.lumOff}"/>`;
          }
          colorXml += `</a:schemeClr>`;
        } else {
          colorXml = `<a:srgbClr val="${stop.color.replace("#", "")}"/>`;
        }
        gsLst += `<a:gs pos="${stop.position}">${colorXml}</a:gs>`;
      }
      const angle = fill.gradientAngle || 0;
      return `<a:gradFill><a:gsLst>${gsLst}</a:gsLst><a:lin ang="${angle}" scaled="1"/></a:gradFill>`;
    }

    const color = ExpDrawing.colorToScheme(fallbackColor || "#ffffff");
    return `<a:solidFill>${color}</a:solidFill>`;
  }

  static isConnectorShape(shape: ShapeInfo): shape is AllConnectorShapeInfo {
    const type = shape.type.toLowerCase();
    return type === "line" || type.includes("connector");
  }
}
