import { ShapeInfo, ShapeFill, StraightConnectorShapeInfo, BentConnectorShapeInfo, CurvedConnectorShapeInfo, LineShapeInfo } from "./types";
import { readZipEntry as readZipEntryUtil } from "./util";
import { DOMParser } from "@xmldom/xmldom";

export class ImpDrawing {
  static async loadWorksheetDrawing(
    zip: any,
    shapes: ShapeInfo[],
    sheetPath: string,
  ) {
    const sheetRelsPath =
      "xl/" + sheetPath.replace("worksheets/", "worksheets/_rels/") + ".rels";
    const relsContent = await readZipEntryUtil(zip, sheetRelsPath);
    if (!relsContent) return;

    const relParser = new DOMParser();
    const relDoc = relParser.parseFromString(relsContent, "text/xml");
    const rels = relDoc.getElementsByTagName("Relationship");

    let drawingTarget = "";
    for (let i = 0; i < rels.length; i++) {
      const rel = rels[i];
      const type = rel.getAttribute("Type") || "";
      if (type.includes("drawing")) {
        drawingTarget = rel.getAttribute("Target") || "";
        break;
      }
    }

    if (!drawingTarget) return;

    let drawingPath: string;
    if (drawingTarget.startsWith("../")) {
      drawingPath = "xl/" + drawingTarget.substring(3);
    } else if (sheetPath.includes("/")) {
      drawingPath = "xl/" + sheetPath.replace(/[^/]+$/, "") + drawingTarget;
    } else {
      drawingPath = "xl/drawings/" + drawingTarget;
    }

    const drawingContent = await readZipEntryUtil(zip, drawingPath);
    if (!drawingContent) return;

    const parser = new DOMParser();
    const doc = parser.parseFromString(drawingContent, "text/xml");

    const twoCellAnchors = doc.getElementsByTagName("xdr:twoCellAnchor");
    for (let i = 0; i < twoCellAnchors.length; i++) {
      const anchor = twoCellAnchors[i];

      const from = anchor.getElementsByTagName("xdr:from")[0];
      const to = anchor.getElementsByTagName("xdr:to")[0];
      if (!from || !to) continue;

      const fromCol = parseInt(
        ImpDrawing.getXmlValue(from, "xdr:col") || "0",
        10,
      );
      const fromColOff = parseInt(
        ImpDrawing.getXmlValue(from, "xdr:colOff") || "0",
        10,
      );
      const fromRow = parseInt(
        ImpDrawing.getXmlValue(from, "xdr:row") || "0",
        10,
      );
      const fromRowOff = parseInt(
        ImpDrawing.getXmlValue(from, "xdr:rowOff") || "0",
        10,
      );
      const toCol = parseInt(ImpDrawing.getXmlValue(to, "xdr:col") || "0", 10);
      const toColOff = parseInt(
        ImpDrawing.getXmlValue(to, "xdr:colOff") || "0",
        10,
      );
      const toRow = parseInt(ImpDrawing.getXmlValue(to, "xdr:row") || "0", 10);
      const toRowOff = parseInt(
        ImpDrawing.getXmlValue(to, "xdr:rowOff") || "0",
        10,
      );

      let shapeType = "rect";
      let lineColor = "#000000";
      let lineWidth = 12700;
      let hasArrowStart = false;
      let hasArrowEnd = false;
      let flipV = false;
      let flipH = false;
      let rotation = 0;
      let cx: number | undefined;
      let cy: number | undefined;

      const sp = anchor.getElementsByTagName("xdr:sp")[0];
      const cxnSp = anchor.getElementsByTagName("xdr:cxnSp")[0];
      const graphicFrame = anchor.getElementsByTagName("xdr:graphicFrame")[0];

      if (graphicFrame) {
        continue;
      }

      let fill: ShapeFill | undefined;

      if (sp || cxnSp) {
        const spPr = (sp || cxnSp)?.getElementsByTagName("xdr:spPr")[0];

        if (spPr) {
          const xfrm = spPr.getElementsByTagName("a:xfrm")[0];
          if (xfrm) {
            const flipVAttr = xfrm.getAttribute("flipV");
            flipV = flipVAttr === "1";
            const flipHAttr = xfrm.getAttribute("flipH");
            flipH = flipHAttr === "1";
            const rotAttr = xfrm.getAttribute("rot");
            if (rotAttr) {
              rotation = parseInt(rotAttr, 10) / 60000;
              if (rotation > 180) {
                rotation = rotation - 360;
              }
            }
            const ext = xfrm.getElementsByTagName("a:ext")[0];
            if (ext) {
              const cxAttr = ext.getAttribute("cx");
              const cyAttr = ext.getAttribute("cy");
              if (cxAttr) cx = parseInt(cxAttr, 10);
              if (cyAttr) cy = parseInt(cyAttr, 10);
            }
          }

          const prstGeom = spPr.getElementsByTagName("a:prstGeom")[0];
          if (prstGeom) {
            shapeType = prstGeom.getAttribute("prst") || "rect";
          }

          const solidFill = spPr.getElementsByTagName("a:solidFill")[0];
          const gradFill = spPr.getElementsByTagName("a:gradFill")[0];

          if (gradFill) {
            fill = ImpDrawing.parseGradFill(gradFill);
          } else if (solidFill) {
            fill = ImpDrawing.parseSolidFill(solidFill);
          } else {
            const style = (sp || cxnSp)?.getElementsByTagName("xdr:style")[0];
            if (style) {
              const fillRef = style.getElementsByTagName("a:fillRef")[0];
              if (fillRef) {
                const schemeClr =
                  fillRef.getElementsByTagName("a:schemeClr")[0];
                if (schemeClr) {
                  fill = {
                    type: "solid",
                    color: schemeClr.getAttribute("val") || "accent1",
                    isSchemeColor: true,
                  };
                } else {
                  const srgbClr = fillRef.getElementsByTagName("a:srgbClr")[0];
                  if (srgbClr) {
                    fill = {
                      type: "solid",
                      color: "#" + (srgbClr.getAttribute("val") || "ffffff"),
                      isSchemeColor: false,
                    };
                  } else {
                    fill = { type: "none" };
                  }
                }
              } else {
                fill = { type: "none" };
              }
            } else {
              fill = { type: "none" };
            }
          }

          const ln = spPr.getElementsByTagName("a:ln")[0];
          if (ln) {
            lineWidth = parseInt(ln.getAttribute("w") || "12700", 10);
            const strokeClr = ln.getElementsByTagName("a:srgbClr")[0];
            const strokeScheme = ln.getElementsByTagName("a:schemeClr")[0];
            if (strokeClr) {
              lineColor = "#" + (strokeClr.getAttribute("val") || "000000");
            } else if (strokeScheme) {
              lineColor = ImpDrawing.parseSchemeColor(strokeScheme);
            }
          }
        }

        const headEnd = (sp || cxnSp)?.getElementsByTagName("a:headEnd")[0];
        if (headEnd) {
          hasArrowStart = true;
        }

        const tailEnd = (sp || cxnSp)?.getElementsByTagName("a:tailEnd")[0];
        if (tailEnd) {
          hasArrowEnd = true;
        }
      }

      const nameEl = anchor.getElementsByTagName("xdr:cNvPr")[0];
      const name = nameEl?.getAttribute("name") || `Shape ${i + 1}`;

      const baseShape = {
        name,
        type: shapeType,
        fromCol,
        fromColOff,
        fromRow,
        fromRowOff,
        toCol,
        toColOff,
        toRow,
        toRowOff,
        fill,
        cx,
        cy,
        lineColor,
        lineWidth,
        flipV,
        flipH,
        rotation,
      };

      const isConnector = cxnSp !== undefined || shapeType === "line" || shapeType.includes("Connector");
      
      if (isConnector) {
        const connectorShape = {
          ...baseShape,
          hasArrowStart,
          hasArrowEnd,
        };
        
        const typeLower = shapeType.toLowerCase();
        if (typeLower === "line") {
          shapes.push(connectorShape as LineShapeInfo);
        } else if (typeLower.includes("bent") || typeLower.includes("elbow")) {
          shapes.push(connectorShape as BentConnectorShapeInfo);
        } else if (typeLower.includes("curved")) {
          shapes.push(connectorShape as CurvedConnectorShapeInfo);
        } else {
          shapes.push(connectorShape as StraightConnectorShapeInfo);
        }
      } else {
        shapes.push(baseShape as ShapeInfo);
      }
    }
  }

  private static getXmlValue(parent: Element, tagName: string): string | null {
    const el = parent.getElementsByTagName(tagName)[0];
    return el?.textContent || null;
  }

  private static parseSchemeColor(el: Element): string {
    const lumMod = el.getElementsByTagName("a:lumMod")[0];
    const lumOff = el.getElementsByTagName("a:lumOff")[0];
    const shade = el.getElementsByTagName("a:shade")[0];

    const val = el.getAttribute("val") || "accent1";
    const colorMap: Record<string, string> = {
      accent1: "4472c4",
      accent2: "ed7d31",
      accent3: "a5a5a5",
      accent4: "ffc000",
      accent5: "5b9bd5",
      accent6: "70ad47",
    };

    let color = colorMap[val] || "4472c4";

    if (shade) {
      const shadeVal = parseInt(shade.getAttribute("val") || "0", 10) / 100000;
      const r = parseInt(color.slice(0, 2), 16);
      const g = parseInt(color.slice(2, 4), 16);
      const b = parseInt(color.slice(4, 6), 16);
      color =
        "#" +
        Math.round(r * (1 - shadeVal))
          .toString(16)
          .padStart(2, "0") +
        Math.round(g * (1 - shadeVal))
          .toString(16)
          .padStart(2, "0") +
        Math.round(b * (1 - shadeVal))
          .toString(16)
          .padStart(2, "0");
    }

    return "#" + color;
  }

  private static parseSolidFill(solidFill: Element): ShapeFill {
    const schemeClr = solidFill.getElementsByTagName("a:schemeClr")[0];
    const srgbClr = solidFill.getElementsByTagName("a:srgbClr")[0];

    let color: string = "#ffffff";
    let isSchemeColor = false;
    if (schemeClr) {
      color = schemeClr.getAttribute("val") || "accent1";
      isSchemeColor = true;
    } else if (srgbClr) {
      color = "#" + (srgbClr.getAttribute("val") || "ffffff");
    }

    return {
      type: "solid",
      color,
      isSchemeColor,
    };
  }

  private static parseGradFill(gradFill: Element): ShapeFill {
    const gsLst = gradFill.getElementsByTagName("a:gsLst")[0];
    const lin = gradFill.getElementsByTagName("a:lin")[0];

    const gradientStops: Array<{
      position: number;
      color: string;
      isSchemeColor?: boolean;
      lumMod?: number;
      lumOff?: number;
    }> = [];

    if (gsLst) {
      const gsElements = gsLst.getElementsByTagName("a:gs");
      for (let i = 0; i < gsElements.length; i++) {
        const gs = gsElements[i];
        const pos = parseInt(gs.getAttribute("pos") || "0", 10);
        const schemeClr = gs.getElementsByTagName("a:schemeClr")[0];
        const srgbClr = gs.getElementsByTagName("a:srgbClr")[0];

        let color: string;
        let isSchemeColor = false;
        let lumMod: number | undefined;
        let lumOff: number | undefined;

        if (schemeClr) {
          color = schemeClr.getAttribute("val") || "accent1";
          isSchemeColor = true;
          const lumModEl = schemeClr.getElementsByTagName("a:lumMod")[0];
          const lumOffEl = schemeClr.getElementsByTagName("a:lumOff")[0];
          if (lumModEl) {
            lumMod = parseInt(lumModEl.getAttribute("val") || "0", 10);
          }
          if (lumOffEl) {
            lumOff = parseInt(lumOffEl.getAttribute("val") || "0", 10);
          }
        } else if (srgbClr) {
          color = "#" + (srgbClr.getAttribute("val") || "ffffff");
        } else {
          color = "#000000";
        }

        gradientStops.push({
          position: pos,
          color,
          isSchemeColor,
          lumMod,
          lumOff,
        });
      }
    }

    const gradientAngle = lin
      ? parseInt(lin.getAttribute("ang") || "0", 10)
      : 0;

    return {
      type: "gradient",
      gradientStops,
      gradientAngle,
    };
  }
}
