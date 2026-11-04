import type PDFDocument from "pdfkit";
import { DOMParser } from "@xmldom/xmldom";
import { Rect } from "./ShapeTypes";
import { RenderContext } from "./RenderContext";
import { PageLayout } from "./PageLayout";
import { ShapeXmlParser, ParsedShape } from "./ShapeData";
import { ShapeRenderer } from "./ShapeRenderer";

/** 命名空间常量。 */
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function attrOf(el: Element | null, name: string): string | null {
  if (!el) return null;
  const a = el.getAttribute(name);
  return a ?? null;
}

function child(el: Element | null, ns: string, localName: string): Element | null {
  if (!el) return null;
  for (let i = 0; i < el.childNodes.length; i++) {
    const c = el.childNodes[i];
    if (c.nodeType === 1) {
      const e = c as Element;
      if (e.namespaceURI === ns && e.localName === localName) return e;
    }
  }
  return null;
}

function children(el: Element | null): Element[] {
  const result: Element[] = [];
  if (!el) return result;
  for (let i = 0; i < el.childNodes.length; i++) {
    const c = el.childNodes[i];
    if (c.nodeType === 1) result.push(c as Element);
  }
  return result;
}

function numberOf(el: Element | null, name: string): number {
  if (!el) return 0;
  const v = parseFloat(attrOf(el, name) ?? "");
  return isNaN(v) || !isFinite(v) ? 0 : v;
}

/** DrawingML 组合形状渲染器：将子形状按锚点坐标定位到框架矩形内。 */
export class GroupShapeRenderer {
  private context: RenderContext;

  constructor(context: RenderContext) {
    this.context = context;
  }

  tryDraw(doc: InstanceType<typeof PDFDocument>, frame: Rect, xml: string | null): boolean {
    if (!xml || xml.indexOf("grpSp") < 0) return false;
    let root: Element;
    try {
      root = new DOMParser().parseFromString(xml, "text/xml").documentElement;
    } catch {
      return false;
    }
    if (root.namespaceURI !== XDR || root.localName !== "grpSp") return false;
    this.drawGroup(doc, frame, root);
    return true;
  }

  private drawGroup(doc: InstanceType<typeof PDFDocument>, frame: Rect, root: Element): void {
    const properties = child(root, XDR, "grpSpPr");
    const transform = properties ? child(properties, A, "xfrm") : null;
    const childOffset = transform ? child(transform, A, "chOff") : null;
    const childExtent = transform ? child(transform, A, "chExt") : null;
    if (!childOffset || !childExtent) return;

    const originX = numberOf(childOffset, "x");
    const originY = numberOf(childOffset, "y");
    const extentX = numberOf(childExtent, "cx");
    const extentY = numberOf(childExtent, "cy");
    if (extentX <= 0 || extentY <= 0) return;

    const renderer = new ShapeRenderer(doc, this.context);
    for (const shapeEl of children(root)) {
      if (shapeEl.namespaceURI !== XDR) continue;
      if (shapeEl.localName !== "sp" && shapeEl.localName !== "pic" && shapeEl.localName !== "grpSp") continue;

      const shapePropEl = shapeEl.localName === "grpSp" ? "grpSpPr" : "spPr";
      const shapeProperties = child(shapeEl, XDR, shapePropEl);
      const childTransform = shapeProperties ? child(shapeProperties, A, "xfrm") : null;
      const off = childTransform ? child(childTransform, A, "off") : null;
      const ext = childTransform ? child(childTransform, A, "ext") : null;
      if (!off || !ext) continue;

      const left = frame.left + ((numberOf(off, "x") - originX) / extentX) * frame.width;
      const top = frame.top + ((numberOf(off, "y") - originY) / extentY) * frame.height;
      const width = (numberOf(ext, "cx") / extentX) * frame.width;
      const height = (numberOf(ext, "cy") / extentY) * frame.height;
      if (width <= 0 || height <= 0) continue;

      const childFrame = new Rect(left, top, left + width, top + height);
      if (shapeEl.localName === "grpSp") {
        this.drawGroup(doc, childFrame, shapeEl);
      } else if (shapeEl.localName === "pic") {
        // 暂不处理图片
      } else {
        const parsed = ShapeXmlParser.parse(shapeEl.toString(), "rect", this.context.colors);
        if (parsed) {
          if (parsed.textWrap && parsed.text) parsed.textSizePt *= 1.18;
          renderer.draw(childFrame, parsed);
        }
      }
    }
  }
}
