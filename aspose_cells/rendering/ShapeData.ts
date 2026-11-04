import { DOMParser } from "@xmldom/xmldom";
import { RgbaColor } from "./ColorValue";
import { RenderColor } from "./RenderColor";

/**
 * 从 DrawingML XML (xdr:sp / xdr:cxnSp) 解析出的形状数据，渲染器所需的最小集合。
 * 几何、填充、线条（宽度 + 箭头）和文本。
 */
export class ParsedShape {
  geometry = "rect";
  flipH = false;
  flipV = false;
  rotationDeg = 0;
  adjusts: number[] | null = null;
  customGeometry: Element | null = null;

  hasFill = false;
  fillColor: RgbaColor = RgbaColor.white;
  fillPattern: string | null = null;
  patternForeground: RgbaColor = RgbaColor.black;
  reflection: Element | null = null;

  hasLine = true;
  lineColor: RgbaColor = RgbaColor.black;
  lineWidthPt = 0.75;

  headEnd: string | null = null; // 起始箭头
  tailEnd: string | null = null; // 末尾箭头

  text: string | null = null;
  textColor: RgbaColor = RgbaColor.black;
  textSizePt = 11;
  textBold = false;
  textVAnchor = "ctr"; // t / ctr / b
  textHAlign = "ctr";  // l / ctr / r
  textVerticalType = "horz";
  textWrap = false;
  textInsetTopPt = 3.6;
  textInsetLeftPt = 7.2;
}

/** SmartArt 形状：携带相对于图表框架的绝对位置和尺寸。 */
export class SmartArtShape {
  constructor(
    public shape: ParsedShape,
    public xPt: number,
    public yPt: number,
    public wPt: number,
    public hPt: number,
  ) {}
}

/** DrawingML XML 解析器。 */
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

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

function children(el: Element | null, ns: string, localName: string): Element[] {
  const result: Element[] = [];
  if (!el) return result;
  for (let i = 0; i < el.childNodes.length; i++) {
    const c = el.childNodes[i];
    if (c.nodeType === 1) {
      const e = c as Element;
      if (e.namespaceURI === ns && e.localName === localName) result.push(e);
    }
  }
  return result;
}

function descendants(el: Element | null, ns: string, localName: string): Element[] {
  const result: Element[] = [];
  if (!el) return result;
  const queue: Element[] = [el];
  while (queue.length > 0) {
    const cur = queue.shift()!;
    for (let i = 0; i < cur.childNodes.length; i++) {
      const c = cur.childNodes[i];
      if (c.nodeType === 1) {
        const e = c as Element;
        if (e.namespaceURI === ns && e.localName === localName) result.push(e);
        queue.push(e);
      }
    }
  }
  return result;
}

function firstDescendant(el: Element | null, ns: string, localName: string): Element | null {
  const result = descendants(el, ns, localName);
  return result.length > 0 ? result[0] : null;
}

function tryDouble(s: string | null): number | null {
  if (!s) return null;
  const v = parseFloat(s);
  if (isNaN(v) || !isFinite(v)) return null;
  return v;
}

function tryInt(s: string | null): number | null {
  if (!s) return null;
  const v = parseInt(s, 10);
  if (isNaN(v)) return null;
  return v;
}

export class ShapeXmlParser {
  /** 解析单个形状元素。 */
  static parse(rawXml: string | null, geometryFallback: string, colors: RenderColor): ParsedShape | null {
    if (!rawXml) return null;
    let doc: Document;
    try {
      doc = new DOMParser().parseFromString(rawXml, "text/xml");
    } catch {
      return ShapeXmlParser.fromFallback(geometryFallback);
    }
    const root = doc.documentElement;
    if (!root) return ShapeXmlParser.fromFallback(geometryFallback);
    return ShapeXmlParser.parseShapeElement(root, XDR, geometryFallback, colors);
  }

  /** 解析 SmartArt 图表 (dsp:drawing) 为带位置的形状列表。 */
  static parseSmartArtDrawing(rawXml: string | null, colors: RenderColor): SmartArtShape[] {
    const result: SmartArtShape[] = [];
    if (!rawXml) return result;
    let doc: Document;
    try {
      doc = new DOMParser().parseFromString(rawXml, "text/xml");
    } catch {
      return result;
    }
    const sps = descendants(doc.documentElement, DSP, "sp");
    for (const sp of sps) {
      const spPr = child(sp, DSP, "spPr");
      const xfrm = spPr ? child(spPr, A, "xfrm") : null;
      if (!xfrm) continue;
      const off = child(xfrm, A, "off");
      const ext = child(xfrm, A, "ext");
      if (!off || !ext) continue;
      const ox = tryDouble(attrOf(off, "x"));
      const oy = tryDouble(attrOf(off, "y"));
      const cx = tryDouble(attrOf(ext, "cx"));
      const cy = tryDouble(attrOf(ext, "cy"));
      if (ox === null || oy === null || cx === null || cy === null) continue;
      if (cx <= 0 || cy <= 0) continue;
      const shape = ShapeXmlParser.parseShapeElement(sp, DSP, "rect", colors);
      if (!shape) continue;
      result.push(new SmartArtShape(shape, ox / 12700, oy / 12700, cx / 12700, cy / 12700));
    }
    return result;
  }

  private static fromFallback(geometry: string): ParsedShape {
    const s = new ParsedShape();
    s.geometry = geometry ?? "rect";
    s.hasFill = false;
    return s;
  }

  private static parseShapeElement(root: Element, shapeNs: string, geometryFallback: string, colors: RenderColor): ParsedShape | null {
    const spPr = child(root, shapeNs, "spPr");
    const style = child(root, shapeNs, "style");
    const txBody = child(root, shapeNs, "txBody");
    const shape = new ParsedShape();

    const xfrm = spPr ? child(spPr, A, "xfrm") : null;
    if (xfrm) {
      if (attrOf(xfrm, "flipH") === "1") shape.flipH = true;
      if (attrOf(xfrm, "flipV") === "1") shape.flipV = true;
      const rot = tryDouble(attrOf(xfrm, "rot"));
      if (rot !== null) shape.rotationDeg = rot / 60000;
    }

    const prstGeom = spPr ? child(spPr, A, "prstGeom") : null;
    const prst = attrOf(prstGeom, "prst");
    shape.geometry = prst ? prst : (geometryFallback ?? "rect");
    ShapeXmlParser.parseAdjusts(prstGeom, shape);
    shape.customGeometry = spPr ? child(spPr, A, "custGeom") : null;

    ShapeXmlParser.parseFill(spPr, style, colors, shape);
    ShapeXmlParser.parseLine(spPr, style, colors, shape);
    ShapeXmlParser.parseText(txBody, style, colors, shape);

    let effects = spPr ? child(spPr, A, "effectLst") : null;
    if (!effects && style && colors) {
      const effectRef = child(style, A, "effectRef");
      const idx = tryInt(attrOf(effectRef, "idx"));
      if (idx !== null) effects = (colors as any).resolveThemeEffects ? (colors as any).resolveThemeEffects(idx) : null;
    }
    shape.reflection = effects ? child(effects, A, "reflection") : null;
    return shape;
  }

  private static parseAdjusts(prstGeom: Element | null, shape: ParsedShape): void {
    const avLst = prstGeom ? child(prstGeom, A, "avLst") : null;
    if (!avLst) return;
    const values: number[] = [];
    for (const gd of children(avLst, A, "gd")) {
      const fmla = attrOf(gd, "fmla");
      if (!fmla) continue;
      if (!fmla.startsWith("val ")) continue;
      const v = tryDouble(fmla.substring(4));
      if (v !== null) values.push(v / 100000);
    }
    shape.adjusts = values.length > 0 ? values : null;
  }

  private static parseFill(spPr: Element | null, style: Element | null, colors: RenderColor, shape: ParsedShape): void {
    if (spPr) {
      if (child(spPr, A, "noFill")) {
        shape.hasFill = false;
        return;
      }
      const pattern = child(spPr, A, "pattFill");
      if (pattern && attrOf(pattern, "prst") === "pct5") {
        shape.hasFill = true;
        shape.fillPattern = "pct5";
        const bgClr = child(pattern, A, "bgClr");
        const fgClr = child(pattern, A, "fgClr");
        const c1 = ShapeXmlParser.tryColor(ShapeXmlParser.fillChild(bgClr), colors);
        const c2 = ShapeXmlParser.tryColor(ShapeXmlParser.fillChild(fgClr), colors);
        shape.fillColor = c1 ?? RgbaColor.white;
        shape.patternForeground = c2 ?? RgbaColor.black;
        return;
      }
      const solid = child(spPr, A, "solidFill");
      if (solid) {
        const c = ShapeXmlParser.tryColor(ShapeXmlParser.fillChild(solid), colors);
        if (c) {
          shape.hasFill = true;
          shape.fillColor = c;
          return;
        }
      }
    }
    const fillRef = style ? child(style, A, "fillRef") : null;
    if (fillRef && attrOf(fillRef, "idx") !== "0") {
      const c = ShapeXmlParser.tryColor(ShapeXmlParser.fillChild2(fillRef), colors);
      if (c) {
        shape.hasFill = true;
        shape.fillColor = c;
      }
    }
  }

  private static parseLine(spPr: Element | null, style: Element | null, colors: RenderColor, shape: ParsedShape): void {
    const ln = spPr ? child(spPr, A, "ln") : null;
    if (ln && child(ln, A, "noFill")) {
      shape.hasLine = false;
      return;
    }
    if (ln) {
      const emu = tryDouble(attrOf(ln, "w"));
      if (emu !== null && emu > 0) {
        shape.lineWidthPt = emu / 12700;
      } else {
        const lnRefIdx = style && child(style, A, "lnRef") ? attrOf(child(style, A, "lnRef")!, "idx") : null;
        const idx = lnRefIdx !== null ? tryInt(lnRefIdx) : null;
        if (idx !== null && colors && (colors as any).resolveThemeLineWidth) {
          shape.lineWidthPt = (colors as any).resolveThemeLineWidth(idx, 0.75);
        } else {
          shape.lineWidthPt = 0.75;
        }
      }
    } else {
      shape.lineWidthPt = 0.75;
    }

    const lnFill = ln ? child(ln, A, "solidFill") : null;
    if (lnFill) {
      const c = ShapeXmlParser.tryColor(ShapeXmlParser.fillChild(lnFill), colors);
      if (c) {
        shape.hasLine = true;
        shape.lineColor = c;
      }
    } else {
      const lnRef = style ? child(style, A, "lnRef") : null;
      if (lnRef) {
        const c = ShapeXmlParser.tryColor(ShapeXmlParser.fillChild2(lnRef), colors);
        if (c) {
          shape.hasLine = true;
          shape.lineColor = c;
        }
      }
    }

    if (ln) {
      const head = child(ln, A, "headEnd");
      const tail = child(ln, A, "tailEnd");
      shape.headEnd = head ? (attrOf(head, "type") ?? "none") : null;
      shape.tailEnd = tail ? (attrOf(tail, "type") ?? "none") : null;
    }
  }

  private static parseText(txBody: Element | null, style: Element | null, colors: RenderColor, shape: ParsedShape): void {
    if (!txBody) return;
    const textNodes = descendants(txBody, A, "t");
    let sb = "";
    for (const t of textNodes) {
      sb += t.textContent ?? "";
    }
    shape.text = sb.length > 0 ? sb : null;

    const bodyPr = child(txBody, A, "bodyPr");
    if (bodyPr) {
      const anchor = attrOf(bodyPr, "anchor");
      if (anchor) shape.textVAnchor = anchor;
      const verticalType = attrOf(bodyPr, "vert");
      if (verticalType) shape.textVerticalType = verticalType;
      const wrap = attrOf(bodyPr, "wrap");
      shape.textWrap = wrap === "square";
      const tIns = tryDouble(attrOf(bodyPr, "tIns"));
      shape.textInsetTopPt = tIns !== null ? tIns / 12700 : 3.6;
      const lIns = tryDouble(attrOf(bodyPr, "lIns"));
      shape.textInsetLeftPt = lIns !== null ? lIns / 12700 : 7.2;
    }

    const firstPPr = firstDescendant(txBody, A, "pPr");
    const algn = attrOf(firstPPr, "algn");
    shape.textHAlign = algn ?? "l";

    const rPr = firstDescendant(txBody, A, "rPr") ?? firstDescendant(txBody, A, "endParaRPr");
    if (rPr) {
      const sz = tryDouble(attrOf(rPr, "sz"));
      if (sz !== null && sz > 0) shape.textSizePt = sz / 100;
      if (attrOf(rPr, "b") === "1") {
        shape.textBold = true;
        const runs = descendants(txBody, A, "r");
        for (const r of runs) {
          const rt = child(r, A, "t");
          const rp = child(r, A, "rPr");
          const rText = rt ? (rt.textContent ?? "") : "";
          if (rText.length > (shape.text?.length ?? 0) / 2
              && rp && attrOf(rp, "b") !== "1") {
            shape.textBold = false;
            break;
          }
        }
      }
      const solid = child(rPr, A, "solidFill");
      if (solid) {
        const c = ShapeXmlParser.tryColor(ShapeXmlParser.fillChild(solid), colors);
        if (c) { shape.textColor = c; return; }
      }
    }

    const fontRef = style ? child(style, A, "fontRef") : null;
    if (fontRef) {
      const c = ShapeXmlParser.tryColor(ShapeXmlParser.fillChild2(fontRef), colors);
      if (c) shape.textColor = c;
    }
  }

  private static fillChild(solidFill: Element | null): Element | null {
    if (!solidFill) return null;
    return child(solidFill, A, "srgbClr") ?? child(solidFill, A, "schemeClr");
  }

  private static fillChild2(refEl: Element | null): Element | null {
    if (!refEl) return null;
    return child(refEl, A, "srgbClr") ?? child(refEl, A, "schemeClr");
  }

  private static tryColor(colorEl: Element | null, colors: RenderColor): RgbaColor | null {
    if (!colorEl) return null;
    const localName = colorEl.localName;
    if (localName === "srgbClr") {
      const h = attrOf(colorEl, "val");
      const c = ShapeXmlParser.tryHex(h);
      if (c) return ShapeXmlParser.applyColorModifiers(colorEl, c);
      return null;
    }
    if (localName === "schemeClr") {
      const name = attrOf(colorEl, "val") ?? "";
      let baseColor = colors.resolveSchemeName(name, new RgbaColor(128, 128, 128));
      const lumModEl = child(colorEl, A, "lumMod");
      const lumOffEl = child(colorEl, A, "lumOff");
      let lumMod = 0, lumOff = 0;
      if (lumModEl) {
        const v = tryDouble(attrOf(lumModEl, "val"));
        if (v !== null) lumMod = v;
      }
      if (lumOffEl) {
        const v = tryDouble(attrOf(lumOffEl, "val"));
        if (v !== null) lumOff = v;
      }
      if (lumMod > 0 || lumOff > 0) {
        baseColor = RenderColor.applyLuma(baseColor, lumMod > 0 ? lumMod / 100000 : 1, lumOff / 100000);
      }
      return ShapeXmlParser.applyColorModifiers(colorEl, baseColor);
    }
    return null;
  }

  private static applyColorModifiers(colorEl: Element | null, color: RgbaColor): RgbaColor {
    if (!colorEl) return color;
    const shadeEl = child(colorEl, A, "shade");
    if (shadeEl) {
      const shadeVal = tryInt(attrOf(shadeEl, "val"));
      if (shadeVal !== null) {
        const factor = shadeVal / 100000;
        color = new RgbaColor(
          ShapeXmlParser.scaleByte(color.r, factor),
          ShapeXmlParser.scaleByte(color.g, factor),
          ShapeXmlParser.scaleByte(color.b, factor),
          color.a,
        );
      }
    }
    const tintEl = child(colorEl, A, "tint");
    if (tintEl) {
      const tintVal = tryInt(attrOf(tintEl, "val"));
      if (tintVal !== null) {
        const amount = tintVal / 100000;
        color = new RgbaColor(
          ShapeXmlParser.tintByte(color.r, amount),
          ShapeXmlParser.tintByte(color.g, amount),
          ShapeXmlParser.tintByte(color.b, amount),
          color.a,
        );
      }
    }
    const alphaEl = child(colorEl, A, "alpha");
    if (alphaEl) {
      const av = tryInt(attrOf(alphaEl, "val"));
      if (av !== null) {
        const clamped = Math.max(0, Math.min(100000, av));
        color = new RgbaColor(color.r, color.g, color.b, Math.round(clamped * 255 / 100000));
      }
    }
    return color;
  }

  private static scaleByte(value: number, factor: number): number {
    const channel = value / 255;
    const linear = channel <= 0.04045 ? channel / 12.92 : Math.pow((channel + 0.055) / 1.055, 2.4);
    const clamped = Math.max(0, Math.min(1, linear * factor));
    const encoded = clamped <= 0.0031308 ? clamped * 12.92 : 1.055 * Math.pow(clamped, 1 / 2.4) - 0.055;
    return Math.round(encoded * 255);
  }

  private static tintByte(value: number, amount: number): number {
    const tinted = value + (255 - value) * amount;
    return Math.max(0, Math.min(255, Math.round(tinted)));
  }

  private static tryHex(hex: string | null): RgbaColor | null {
    if (!hex || hex.length < 6) return null;
    let h = hex.length === 8 ? hex.substring(2) : hex;
    const r = parseInt(h.substring(0, 2), 16);
    const g = parseInt(h.substring(2, 4), 16);
    const b = parseInt(h.substring(4, 6), 16);
    if (isNaN(r) || isNaN(g) || isNaN(b)) return null;
    return new RgbaColor(r, g, b, 255);
  }
}
