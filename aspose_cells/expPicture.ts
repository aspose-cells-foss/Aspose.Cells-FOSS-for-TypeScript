import { PictureInfo } from "./types";
import { escapeXml } from "./util";

export class ExpPicture {
  static generatePictureXml(pictures: PictureInfo[], imageMap: Map<string, number>): string {
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`;

    xml += ExpPicture.generatePictureAnchors(pictures, imageMap);

    xml += "</xdr:wsDr>";
    return xml;
  }

  static generatePictureAnchors(pictures: PictureInfo[], imageMap: Map<string, number>): string {
    let xml = "";

    const sortedPictures = [...pictures].sort((a, b) => (a.zIndex || 0) - (b.zIndex || 0));

    for (let i = 0; i < sortedPictures.length; i++) {
      const picture = sortedPictures[i];
      xml += ExpPicture.pictureToDrawingXml(picture, i, imageMap);
    }

    return xml;
  }

  static generateUniqueId(): string {
    return "{" + ExpPicture.generateGuid() + "}";
  }

  static generateGuid(): string {
    return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
      const r = (Math.random() * 16) | 0;
      const v = c === "x" ? r : (r & 0x3) | 0x8;
      return v.toString(16).toUpperCase();
    });
  }

  static pictureToDrawingXml(picture: PictureInfo, index: number, imageMap: Map<string, number>): string {
    const fromCol = picture.fromCol;
    const fromColOff = picture.fromColOff ?? 0;
    const fromRow = picture.fromRow;
    const fromRowOff = picture.fromRowOff ?? 0;
    const toCol = picture.toCol;
    const toColOff = picture.toColOff ?? 0;
    const toRow = picture.toRow;
    const toRowOff = picture.toRowOff ?? 0;

    const x = picture.x ?? 0;
    const y = picture.y ?? 0;
    const cx = picture.width ?? 0;
    const cy = picture.height ?? 0;

    const uniqueId = ExpPicture.generateUniqueId();
    const extLst = `<a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="${uniqueId}"/></a:ext></a:extLst>`;

    const imageId = imageMap.get(picture.src) ?? 1;
    const rId = `rId${imageId}`;

    let anchor = `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${fromCol}</xdr:col><xdr:colOff>${fromColOff}</xdr:colOff><xdr:row>${fromRow}</xdr:row><xdr:rowOff>${fromRowOff}</xdr:rowOff></xdr:from><xdr:to><xdr:col>${toCol}</xdr:col><xdr:colOff>${toColOff}</xdr:colOff><xdr:row>${toRow}</xdr:row><xdr:rowOff>${toRowOff}</xdr:rowOff></xdr:to>`;

    anchor += `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${index + 3}" name="${escapeXml(picture.name)}">${extLst}</xdr:cNvPr><xdr:cNvPicPr><a:picLocks noChangeAspect="1" /></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${rId}" /><a:stretch><a:fillRect /></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="${x}" y="${y}" /><a:ext cx="${cx}" cy="${cy}" /></a:xfrm><a:prstGeom prst="rect"><a:avLst /></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData /></xdr:twoCellAnchor>`;

    return anchor;
  }

  static generateDrawingRels(pictures: PictureInfo[]): { [key: string]: string } {
    const rels: { [key: string]: string } = {};
    const uniqueImages = new Set<string>();

    for (const picture of pictures) {
      if (!uniqueImages.has(picture.src)) {
        uniqueImages.add(picture.src);
        const id = uniqueImages.size;
        rels[`rId${id}`] = picture.src;
      }
    }

    return rels;
  }

  static generateDrawingRelsXml(pictures: PictureInfo[], imageMap: Map<string, number>): string {
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;

    for (const [src, id] of imageMap) {
      const fileName = src.split("/").pop() || src;
      xml += `<Relationship Id="rId${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${fileName}" />`;
    }

    xml += "</Relationships>";
    return xml;
  }
}
