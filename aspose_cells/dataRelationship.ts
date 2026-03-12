export class DataRelationship {
  static worksheetRels(sheetIndex: number): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${sheetIndex}.xml"/></Relationships>`;
  }

  static drawingRels(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  }

  static drawingsRels(shapesArray: { length: number }[]): string {
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;

    let drawingCount = 0;
    for (const item of shapesArray) {
      if (item.length > 0) {
        drawingCount++;
        xml += `<Relationship Id="rId${drawingCount}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="drawing${drawingCount}.xml" />`;
      }
    }

    xml += "</Relationships>";
    return xml;
  }
}
