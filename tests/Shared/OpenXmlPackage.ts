import * as fs from "fs";
import {
  attribute,
  documentRoot,
  elementChildren,
  firstElement,
  localName,
  parseXml,
  XmlNode,
} from "../../aspose_cells/Xml/XmlReader";
import { ZipReader, ZipWriter } from "../../aspose_cells/Packaging/Archive";
import { CONTENT_TYPE_NS, PACKAGE_RELATIONSHIP_NS, RELATIONSHIP_NS } from "../../aspose_cells/XlsxWorkbookSerializerCommon";
import { XmlElement, xmlDocument } from "../../aspose_cells/Xml/XmlWriter";

const WORKSHEET_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const WORKBOOK_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const OFFICE_DOCUMENT_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const WORKSHEET_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

/**
 * Reader-side replacement for the closed-source OpenXML SDK used by the .NET tests.
 * Exposes worksheet parts resolved through workbook.xml relationship ids.
 */
export class OpenXmlWorkbookReader {
  private readonly archive: ZipReader;
  private readonly workbookDocument: XmlNode;
  private readonly relationshipTargets: Map<string, string>;

  public constructor(public readonly path: string) {
    this.archive = ZipReader.fromFile(path);
    this.workbookDocument = parseXml(this.archive.readText("xl/workbook.xml"));
    this.relationshipTargets = readRelationshipTargets(this.archive, "xl/_rels/workbook.xml.rels");
  }

  public get workbookRoot(): XmlNode {
    const root = documentRoot(this.workbookDocument);
    if (!root) {
      throw new Error("Workbook XML root element is missing.");
    }

    return root;
  }

  public get sheetNames(): string[] {
    return elementChildren(firstElement(this.workbookRoot, "sheets"), "sheet").map(
      (sheet) => attribute(sheet, "name") ?? "",
    );
  }

  public sheetPartPathByName(sheetName: string): string {
    const sheets = elementChildren(firstElement(this.workbookRoot, "sheets"), "sheet");
    for (const sheet of sheets) {
      if ((attribute(sheet, "name") ?? "") === sheetName) {
        const relationshipId = attribute(sheet, "id");
        if (relationshipId === null) {
          throw new Error("Worksheet relationship id is missing.");
        }

        const target = this.relationshipTargets.get(relationshipId);
        if (target === undefined) {
          throw new Error(`Worksheet relationship '${relationshipId}' was not found.`);
        }

        return resolvePartPath("xl", target);
      }
    }

    throw new Error(`Worksheet '${sheetName}' was not found.`);
  }

  public sheetPartPathByIndex(index: number): string {
    const sheets = elementChildren(firstElement(this.workbookRoot, "sheets"), "sheet");
    if (index < 0 || index >= sheets.length) {
      throw new Error("Worksheet index was out of range.");
    }

    return this.sheetPartPathByName(attribute(sheets[index], "name") ?? "");
  }

  public sheetXmlByName(sheetName: string): XmlNode {
    return parseXml(this.archive.readText(this.sheetPartPathByName(sheetName)));
  }

  public sheetXmlByIndex(index: number): XmlNode {
    return parseXml(this.archive.readText(this.sheetPartPathByIndex(index)));
  }

  public partXml(partPath: string): XmlNode {
    return parseXml(this.archive.readText(partPath));
  }

  public readPartText(partPath: string): string {
    return this.archive.readText(partPath);
  }

  public hasPart(partPath: string): boolean {
    return this.archive.hasEntry(partPath);
  }

  public listEntries(): string[] {
    return this.archive.listEntries();
  }

  public get archiveReader(): ZipReader {
    return this.archive;
  }
}

/** Reads a relationship file into an Id -> Target map. */
export function readRelationshipTargets(archive: ZipReader, relationshipPartPath: string): Map<string, string> {
  const targets = new Map<string, string>();
  if (!archive.hasEntry(relationshipPartPath)) {
    return targets;
  }

  const root = documentRoot(parseXml(archive.readText(relationshipPartPath)));
  if (!root) {
    return targets;
  }

  for (const relationship of elementChildren(root, "Relationship")) {
    const id = attribute(relationship, "Id");
    const target = attribute(relationship, "Target");
    if (id !== null && target !== null) {
      targets.set(id, target);
    }
  }

  return targets;
}

/** Resolves a part path relative to the folder that owns the relationship part. */
export function resolvePartPath(baseFolder: string, target: string): string {
  if (target.startsWith("/")) {
    return target.replace(/^\/+/, "");
  }

  const segments = (baseFolder.length === 0 ? "" : baseFolder + "/").split("/").filter((s) => s.length > 0);
  for (const segment of target.split("/")) {
    if (segment === ".") {
      continue;
    }

    if (segment === "..") {
      segments.pop();
      continue;
    }

    segments.push(segment);
  }

  return segments.join("/");
}

export interface PackagePart {
  readonly partPath: string;
  readonly xml: string;
  readonly contentType?: string;
  readonly relationshipType?: string;
  readonly relationshipId?: string;
}

/**
 * Writer-side replacement for the OpenXML SDK package builder.
 * Builds a minimal, valid workbook package from raw part XML.
 */
export class OpenXmlPackageBuilder {
  private workbookXml: string | null = null;
  private readonly worksheetParts: Array<{ relationshipId: string; xml: string }> = [];
  private readonly extraParts: PackagePart[] = [];

  /** Registers a worksheet part and returns the relationship id assigned to it. */
  public addWorksheet(xml: string, relationshipId?: string): string {
    const id = relationshipId ?? `rId${this.worksheetParts.length + 1}`;
    this.worksheetParts.push({ relationshipId: id, xml });
    return id;
  }

  public setWorkbookXml(xml: string): void {
    this.workbookXml = xml;
  }

  public addPart(part: PackagePart): void {
    this.extraParts.push(part);
  }

  public save(path: string): void {
    fs.writeFileSync(path, this.toBuffer());
  }

  public toBuffer(): Buffer {
    const archive = new ZipWriter();
    this.writeContentTypes(archive);
    this.writeRootRelationships(archive);
    this.writeWorkbook(archive);
    this.writeWorkbookRelationships(archive);

    for (let index = 0; index < this.worksheetParts.length; index++) {
      archive.addText(`xl/worksheets/sheet${index + 1}.xml`, this.worksheetParts[index].xml);
    }

    for (const part of this.extraParts) {
      archive.addText(part.partPath, part.xml);
    }

    return archive.toBuffer();
  }

  private writeContentTypes(archive: ZipWriter): void {
    const root = XmlElement.createWithNamespace("Types", CONTENT_TYPE_NS)
      .child(
        XmlElement.create("Default")
          .attr("Extension", "rels")
          .attr("ContentType", "application/vnd.openxmlformats-package.relationships+xml"),
      )
      .child(
        XmlElement.create("Default")
          .attr("Extension", "xml")
          .attr("ContentType", "application/xml"),
      )
      .child(
        XmlElement.create("Override")
          .attr("PartName", "/xl/workbook.xml")
          .attr("ContentType", WORKBOOK_CONTENT_TYPE),
      );

    for (let index = 0; index < this.worksheetParts.length; index++) {
      root.child(
        XmlElement.create("Override")
          .attr("PartName", `/xl/worksheets/sheet${index + 1}.xml`)
          .attr("ContentType", WORKSHEET_CONTENT_TYPE),
      );
    }

    for (const part of this.extraParts) {
      if (part.contentType !== undefined) {
        root.child(
          XmlElement.create("Override")
            .attr("PartName", `/${part.partPath}`)
            .attr("ContentType", part.contentType),
        );
      }
    }

    archive.addText("[Content_Types].xml", xmlDocument(root));
  }

  private writeRootRelationships(archive: ZipWriter): void {
    const root = XmlElement.createWithNamespace("Relationships", PACKAGE_RELATIONSHIP_NS).child(
      XmlElement.create("Relationship")
        .attr("Id", "rId1")
        .attr("Type", OFFICE_DOCUMENT_RELATIONSHIP)
        .attr("Target", "xl/workbook.xml"),
    );
    archive.addText("_rels/.rels", xmlDocument(root));
  }

  private writeWorkbook(archive: ZipWriter): void {
    if (this.workbookXml !== null) {
      archive.addText("xl/workbook.xml", this.workbookXml);
      return;
    }

    const root = XmlElement.createWithNamespace("workbook", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
      .attr("xmlns:r", RELATIONSHIP_NS);
    const sheets = XmlElement.create("sheets");
    for (let index = 0; index < this.worksheetParts.length; index++) {
      sheets.child(
        XmlElement.create("sheet")
          .attr("name", `Sheet${index + 1}`)
          .attr("sheetId", index + 1)
          .attr("r:id", this.worksheetParts[index].relationshipId),
      );
    }

    root.child(sheets);
    archive.addText("xl/workbook.xml", xmlDocument(root));
  }

  private writeWorkbookRelationships(archive: ZipWriter): void {
    const root = XmlElement.createWithNamespace("Relationships", PACKAGE_RELATIONSHIP_NS);
    for (let index = 0; index < this.worksheetParts.length; index++) {
      root.child(
        XmlElement.create("Relationship")
          .attr("Id", this.worksheetParts[index].relationshipId)
          .attr("Type", WORKSHEET_RELATIONSHIP)
          .attr("Target", `worksheets/sheet${index + 1}.xml`),
      );
    }

    for (const part of this.extraParts) {
      const id = part.relationshipId ?? `rId${indexOfRelationship(root) + 1}`;
      if (part.relationshipType !== undefined) {
        root.child(
          XmlElement.create("Relationship")
            .attr("Id", id)
            .attr("Type", part.relationshipType)
            .attr("Target", part.partPath.replace(/^xl\//, "")),
        );
      }
    }

    archive.addText("xl/_rels/workbook.xml.rels", xmlDocument(root));
  }
}

function indexOfRelationship(root: XmlElement): number {
  return root.childCount;
}

export interface CreateOpenXmlWorkbookOptions {
  /** Overrides the generated workbook.xml part, for tests that add defined names. */
  readonly workbookXml?: string;
  /** Additional parts (for example styles.xml) written next to the worksheet. */
  readonly extraParts?: readonly PackagePart[];
}

/** Convenience helper: creates a package containing one worksheet with the given XML. */
export function createOpenXmlWorkbook(
  path: string,
  sheetName: string,
  worksheetXml: string,
  options?: CreateOpenXmlWorkbookOptions,
): void {
  const builder = new OpenXmlPackageBuilder();
  builder.addWorksheet(worksheetXml, "rId1");
  builder.setWorkbookXml(
    options?.workbookXml ??
      xmlDocument(
        XmlElement.createWithNamespace(
          "workbook",
          "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        )
          .attr("xmlns:r", RELATIONSHIP_NS)
          .child(
            XmlElement.create("sheets").child(
              XmlElement.create("sheet").attr("name", sheetName).attr("sheetId", 1).attr("r:id", "rId1"),
            ),
          ),
      ),
  );
  for (const part of options?.extraParts ?? []) {
    builder.addPart(part);
  }

  builder.save(path);
}

export { localName };
