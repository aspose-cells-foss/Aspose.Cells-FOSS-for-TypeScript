import { PageOrientation } from "./Core/PageOrientation";
import { PageSetupModel } from "./Core/PageSetupModel";
import { DiagnosticSeverity } from "./enums";
import { LoadDiagnostics } from "./LoadDiagnostics";
import { LoadIssue } from "./LoadIssue";
import { LoadOptions } from "./LoadOptions";
import { WarningInfo } from "./WarningInfo";
import { ZipReader } from "./Packaging/Archive";
import { attribute, documentRoot, elementChildren, firstElement, parseXml, XmlNode } from "./Xml/XmlReader";

/**
 * Reads portable page defaults from the public portion of a saved DEVMODEW.
 */
export function loadPrinterSettings(
  setup: PageSetupModel,
  root: XmlNode,
  archive: ZipReader,
  worksheetUri: string,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  sheetName: string,
): void {
  const pageSetupElement = firstElement(root, "pageSetup");
  const id = pageSetupElement === null ? null : attributeOrNull(pageSetupElement, "id");
  if (id === null || id.length === 0) {
    return;
  }

  if (setup.PaperSize !== 0 && setup.Orientation !== PageOrientation.Default) {
    return;
  }

  const slash = worksheetUri.lastIndexOf("/");
  const relsUri =
    worksheetUri.substring(0, slash + 1) + "_rels/" + worksheetUri.substring(slash + 1) + ".rels";
  const relationships = loadRelationships(archive, relsUri, worksheetUri);

  const target = relationships.get(id.toLowerCase());
  if (target === undefined) {
    warn(diagnostics, options, sheetName);
    return;
  }

  if (!archive.hasEntry(target)) {
    warn(diagnostics, options, sheetName);
    return;
  }

  const data = new Uint8Array(archive.readBinary(target));
  if (data.length < 220) {
    warn(diagnostics, options, sheetName);
    return;
  }

  const header = data.slice(0, 220);
  if (!tryApply(header, data.length, setup)) {
    warn(diagnostics, options, sheetName);
  }
}

export function tryApply(header: Uint8Array | null, length: number, setup: PageSetupModel): boolean {
  if (header === null || header.length < 220) {
    return false;
  }

  const size = readUInt16(header, 68);
  const extra = readUInt16(header, 70);
  if (size < 220 || size + extra > length) {
    return false;
  }

  // dmFields marks which fields are meaningful. Never interpret driver-private bytes.
  const fields = readUInt16(header, 72);
  const paper = readUInt16(header, 78);
  if ((fields & 2) !== 0 && paper > 0 && paper < 256 && setup.PaperSize === 0) {
    setup.PaperSize = paper;
  }

  const orientation = readUInt16(header, 76);
  if ((fields & 1) !== 0 && setup.Orientation === PageOrientation.Default) {
    if (orientation === 1) {
      setup.Orientation = PageOrientation.Portrait;
    }

    if (orientation === 2) {
      setup.Orientation = PageOrientation.Landscape;
    }
  }

  return true;
}

function readUInt16(data: Uint8Array, offset: number): number {
  return data[offset] | (data[offset + 1] << 8);
}

function loadRelationships(archive: ZipReader, relsUri: string, sourceUri: string): Map<string, string> {
  const relationships = new Map<string, string>();
  if (!archive.hasEntry(relsUri)) {
    return relationships;
  }

  const document = parseXml(archive.readText(relsUri));
  const root = documentRoot(document);
  if (root === null) {
    return relationships;
  }

  for (const relationship of elementChildren(root, "Relationship")) {
    const id = attribute(relationship, "Id");
    const target = attribute(relationship, "Target");
    if (id !== null && id.length > 0 && target !== null && target.length > 0) {
      relationships.set(id.toLowerCase(), resolvePartUri(sourceUri, target));
    }
  }

  return relationships;
}

function resolvePartUri(sourceUri: string, target: string): string {
  if (target.startsWith("/")) {
    return target;
  }

  const resolved = new URL(target, `http://package${sourceUri}`);
  return resolved.pathname;
}

function attributeOrNull(node: XmlNode, name: string): string | null {
  const attributes = (node as unknown as { attributes?: ArrayLike<{ nodeName: string; value: string }> }).attributes;
  if (attributes === undefined) {
    return null;
  }

  for (let index = 0; index < attributes.length; index++) {
    const attributeNode = attributes[index];
    const separator = attributeNode.nodeName.indexOf(":");
    const local = separator >= 0 ? attributeNode.nodeName.substring(separator + 1) : attributeNode.nodeName;
    if (local === name) {
      return attributeNode.value;
    }
  }

  return null;
}

function warn(diagnostics: LoadDiagnostics, options: LoadOptions, sheetName: string): void {
  const issue = new LoadIssue(
    "PS-L001",
    DiagnosticSeverity.Warning,
    "Saved printer settings are missing or invalid; explicit page settings and default paper size were retained.",
    false,
    true,
  );
  issue.SheetName = sheetName;
  diagnostics.Add(issue);

  if (options.WarningCallback !== null) {
    const info = new WarningInfo();
    info.Code = issue.Code;
    info.Severity = issue.Severity;
    info.Message = issue.Message;
    info.DataLossRisk = issue.DataLossRisk;
    info.PartUri = issue.PartUri;
    info.SheetName = issue.SheetName;
    info.CellRef = issue.CellRef;
    info.RowIndex = issue.RowIndex;
    options.WarningCallback.Warning(info);
  }
}
