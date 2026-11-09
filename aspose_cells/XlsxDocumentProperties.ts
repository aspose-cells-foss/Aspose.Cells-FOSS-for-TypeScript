import { WorkbookModel } from "./Core/WorkbookModel";
import { DiagnosticSeverity } from "./enums";
import { LoadDiagnostics } from "./LoadDiagnostics";
import { LoadIssue } from "./LoadIssue";
import { LoadOptions } from "./LoadOptions";
import { addIssue, resolvePartUri } from "./XlsxWorkbookArchiveHelpers";
import { ZipReader } from "./Packaging/Archive";
import { XmlElement } from "./Xml/XmlWriter";
import {
  XmlNode,
  attribute,
  documentRoot,
  elementChildren,
  firstElement,
  parseXml,
  textOf,
} from "./Xml/XmlReader";

const CORE_PROPERTIES_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const DUBLIN_CORE_NS = "http://purl.org/dc/elements/1.1/";
const DUBLIN_CORE_TERMS_NS = "http://purl.org/dc/terms/";
const XSI_NS = "http://www.w3.org/2001/XMLSchema-instance";
const EXTENDED_PROPERTIES_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
const VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

export const CORE_PROPERTIES_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
export const EXTENDED_PROPERTIES_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";

/** Builds the docProps/core.xml document, or null when core properties have no stored state. */
export function buildCorePropertiesDocument(model: WorkbookModel): XmlElement | null {
  if (!model.DocumentProperties.Core.HasStoredState()) {
    return null;
  }

  const core = model.DocumentProperties.Core;
  const root = XmlElement.create("cp:coreProperties")
    .attr("xmlns:cp", CORE_PROPERTIES_NS)
    .attr("xmlns:dc", DUBLIN_CORE_NS)
    .attr("xmlns:dcterms", DUBLIN_CORE_TERMS_NS)
    .attr("xmlns:xsi", XSI_NS);

  addStringElement(root, "dc:title", core.Title);
  addStringElement(root, "dc:subject", core.Subject);
  addStringElement(root, "dc:creator", core.Creator);
  addStringElement(root, "cp:keywords", core.Keywords);
  addStringElement(root, "dc:description", core.Description);
  addStringElement(root, "cp:lastModifiedBy", core.LastModifiedBy);
  addStringElement(root, "cp:revision", core.Revision);
  addStringElement(root, "cp:category", core.Category);
  addStringElement(root, "cp:contentStatus", core.ContentStatus);
  addDateElement(root, "dcterms:created", core.Created);
  addDateElement(root, "dcterms:modified", core.Modified);

  return root;
}

/** Builds the docProps/app.xml document, or null when extended properties have no stored state. */
export function buildExtendedPropertiesDocument(model: WorkbookModel): XmlElement | null {
  if (!model.DocumentProperties.Extended.HasStoredState()) {
    return null;
  }

  const extended = model.DocumentProperties.Extended;
  const root = XmlElement.create("Properties")
    .attr("xmlns", EXTENDED_PROPERTIES_NS)
    .attr("xmlns:vt", VT_NS);

  addStringElement(root, "Application", extended.Application);
  addStringElement(root, "AppVersion", extended.AppVersion);
  addStringElement(root, "Company", extended.Company);
  addStringElement(root, "Manager", extended.Manager);
  addIntElement(root, "DocSecurity", extended.DocSecurity);
  addStringElement(root, "HyperlinkBase", extended.HyperlinkBase);
  addBoolElement(root, "ScaleCrop", extended.ScaleCrop);
  addBoolElement(root, "LinksUpToDate", extended.LinksUpToDate);
  addBoolElement(root, "SharedDoc", extended.SharedDoc);

  return root;
}

/** Loads core and extended document properties resolved from the package root relationships. */
export function loadDocumentProperties(
  archive: ZipReader,
  model: WorkbookModel,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
): void {
  const relationshipTargets = loadRootDocumentPropertiesRelationshipTargets(archive, diagnostics, options);

  let corePropertiesPartUri = "";
  let extendedPropertiesPartUri = "";
  for (const [type, target] of relationshipTargets) {
    if (type.toLowerCase() === CORE_PROPERTIES_RELATIONSHIP_TYPE.toLowerCase()) {
      corePropertiesPartUri = target;
    } else if (type.toLowerCase() === EXTENDED_PROPERTIES_RELATIONSHIP_TYPE.toLowerCase()) {
      extendedPropertiesPartUri = target;
    }
  }

  loadCoreProperties(archive, corePropertiesPartUri, model, diagnostics, options);
  loadExtendedProperties(archive, extendedPropertiesPartUri, model, diagnostics, options);
}

function loadRootDocumentPropertiesRelationshipTargets(
  archive: ZipReader,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
): Map<string, string> {
  const relationshipTargets = new Map<string, string>();
  if (!archive.hasEntry("/_rels/.rels")) {
    return relationshipTargets;
  }

  let root: XmlNode | null;
  try {
    root = documentRoot(parseXml(archive.readText("/_rels/.rels")));
  } catch {
    addDocumentPropertiesIssue(
      diagnostics,
      options,
      "/_rels/.rels",
      "Package root relationships were malformed and document properties were ignored.",
    );
    return relationshipTargets;
  }

  if (root === null) {
    return relationshipTargets;
  }

  for (const relationship of elementChildren(root, "Relationship")) {
    const type = attribute(relationship, "Type");
    const target = attribute(relationship, "Target");
    const targetMode = attribute(relationship, "TargetMode");
    if (type === null || target === null || type.length === 0 || target.length === 0) {
      continue;
    }

    if (targetMode !== null && targetMode.toLowerCase() === "external") {
      continue;
    }

    if (
      type.toLowerCase() !== CORE_PROPERTIES_RELATIONSHIP_TYPE.toLowerCase() &&
      type.toLowerCase() !== EXTENDED_PROPERTIES_RELATIONSHIP_TYPE.toLowerCase()
    ) {
      continue;
    }

    let alreadyStored = false;
    for (const existingType of relationshipTargets.keys()) {
      if (existingType.toLowerCase() === type.toLowerCase()) {
        alreadyStored = true;
        break;
      }
    }

    if (!alreadyStored) {
      relationshipTargets.set(type, resolvePartUri("/", target));
    }
  }

  return relationshipTargets;
}

function loadCoreProperties(
  archive: ZipReader,
  partUri: string,
  model: WorkbookModel,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
): void {
  if (partUri.length === 0 || !archive.hasEntry(partUri)) {
    return;
  }

  let root: XmlNode | null;
  try {
    root = documentRoot(parseXml(archive.readText(partUri)));
  } catch {
    addDocumentPropertiesIssue(
      diagnostics,
      options,
      partUri,
      "Core document properties part was malformed and was ignored.",
    );
    return;
  }

  if (root === null) {
    return;
  }

  const core = model.DocumentProperties.Core;
  core.Title = readElementValue(root, "title");
  core.Subject = readElementValue(root, "subject");
  core.Creator = readElementValue(root, "creator");
  core.Keywords = readElementValue(root, "keywords");
  core.Description = readElementValue(root, "description");
  core.LastModifiedBy = readElementValue(root, "lastModifiedBy");
  core.Revision = readElementValue(root, "revision");
  core.Category = readElementValue(root, "category");
  core.ContentStatus = readElementValue(root, "contentStatus");
  core.Created = readDateElement(root, "created", diagnostics, options, partUri);
  core.Modified = readDateElement(root, "modified", diagnostics, options, partUri);
}

function loadExtendedProperties(
  archive: ZipReader,
  partUri: string,
  model: WorkbookModel,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
): void {
  if (partUri.length === 0 || !archive.hasEntry(partUri)) {
    return;
  }

  let root: XmlNode | null;
  try {
    root = documentRoot(parseXml(archive.readText(partUri)));
  } catch {
    addDocumentPropertiesIssue(
      diagnostics,
      options,
      partUri,
      "Extended document properties part was malformed and was ignored.",
    );
    return;
  }

  if (root === null) {
    return;
  }

  const extended = model.DocumentProperties.Extended;
  extended.Application = readElementValue(root, "Application");
  extended.AppVersion = readElementValue(root, "AppVersion");
  extended.Company = readElementValue(root, "Company");
  extended.Manager = readElementValue(root, "Manager");
  extended.DocSecurity = readIntElement(root, "DocSecurity", diagnostics, options, partUri);
  extended.HyperlinkBase = readElementValue(root, "HyperlinkBase");
  extended.ScaleCrop = readBoolElement(root, "ScaleCrop", diagnostics, options, partUri);
  extended.LinksUpToDate = readBoolElement(root, "LinksUpToDate", diagnostics, options, partUri);
  extended.SharedDoc = readBoolElement(root, "SharedDoc", diagnostics, options, partUri);
}

function addStringElement(parent: XmlElement, name: string, value: string): void {
  if (value !== null && value !== undefined && value.length > 0) {
    parent.child(XmlElement.create(name).child(value));
  }
}

function addIntElement(parent: XmlElement, name: string, value: number | null): void {
  if (value !== null) {
    parent.child(XmlElement.create(name).child(value.toString()));
  }
}

function addBoolElement(parent: XmlElement, name: string, value: boolean | null): void {
  if (value !== null) {
    parent.child(XmlElement.create(name).child(value ? "true" : "false"));
  }
}

function addDateElement(parent: XmlElement, name: string, value: Date | null): void {
  if (value === null) {
    return;
  }

  parent.child(XmlElement.create(name).attr("xsi:type", "dcterms:W3CDTF").child(formatXmlDate(value)));
}

function formatXmlDate(value: Date): string {
  const iso = value.toISOString();
  if (iso.endsWith(".000Z")) {
    return iso.substring(0, iso.length - 5) + "Z";
  }

  const match = iso.match(/\.\d+Z$/);
  if (match !== null) {
    const fraction = match[0].substring(0, match[0].length - 1).replace(/0+$/, "");
    return iso.substring(0, match.index) + fraction + "Z";
  }

  return iso;
}

function readElementValue(parent: XmlNode, localNameValue: string): string {
  const element = firstElement(parent, localNameValue);
  return element === null ? "" : textOf(element).trim();
}

function readDateElement(
  parent: XmlNode,
  localNameValue: string,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  partUri: string,
): Date | null {
  const element = firstElement(parent, localNameValue);
  if (element === null) {
    return null;
  }

  const parsed = new Date(textOf(element).trim());
  if (Number.isFinite(parsed.getTime())) {
    return parsed;
  }

  addDocumentPropertiesIssue(
    diagnostics,
    options,
    partUri,
    "Document property '" + localNameValue + "' had an invalid timestamp and was ignored.",
  );
  return null;
}

function readIntElement(
  parent: XmlNode,
  localNameValue: string,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  partUri: string,
): number | null {
  const element = firstElement(parent, localNameValue);
  if (element === null) {
    return null;
  }

  const value = textOf(element).trim();
  const parsed = value.length === 0 ? Number.NaN : Number(value);
  if (/^[+-]?\d+$/.test(value) && Number.isInteger(parsed) && parsed >= 0) {
    return parsed;
  }

  addDocumentPropertiesIssue(
    diagnostics,
    options,
    partUri,
    "Document property '" + localNameValue + "' had an invalid integer value and was ignored.",
  );
  return null;
}

function readBoolElement(
  parent: XmlNode,
  localNameValue: string,
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  partUri: string,
): boolean | null {
  const element = firstElement(parent, localNameValue);
  if (element === null) {
    return null;
  }

  const rawValue = textOf(element);
  if (rawValue === "1" || rawValue.toLowerCase() === "true") {
    return true;
  }

  if (rawValue === "0" || rawValue.toLowerCase() === "false") {
    return false;
  }

  addDocumentPropertiesIssue(
    diagnostics,
    options,
    partUri,
    "Document property '" + localNameValue + "' had an invalid Boolean value and was ignored.",
  );
  return null;
}

function addDocumentPropertiesIssue(
  diagnostics: LoadDiagnostics,
  options: LoadOptions,
  partUri: string,
  message: string,
): void {
  const issue = new LoadIssue("WB-L004", DiagnosticSeverity.Warning, message);
  issue.PartUri = partUri;
  addIssue(diagnostics, options, issue);
}
