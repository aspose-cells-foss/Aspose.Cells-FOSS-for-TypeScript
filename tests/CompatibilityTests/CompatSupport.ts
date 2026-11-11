import * as fs from "fs";
import * as path from "path";
import { AssertEx, resolveRepositoryRoot } from "../Shared/TestInfrastructure";
import { localName, XmlNode } from "../../aspose_cells/Xml/XmlReader";

/** Returns the local names of the direct child elements of the supplied node. */
export function childLocalNames(node: XmlNode | null): string[] {
  const names: string[] = [];
  if (node === null || node.childNodes === null) {
    return names;
  }

  for (let index = 0; index < node.childNodes.length; index++) {
    const child = node.childNodes[index];
    const name = localName(child.nodeName);
    if (name.startsWith("#")) {
      continue;
    }

    names.push(name);
  }

  return names;
}

/** Asserts that the first element appears before the second one in document order. */
export function assertElementOrder(
  elementNames: readonly string[],
  expectedBefore: string,
  expectedAfter: string,
): void {
  const beforeIndex = findElementIndex(elementNames, expectedBefore);
  const afterIndex = findElementIndex(elementNames, expectedAfter);

  AssertEx.true(beforeIndex >= 0, "Missing worksheet element: " + expectedBefore);
  AssertEx.true(afterIndex >= 0, "Missing worksheet element: " + expectedAfter);
  AssertEx.true(
    beforeIndex < afterIndex,
    "Worksheet element order is invalid. Expected '" + expectedBefore + "' before '" + expectedAfter + "'.",
  );
}

/** Returns the position of the supplied element name, or -1 when it is absent. */
export function findElementIndex(elementNames: readonly string[], expectedName: string): number {
  for (let index = 0; index < elementNames.length; index++) {
    if (elementNames[index] === expectedName) {
      return index;
    }
  }

  return -1;
}

/** Returns the first direct child whose qualified name (prefix plus local name) matches. */
export function firstQualifiedChild(node: XmlNode | null, qualifiedName: string): XmlNode | null {
  if (node === null || node.childNodes === null) {
    return null;
  }

  for (let index = 0; index < node.childNodes.length; index++) {
    const child = node.childNodes[index];
    if (child.nodeName === qualifiedName) {
      return child;
    }
  }

  return null;
}

/** Follows a chain of direct children by local name and returns the final node. */
export function childPath(node: XmlNode | null, ...names: string[]): XmlNode | null {
  let current = node;
  for (const name of names) {
    if (current === null) {
      return null;
    }

    current = firstChildByLocalName(current, name);
  }

  return current;
}

function firstChildByLocalName(node: XmlNode, name: string): XmlNode | null {
  if (node.childNodes === null) {
    return null;
  }

  for (let index = 0; index < node.childNodes.length; index++) {
    const child = node.childNodes[index];
    if (localName(child.nodeName) === name) {
      return child;
    }
  }

  return null;
}

/** Resolves a repository-relative input file. */
export function resolveRepositoryFile(firstSegment: string, secondSegment: string): string {
  return path.join(resolveRepositoryRoot(), firstSegment, secondSegment);
}

/** Appends the roundtrip output suffix before the file extension. */
export function appendOutSuffix(fileName: string): string {
  const extension = path.extname(fileName);
  const baseName = path.basename(fileName, extension);
  if (extension.length === 0) {
    return baseName + "out";
  }

  return baseName + "out" + extension;
}

/** Lists all .xlsx files in the directory, mirroring Directory.GetFiles(dir, "*.xlsx"). */
export function listXlsxFiles(directory: string): string[] {
  if (!fs.existsSync(directory) || !fs.statSync(directory).isDirectory()) {
    throw new Error("Input directory not found: " + directory);
  }

  const files = fs
    .readdirSync(directory)
    .filter((name) => name.toLowerCase().endsWith(".xlsx"))
    .filter((name) => !name.startsWith("~$"))
    .filter((name) => fs.statSync(path.join(directory, name)).isFile())
    .sort((left, right) => left.localeCompare(right));

  if (files.length === 0) {
    throw new Error("No .xlsx files found in: " + directory);
  }

  return files.map((name) => path.join(directory, name));
}
