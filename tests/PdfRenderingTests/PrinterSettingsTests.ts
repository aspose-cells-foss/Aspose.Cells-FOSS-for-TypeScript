import { PageOrientation } from "../../aspose_cells/Core/PageOrientation";
import { PageSetupModel } from "../../aspose_cells/Core/PageSetupModel";
import { LoadDiagnostics } from "../../aspose_cells/LoadDiagnostics";
import { LoadOptions } from "../../aspose_cells/LoadOptions";
import { ZipReader, ZipWriter } from "../../aspose_cells/Packaging/Archive";
import { loadPrinterSettings, tryApply } from "../../aspose_cells/XlsxWorkbookPrinterSettings";
import { documentRoot, parseXml } from "../../aspose_cells/Xml/XmlReader";
import { AssertEx } from "../Shared/TestInfrastructure";
import { guardedTest, requireInput, RASTER_REASON } from "./Support";

const MAIN_NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELATIONSHIP_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_RELATIONSHIP_NAMESPACE =
  "http://schemas.openxmlformats.org/package/2006/relationships";

/**
 * Port of Aspose.Cells_FOSS.PdfRenderingTests.PrinterSettingsTests.
 */
export class PrinterSettingsTests {
  /** Public DEVMODEW header parsing plus recovery when the printer relation is missing. */
  static PublicHeaderAndRecovery(): void {
    const header = new Uint8Array(220);
    header[68] = 220; // dmSize
    header[72] = 3; // DM_ORIENTATION | DM_PAPERSIZE
    header[76] = 2; // landscape
    header[78] = 1; // Letter
    let setup = new PageSetupModel();
    AssertEx.true(tryApply(header, header.length, setup));
    AssertEx.equal(1, setup.PaperSize);
    AssertEx.equal(PageOrientation.Landscape, setup.Orientation);

    setup.PaperSize = 9;
    setup.Orientation = PageOrientation.Portrait;
    tryApply(header, header.length, setup);
    AssertEx.equal(9, setup.PaperSize, "Explicit worksheet paper size wins.");
    AssertEx.equal(PageOrientation.Portrait, setup.Orientation);

    header[72] = 0;
    setup = new PageSetupModel();
    tryApply(header, header.length, setup);
    AssertEx.equal(0, setup.PaperSize, "Ignore fields without their validity flags.");
    AssertEx.equal(PageOrientation.Default, setup.Orientation);

    header[70] = 1;
    AssertEx.true(!tryApply(header, header.length, setup));
    AssertEx.true(!tryApply(new Uint8Array(219), 219, setup));
    header[70] = 0;
    header[72] = 3;

    const writer = new ZipWriter();
    writer.addText(
      "xl/worksheets/_rels/sheet1.xml.rels",
      `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${PACKAGE_RELATIONSHIP_NAMESPACE}">` +
        `<Relationship Id="rId1" Type="${RELATIONSHIP_NAMESPACE}/printerSettings" ` +
        `Target="../printerSettings/printerSettings1.bin"/></Relationships>`,
    );
    writer.addBinary("xl/printerSettings/printerSettings1.bin", header);
    const reader = new ZipReader(writer.toBuffer());

    const diagnostics = new LoadDiagnostics();
    setup = new PageSetupModel();
    setup.PaperSize = 9;

    const rootXml =
      `<worksheet xmlns="${MAIN_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}">` +
      `<pageSetup r:id="rId1"/></worksheet>`;
    loadPrinterSettings(
      setup,
      PrinterSettingsTests.root(rootXml),
      reader,
      "/xl/worksheets/sheet1.xml",
      diagnostics,
      new LoadOptions(),
      "Sheet1",
    );
    AssertEx.equal(9, setup.PaperSize);
    AssertEx.equal(PageOrientation.Landscape, setup.Orientation);
    AssertEx.equal(0, diagnostics.Issues.length);

    const missingRootXml =
      `<worksheet xmlns="${MAIN_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}">` +
      `<pageSetup r:id="missing"/></worksheet>`;
    loadPrinterSettings(
      new PageSetupModel(),
      PrinterSettingsTests.root(missingRootXml),
      reader,
      "/xl/worksheets/sheet1.xml",
      diagnostics,
      new LoadOptions(),
      "Sheet1",
    );
    AssertEx.equal(1, diagnostics.Issues.length);
    AssertEx.equal("PS-L001", diagnostics.Issues[0].Code);
  }

  /** Fixture 0011: saved printer settings survive an XLSX round trip. */
  static FixtureRoundTrip(): void {
    requireInput("0011.xlsx");
    guardedTest([], RASTER_REASON);
  }

  /** Fixture 0011: geometry of the Excel reference output. */
  static ReferenceOutputGeometry(): void {
    requireInput("0011.xlsx");
    guardedTest([], RASTER_REASON);
  }

  private static root(xml: string) {
    const root = documentRoot(parseXml(xml));
    if (!root) {
      throw new Error("Expected a document root element.");
    }

    return root;
  }
}
