import { TableStyleType, TotalsCalculation, Workbook } from "../../aspose_cells";
import { AssertEx } from "./TestInfrastructure";

/** Scenario factory for Excel table workbooks. */
export class ListObjectScenarioFactory {
  public static CreateTableWorkbook(): Workbook {
    const workbook = new Workbook();
    const sheet = workbook.Worksheets[0];
    sheet.Name = "SalesData";

    sheet.Cells["A1"].PutValue("Product");
    sheet.Cells["B1"].PutValue("Region");
    sheet.Cells["C1"].PutValue("Revenue");
    sheet.Cells["A2"].PutValue("Widget");
    sheet.Cells["B2"].PutValue("North");
    sheet.Cells["C2"].PutValue(1200);
    sheet.Cells["A3"].PutValue("Gadget");
    sheet.Cells["B3"].PutValue("South");
    sheet.Cells["C3"].PutValue(850);
    sheet.Cells["A4"].PutValue("Doohickey");
    sheet.Cells["B4"].PutValue("East");
    sheet.Cells["C4"].PutValue(2100);

    const index = sheet.ListObjects.Add("A1", "C4", true);
    const table = sheet.ListObjects.get(index);
    table.TableStyleType = TableStyleType.TableStyleMedium2;
    table.ShowTableStyleRowStripes = true;
    table.ShowTableStyleFirstColumn = false;
    table.ShowTableStyleLastColumn = false;
    table.ShowTableStyleColumnStripes = false;
    table.ShowTotals = true;
    table.ListColumns.get(2).TotalsCalculation = TotalsCalculation.Sum;
    table.ListColumns.get(0).TotalsRowLabel = "Total";

    return workbook;
  }

  public static AssertTableWorkbook(workbook: Workbook): void {
    const sheet = workbook.Worksheets[0];
    AssertEx.equal(1, sheet.ListObjects.Count);

    const table = sheet.ListObjects.get(0);
    AssertEx.equal("Table1", table.DisplayName);
    AssertEx.equal(0, table.StartRow);
    AssertEx.equal(0, table.StartColumn);
    AssertEx.equal(4, table.EndRow);
    AssertEx.equal(2, table.EndColumn);
    AssertEx.true(table.ShowHeaderRow);
    AssertEx.true(table.ShowTotals);
    AssertEx.equal(TableStyleType.TableStyleMedium2, table.TableStyleType);
    AssertEx.equal("TableStyleMedium2", table.TableStyleName);
    AssertEx.true(table.ShowTableStyleRowStripes);
    AssertEx.false(table.ShowTableStyleFirstColumn);
    AssertEx.false(table.ShowTableStyleLastColumn);
    AssertEx.false(table.ShowTableStyleColumnStripes);
    AssertEx.equal(3, table.ListColumns.Count);
    AssertEx.equal("Product", table.ListColumns.get(0).Name);
    AssertEx.equal("Region", table.ListColumns.get(1).Name);
    AssertEx.equal("Revenue", table.ListColumns.get(2).Name);
    AssertEx.equal(TotalsCalculation.Sum, table.ListColumns.get(2).TotalsCalculation);
    AssertEx.equal(TotalsCalculation.None, table.ListColumns.get(0).TotalsCalculation);
    AssertEx.equal("Total", table.ListColumns.get(0).TotalsRowLabel);
  }
}
