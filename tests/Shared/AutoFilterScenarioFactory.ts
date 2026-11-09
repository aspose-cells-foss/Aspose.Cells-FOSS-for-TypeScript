import {
  CellArea,
  Color,
  FillPattern,
  FilterOperatorType,
  FormatConditionType,
  OperatorType,
  Workbook,
  Worksheet,
} from "../../aspose_cells";
import { AssertEx } from "./TestInfrastructure";

/** Scenario factory for worksheet auto filter settings. */
export class AutoFilterScenarioFactory {
  public static CreateAutoFilterWorkbook(): Workbook {
    const workbook = new Workbook();
    const sheet = workbook.Worksheets[0];
    sheet.Name = "Filtered";

    sheet.Cells.get(0, 0).PutValue("Status");
    sheet.Cells.get(0, 1).PutValue("Amount");
    sheet.Cells.get(0, 2).PutValue("Color");
    sheet.Cells.get(0, 3).PutValue("Date");
    sheet.Cells.get(0, 4).PutValue("Score");
    sheet.Cells.get(1, 0).PutValue("Open");
    sheet.Cells.get(1, 1).PutValue(10);
    sheet.Cells.get(1, 2).PutValue("Red");
    sheet.Cells.get(1, 3).PutValue(new Date(Date.UTC(2024, 4, 1)));
    sheet.Cells.get(1, 4).PutValue(70);
    sheet.Cells.get(2, 0).PutValue("Closed");
    sheet.Cells.get(2, 1).PutValue(20);
    sheet.Cells.get(2, 2).PutValue("Blue");
    sheet.Cells.get(2, 3).PutValue(new Date(Date.UTC(2024, 4, 2)));
    sheet.Cells.get(2, 4).PutValue(80);
    sheet.Cells.get(3, 0).PutValue("Open");
    sheet.Cells.get(3, 1).PutValue(30);
    sheet.Cells.get(3, 2).PutValue("Green");
    sheet.Cells.get(3, 3).PutValue(new Date(Date.UTC(2024, 4, 3)));
    sheet.Cells.get(3, 4).PutValue(90);
    sheet.Cells.get(4, 0).PutValue("Pending");
    sheet.Cells.get(4, 1).PutValue(40);
    sheet.Cells.get(4, 2).PutValue("Yellow");
    sheet.Cells.get(4, 3).PutValue(new Date(Date.UTC(2024, 4, 4)));
    sheet.Cells.get(4, 4).PutValue(60);
    sheet.Cells.get(5, 0).PutValue("Closed");
    sheet.Cells.get(5, 1).PutValue(50);
    sheet.Cells.get(5, 2).PutValue("Black");
    sheet.Cells.get(5, 3).PutValue(new Date(Date.UTC(2024, 4, 5)));
    sheet.Cells.get(5, 4).PutValue(50);
    AutoFilterScenarioFactory.AddDifferentialStyles(sheet);

    sheet.AutoFilter.Range = "A1:E6";

    const statusColumn = sheet.AutoFilter.FilterColumns.get(sheet.AutoFilter.FilterColumns.Add(0));
    statusColumn.HiddenButton = true;
    statusColumn.Filters.Add("Open");
    statusColumn.Filters.Add("Closed");

    const amountColumn = sheet.AutoFilter.FilterColumns.get(sheet.AutoFilter.FilterColumns.Add(1));
    amountColumn.CustomFilters.MatchAll = true;
    amountColumn.CustomFilters.Add(FilterOperatorType.GreaterOrEqual, "10");
    amountColumn.CustomFilters.Add(FilterOperatorType.LessOrEqual, "50");

    const colorColumn = sheet.AutoFilter.FilterColumns.get(sheet.AutoFilter.FilterColumns.Add(2));
    colorColumn.ColorFilter.Enabled = true;
    colorColumn.ColorFilter.DifferentialStyleId = 3;
    colorColumn.ColorFilter.CellColor = true;

    const dateColumn = sheet.AutoFilter.FilterColumns.get(sheet.AutoFilter.FilterColumns.Add(3));
    dateColumn.DynamicFilter.Enabled = true;
    dateColumn.DynamicFilter.Type = "thisMonth";
    dateColumn.DynamicFilter.Value = 1;
    dateColumn.DynamicFilter.MaxValue = 31;

    const scoreColumn = sheet.AutoFilter.FilterColumns.get(sheet.AutoFilter.FilterColumns.Add(4));
    scoreColumn.Top10.Enabled = true;
    scoreColumn.Top10.Top = false;
    scoreColumn.Top10.Percent = true;
    scoreColumn.Top10.Value = 10;
    scoreColumn.Top10.FilterValue = 2.5;

    sheet.AutoFilter.SortState.Ref = "A2:E6";
    sheet.AutoFilter.SortState.CaseSensitive = true;
    sheet.AutoFilter.SortState.SortMethod = "pinYin";

    const valueSort = sheet.AutoFilter.SortState.SortConditions.get(
      sheet.AutoFilter.SortState.SortConditions.Add("B2:B6"),
    );
    valueSort.Descending = true;
    valueSort.SortBy = "value";
    valueSort.CustomList = "High,Medium,Low";

    const colorSort = sheet.AutoFilter.SortState.SortConditions.get(
      sheet.AutoFilter.SortState.SortConditions.Add("C2:C6"),
    );
    colorSort.SortBy = "cellColor";
    colorSort.DifferentialStyleId = 4;

    const iconSort = sheet.AutoFilter.SortState.SortConditions.get(
      sheet.AutoFilter.SortState.SortConditions.Add("E2:E6"),
    );
    iconSort.SortBy = "icon";
    iconSort.IconSet = "3TrafficLights1";
    iconSort.IconId = 2;

    return workbook;
  }

  public static AssertAutoFilter(workbook: Workbook): void {
    const sheet = workbook.Worksheets[0];
    AssertEx.equal("A1:E6", sheet.AutoFilter.Range);
    AssertEx.equal(5, sheet.AutoFilter.FilterColumns.Count);

    const statusColumn = sheet.AutoFilter.FilterColumns.get(0);
    AssertEx.equal(0, statusColumn.ColumnIndex);
    AssertEx.true(statusColumn.HiddenButton);
    AssertEx.equal(2, statusColumn.Filters.Count);
    AssertEx.equal("Open", statusColumn.Filters.get(0));
    AssertEx.equal("Closed", statusColumn.Filters.get(1));

    const amountColumn = sheet.AutoFilter.FilterColumns.get(1);
    AssertEx.equal(1, amountColumn.ColumnIndex);
    AssertEx.true(amountColumn.CustomFilters.MatchAll);
    AssertEx.equal(2, amountColumn.CustomFilters.Count);
    AssertEx.equal(FilterOperatorType.GreaterOrEqual, amountColumn.CustomFilters.get(0).Operator);
    AssertEx.equal("10", amountColumn.CustomFilters.get(0).Value);
    AssertEx.equal(FilterOperatorType.LessOrEqual, amountColumn.CustomFilters.get(1).Operator);
    AssertEx.equal("50", amountColumn.CustomFilters.get(1).Value);

    const colorColumn = sheet.AutoFilter.FilterColumns.get(2);
    AssertEx.equal(2, colorColumn.ColumnIndex);
    AssertEx.true(colorColumn.ColorFilter.Enabled);
    AssertEx.equal(3, colorColumn.ColorFilter.DifferentialStyleId ?? -1);
    AssertEx.true(colorColumn.ColorFilter.CellColor);

    const dateColumn = sheet.AutoFilter.FilterColumns.get(3);
    AssertEx.equal(3, dateColumn.ColumnIndex);
    AssertEx.true(dateColumn.DynamicFilter.Enabled);
    AssertEx.equal("thisMonth", dateColumn.DynamicFilter.Type);
    AssertEx.equal(1, dateColumn.DynamicFilter.Value ?? 0);
    AssertEx.equal(31, dateColumn.DynamicFilter.MaxValue ?? 0);

    const scoreColumn = sheet.AutoFilter.FilterColumns.get(4);
    AssertEx.equal(4, scoreColumn.ColumnIndex);
    AssertEx.true(scoreColumn.Top10.Enabled);
    AssertEx.false(scoreColumn.Top10.Top);
    AssertEx.true(scoreColumn.Top10.Percent);
    AssertEx.equal(10, scoreColumn.Top10.Value ?? 0);
    AssertEx.equal(2.5, scoreColumn.Top10.FilterValue ?? 0);

    AssertEx.equal("A2:E6", sheet.AutoFilter.SortState.Ref);
    AssertEx.true(sheet.AutoFilter.SortState.CaseSensitive);
    AssertEx.equal("pinYin", sheet.AutoFilter.SortState.SortMethod);
    AssertEx.equal(3, sheet.AutoFilter.SortState.SortConditions.Count);

    const valueSort = sheet.AutoFilter.SortState.SortConditions.get(0);
    AssertEx.equal("B2:B6", valueSort.Ref);
    AssertEx.true(valueSort.Descending);
    AssertEx.equal("value", valueSort.SortBy);
    AssertEx.equal("High,Medium,Low", valueSort.CustomList);
    AssertEx.null(valueSort.DifferentialStyleId);
    AssertEx.equal("", valueSort.IconSet);
    AssertEx.null(valueSort.IconId);

    const colorSort = sheet.AutoFilter.SortState.SortConditions.get(1);
    AssertEx.equal("C2:C6", colorSort.Ref);
    AssertEx.false(colorSort.Descending);
    AssertEx.equal("cellColor", colorSort.SortBy);
    AssertEx.equal("", colorSort.CustomList);
    AssertEx.equal(4, colorSort.DifferentialStyleId ?? -1);
    AssertEx.equal("", colorSort.IconSet);
    AssertEx.null(colorSort.IconId);

    const iconSort = sheet.AutoFilter.SortState.SortConditions.get(2);
    AssertEx.equal("E2:E6", iconSort.Ref);
    AssertEx.false(iconSort.Descending);
    AssertEx.equal("icon", iconSort.SortBy);
    AssertEx.equal("", iconSort.CustomList);
    AssertEx.null(iconSort.DifferentialStyleId);
    AssertEx.equal("3TrafficLights1", iconSort.IconSet);
    AssertEx.equal(2, iconSort.IconId ?? -1);
  }

  private static AddDifferentialStyles(sheet: Worksheet): void {
    const colors = [
      Color.FromArgb(255, 255, 235, 156),
      Color.FromArgb(255, 198, 239, 206),
      Color.FromArgb(255, 221, 235, 247),
      Color.FromArgb(255, 255, 199, 206),
      Color.FromArgb(255, 189, 215, 238),
    ];

    for (let index = 0; index < colors.length; index++) {
      const collection = sheet.ConditionalFormattings.get(sheet.ConditionalFormattings.Add());
      collection.AddArea(CellArea.CreateCellArea(index + 1, 0, index + 1, 0));
      const condition = collection.get(
        collection.AddCondition(FormatConditionType.Expression, OperatorType.None, "TRUE", ""),
      );
      const style = condition.Style;
      style.Pattern = FillPattern.Solid;
      style.ForegroundColor = colors[index];
      condition.Style = style;
    }
  }
}
