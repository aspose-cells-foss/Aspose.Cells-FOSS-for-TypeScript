import { Workbook } from "../../aspose_cells";
import { AssertEx } from "../Shared/TestInfrastructure";

export function commentCollectionEmptyOnNewWorksheet(): void {
  const workbook = new Workbook();
  AssertEx.equal(0, workbook.Worksheets[0].Comments.Count);
}

export function commentAddByRowColSetsDefaults(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];

  const comment = sheet.Comments.Add(2, 3);
  AssertEx.equal(1, sheet.Comments.Count);
  AssertEx.equal(2, comment.Row);
  AssertEx.equal(3, comment.Column);
  AssertEx.equal("", comment.Note);
  AssertEx.equal("", comment.Author);
  AssertEx.false(comment.IsVisible);
  AssertEx.equal(129, comment.Width);
  AssertEx.equal(75, comment.Height);
}

export function commentAddByCellNameParsesReference(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];

  const comment = sheet.Comments.Add("B3");
  AssertEx.equal(2, comment.Row);
  AssertEx.equal(1, comment.Column);
  AssertEx.equal(1, sheet.Comments.Count);

  const byName = sheet.Comments.get("B3");
  AssertEx.notNull(byName, "Should find comment by cell name B3");
  AssertEx.equal(2, byName!.Row);
}

export function commentAddDuplicateCellThrows(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  sheet.Comments.Add(0, 0);
  AssertEx.throws(() => {
    sheet.Comments.Add(0, 0);
  });
}

export function commentAddNegativeRowThrows(): void {
  const workbook = new Workbook();
  AssertEx.throws(() => {
    workbook.Worksheets[0].Comments.Add(-1, 0);
  });
}

export function commentAddNegativeColumnThrows(): void {
  const workbook = new Workbook();
  AssertEx.throws(() => {
    workbook.Worksheets[0].Comments.Add(0, -1);
  });
}

export function commentAddInvalidCellNameThrows(): void {
  const workbook = new Workbook();
  AssertEx.throws(() => {
    workbook.Worksheets[0].Comments.Add("not_a_cell");
  });
}

export function commentCountIncrementsAfterAdd(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];

  sheet.Comments.Add(0, 0);
  AssertEx.equal(1, sheet.Comments.Count);
  sheet.Comments.Add(1, 0);
  AssertEx.equal(2, sheet.Comments.Count);
  sheet.Comments.Add(0, 1);
  AssertEx.equal(3, sheet.Comments.Count);
}

export function commentRemoveAtIndexDecrementsCount(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  sheet.Comments.Add(0, 0);
  sheet.Comments.Add(1, 0);
  AssertEx.equal(2, sheet.Comments.Count);

  sheet.Comments.RemoveAt(0);
  AssertEx.equal(1, sheet.Comments.Count);
  AssertEx.equal(1, sheet.Comments.get(0).Row);

  sheet.Comments.RemoveAt(0);
  AssertEx.equal(0, sheet.Comments.Count);
}

export function commentRemoveAtIndexOutOfRangeThrows(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  AssertEx.throws(() => {
    sheet.Comments.RemoveAt(0);
  });
  sheet.Comments.Add(0, 0);
  AssertEx.throws(() => {
    sheet.Comments.RemoveAt(1);
  });
  AssertEx.throws(() => {
    sheet.Comments.RemoveAt(-1);
  });
}

export function commentRemoveAtCellNameNoopWhenNotFound(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  sheet.Comments.Add(0, 0);
  sheet.Comments.RemoveAt("Z99");
  AssertEx.equal(1, sheet.Comments.Count);
}

export function commentIndexerByCellNameReturnsNullWhenNotFound(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  sheet.Comments.Add(0, 0);
  const found = sheet.Comments.get("A1");
  AssertEx.notNull(found, "A1 should return the comment");
  const missing = sheet.Comments.get("B2");
  AssertEx.null(missing, "B2 has no comment; should return null");
}

export function commentNoteAndAuthorMutateViaProperty(): void {
  const workbook = new Workbook();
  const sheet = workbook.Worksheets[0];
  const comment = sheet.Comments.Add(0, 0);

  comment.Note = "Review this";
  comment.Author = "Alice";
  comment.IsVisible = true;

  AssertEx.equal("Review this", sheet.Comments.get(0).Note);
  AssertEx.equal("Alice", sheet.Comments.get(0).Author);
  AssertEx.true(sheet.Comments.get(0).IsVisible);
}

export function commentIsVisibleDefaultsFalse(): void {
  const workbook = new Workbook();
  const comment = workbook.Worksheets[0].Comments.Add(0, 0);
  AssertEx.false(comment.IsVisible);
}

export function commentWidthHeightDefaultsAre129And75(): void {
  const workbook = new Workbook();
  const comment = workbook.Worksheets[0].Comments.Add(0, 0);
  AssertEx.equal(129, comment.Width);
  AssertEx.equal(75, comment.Height);
}

export function commentWidthZeroThrows(): void {
  const workbook = new Workbook();
  const comment = workbook.Worksheets[0].Comments.Add(0, 0);
  AssertEx.throws(() => {
    comment.Width = 0;
  });
  AssertEx.throws(() => {
    comment.Width = -5;
  });
}

export function commentHeightZeroThrows(): void {
  const workbook = new Workbook();
  const comment = workbook.Worksheets[0].Comments.Add(0, 0);
  AssertEx.throws(() => {
    comment.Height = 0;
  });
  AssertEx.throws(() => {
    comment.Height = -1;
  });
}
