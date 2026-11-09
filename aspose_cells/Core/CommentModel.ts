export class CommentModel {
  public Row: number = 0;
  public Column: number = 0;
  public Author: string;
  public Note: string;
  public IsVisible: boolean = false;
  public Width: number;
  public Height: number;
  public RawVmlShapeXml: string | null = null;

  public constructor() {
    this.Author = "";
    this.Note = "";
    this.Width = 129;
    this.Height = 75;
  }
}
