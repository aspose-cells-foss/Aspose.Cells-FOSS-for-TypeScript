import { RichTextRunValue } from "./Core/RichTextRunValue";

export class SharedStringEntry {
  public Text: string = "";
  public Runs: RichTextRunValue[] | null = null;

  public clone(): SharedStringEntry {
    const clone = new SharedStringEntry();
    clone.Text = this.Text ?? "";
    if (this.Runs != null && this.Runs.length > 0) {
      clone.Runs = new Array<RichTextRunValue>(this.Runs.length);
      for (let index = 0; index < this.Runs.length; index++) {
        clone.Runs[index] = this.Runs[index]!.clone();
      }
    }

    return clone;
  }
}
