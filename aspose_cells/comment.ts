import type { Comment as CommentType } from "./types"
import { cellRef, escapeXml, generateUuid } from "./util"

export class LegacyComment {
  private _row: number
  private _col: number
  private _text: string
  private _author?: string
  private _width = 150
  private _height = 100

  constructor(row: number, col: number, text: string, author?: string) {
    this._row = row
    this._col = col
    this._text = text
    this._author = author
  }

  get row(): number {
    return this._row
  }

  get col(): number {
    return this._col
  }

  get ref(): string {
    return cellRef(this._row, this._col)
  }

  get text(): string {
    return this._text
  }

  set text(value: string) {
    this._text = value
  }

  get author(): string | undefined {
    return this._author
  }

  set author(value: string) {
    this._author = value
  }

  get width(): number {
    return this._width
  }

  set width(value: number) {
    this._width = value
  }

  get height(): number {
    return this._height
  }

  set height(value: number) {
    this._height = value
  }

  toXml(): string {
    return `<comment ref="${this.ref}" author="${escapeXml(this._author || "")}" guid="${generateUuid()}">
    <text>${escapeXml(this._text)}</text>
  </comment>`
  }
}

export class LegacyCommentCollection {
  private _comments: LegacyComment[] = []

  add(row: number, col: number, text: string, author?: string): LegacyComment {
    const comment = new LegacyComment(row, col, text, author)
    this._comments.push(comment)
    return comment
  }

  get(index: number): LegacyComment | undefined {
    return this._comments[index]
  }

  getByCell(row: number, col: number): LegacyComment | undefined {
    return this._comments.find((c) => c.row === row && c.col === col)
  }

  remove(index: number) {
    this._comments.splice(index, 1)
  }

  clear() {
    this._comments = []
  }

  get count(): number {
    return this._comments.length
  }

  toXml(): string {
    return this._comments.map((c) => c.toXml()).join("\n")
  }
}

import { CellsException } from "./CellsException";
import { CommentModel } from "./Core/CommentModel";

// Implemented by serializer module.
/** Represents a cell comment (legacy note). */
export class Comment {
  public constructor(private readonly model: CommentModel) {}

  public get Model(): CommentModel {
    return this.model;
  }

  public get Row(): number {
    return this.model.Row;
  }

  public get Column(): number {
    return this.model.Column;
  }

  public get Author(): string {
    return this.model.Author;
  }

  public set Author(value: string) {
    this.model.Author = value ?? "";
  }

  public get Note(): string {
    return this.model.Note;
  }

  public set Note(value: string) {
    this.model.Note = value ?? "";
  }

  public get IsVisible(): boolean {
    return this.model.IsVisible;
  }

  public set IsVisible(value: boolean) {
    this.model.IsVisible = value;
  }

  public get Width(): number {
    return this.model.Width;
  }

  public set Width(value: number) {
    if (value <= 0) {
      throw new CellsException("Comment width must be positive.");
    }

    this.model.Width = value;
  }

  public get Height(): number {
    return this.model.Height;
  }

  public set Height(value: number) {
    if (value <= 0) {
      throw new CellsException("Comment height must be positive.");
    }

    this.model.Height = value;
  }
}
