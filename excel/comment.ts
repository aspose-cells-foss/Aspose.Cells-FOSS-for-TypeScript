import type { Comment as CommentType } from "./types"
import { cellRef, escapeXml, generateUuid } from "./util"

export class Comment {
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

export class CommentCollection {
  private _comments: Comment[] = []

  add(row: number, col: number, text: string, author?: string): Comment {
    const comment = new Comment(row, col, text, author)
    this._comments.push(comment)
    return comment
  }

  get(index: number): Comment | undefined {
    return this._comments[index]
  }

  getByCell(row: number, col: number): Comment | undefined {
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
