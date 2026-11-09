import { StyleValue } from "./StyleValue";

export class StyleRepository {
  public Normalize(style: StyleValue): StyleValue {
    return style.clone();
  }
}
