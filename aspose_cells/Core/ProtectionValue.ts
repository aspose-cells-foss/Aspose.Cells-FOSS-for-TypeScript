export class ProtectionValue {
  public IsLocked: boolean = true;
  public IsHidden: boolean = false;

  public clone(): ProtectionValue {
    const clone = new ProtectionValue();
    clone.IsLocked = this.IsLocked;
    clone.IsHidden = this.IsHidden;
    return clone;
  }
}
