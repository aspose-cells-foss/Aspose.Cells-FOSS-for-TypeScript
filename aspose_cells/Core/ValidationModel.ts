import { CellArea } from "../CellArea";
import { OperatorType } from "../OperatorType";
import { ValidationAlertType } from "../ValidationAlertType";
import { ValidationType } from "../ValidationType";

export class ValidationModel {
  public readonly Areas: CellArea[];
  public Type: ValidationType = ValidationType.AnyValue;
  public AlertStyle: ValidationAlertType = ValidationAlertType.Stop;
  public Operator: OperatorType = OperatorType.None;
  public Formula1: string | null = null;
  public Formula2: string | null = null;
  public IgnoreBlank: boolean = false;
  public InCellDropDown: boolean = true;
  public InputTitle: string | null = null;
  public InputMessage: string | null = null;
  public ErrorTitle: string | null = null;
  public ErrorMessage: string | null = null;
  public ShowInput: boolean = false;
  public ShowError: boolean = false;

  public constructor() {
    this.Areas = [];
  }
}
