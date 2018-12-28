package Microsoft.Office.Interop.Excel;

public enum XlPasteSpecialOperation {
  xlPasteSpecialOperationAdd(2L),
  xlPasteSpecialOperationSubtract(3L),
  xlPasteSpecialOperationMultiply(4L),
  xlPasteSpecialOperationDivide(5L),
  xlPasteSpecialOperationNone(-4142L),
  ;
  private long numVal;

  XlPasteSpecialOperation(long numVal) {
    this.numVal = numVal;
  }

  public long getNumVal() {
    return numVal;
  }
}
