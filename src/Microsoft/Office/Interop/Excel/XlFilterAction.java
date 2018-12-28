package Microsoft.Office.Interop.Excel;

public enum XlFilterAction {
  xlFilterInPlace(1L),
  xlFilterCopy(2L),
  ;
  private long numVal;

  XlFilterAction(long numVal) {
    this.numVal = numVal;
  }

  public long getNumVal() {
    return numVal;
  }
}
