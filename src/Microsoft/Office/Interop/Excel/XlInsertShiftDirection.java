package Microsoft.Office.Interop.Excel;

public enum XlInsertShiftDirection {
  xlShiftToRight(-4161L),
  xlShiftDown(-4121L),
  ;
  private long numVal;

  XlInsertShiftDirection(long numVal) {
    this.numVal = numVal;
  }

  public long getNumVal() {
    return numVal;
  }
}
