package Microsoft.Office.Interop.Excel;

public enum XlPivotFieldOrientation {
  xlHidden(0L),
  xlRowField(1L),
  xlColumnField(2L),
  xlPageField(3L),
  xlDataField(4L),
  ;
  private long numVal;

  XlPivotFieldOrientation(long numVal) {
    this.numVal = numVal;
  }

  public long getNumVal() {
    return numVal;
  }
}
