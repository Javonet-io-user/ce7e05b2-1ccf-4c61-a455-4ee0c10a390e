package Microsoft.Office.Interop.Excel;

public enum XlSortOrientation {
  xlSortColumns(1L),
  xlSortRows(2L),
  ;
  private long numVal;

  XlSortOrientation(long numVal) {
    this.numVal = numVal;
  }

  public long getNumVal() {
    return numVal;
  }
}
