package Microsoft.Office.Interop.Excel;

public enum XlSortOrder {
  xlAscending(1L),
  xlDescending(2L),
  ;
  private long numVal;

  XlSortOrder(long numVal) {
    this.numVal = numVal;
  }

  public long getNumVal() {
    return numVal;
  }
}
