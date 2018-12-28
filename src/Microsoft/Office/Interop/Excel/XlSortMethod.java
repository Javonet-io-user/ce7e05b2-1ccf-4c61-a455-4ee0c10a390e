package Microsoft.Office.Interop.Excel;

public enum XlSortMethod {
  xlPinYin(1L),
  xlStroke(2L),
  ;
  private long numVal;

  XlSortMethod(long numVal) {
    this.numVal = numVal;
  }

  public long getNumVal() {
    return numVal;
  }
}
