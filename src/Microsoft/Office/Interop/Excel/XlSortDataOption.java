package Microsoft.Office.Interop.Excel;

public enum XlSortDataOption {
  xlSortNormal(0L),
  xlSortTextAsNumbers(1L),
  ;
  private long numVal;

  XlSortDataOption(long numVal) {
    this.numVal = numVal;
  }

  public long getNumVal() {
    return numVal;
  }
}
