package Microsoft.Office.Interop.Excel;

public enum XlWindowState {
  xlNormal(-4143L),
  xlMinimized(-4140L),
  xlMaximized(-4137L),
  ;
  private long numVal;

  XlWindowState(long numVal) {
    this.numVal = numVal;
  }

  public long getNumVal() {
    return numVal;
  }
}
