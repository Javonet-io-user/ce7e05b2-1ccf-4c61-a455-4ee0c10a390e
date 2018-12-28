package Microsoft.Office.Interop.Excel;

public enum XlYesNoGuess {
  xlGuess(0L),
  xlYes(1L),
  xlNo(2L),
  ;
  private long numVal;

  XlYesNoGuess(long numVal) {
    this.numVal = numVal;
  }

  public long getNumVal() {
    return numVal;
  }
}
