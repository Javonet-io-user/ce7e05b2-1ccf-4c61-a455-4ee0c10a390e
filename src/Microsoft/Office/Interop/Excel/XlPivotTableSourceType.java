package Microsoft.Office.Interop.Excel;

public enum XlPivotTableSourceType {
  xlDatabase(1L),
  xlExternal(2L),
  xlConsolidation(3L),
  xlScenario(4L),
  xlPivotTable(-4148L),
  ;
  private long numVal;

  XlPivotTableSourceType(long numVal) {
    this.numVal = numVal;
  }

  public long getNumVal() {
    return numVal;
  }
}
